"""
Telegram bot for indexing and browsing an Excel-file archive.

Install with: pip install -r requirements.txt
"""

from __future__ import annotations

import logging
import re
import os
import sqlite3
import sys
from pathlib import Path
from typing import Any, Optional

import psycopg
from telegram import BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.error import TelegramError
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)


# Bot configuration: replace TOKEN and add the numeric Telegram user IDs.
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "").strip()
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "").strip()
ADMIN_IDS = {5230838728}
# Set to None to allow everyone to browse, or use e.g. {123456789, 987654321}.
ALLOWED_IDS: Optional[set[int]] = None

# Storage and labels.
DATABASE_PATH = Path(__file__).with_name("excel_archive.db")
WEBHOOK_PATH = "telegram-webhook"
SEASONS = ("spring", "summer", "autumn", "winter", "bts")
CATEGORIES = ("kvi", "discount", "best sell")
CATEGORY_CALLBACKS = {"kvi": "kvi", "discount": "discount", "best sell": "best_sell"}
CATEGORY_LABELS = {value: key for key, value in CATEGORY_CALLBACKS.items()}
SEASON_LABELS = {"bts": "Back to School"}

logging.basicConfig(
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

USAGE = (
    "Send an Excel file with the caption in this format:\n"
    "year|season|category\n"
    "Example: 2025|bts|best sell\n"
    "Seasons: spring, summer, autumn, winter, bts\n"
    "Categories: kvi, discount, best sell"
)


# Database helpers. Each operation owns and closes its PostgreSQL connection.
def _database_connection() -> psycopg.Connection[tuple[Any, ...]]:
    if not DATABASE_URL:
        raise RuntimeError("Set DATABASE_URL to your PostgreSQL connection string.")
    return psycopg.connect(DATABASE_URL, connect_timeout=10)


def initialize_database() -> None:
    """Create the archive table automatically on the first run."""
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS files (
                    year TEXT NOT NULL,
                    season TEXT NOT NULL,
                    category TEXT NOT NULL,
                    file_id TEXT NOT NULL,
                    file_name TEXT NOT NULL,
                    PRIMARY KEY (year, season, category)
                )
                """
            )


def migrate_sqlite_to_postgres() -> int:
    """Copy the local SQLite archive to PostgreSQL without deleting the source."""
    if not DATABASE_PATH.is_file():
        raise FileNotFoundError(f"SQLite archive not found: {DATABASE_PATH}")

    initialize_database()
    sqlite_connection = sqlite3.connect(DATABASE_PATH)
    try:
        rows = sqlite_connection.execute(
            """
            SELECT year, season, category, file_id, file_name
            FROM files
            """
        ).fetchall()
    finally:
        sqlite_connection.close()

    if not rows:
        logger.info("The local SQLite archive contains no entries to migrate")
        return 0

    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.executemany(
                """
                INSERT INTO files (year, season, category, file_id, file_name)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT(year, season, category) DO UPDATE SET
                    file_id = EXCLUDED.file_id,
                    file_name = EXCLUDED.file_name
                """,
                rows,
            )
    logger.info("Migrated %d archive entries from SQLite to PostgreSQL", len(rows))
    return len(rows)


def fetch_years() -> list[str]:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT DISTINCT year FROM files ORDER BY year DESC"
            )
            return [row[0] for row in cursor.fetchall()]


def fetch_seasons(year: str) -> list[str]:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                "SELECT DISTINCT season FROM files WHERE year = %s", (year,)
            )
            available = {row[0] for row in cursor.fetchall()}
    return [season for season in SEASONS if season in available]


def fetch_categories(year: str, season: str) -> list[str]:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT DISTINCT category FROM files
                WHERE year = %s AND season = %s
                """,
                (year, season),
            )
            available = {row[0] for row in cursor.fetchall()}
    return [category for category in CATEGORIES if category in available]


def fetch_file(year: str, season: str, category: str) -> Optional[tuple[str, str]]:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT file_id, file_name FROM files
                WHERE year = %s AND season = %s AND category = %s
                """,
                (year, season, category),
            )
            return cursor.fetchone()


def save_file(
    year: str, season: str, category: str, file_id: str, file_name: str
) -> None:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                INSERT INTO files (year, season, category, file_id, file_name)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT(year, season, category) DO UPDATE SET
                    file_id = EXCLUDED.file_id,
                    file_name = EXCLUDED.file_name
                """,
                (year, season, category, file_id, file_name),
            )


def delete_file(year: str, season: str, category: str) -> bool:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                "DELETE FROM files WHERE year = %s AND season = %s AND category = %s",
                (year, season, category),
            )
            return cursor.rowcount > 0


def fetch_all_files() -> list[tuple[str, str, str, str]]:
    with _database_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT year, season, category, file_name FROM files
                ORDER BY year DESC, season, category
                """
            )
            return cursor.fetchall()


# Shared access and path validation.
def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def can_browse(user_id: int) -> bool:
    return ALLOWED_IDS is None or user_id in ALLOWED_IDS


def parse_path(value: str) -> Optional[tuple[str, str, str]]:
    """Parse and validate a year|season|category path."""
    parts = [part.strip().lower() for part in value.split("|")]
    if len(parts) != 3:
        return None
    year, season, category = parts
    if not re.fullmatch(r"[0-9]{4}", year):
        return None
    if season not in SEASONS or category not in CATEGORIES:
        return None
    return year, season, category


def season_label(season: str) -> str:
    return SEASON_LABELS.get(season, season)


def category_label(category: str) -> str:
    return category.replace("_", " ")


def make_keyboard(items: list[tuple[str, str]], back_data: Optional[str] = None):
    rows = [[InlineKeyboardButton(label, callback_data=data)] for label, data in items]
    if back_data is not None:
        rows.append([InlineKeyboardButton("⬅️ Back", callback_data=back_data)])
    return InlineKeyboardMarkup(rows) if rows else None


# Admin uploads: Telegram keeps the file; the database stores only its file_id.
async def handle_document(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    user = update.effective_user
    message = update.effective_message
    document = message.document if message else None
    if user is None or message is None or document is None:
        return

    file_name = document.file_name or ""
    if Path(file_name).suffix.lower() not in {".xlsx", ".xls"}:
        await message.reply_text("Please send an Excel file (.xlsx or .xls).")
        return
    if not is_admin(user.id):
        await message.reply_text("Only archive admins can upload files.")
        return

    if not message.caption:
        await message.reply_text(USAGE)
        return
    path = parse_path(message.caption)
    if path is None:
        await message.reply_text(f"Invalid caption.\n\n{USAGE}")
        return

    year, season, category = path
    try:
        save_file(year, season, category, document.file_id, file_name or "Excel file")
    except psycopg.Error:
        logger.exception("Could not save uploaded file metadata")
        await message.reply_text("The file could not be added because of a database error.")
        return

    await message.reply_text(
        f"Saved: {year} -> {season_label(season)} -> {category_label(category)}"
    )


# Browsing uses short callback payloads; even the deepest path stays under 64 bytes.
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    message = update.effective_message
    if user is None or message is None:
        return
    if not can_browse(user.id):
        await message.reply_text("You are not allowed to browse this archive.")
        return

    try:
        years = fetch_years()
    except psycopg.Error:
        logger.exception("Could not load archive years")
        await message.reply_text("The archive could not be loaded due to a database error.")
        return

    if not years:
        await message.reply_text("Nothing uploaded yet.")
        return
    buttons = [(year, f"y:{year}") for year in years]
    await message.reply_text(
        "Choose a year:", reply_markup=make_keyboard(buttons)
    )


async def browse_callback(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    query = update.callback_query
    user = update.effective_user
    if query is None or user is None:
        return
    if not can_browse(user.id):
        await query.answer("You are not allowed to browse this archive.", show_alert=True)
        return

    data = query.data or ""
    try:
        if data.startswith("y:"):
            year = data[2:]
            seasons = fetch_seasons(year)
            if not seasons:
                await query.answer()
                await query.edit_message_text(
                    "Nothing uploaded yet.",
                    reply_markup=make_keyboard([], "home"),
                )
                return
            buttons = [
                (season_label(season), f"s:{year}|{season}") for season in seasons
            ]
            await query.answer()
            await query.edit_message_text(
                f"{year} — choose a season:",
                reply_markup=make_keyboard(buttons, "home"),
            )
            return

        if data.startswith("s:"):
            parts = data[2:].split("|")
            if len(parts) != 2:
                await query.answer("Invalid selection.", show_alert=True)
                return
            year, season = parts
            if not re.fullmatch(r"[0-9]{4}", year) or season not in SEASONS:
                await query.answer("Invalid selection.", show_alert=True)
                return
            categories = fetch_categories(year, season)
            if not categories:
                await query.answer()
                await query.edit_message_text(
                    "Nothing uploaded yet.",
                    reply_markup=make_keyboard([], f"y:{year}"),
                )
                return
            buttons = [
                (
                    category_label(category),
                    f"n:{year}|{season}|{CATEGORY_CALLBACKS[category]}",
                )
                for category in categories
            ]
            await query.answer()
            await query.edit_message_text(
                f"{year} -> {season_label(season)} — choose a category:",
                reply_markup=make_keyboard(buttons, f"y:{year}"),
            )
            return

        if data.startswith("n:"):
            parts = data[2:].split("|")
            if len(parts) != 3:
                await query.answer("Invalid selection.", show_alert=True)
                return
            year, season, category_code = parts
            category = CATEGORY_LABELS.get(category_code)
            if (
                not re.fullmatch(r"[0-9]{4}", year)
                or season not in SEASONS
                or category is None
            ):
                await query.answer("Invalid selection.", show_alert=True)
                return
            stored = fetch_file(year, season, category)
            await query.answer()
            if stored is None:
                await query.edit_message_text(
                    "Nothing uploaded yet.",
                    reply_markup=make_keyboard([], f"s:{year}|{season}"),
                )
                return

            caption = (
                f"{year} -> {season_label(season)} -> {category_label(category)}"
            )
            try:
                await query.message.reply_document(
                    document=stored[0],
                    caption=caption,
                    filename=stored[1],
                )
            except TelegramError:
                logger.exception("Could not send archived file %s", caption)
                await query.message.reply_text(
                    "The stored file could not be retrieved. Please contact an admin."
                )
            return

        if data == "home":
            years = fetch_years()
            await query.answer()
            if not years:
                await query.edit_message_text("Nothing uploaded yet.")
                return
            buttons = [(year, f"y:{year}") for year in years]
            await query.edit_message_text(
                "Choose a year:", reply_markup=make_keyboard(buttons)
            )
            return

        await query.answer("Invalid selection.", show_alert=True)
    except psycopg.Error:
        logger.exception("Could not browse the archive")
        await query.answer("The archive is temporarily unavailable.", show_alert=True)


# Admin commands for removing one entry and listing the current index.
async def delete_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    user = update.effective_user
    message = update.effective_message
    if user is None or message is None:
        return
    if not is_admin(user.id):
        await message.reply_text("Only archive admins can delete files.")
        return
    if not context.args:
        await message.reply_text(
            "Usage: /delete year|season|category\nExample: /delete 2025|bts|best sell"
        )
        return

    path = parse_path(" ".join(context.args))
    if path is None:
        await message.reply_text(
            f"Invalid path.\nAllowed seasons: {', '.join(SEASONS)}\n"
            f"Allowed categories: {', '.join(CATEGORIES)}\n"
            "Usage: /delete 2025|bts|best sell"
        )
        return
    try:
        removed = delete_file(*path)
    except psycopg.Error:
        logger.exception("Could not delete archive entry")
        await message.reply_text("The file could not be deleted due to a database error.")
        return
    if removed:
        await message.reply_text(
            f"Deleted: {path[0]} -> {season_label(path[1])} -> "
            f"{category_label(path[2])}"
        )
    else:
        await message.reply_text("No file is stored at that path.")


async def list_command(
    update: Update, context: ContextTypes.DEFAULT_TYPE
) -> None:
    user = update.effective_user
    message = update.effective_message
    if user is None or message is None:
        return
    if not is_admin(user.id):
        await message.reply_text("Only archive admins can list stored files.")
        return
    try:
        files = fetch_all_files()
    except psycopg.Error:
        logger.exception("Could not list archive entries")
        await message.reply_text("The archive could not be listed due to a database error.")
        return
    if not files:
        await message.reply_text("Nothing uploaded yet.")
        return

    lines = [
        f"{year} -> {season_label(season)} -> {category_label(category)} "
        f"({file_name})"
        for year, season, category, file_name in files
    ]
    # Split long listings to stay within Telegram's message-size limit.
    chunk = "Stored files:\n"
    for line in lines:
        addition = f"\n{line}"
        if len(chunk) + len(addition) > 3500:
            await message.reply_text(chunk)
            chunk = "Stored files (continued):\n"
        chunk += addition
    await message.reply_text(chunk)


# Log unexpected handler failures instead of silently swallowing them.
async def error_handler(
    update: object, context: ContextTypes.DEFAULT_TYPE
) -> None:
    if context.error is None:
        logger.error("Unhandled bot error without exception details")
        return
    logger.error(
        "Unhandled bot error while processing update",
        exc_info=(
            type(context.error),
            context.error,
            context.error.__traceback__,
        ),
    )


async def set_bot_commands(application: Application) -> None:
    await application.bot.set_my_commands(
        [BotCommand("start", "Browse the Excel archive")]
    )


def main() -> None:
    if not TOKEN:
        raise RuntimeError(
            "Telegram bot token is missing. Set TELEGRAM_BOT_TOKEN in the "
            "service environment before starting the bot."
        )
    if RENDER_EXTERNAL_URL and (
        not WEBHOOK_SECRET
        or len(WEBHOOK_SECRET) > 256
        or re.fullmatch(r"[A-Za-z0-9_-]+", WEBHOOK_SECRET) is None
    ):
        raise RuntimeError(
            "Set WEBHOOK_SECRET to 1-256 letters, numbers, underscores, or "
            "hyphens when using RENDER_EXTERNAL_URL."
        )
    initialize_database()
    application = (
        Application.builder()
        .token(TOKEN)
        .post_init(set_bot_commands)
        .build()
    )
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("delete", delete_command))
    application.add_handler(CommandHandler("list", list_command))
    application.add_handler(CallbackQueryHandler(browse_callback))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    application.add_error_handler(error_handler)
    if RENDER_EXTERNAL_URL:
        try:
            port = int(os.getenv("PORT", "10000"))
        except ValueError as error:
            raise RuntimeError("PORT must be a valid integer.") from error
        if port < 1 or port > 65535:
            raise RuntimeError("PORT must be between 1 and 65535.")
        webhook_url = f"{RENDER_EXTERNAL_URL.rstrip('/')}/{WEBHOOK_PATH}"
        logger.info("Starting Excel archive bot webhook on port %d", port)
        application.run_webhook(
            listen="0.0.0.0",
            port=port,
            url_path=WEBHOOK_PATH,
            webhook_url=webhook_url,
            secret_token=WEBHOOK_SECRET,
        )
    else:
        logger.info("Starting Excel archive bot with polling")
        application.run_polling()


if __name__ == "__main__":
    if sys.argv[1:] == ["--migrate-sqlite"]:
        migrated_count = migrate_sqlite_to_postgres()
        print(f"Migrated {migrated_count} archive entries.")
    elif len(sys.argv) > 1:
        raise SystemExit("Usage: python bot.py [--migrate-sqlite]")
    else:
        main()
