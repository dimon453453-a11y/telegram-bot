import os
import tempfile

from telegram import Update
from telegram.ext import ContextTypes

from utils.excel_editor import ExcelEditor

# Maximum number of rows displayed inline to keep the reply readable.
_MAX_DISPLAY_ROWS = 50


async def handle_excel_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle an Excel (.xlsx) document sent by the user.

    The bot downloads the file, reads every sheet with ExcelEditor, and
    replies with a plain-text table showing the contents.
    """
    document = update.message.document

    if document is None:
        await update.message.reply_text("Пожалуйста, отправьте файл Excel (.xlsx).")
        return

    # Accept only Excel files.
    mime = document.mime_type or ""
    filename = document.file_name or ""
    is_excel = mime in (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
    ) or filename.lower().endswith((".xlsx", ".xls"))

    if not is_excel:
        await update.message.reply_text(
            "Этот файл не является файлом Excel. "
            "Пожалуйста, отправьте файл с расширением .xlsx."
        )
        return

    await update.message.reply_text("📥 Получаю файл, подождите...")

    # Download the file into a temporary directory.
    with tempfile.TemporaryDirectory() as tmp_dir:
        filepath = os.path.join(tmp_dir, filename or "upload.xlsx")
        tg_file = await context.bot.get_file(document.file_id)
        try:
            await tg_file.download_to_drive(filepath)
        except Exception:
            await update.message.reply_text(
                "❌ Не удалось скачать файл. Попробуйте ещё раз."
            )
            return

        editor = ExcelEditor(filepath)
        reply = _format_workbook(editor)

    await update.message.reply_text(reply, parse_mode=None)


def _format_workbook(editor: ExcelEditor) -> str:
    """Return a human-readable text representation of every sheet."""
    parts: list[str] = []

    for sheet_name in editor.sheet_names:
        editor.select_sheet(sheet_name)
        rows = editor.read_all()

        if not rows or all(all(v is None for v in r) for r in rows):
            parts.append(f"📋 Лист: {sheet_name}\n(пустой лист)")
            continue

        total = len(rows)
        display = rows[:_MAX_DISPLAY_ROWS]

        table_lines = [_format_row(r) for r in display]
        table = "\n".join(table_lines)

        truncation = ""
        if total > _MAX_DISPLAY_ROWS:
            truncation = f"\n... и ещё {total - _MAX_DISPLAY_ROWS} строк не показаны."

        parts.append(f"📋 Лист: {sheet_name}\n{table}{truncation}")

    return "\n\n".join(parts) if parts else "Файл не содержит данных."


def _format_row(row: list) -> str:
    """Format a single row as tab-separated values, replacing None with a dash."""
    return "\t".join("-" if v is None else str(v) for v in row)
