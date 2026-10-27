"""
Telegram PDF toolbox bot

Features
- Convert Word / Excel / PowerPoint / text / images to PDF
- Merge several photos / images / PDFs into ONE PDF  (/merge ... /done)
- PDF -> Word, JPG images (zip), plain text
- Extract pages from a PDF (/split 1-3)
- Compress PDF (Ghostscript)
- Friendly progress + error messages
"""
import asyncio
import logging
import os
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path

import pymupdf as fitz  # PyMuPDF
from PIL import Image, ImageOps
from telegram import (
    BotCommand,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Update,
)
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("pdf-bot")

# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
# Paste your token between the quotes, OR leave it empty and set the BOT_TOKEN
# environment variable instead.
BOT_TOKEN = ""
BOT_TOKEN = BOT_TOKEN or os.environ.get("BOT_TOKEN", "")
if not BOT_TOKEN:
    raise SystemExit("No bot token found. Paste it into BOT_TOKEN in bot.py or set the BOT_TOKEN environment variable.")

# LibreOffice (Office -> PDF). On Windows, e.g.:
#   C:\Program Files\LibreOffice\program\soffice.exe
SOFFICE = os.environ.get("SOFFICE_PATH", "soffice")
# Ghostscript (PDF compression). On Windows it is usually gswin64c
GS = os.environ.get("GS_PATH", "gswin64c" if os.name == "nt" else "gs")

MAX_DOWNLOAD_MB = 20      # Telegram limit for files a bot can download
MAX_UPLOAD_MB = 50        # Telegram limit for files a bot can send
MAX_MERGE_FILES = 30
MAX_IMAGE_PAGES = 30      # max pages for PDF -> images
CONVERT_TIMEOUT = 120     # seconds

OFFICE_EXT = {
    ".doc", ".docx", ".odt", ".rtf", ".txt", ".html", ".htm",
    ".xls", ".xlsx", ".ods", ".csv",
    ".ppt", ".pptx", ".odp",
}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff", ".tif"}

# LibreOffice doesn't like many instances at once
convert_lock = asyncio.Semaphore(2)


class UserError(Exception):
    """An error whose message is safe and friendly to show to the user."""


TOO_BIG_MSG = (
    f"📦 That file is larger than {MAX_DOWNLOAD_MB} MB.\n\n"
    f"Telegram only lets bots download files up to {MAX_DOWNLOAD_MB} MB, so I can't open it. "
    "Try compressing it on your device or splitting it into smaller parts, then send it again."
)

UNSUPPORTED_MSG = (
    "🤔 I can't work with that type of file yet.\n\n"
    "I can convert: Word, Excel, PowerPoint, TXT, HTML and images to PDF, "
    "and I can work with PDFs. Send /help to see everything."
)

PROGRESS = {
    "topdf": "⚙️ Converting to PDF...",
    "word": "📝 Converting PDF to Word (this can take a while for long files)...",
    "images": "🖼️ Rendering pages as images...",
    "text": "🔤 Extracting text...",
    "compress_ebook": "🗜️ Compressing your PDF...",
    "compress_screen": "🗜️ Compressing your PDF (strong)...",
    "split": "✂️ Extracting pages...",
}


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def human_size(n: int) -> str:
    if n < 1024 * 1024:
        return f"{n / 1024:.0f} KB"
    return f"{n / 1024 / 1024:.1f} MB"


def new_workdir() -> Path:
    return Path(tempfile.mkdtemp(prefix="pdfbot_"))


def cleanup_stale_dirs():
    """Remove temp folders left behind by a previous crash."""
    for p in Path(tempfile.gettempdir()).glob("pdfbot_*"):
        shutil.rmtree(p, ignore_errors=True)


def is_too_big(size) -> bool:
    return bool(size) and size > MAX_DOWNLOAD_MB * 1024 * 1024


def open_pdf(path: Path):
    try:
        doc = fitz.open(path)
    except Exception:
        raise UserError("This PDF looks corrupt or isn't a valid PDF file.")
    if doc.needs_pass:
        doc.close()
        raise UserError("This PDF is password-protected. Please remove the password and send it again.")
    if doc.page_count == 0:
        doc.close()
        raise UserError("This PDF has no pages.")
    return doc


async def download(context: ContextTypes.DEFAULT_TYPE, file_id: str, dest: Path):
    tg_file = await context.bot.get_file(file_id)
    await tg_file.download_to_drive(custom_path=dest)


async def send_file(msg, path: Path, filename: str, caption: str):
    if path.stat().st_size > MAX_UPLOAD_MB * 1024 * 1024:
        raise UserError(
            f"The result is larger than {MAX_UPLOAD_MB} MB, which is the most Telegram lets me send."
        )
    await msg.chat.send_action(ChatAction.UPLOAD_DOCUMENT)
    with open(path, "rb") as f:
        await msg.reply_document(document=f, filename=filename, caption=caption)


# ----------------------------------------------------------------------------
# Conversion functions (blocking ones are run with asyncio.to_thread)
# ----------------------------------------------------------------------------
def image_to_pdf(src: Path, out_dir: Path) -> Path:
    pdf = out_dir / (src.stem + ".pdf")
    try:
        with Image.open(src) as img:
            img = ImageOps.exif_transpose(img).convert("RGB")  # fixes rotated phone photos
            img.save(pdf, "PDF", resolution=100.0)
    except Exception:
        raise UserError("I couldn't read that image. It may be corrupt or in an unsupported format.")
    return pdf


async def office_to_pdf(src: Path, out_dir: Path) -> Path:
    profile = Path(tempfile.gettempdir()) / f"lo_profile_{uuid.uuid4().hex}"
    cmd = [
        SOFFICE,
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(src),
    ]
    try:
        async with convert_lock:
            try:
                proc = await asyncio.create_subprocess_exec(
                    *cmd,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                )
            except FileNotFoundError:
                raise UserError(
                    "Office conversion isn't available right now (LibreOffice is not installed "
                    "or its path is wrong on the server)."
                )
            try:
                await asyncio.wait_for(proc.communicate(), timeout=CONVERT_TIMEOUT)
            except asyncio.TimeoutError:
                proc.kill()
                raise UserError("The conversion took too long and was cancelled. Try a smaller file.")
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    pdf = out_dir / (src.stem + ".pdf")
    if not pdf.exists():
        raise UserError(
            "I couldn't convert that file. It may be corrupt, password-protected or in an unsupported format."
        )
    return pdf


def merge_files(files: list, out_path: Path, work_dir: Path):
    """Merge images and PDFs, in order, into one PDF."""
    merged = fitz.open()
    try:
        for f in files:
            if f.suffix.lower() in IMAGE_EXT:
                f = image_to_pdf(f, work_dir)
            with open_pdf(f) as d:
                merged.insert_pdf(d)
        merged.save(out_path, garbage=3, deflate=True)
    finally:
        merged.close()


def pdf_to_docx(src: Path, out: Path):
    open_pdf(src).close()  # validates (corrupt / password)
    try:
        from pdf2docx import Converter
    except ImportError:
        raise UserError("Word conversion isn't installed on the server.")
    cv = Converter(str(src))
    try:
        cv.convert(str(out))
    finally:
        cv.close()


def pdf_to_images_zip(src: Path, out: Path, stem: str):
    with open_pdf(src) as doc:
        if doc.page_count > MAX_IMAGE_PAGES:
            raise UserError(
                f"This PDF has {doc.page_count} pages. For images I can handle up to {MAX_IMAGE_PAGES} pages. "
                "Use 'Extract pages' first to pick the pages you need."
            )
        with zipfile.ZipFile(out, "w", zipfile.ZIP_STORED) as z:
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=150)
                z.writestr(f"{stem}_page_{i}.jpg", pix.tobytes("jpg"))
        return doc.page_count


def pdf_to_text(src: Path, out: Path):
    with open_pdf(src) as doc:
        text = "\n\n".join(page.get_text() for page in doc)
    if not text.strip():
        raise UserError(
            "I couldn't find any text in this PDF. It's probably a scan (images of pages), "
            "so there's no text to extract."
        )
    out.write_text(text, encoding="utf-8")


def parse_pages(spec: str, total: int) -> list:
    spec = spec.replace(" ", "")
    if not spec:
        raise UserError("Please tell me which pages, e.g. 1-3 or 2,5,7-9")
    pages = []
    for part in spec.split(","):
        try:
            if "-" in part:
                a, b = part.split("-", 1)
                start = int(a) if a else 1
                end = int(b) if b else total
            else:
                start = end = int(part)
        except ValueError:
            raise UserError("I didn't understand that. Use a format like 1-3 or 2,5,7-9")
        if start < 1 or end > total or start > end:
            raise UserError(f"'{part}' isn't a valid page range. This PDF has {total} page(s).")
        pages.extend(range(start - 1, end))
    return pages


def extract_pages(src: Path, out: Path, spec: str) -> int:
    with open_pdf(src) as doc:
        pages = parse_pages(spec, doc.page_count)
        doc.select(pages)
        doc.save(out, garbage=3, deflate=True)
    return len(pages)


async def compress_pdf(src: Path, out: Path, level: str):
    """level: 'ebook' (good quality) or 'screen' (smallest)."""
    open_pdf(src).close()
    cmd = [
        GS, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{level}",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-sOutputFile={out}", str(src),
    ]
    try:
        proc = await asyncio.create_subprocess_exec(
            *cmd, stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE
        )
    except FileNotFoundError:
        raise UserError(
            "Compression isn't available right now (Ghostscript is not installed or its path is wrong)."
        )
    try:
        await asyncio.wait_for(proc.communicate(), timeout=CONVERT_TIMEOUT * 2)
    except asyncio.TimeoutError:
        proc.kill()
        raise UserError("Compression took too long and was cancelled.")
    if proc.returncode != 0 or not out.exists():
        raise UserError("I couldn't compress this PDF. It may be damaged.")


# ----------------------------------------------------------------------------
# Core job runner: download -> process -> send, with progress + error messages
# ----------------------------------------------------------------------------
async def process(update: Update, context: ContextTypes.DEFAULT_TYPE, *,
                  file_id: str, filename: str, action: str,
                  arg: str = "", caption: str = "✅ Done!") -> bool:
    msg = update.effective_message
    status = await msg.reply_text("⏳ Downloading your file...")
    work = new_workdir()
    try:
        ext = Path(filename).suffix.lower()
        stem = Path(filename).stem or "document"
        src = work / f"input_{uuid.uuid4().hex[:8]}{ext}"
        await download(context, file_id, src)
        await status.edit_text(PROGRESS[action])
        await msg.chat.send_action(ChatAction.TYPING)

        if action == "topdf":
            if ext in IMAGE_EXT:
                out = await asyncio.to_thread(image_to_pdf, src, work)
            else:
                out = await office_to_pdf(src, work)
            out_name = f"{stem}.pdf"

        elif action == "word":
            out = work / "output.docx"
            await asyncio.to_thread(pdf_to_docx, src, out)
            out_name = f"{stem}.docx"

        elif action == "images":
            out = work / "output.zip"
            n = await asyncio.to_thread(pdf_to_images_zip, src, out, stem)
            out_name = f"{stem}_images.zip"
            caption = f"✅ {n} page(s) as JPG images"

        elif action == "text":
            out = work / "output.txt"
            await asyncio.to_thread(pdf_to_text, src, out)
            out_name = f"{stem}.txt"

        elif action.startswith("compress_"):
            level = action.split("_", 1)[1]
            out = work / "compressed.pdf"
            await compress_pdf(src, out, level)
            before, after = src.stat().st_size, out.stat().st_size
            if after >= before:
                await status.edit_text(
                    "ℹ️ This PDF is already well optimised, so I couldn't make it any smaller."
                )
                return True
            saved = round((1 - after / before) * 100)
            caption = f"✅ {human_size(before)} → {human_size(after)} (saved {saved}%)"
            out_name = f"{stem}_compressed.pdf"

        elif action == "split":
            out = work / "pages.pdf"
            n = await asyncio.to_thread(extract_pages, src, out, arg)
            out_name = f"{stem}_pages.pdf"
            caption = f"✅ Extracted {n} page(s)"

        else:
            raise UserError("Unknown action.")

        await send_file(msg, out, out_name, caption)
        await status.delete()
        return True

    except UserError as e:
        text = f"⚠️ {e}"
        if action == "split":
            text += "\n\nSend the pages again, or /cancel."
        await status.edit_text(text)
        return False
    except Exception:
        log.exception("Job failed (%s)", action)
        await status.edit_text(
            "❌ Something went wrong while processing your file. "
            "It may be corrupt or in an unusual format. Please try again or send a different file."
        )
        return False
    finally:
        shutil.rmtree(work, ignore_errors=True)


# ----------------------------------------------------------------------------
# PDF action buttons
# ----------------------------------------------------------------------------
def pdf_keyboard(msg_id: int) -> InlineKeyboardMarkup:
    b = InlineKeyboardButton
    return InlineKeyboardMarkup([
        [b("📝 To Word", callback_data=f"act:word:{msg_id}"),
         b("🖼️ To images", callback_data=f"act:images:{msg_id}"),
         b("🔤 To text", callback_data=f"act:text:{msg_id}")],
        [b("🗜️ Compress", callback_data=f"act:compress_ebook:{msg_id}"),
         b("🗜️ Compress (strong)", callback_data=f"act:compress_screen:{msg_id}")],
        [b("✂️ Extract pages", callback_data=f"act:split:{msg_id}")],
    ])


def remember_pdf(context, msg_id: int, file_id: str, name: str) -> dict:
    info = {"file_id": file_id, "name": name}
    files = context.user_data.setdefault("files", {})
    files[msg_id] = info
    while len(files) > 20:
        files.pop(next(iter(files)))
    context.user_data["last_pdf"] = info
    context.user_data.pop("awaiting_split", None)
    return info


async def on_action(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    try:
        _, action, mid = q.data.split(":")
        info = context.user_data.get("files", {}).get(int(mid))
    except ValueError:
        info = None
    if not info:
        await q.message.reply_text("I no longer have that file. Please send the PDF again.")
        return
    if action == "split":
        context.user_data["awaiting_split"] = info
        await q.message.reply_text(
            "✂️ Which pages do you want to keep?\n\n"
            "Examples:\n1-3\n2,5,7\n4-  (page 4 to the end)\n1-3,8-10\n\n"
            "Send /cancel to stop."
        )
        return
    await process(update, context, file_id=info["file_id"], filename=info["name"], action=action)


# ----------------------------------------------------------------------------
# Merge mode (/merge ... /done)
# ----------------------------------------------------------------------------
def drop_merge_session(context):
    s = context.user_data.pop("merge", None)
    if s:
        shutil.rmtree(s["dir"], ignore_errors=True)


async def merge_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    drop_merge_session(context)
    context.user_data.pop("awaiting_split", None)
    context.user_data["merge"] = {"dir": new_workdir(), "files": []}
    await update.message.reply_text(
        "📎 Merge mode is ON.\n\n"
        "Send me photos, images or PDFs in the order you want them. "
        f"(up to {MAX_MERGE_FILES} files)\n\n"
        "When you're finished, send /done. To abort, send /cancel."
    )


async def add_to_merge(update: Update, context: ContextTypes.DEFAULT_TYPE,
                       file_id: str, size, ext: str):
    msg = update.message
    session = context.user_data["merge"]
    if ext not in IMAGE_EXT and ext != ".pdf":
        await msg.reply_text("⚠️ In merge mode I can only take images and PDFs. Send another file, or /done.")
        return
    if is_too_big(size):
        await msg.reply_text(TOO_BIG_MSG)
        return
    if len(session["files"]) >= MAX_MERGE_FILES:
        await msg.reply_text(f"⚠️ You've reached the limit of {MAX_MERGE_FILES} files. Send /done to build the PDF.")
        return
    dest = session["dir"] / f"{len(session['files']) + 1:03d}{ext}"
    try:
        await download(context, file_id, dest)
    except Exception:
        log.exception("Merge download failed")
        await msg.reply_text("❌ I couldn't download that file. Please send it again.")
        return
    session["files"].append(dest)
    await msg.reply_text(f"✅ Added #{len(session['files'])}. Send more, or /done to finish.")


async def merge_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    session = context.user_data.get("merge")
    if not session:
        await msg.reply_text("You're not in merge mode. Send /merge to start.")
        return
    if not session["files"]:
        await msg.reply_text("You haven't sent anything yet. Send some images or PDFs first, or /cancel.")
        return
    status = await msg.reply_text(f"⚙️ Merging {len(session['files'])} file(s) into one PDF...")
    try:
        out = session["dir"] / "merged.pdf"
        await asyncio.to_thread(merge_files, session["files"], out, session["dir"])
        await send_file(msg, out, "merged.pdf", f"✅ Merged {len(session['files'])} file(s)")
        await status.delete()
        drop_merge_session(context)
    except UserError as e:
        await status.edit_text(f"⚠️ {e}\n\nYou can fix this and send /done again, or /cancel.")
    except Exception:
        log.exception("Merge failed")
        await status.edit_text("❌ Something went wrong while merging. One of the files may be corrupt. Try /merge again.")
        drop_merge_session(context)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    drop_merge_session(context)
    context.user_data.pop("awaiting_split", None)
    await update.message.reply_text("🚫 Cancelled.")


# ----------------------------------------------------------------------------
# /split command (alternative to the button)
# ----------------------------------------------------------------------------
async def split_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    spec = " ".join(context.args)
    if not spec:
        await msg.reply_text(
            "Usage: /split 1-3\n\nSend a PDF first (or reply to one) and then use /split with the pages, "
            "e.g. /split 2,5,7-9"
        )
        return
    info = None
    if msg.reply_to_message:
        info = context.user_data.get("files", {}).get(msg.reply_to_message.message_id)
    info = info or context.user_data.get("last_pdf")
    if not info:
        await msg.reply_text("Please send me a PDF first, then use /split.")
        return
    await process(update, context, file_id=info["file_id"], filename=info["name"], action="split", arg=spec)


# ----------------------------------------------------------------------------
# Message handlers
# ----------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Hi! I'm your PDF toolbox.\n\n"
        "• Send a Word / Excel / PowerPoint / TXT / image file → I convert it to PDF\n"
        "• Send a PDF → I can turn it into Word, images or text, compress it or extract pages\n"
        "• /merge → combine several photos or PDFs into one PDF\n\n"
        "🔒 Your files are deleted right after processing.\n"
        "Send /help for details."
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📄 To PDF: doc, docx, odt, rtf, xls, xlsx, ods, csv, ppt, pptx, odp, txt, html, "
        "jpg, png, bmp, gif, webp, tiff\n\n"
        "📚 Commands:\n"
        "/merge – start merging photos/images/PDFs into one PDF\n"
        "/done – finish and get the merged PDF\n"
        "/split 1-3 – keep only certain pages of your last PDF\n"
        "/cancel – cancel the current action\n\n"
        "📎 Send any PDF and tap a button to convert it to Word / images / text, "
        "compress it, or extract pages.\n\n"
        f"Max file size: {MAX_DOWNLOAD_MB} MB (a Telegram limit for bots).\n"
        "Tip: send photos as 'File' to keep full quality."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    doc = msg.document
    name = doc.file_name or "document"
    ext = Path(name).suffix.lower()

    if "merge" in context.user_data:
        await add_to_merge(update, context, doc.file_id, doc.file_size, ext)
        return

    if is_too_big(doc.file_size):
        await msg.reply_text(TOO_BIG_MSG)
        return

    if ext == ".pdf":
        info = remember_pdf(context, msg.message_id, doc.file_id, name)
        await msg.reply_text("📄 Got your PDF! What would you like to do?",
                             reply_markup=pdf_keyboard(msg.message_id))
        return

    if ext in IMAGE_EXT or ext in OFFICE_EXT:
        tip = "\n\nTip: use /merge to combine several images into one PDF." if ext in IMAGE_EXT else ""
        await process(update, context, file_id=doc.file_id, filename=name,
                      action="topdf", caption="✅ Done!" + tip)
        return

    await msg.reply_text(UNSUPPORTED_MSG)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    photo = msg.photo[-1]  # highest resolution

    if "merge" in context.user_data:
        await add_to_merge(update, context, photo.file_id, photo.file_size, ".jpg")
        return

    await process(update, context, file_id=photo.file_id, filename="photo.jpg", action="topdf",
                  caption="✅ Done!\n\nTip: use /merge to combine several photos into one PDF.")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    info = context.user_data.get("awaiting_split")
    if not info:
        await update.message.reply_text("Send me a file to convert, or /help to see what I can do.")
        return
    ok = await process(update, context, file_id=info["file_id"], filename=info["name"],
                       action="split", arg=update.message.text)
    if ok:
        context.user_data.pop("awaiting_split", None)


async def error_handler(update, context: ContextTypes.DEFAULT_TYPE):
    log.error("Unhandled error", exc_info=context.error)


async def post_init(app: Application):
    await app.bot.set_my_commands([
        BotCommand("start", "Start the bot"),
        BotCommand("help", "Supported formats and commands"),
        BotCommand("merge", "Merge photos/PDFs into one PDF"),
        BotCommand("done", "Finish merging and get the PDF"),
        BotCommand("split", "Extract pages, e.g. /split 1-3"),
        BotCommand("cancel", "Cancel current action"),
    ])


def main():
    cleanup_stale_dirs()
    app = Application.builder().token(BOT_TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler(["merge", "images"], merge_start))
    app.add_handler(CommandHandler("done", merge_done))
    app.add_handler(CommandHandler("cancel", cancel))
    app.add_handler(CommandHandler("split", split_cmd))
    app.add_handler(CallbackQueryHandler(on_action, pattern=r"^act:"))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)

    log.info("Bot is running...")
    app.run_polling()


if __name__ == "__main__":
    main()
