import utils.session
import io
import os
import re
import csv
import shutil
import sqlite3
import logging
import asyncio
from datetime import datetime
from PIL import Image, ExifTags
import easyocr
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="torch")
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from telegram.request import HTTPXRequest

from utils.gdrive import upload_image_to_drive, download_image_as_stream, delete_drive_file

# Logging configuration
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logging.getLogger("httpx").setLevel(logging.WARNING)

# Set up absolute paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.abspath(os.environ.get("BOT_DATA_DIR", BASE_DIR))
DOWNLOAD_DIR = os.path.join(DATA_DIR, "images")
DB_PATH = os.path.join(DATA_DIR, "reports.db")
TOKEN_PATH = os.environ.get("BOT_TOKEN_FILE", os.path.join(BASE_DIR, "token.txt"))

os.makedirs(DOWNLOAD_DIR, exist_ok=True)

AVAILABLE_BLOCKS = ["A", "B", "C", "D"]

# Configurable monthly quota target per worker
MONTHLY_TASK_TARGET = 50

# Global EasyOCR reader instance
OCR_READER = None


def get_ocr_reader():
    """Initializes the EasyOCR reader lazily to avoid delay at script load."""
    global OCR_READER
    if OCR_READER is None:
        # gpu=False ensures stability across CPU-based Windows instances
        OCR_READER = easyocr.Reader(['en'], gpu=False)
    return OCR_READER


# --- LOAD SECURE TOKEN ---
if not os.path.exists(TOKEN_PATH):
    raise FileNotFoundError(f"Could not find 'token.txt' at {TOKEN_PATH}. Please create it and paste your bot token inside.")

with open(TOKEN_PATH, "r", encoding="utf-8") as f:
    TOKEN = f.read().strip()

if not TOKEN:
    raise ValueError(f"'token.txt' at {TOKEN_PATH} is empty. Please add your Telegram Bot Token.")


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS task_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            image_path TEXT,
            description TEXT,
            date_logged TIMESTAMP,
            block_name TEXT DEFAULT 'GENERAL',
            task_type TEXT DEFAULT 'SINGLE',
            pair_image_path TEXT DEFAULT NULL,
            pair_status TEXT DEFAULT 'COMPLETED'
        )
    ''')

    # Automatic schema migration for existing databases
    cursor.execute("PRAGMA table_info(task_reports)")
    columns = [col[1] for col in cursor.fetchall()]
    if "block_name" not in columns:
        cursor.execute("ALTER TABLE task_reports ADD COLUMN block_name TEXT DEFAULT 'GENERAL'")
    if "task_type" not in columns:
        cursor.execute("ALTER TABLE task_reports ADD COLUMN task_type TEXT DEFAULT 'SINGLE'")
    if "pair_image_path" not in columns:
        cursor.execute("ALTER TABLE task_reports ADD COLUMN pair_image_path TEXT DEFAULT NULL")
    if "pair_status" not in columns:
        cursor.execute("ALTER TABLE task_reports ADD COLUMN pair_status TEXT DEFAULT 'COMPLETED'")

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_reports_date ON task_reports(date_logged)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_reports_user_date ON task_reports(user_id, date_logged)")
    cursor.execute('''
        CREATE INDEX IF NOT EXISTS idx_reports_pending_pair
        ON task_reports(user_id, block_name, task_type, pair_status, id)
    ''')

    conn.commit()
    conn.close()


def extract_image_capture_date(file_path):
    """
    Attempts to read EXIF 'DateTimeOriginal' or 'DateTime' from an image.
    Returns string 'YYYY-MM-DD HH:MM:SS' if found, else None.
    """
    try:
        with Image.open(file_path) as img:
            exif = img.getexif()
            if not exif:
                return None

            exif_table = {
                ExifTags.TAGS.get(k, k): v 
                for k, v in exif.items()
            }

            raw_date = exif_table.get('DateTimeOriginal') or exif_table.get('DateTime')

            if not raw_date:
                try:
                    exif_ifd = exif.get_ifd(0x8769)
                    if exif_ifd:
                        ifd_table = {
                            ExifTags.TAGS.get(k, k): v 
                            for k, v in exif_ifd.items()
                        }
                        raw_date = ifd_table.get('DateTimeOriginal') or ifd_table.get('DateTime')
                except Exception:
                    pass

            if raw_date and isinstance(raw_date, str):
                parts = raw_date.strip().split(" ")
                date_part = parts[0].replace(":", "-")
                time_part = parts[1] if len(parts) > 1 else "12:00:00"
                if re.match(r"^\d{4}-\d{2}-\d{2}$", date_part):
                    return f"{date_part} {time_part}"
    except Exception as e:
        logging.debug(f"Could not read EXIF data for {file_path}: {e}")

    return None


def extract_visual_timestamp_ocr(file_path):
    """
    Scans the image using EasyOCR to detect printed date watermarks.
    Parses timestamps like '29 Oct 2026 at 5.31.48 PM' as well as standard formats.
    """
    month_map = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'jun': '06',
        'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }

    try:
        reader = get_ocr_reader()
        with Image.open(file_path) as img:
            w, h = img.size
            # Check bottom 35%, top 25%, then entire image
            bottom_strip = img.crop((0, int(h * 0.65), w, h))
            top_strip = img.crop((0, 0, w, int(h * 0.25)))
            full_img = img.copy()

            targets = [bottom_strip, top_strip, full_img]

        for target_img in targets:
            buf = io.BytesIO()
            target_img.save(buf, format="JPEG", quality=90)
            buf.seek(0)
            
            ocr_results = reader.readtext(buf.read(), detail=0)
            if not ocr_results:
                continue

            raw_text = " ".join(ocr_results)
            # Normalize common OCR noise: replace letters 'o'/'O' where digits are expected
            normalized_text = re.sub(r'(?<=\d)[oO]|(?<=[/\.-])[oO]|[oO](?=\d)', '0', raw_text)

            # --- Pattern 1: GPS Map / Timestamp Camera format (e.g., '29 Oct 2026 at 5.31.48 PM') ---
            match_camera_stamp = re.search(
                r'\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(20\d{2})(?:\s+at\s+(\d{1,2})[\.:](\d{2})(?:[\.:](\d{2}))?\s*(AM|PM)?)?',
                normalized_text,
                re.IGNORECASE
            )
            if match_camera_stamp:
                d, mon_str, y, hr, mn, sc, meridiem = match_camera_stamp.groups()
                mon_key = mon_str[:3].lower()
                if mon_key in month_map:
                    month_num = month_map[mon_key]
                    
                    # Convert 12-hour AM/PM to 24-hour SQL time
                    if hr and mn:
                        h_val = int(hr)
                        m_val = int(mn)
                        s_val = int(sc) if sc else 0
                        if meridiem:
                            meridiem = meridiem.upper()
                            if meridiem == 'PM' and h_val < 12:
                                h_val += 12
                            elif meridiem == 'AM' and h_val == 12:
                                h_val = 0
                        time_str = f"{h_val:02d}:{m_val:02d}:{s_val:02d}"
                    else:
                        time_str = "12:00:00"
                        
                    return f"{y}-{month_num}-{int(d):02d} {time_str}"

            # --- Pattern 2: Numeric YYYY-MM-DD or YYYY/MM/DD ---
            match_ymd = re.search(r'\b(20\d{2})[/\.-](\d{1,2})[/\.-](\d{1,2})\b', normalized_text)
            if match_ymd:
                y, m, d = match_ymd.groups()
                if 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                    time_match = re.search(r'\b(\d{1,2})[\.:](\d{2})(?:[\.:](\d{2}))?\b', normalized_text)
                    t_str = f"{time_match.group(1).zfill(2)}:{time_match.group(2)}:{time_match.group(3) or '00'}" if time_match else "12:00:00"
                    return f"{y}-{int(m):02d}-{int(d):02d} {t_str}"

            # --- Pattern 3: Numeric DD/MM/YYYY or DD-MM-YYYY ---
            match_dmy = re.search(r'\b(\d{1,2})[/\.-](\d{1,2})[/\.-](20\d{2})\b', normalized_text)
            if match_dmy:
                d, m, y = match_dmy.groups()
                if 1 <= int(m) <= 12 and 1 <= int(d) <= 31:
                    time_match = re.search(r'\b(\d{1,2})[\.:](\d{2})(?:[\.:](\d{2}))?\b', normalized_text)
                    t_str = f"{time_match.group(1).zfill(2)}:{time_match.group(2)}:{time_match.group(3) or '00'}" if time_match else "12:00:00"
                    return f"{y}-{int(m):02d}-{int(d):02d} {t_str}"

    except Exception as e:
        logging.warning(f"OCR extraction failed on {file_path}: {e}")

    return None


def compress_image(file_path, max_dim=1280, quality=80):
    try:
        with Image.open(file_path) as img:
            img = img.convert("RGB")
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            img.save(file_path, "JPEG", optimize=True, quality=quality)
    except Exception as e:
        print(f"Image compression failed for {file_path}: {e}")


def insert_image_into_paragraph(paragraph, img_reference, width, height, missing_text="[Image Missing]"):
    """
    Inserts an image into a python-docx paragraph.
    Handles in-memory streams from Google Drive, existing local files, or missing fallbacks.
    """
    if not img_reference:
        paragraph.add_run(missing_text)
        return

    stream = download_image_as_stream(img_reference)
    if stream:
        try:
            paragraph.add_run().add_picture(stream, width=width, height=height)
            return
        except Exception as e:
            print(f"Failed to add image stream {img_reference}: {e}")
            paragraph.add_run("[Corrupted Image]")
            return

    if os.path.exists(img_reference):
        try:
            paragraph.add_run().add_picture(img_reference, width=width, height=height)
            return
        except Exception:
            paragraph.add_run("[Corrupted Image]")
            return

    paragraph.add_run(missing_text)


def generate_monthly_docx(output_docx_path, year_month_filter, target_block=None):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    if target_block:
        cursor.execute('''
            SELECT image_path, description, date_logged, block_name, task_type, pair_image_path, pair_status 
            FROM task_reports 
            WHERE date_logged LIKE ? 
              AND REPLACE(UPPER(block_name), ' ', '') = REPLACE(UPPER(?), ' ', '')
            ORDER BY date_logged ASC, block_name ASC
        ''', (f"{year_month_filter}%", target_block))
    else:
        cursor.execute('''
            SELECT image_path, description, date_logged, block_name, task_type, pair_image_path, pair_status 
            FROM task_reports 
            WHERE date_logged LIKE ? 
            ORDER BY date_logged ASC, block_name ASC
        ''', (f"{year_month_filter}%",))
        
    records = cursor.fetchall()
    conn.close()

    if not records:
        return None

    grouped = {}
    for img_ref, desc, dt, block, t_type, pair_img_ref, p_status in records:
        day_str = dt.split(" ")[0]
        item = {
            "img": img_ref,
            "desc": desc,
            "dt": dt,
            "type": t_type,
            "pair_img": pair_img_ref,
            "status": p_status
        }
        grouped.setdefault(day_str, {}).setdefault(block, []).append(item)

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    month_dt = datetime.strptime(year_month_filter, "%Y-%m")
    month_name_upper = month_dt.strftime("%B %Y").upper()

    title_suffix = f" - {target_block.upper()}" if target_block else ""
    title = doc.add_heading(f"MONTHLY CLEANING ACTIVITY REPORT - {month_name_upper}{title_suffix}", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for day_str, blocks in grouped.items():
        formatted_date = ".".join(day_str.split("-")[::-1])
        
        h2 = doc.add_heading(f"Date: {formatted_date}", level=2)
        h2.paragraph_format.space_before = Pt(14)
        h2.paragraph_format.space_after = Pt(4)

        for block_name, items in blocks.items():
            h3 = doc.add_heading(f"{block_name}", level=3)
            h3.paragraph_format.space_before = Pt(8)
            h3.paragraph_format.space_after = Pt(4)

            paired_tasks = [it for it in items if it["type"] == "PAIR"]
            single_tasks = [it for it in items if it["type"] != "PAIR"]

            for p_task in paired_tasks:
                comp_table = doc.add_table(rows=2, cols=2)
                comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER

                comp_table.cell(0, 0).paragraphs[0].text = "🔴 BEFORE"
                comp_table.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                comp_table.cell(0, 1).paragraphs[0].text = "🟢 AFTER"
                comp_table.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                c_before = comp_table.cell(1, 0)
                p_b = c_before.paragraphs[0]
                p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
                insert_image_into_paragraph(p_b, p_task["img"], width=Inches(3.2), height=Inches(2.4))

                c_after = comp_table.cell(1, 1)
                p_a = c_after.paragraphs[0]
                p_a.alignment = WD_ALIGN_PARAGRAPH.CENTER
                insert_image_into_paragraph(
                    p_a,
                    p_task["pair_img"],
                    width=Inches(3.2),
                    height=Inches(2.4),
                    missing_text="[Pending After Photo]"
                )

                caption_p = doc.add_paragraph()
                caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_cap = caption_p.add_run(f"{p_task['desc']} ({p_task['status']})")
                r_cap.font.size = Pt(8.5)
                r_cap.font.color.rgb = RGBColor(90, 90, 90)

            if single_tasks:
                cols = 4
                num_rows = (len(single_tasks) + cols - 1) // cols
                table = doc.add_table(rows=num_rows, cols=cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER

                for idx, s_task in enumerate(single_tasks):
                    r = idx // cols
                    c = idx % cols
                    cell = table.cell(r, c)
                    cell_p = cell.paragraphs[0]
                    cell_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    insert_image_into_paragraph(cell_p, s_task["img"], width=Inches(1.75), height=Inches(1.75))

                    desc_p = cell.add_paragraph()
                    desc_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_desc = desc_p.add_run(s_task["desc"])
                    run_desc.font.size = Pt(8)
                    run_desc.font.color.rgb = RGBColor(90, 90, 90)

    doc.save(output_docx_path)
    return output_docx_path


def generate_monthly_csv(output_csv_path, year_month_filter, target_block=None):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    if target_block:
        cursor.execute('''
            SELECT id, date_logged, user_id, block_name, description, task_type, pair_status 
            FROM task_reports 
            WHERE date_logged LIKE ? 
              AND REPLACE(UPPER(block_name), ' ', '') = REPLACE(UPPER(?), ' ', '')
            ORDER BY date_logged ASC
        ''', (f"{year_month_filter}%", target_block))
    else:
        cursor.execute('''
            SELECT id, date_logged, user_id, block_name, description, task_type, pair_status 
            FROM task_reports 
            WHERE date_logged LIKE ? 
            ORDER BY date_logged ASC
        ''', (f"{year_month_filter}%",))

    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return None

    with open(output_csv_path, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(["Task ID", "Date", "Time", "User ID", "Worker Name", "Block", "Type", "Status", "Description"])

        for task_id, dt_str, user_id, block, desc, t_type, p_status in rows:
            dt_parts = dt_str.split(" ")
            date_part = dt_parts[0]
            time_part = dt_parts[1] if len(dt_parts) > 1 else ""

            worker_match = re.match(r"^\[(.*?)\]\s*(.*)", desc)
            if worker_match:
                worker_name = worker_match.group(1)
                task_desc = worker_match.group(2)
            else:
                worker_name = f"User_{user_id}"
                task_desc = desc

            writer.writerow([task_id, date_part, time_part, user_id, worker_name, block, t_type, p_status, task_desc])

    return output_csv_path


async def check_admin_rights(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """Allows all users to access report and management features."""
    return True


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    menu_text = (
        "🧹 **FACILITY CLEANING REPORT BOT**\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "Welcome! Use this bot to submit daily cleaning logs, track progress, and generate client reports.\n\n"

        "📸 **HOW TO SUBMIT A TASK**\n"
        "1. Attach a photo (or uncompressed file).\n"
        "2. Add your description and `#block` tag in the caption.\n"
        "   • Example: `#blockA Restroom 2 sanitized`\n"
        "   • *(If you forget `#block`, tap the inline button that appears!)*\n"
        "   • *Date Note: The bot reads photo EXIF metadata, scans visual timestamp watermarks (OCR), or accepts `YYYY-MM-DD` in the caption.*\n\n"

        "🔄 **BEFORE & AFTER PHOTO PAIRING (Optional)**\n"
        "• **Step 1 (Dirty):** Send photo with `#before #blockA Restroom 2 sink`\n"
        "• **Step 2 (Clean):** Send photo with `#after #blockA Restroom 2 sink`\n\n"

        "📋 **GENERAL COMMANDS**\n"
        "• `/mystats` — View monthly task targets and progress bar\n"
        "• `/dailyreport` — Check today's shift activity summary\n"
        "• `/dailyreport YYYY-MM-DD` — View summary for a specific date\n"
        "• `/undo` — Delete your own most recent task submission\n"
        "• `/undo <task_id>` — Delete a specific task by its ID\n"
        "• `/myid` — View your numeric Telegram user ID\n\n"

        "📊 **REPORTS & EXPORTS**\n"
        "• `/reportmonthly [MM-YYYY] [block]` — Generate Word (.docx) report\n"
        "   └ *Example:* `/reportmonthly 10-2026 blockA`\n"
        "• `/exportcsv [MM-YYYY] [block]` — Export data spreadsheet (.csv)\n"
        "• `/backup` — Download instant SQLite (`reports.db`) snapshot\n"
        "• `/cleardata MM-YYYY` — Purge Drive images and database logs for month\n"
    )

    await update.message.reply_text(menu_text, parse_mode="Markdown")


async def my_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    await update.message.reply_text(f"Your Telegram User ID is: `{user_id}`", parse_mode="Markdown")


async def handle_photo_or_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_user or not update.message:
        return

    user_id = update.effective_user.id
    user_name = update.effective_user.first_name or f"User_{user_id}"

    is_photo = bool(update.message.photo)
    is_document = bool(update.message.document)

    if not is_photo and not is_document:
        return

    if is_document:
        doc = update.message.document
        mime = doc.mime_type or ""
        valid_mimes = ["image/jpeg", "image/png", "image/webp"]
        if not any(mime.startswith(m) for m in valid_mimes):
            await update.message.reply_text(
                f"⚠️ File rejected: `{doc.file_name}` is not an image.\n"
                f"Please upload photos or image files (.jpg, .png, .webp).",
                parse_mode="Markdown",
                reply_to_message_id=update.message.message_id
            )
            return

    raw_caption = update.message.caption
    if not raw_caption or not raw_caption.strip():
        await update.message.reply_text(
            f"❌ Submission rejected ({user_name}): Description is required.\n"
            f"Please send the photo with a caption (e.g., `#block b Cleaned restroom`).",
            reply_to_message_id=update.message.message_id
        )
        return

    cleaned_text = raw_caption.strip()

    if is_photo:
        tg_media = update.message.photo[-1]
    else:
        tg_media = update.message.document

    # 1. Download file locally to check EXIF / OCR and compress
    file = await context.bot.get_file(tg_media.file_id)
    file_name = f"{tg_media.file_unique_id}.jpg"
    file_path = os.path.join(DOWNLOAD_DIR, file_name)
    await file.download_to_drive(file_path)

    # 2. Extract original capture date from EXIF before compression
    exif_timestamp = extract_image_capture_date(file_path)

    # 3. Determine date priority: Manual caption > Camera EXIF > OCR Watermark > Upload time
    date_match = re.match(r"^(\d{4}-\d{2}-\d{2})\s*(.*)", cleaned_text)
    if date_match:
        test_date = date_match.group(1)
        remaining_caption = date_match.group(2).strip()
        timestamp = f"{test_date} 12:00:00"
        date_source_label = "manual date"
    else:
        remaining_caption = cleaned_text
        if exif_timestamp:
            timestamp = exif_timestamp
            date_source_label = "photo capture date (EXIF)"
        else:
            # Run OCR in background executor so it doesn't block the async event loop
            loop = asyncio.get_running_loop()
            ocr_timestamp = await loop.run_in_executor(None, extract_visual_timestamp_ocr, file_path)

            if ocr_timestamp:
                timestamp = ocr_timestamp
                date_source_label = "camera watermark (OCR)"
            else:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                date_source_label = "upload time"

    # 4. Compress image
    compress_image(file_path)

    # 5. Upload to Google Drive; ONLY delete local file if Drive upload genuinely succeeded
    drive_file_id = upload_image_to_drive(file_path, file_name)

    if drive_file_id:
        stored_image_reference = drive_file_id
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass
    else:
        # Keep local copy permanently if Drive upload is inactive/failing
        stored_image_reference = file_path

    # Parse task details
    is_before = bool(re.search(r"#before\b", remaining_caption, re.IGNORECASE))
    is_after = bool(re.search(r"#after\b", remaining_caption, re.IGNORECASE))
    remaining_caption = re.sub(r"#(before|after)\b", "", remaining_caption, flags=re.IGNORECASE).strip()

    block_match = re.search(r"#block[\s\-_]*([a-zA-Z0-9]+)", remaining_caption, re.IGNORECASE)
    has_block = bool(block_match)

    if has_block:
        block_identifier = block_match.group(1).upper()
        block_name = f"BLOCK {block_identifier}"
        final_caption = re.sub(r"#block[\s\-_]*[a-zA-Z0-9]+", "", remaining_caption, flags=re.IGNORECASE).strip()
    else:
        block_name = "UNASSIGNED"
        final_caption = remaining_caption.strip()

    caption = final_caption or "Task completed"
    full_description = f"[{user_name}] {caption}"

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    date_label = timestamp.split(' ')[0]

    # --- HANDLE #AFTER PHOTO PAIRING ---
    if is_after:
        cursor.execute('''
            SELECT id, image_path, description FROM task_reports 
            WHERE user_id = ? AND block_name = ? AND task_type = 'PAIR' AND pair_status = 'PENDING'
            ORDER BY id DESC LIMIT 1
        ''', (user_id, block_name))
        open_task = cursor.fetchone()

        if open_task:
            task_id, before_img, orig_desc = open_task
            cursor.execute('''
                UPDATE task_reports 
                SET pair_image_path = ?, pair_status = 'COMPLETED', description = ?
                WHERE id = ?
            ''', (stored_image_reference, full_description, task_id))
            conn.commit()
            conn.close()

            await update.message.reply_text(
                f"🟢 **Paired Successfully!**\n"
                f"Linked AFTER photo to Task #{task_id} [{block_name}].\n"
                f"• Assigned Date: `{date_label}` ({date_source_label})\n"
                f"Task status is now **COMPLETED**.",
                parse_mode="Markdown",
                reply_to_message_id=update.message.message_id
            )
            return
        else:
            cursor.execute('''
                INSERT INTO task_reports (user_id, image_path, description, date_logged, block_name, task_type, pair_image_path, pair_status)
                VALUES (?, ?, ?, ?, ?, 'SINGLE', NULL, 'COMPLETED')
            ''', (user_id, stored_image_reference, full_description + " (Unmatched #after)", timestamp, block_name))
            task_id = cursor.lastrowid
            conn.commit()
            conn.close()

            await update.message.reply_text(
                f"⚠️ Logged as Single Task #{task_id} [{block_name}].\n"
                f"No open `#before` task was found to link this `#after` photo.\n"
                f"• Assigned Date: `{date_label}` ({date_source_label})",
                reply_to_message_id=update.message.message_id
            )
            return

    # --- HANDLE #BEFORE PHOTO LOGGING ---
    if is_before:
        cursor.execute('''
            INSERT INTO task_reports (user_id, image_path, description, date_logged, block_name, task_type, pair_image_path, pair_status)
            VALUES (?, ?, ?, ?, ?, 'PAIR', NULL, 'PENDING')
        ''', (user_id, stored_image_reference, full_description, timestamp, block_name))
        task_id = cursor.lastrowid
        conn.commit()
        conn.close()

        await update.message.reply_text(
            f"🔴 **Before Photo Logged!** (Task #{task_id} [{block_name}])\n"
            f"• Assigned Date: `{date_label}` ({date_source_label})\n"
            f"Waiting for clean photo. Send next photo with `#after #{block_name.lower().replace(' ', '')}` to complete.",
            reply_to_message_id=update.message.message_id
        )
        return

    # --- STANDARD SINGLE TASK LOGGING ---
    cursor.execute('''
        INSERT INTO task_reports (user_id, image_path, description, date_logged, block_name, task_type, pair_image_path, pair_status)
        VALUES (?, ?, ?, ?, ?, 'SINGLE', NULL, 'COMPLETED')
    ''', (user_id, stored_image_reference, full_description, timestamp, block_name))
    task_id = cursor.lastrowid
    conn.commit()
    conn.close()

    if has_block:
        await update.message.reply_text(
            f"✅ Task logged by {user_name} [{block_name}] (ID: #{task_id})\n"
            f"• Date: `{date_label}` ({date_source_label})\n"
            f"Tip: Made a mistake? Type /undo to cancel this submission.",
            reply_to_message_id=update.message.message_id
        )
    else:
        keyboard = []
        row = []
        for blk in AVAILABLE_BLOCKS:
            row.append(InlineKeyboardButton(f"🏢 Block {blk}", callback_data=f"setblk:{task_id}:{user_id}:{blk}"))
            if len(row) == 2:
                keyboard.append(row)
                row = []
        if row:
            keyboard.append(row)

        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text(
            f"⚠️ Task logged by {user_name} (ID: #{task_id})\n"
            f"• Date: `{date_label}` ({date_source_label})\n"
            f"No `#block` tag detected. Please select the correct block:",
            reply_markup=reply_markup,
            reply_to_message_id=update.message.message_id
        )


async def block_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    data_parts = query.data.split(":")
    if len(data_parts) != 4 or data_parts[0] != "setblk":
        return

    task_id = int(data_parts[1])
    selected_block_letter = data_parts[3].upper()
    selected_block = f"BLOCK {selected_block_letter}"

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE task_reports SET block_name = ? WHERE id = ?", (selected_block, task_id))
    cursor.execute("SELECT date_logged, user_id FROM task_reports WHERE id = ?", (task_id,))
    row = cursor.fetchone()
    conn.commit()
    conn.close()

    date_label = row[0].split(" ")[0] if row else "Today"
    author_name = query.from_user.first_name

    await query.edit_message_text(
        f"✅ Task #{task_id} assigned to [{selected_block}] by {author_name} for date: {date_label}"
    )


async def daily_report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.args:
        raw_date = context.args[0].strip()
        if re.match(r"^\d{4}-\d{2}-\d{2}$", raw_date):
            target_date = raw_date
        elif re.match(r"^\d{2}-\d{2}-\d{4}$", raw_date):
            p = raw_date.split("-")
            target_date = f"{p[2]}-{p[1]}-{p[0]}"
        else:
            await update.message.reply_text("Invalid date format. Use `/dailyreport YYYY-MM-DD` (e.g. `/dailyreport 2026-10-12`).")
            return
    else:
        target_date = datetime.now().strftime("%Y-%m-%d")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM task_reports WHERE date_logged LIKE ?", (f"{target_date}%",))
    total_tasks = cursor.fetchone()[0]

    if total_tasks == 0:
        conn.close()
        formatted_date = ".".join(target_date.split("-")[::-1])
        await update.message.reply_text(f"ℹ️ No tasks recorded for date `{formatted_date}`.")
        return

    cursor.execute('''
        SELECT block_name, COUNT(*) FROM task_reports 
        WHERE date_logged LIKE ? 
        GROUP BY block_name 
        ORDER BY COUNT(*) DESC
    ''', (f"{target_date}%",))
    block_counts = cursor.fetchall()

    cursor.execute("SELECT description FROM task_reports WHERE date_logged LIKE ?", (f"{target_date}%",))
    rows = cursor.fetchall()
    conn.close()

    worker_counts = {}
    for (desc,) in rows:
        m = re.match(r"^\[(.*?)\]", desc)
        name = m.group(1) if m else "Unknown"
        worker_counts[name] = worker_counts.get(name, 0) + 1

    formatted_date = ".".join(target_date.split("-")[::-1])
    lines = [
        f"📊 **DAILY SHIFT SUMMARY ({formatted_date})**",
        f"━━━━━━━━━━━━━━━━━━",
        f"• **Total Tasks Logged:** {total_tasks}",
        "",
        "📍 **Breakdown by Location:**"
    ]
    for block_name, count in block_counts:
        lines.append(f"  ▫️ {block_name}: **{count}** tasks")

    lines.append("")
    lines.append("👷‍♂️ **Active Workers:**")
    for worker_name, count in sorted(worker_counts.items(), key=lambda x: x[1], reverse=True):
        lines.append(f"  ▫️ {worker_name}: {count} tasks")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def undo_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not update.effective_user:
        return

    user_id = update.effective_user.id
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    if context.args:
        target_arg = context.args[0].replace("#", "").strip()
        if not target_arg.isdigit():
            await update.message.reply_text("Usage: `/undo` or `/undo <task_id>`.")
            conn.close()
            return
        
        target_task_id = int(target_arg)
        cursor.execute("SELECT id, user_id, image_path, description, block_name, pair_image_path FROM task_reports WHERE id = ?", (target_task_id,))
        task = cursor.fetchone()
    else:
        cursor.execute(
            "SELECT id, user_id, image_path, description, block_name, pair_image_path FROM task_reports WHERE user_id = ? ORDER BY id DESC LIMIT 1",
            (user_id,)
        )
        task = cursor.fetchone()

    if not task:
        conn.close()
        await update.message.reply_text("ℹ️ No task found to undo.")
        return

    task_id, task_author_id, img_ref, desc, block, pair_img_ref = task

    for ref in [img_ref, pair_img_ref]:
        if ref:
            delete_drive_file(ref)
            if os.path.exists(ref):
                try:
                    os.remove(ref)
                except Exception:
                    pass

    cursor.execute("DELETE FROM task_reports WHERE id = ?", (task_id,))
    conn.commit()
    conn.close()

    await update.message.reply_text(
        f"🗑️ **Task Undone & Deleted Successfully**\n"
        f"• **Task ID:** #{task_id}\n"
        f"• **Location:** {block}\n"
        f"• **Details:** {desc}\n"
        f"• **Images:** Purged from storage.",
        parse_mode="Markdown"
    )


async def backup_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not os.path.exists(DB_PATH):
        await update.message.reply_text("⚠️ Database file `reports.db` does not exist yet.")
        return

    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"reports_backup_{timestamp_str}.db"
    backup_temp_path = os.path.join(DATA_DIR, backup_filename)

    status_msg = await update.message.reply_text("📦 Generating database snapshot...")

    try:
        shutil.copy2(DB_PATH, backup_temp_path)

        conn = sqlite3.connect(backup_temp_path)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM task_reports")
        total_rows = cursor.fetchone()[0]
        conn.close()

        file_size_kb = round(os.path.getsize(backup_temp_path) / 1024, 2)

        with open(backup_temp_path, "rb") as db_file:
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=db_file,
                filename=backup_filename,
                caption=(
                    f"🛡️ **Database Backup**\n"
                    f"• **Timestamp:** `{timestamp_str}`\n"
                    f"• **Total Records:** {total_rows} tasks\n"
                    f"• **Size:** {file_size_kb} KB"
                ),
                parse_mode="Markdown"
            )

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"❌ Failed to generate database backup: {e}")
    finally:
        if os.path.exists(backup_temp_path):
            try:
                os.remove(backup_temp_path)
            except Exception:
                pass


async def my_stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Displays monthly stats with a visual progress bar towards the quota target."""
    user_id = update.effective_user.id
    user_name = update.effective_user.first_name or "Worker"

    current_ym = datetime.now().strftime("%Y-%m")
    month_name = datetime.now().strftime("%B %Y")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        SELECT COUNT(*) FROM task_reports 
        WHERE user_id = ? AND date_logged LIKE ?
    ''', (user_id, f"{current_ym}%"))
    total_count = cursor.fetchone()[0]

    if total_count == 0:
        conn.close()
        await update.message.reply_text(
            f"📊 {user_name}, you have not logged any tasks for {month_name} yet.\n"
            f"🎯 Monthly Target: {MONTHLY_TASK_TARGET} tasks."
        )
        return

    cursor.execute('''
        SELECT block_name, COUNT(*) FROM task_reports 
        WHERE user_id = ? AND date_logged LIKE ?
        GROUP BY block_name 
        ORDER BY COUNT(*) DESC
    ''', (user_id, f"{current_ym}%"))
    breakdown = cursor.fetchall()
    conn.close()

    pct = (total_count / MONTHLY_TASK_TARGET) * 100
    filled_blocks = min(10, int(round(pct / 10)))
    empty_blocks = 10 - filled_blocks
    progress_bar = "█" * filled_blocks + "░" * empty_blocks

    if total_count >= MONTHLY_TASK_TARGET:
        pacing_text = f"🏆 **Target Achieved!** ({total_count - MONTHLY_TASK_TARGET} tasks above goal)"
    else:
        remaining = MONTHLY_TASK_TARGET - total_count
        pacing_text = f"🎯 **{remaining} more tasks** to reach your monthly target."

    msg_lines = [
        f"📊 **Monthly Stats for {user_name} ({month_name})**",
        f"━━━━━━━━━━━━━━━━━━",
        f"• **Total Tasks Logged:** {total_count} / {MONTHLY_TASK_TARGET}",
        f"• **Progress:** `[{progress_bar}]` {pct:.1f}%",
        f"• {pacing_text}",
        "",
        "📍 **Breakdown by Location:**"
    ]
    for block_name, count in breakdown:
        msg_lines.append(f"  ▫️ {block_name}: **{count}** tasks")

    await update.message.reply_text("\n".join(msg_lines), parse_mode="Markdown")


def parse_month_and_block_args(args):
    month_lookup = {
        'jan': '01', 'january': '01',
        'feb': '02', 'february': '02',
        'mar': '03', 'march': '03',
        'apr': '04', 'april': '04',
        'may': '05',
        'jun': '06', 'june': '06',
        'jul': '07', 'july': '07',
        'aug': '08', 'august': '08',
        'sep': '09', 'september': '09',
        'oct': '10', 'october': '10',
        'nov': '11', 'november': '11',
        'dec': '12', 'december': '12',
    }

    target_ym = datetime.now().strftime("%Y-%m")
    target_block = None

    joined_args = " ".join(args).strip().lower()

    # Detect written month name + year (e.g., 'october 2026' or 'oct 2026')
    written_match = re.search(r'\b([a-z]{3,9})\s+(20\d{2})\b', joined_args)
    if written_match:
        m_str, y_str = written_match.groups()
        if m_str in month_lookup:
            target_ym = f"{y_str}-{month_lookup[m_str]}"
            joined_args = re.sub(r'\b[a-z]{3,9}\s+20\d{2}\b', '', joined_args).strip()

    # Detect remaining standalone block arguments
    for val in joined_args.split():
        if re.match(r"^\d{2}-\d{4}$", val):
            parts = val.split("-")
            target_ym = f"{parts[1]}-{parts[0]}"
        elif re.match(r"^\d{4}-\d{2}$", val):
            target_ym = val
        elif val:
            m = re.search(r"(?:#?block[\s\-_]*)?([a-zA-Z0-9]+)", val)
            if m:
                target_block = f"BLOCK {m.group(1).upper()}"

    return target_ym, target_block


async def report_monthly_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    target_ym, target_block = parse_month_and_block_args(context.args)
    filter_desc = f"`{target_ym}`" + (f" (Filtered: `{target_block}`)" if target_block else "")
    status_msg = await update.message.reply_text(f"Compiling Word report for {filter_desc}...")

    block_suffix = f"_{target_block.replace(' ', '_')}" if target_block else ""
    output_docx = os.path.join(DATA_DIR, f"Monthly_Report_{target_ym}{block_suffix}.docx")

    loop = asyncio.get_running_loop()
    generated_path = await loop.run_in_executor(
        None, generate_monthly_docx, output_docx, target_ym, target_block
    )

    if not generated_path or not os.path.exists(generated_path):
        await status_msg.edit_text(f"No records found in database for {filter_desc}.")
        return

    with open(generated_path, 'rb') as doc:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=doc,
            filename=f"Report_{target_ym}{block_suffix}.docx",
            caption=f"Here is your Task Activity Word Report for {target_ym}."
        )

    await status_msg.delete()


async def export_csv_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    target_ym, target_block = parse_month_and_block_args(context.args)
    filter_desc = f"`{target_ym}`" + (f" (Filtered: `{target_block}`)" if target_block else "")
    status_msg = await update.message.reply_text(f"Exporting CSV records for {filter_desc}...")

    block_suffix = f"_{target_block.replace(' ', '_')}" if target_block else ""
    output_csv = os.path.join(DATA_DIR, f"Task_Log_{target_ym}{block_suffix}.csv")
    generated_path = generate_monthly_csv(output_csv, target_ym, target_block)

    if not generated_path or not os.path.exists(generated_path):
        await status_msg.edit_text(f"No records found in database for {filter_desc}.")
        return

    with open(generated_path, 'rb') as doc:
        await context.bot.send_document(
            chat_id=update.effective_chat.id,
            document=doc,
            filename=f"Tasks_{target_ym}{block_suffix}.csv",
            caption=f"📊 Task data spreadsheet export for {target_ym}."
        )

    await status_msg.delete()


async def clear_data_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text("Usage: `/cleardata MM-YYYY`\nExample: `/cleardata 10-2026`")
        return

    user_arg = context.args[0].strip()
    if re.match(r"^\d{2}-\d{4}$", user_arg):
        parts = user_arg.split("-")
        target_ym = f"{parts[1]}-{parts[0]}"
    elif re.match(r"^\d{4}-\d{2}$", user_arg):
        target_ym = user_arg
    else:
        await update.message.reply_text("Invalid format. Use `/cleardata MM-YYYY` (e.g., `/cleardata 10-2026`).")
        return

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('SELECT image_path, pair_image_path FROM task_reports WHERE date_logged LIKE ?', (f"{target_ym}%",))
    rows = cursor.fetchall()

    if not rows:
        conn.close()
        await update.message.reply_text(f"No records found for month `{target_ym}` to delete.")
        return

    deleted_drive_count = 0
    for img1, img2 in rows:
        for ref in [img1, img2]:
            if ref:
                if delete_drive_file(ref):
                    deleted_drive_count += 1
                if os.path.exists(ref):
                    try:
                        os.remove(ref)
                    except Exception:
                        pass

    cursor.execute('DELETE FROM task_reports WHERE date_logged LIKE ?', (f"{target_ym}%",))
    conn.commit()
    conn.close()

    await update.message.reply_text(
        f"🧹 Cleaned data for `{target_ym}`:\n"
        f"• Database records removed: {len(rows)}\n"
        f"• Images deleted: {deleted_drive_count}"
    )


async def handle_unrecognized_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Alerts users who send plain text or unrecognized commands."""
    if not update.effective_message:
        return

    chat_type = update.effective_chat.type
    user_text = update.effective_message.text or ""

    if chat_type in ["group", "supergroup"] and not user_text.startswith("/"):
        return

    alert_msg = (
        "⚠️ **Action Not Recognized**\n\n"
        "This bot only accepts **task photos** or designated **commands**.\n\n"
        "• **To log work:** Send a photo with a caption (e.g., `#blockA Cleaned lobby`).\n"
        "• **View valid commands:** Send /start."
    )

    await update.message.reply_text(
        alert_msg,
        parse_mode="Markdown",
        reply_to_message_id=update.message.message_id
    )


async def shift_reminder_worker(app):
    """Background task that checks every 30s and fires at 8:30 AM local time."""
    while True:
        now = datetime.now()
        if now.hour == 8 and now.minute == 30:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT user_id FROM task_reports")
            target_chat_ids = [row[0] for row in cursor.fetchall()]
            conn.close()

            reminder_text = (
                "☀️ **MORNING SHIFT REMINDER**\n\n"
                "Good morning team! Please remember to log your tasks today:\n"
                "• Send your task photos with captions\n"
                "• Tag your location (e.g., `#blockA Cleaned hallway`)\n"
                "• Use `#before` and `#after` for deep-cleaning tasks!"
            )

            for chat_id in target_chat_ids:
                try:
                    await app.bot.send_message(chat_id=chat_id, text=reminder_text, parse_mode="Markdown")
                except Exception as e:
                    print(f"Failed to send reminder to {chat_id}: {e}")

            await asyncio.sleep(65)

        await asyncio.sleep(30)


async def post_init(application):
    """Launches the background reminder worker as soon as the event loop starts."""
    asyncio.create_task(shift_reminder_worker(application))


if __name__ == '__main__':
    init_db()
    
    # Configure generous network timeouts for media downloads & uploads
    request_config = HTTPXRequest(
        connect_timeout=20.0,
        read_timeout=60.0,
        write_timeout=60.0,
        pool_timeout=10.0,
    )

    app = ApplicationBuilder().token(TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("myid", my_id))
    app.add_handler(CommandHandler("mystats", my_stats_command))
    app.add_handler(CommandHandler("dailyreport", daily_report_command))
    app.add_handler(CommandHandler("undo", undo_command))
    app.add_handler(CommandHandler("backup", backup_command))
    app.add_handler(CommandHandler("reportmonthly", report_monthly_command))
    app.add_handler(CommandHandler("exportcsv", export_csv_command))
    app.add_handler(CommandHandler("cleardata", clear_data_command))
    app.add_handler(CallbackQueryHandler(block_callback_handler, pattern=r"^setblk:"))

    media_filter = filters.PHOTO | (
        filters.Document.IMAGE
        | filters.Document.MimeType("image/jpeg")
        | filters.Document.MimeType("image/png")
        | filters.Document.MimeType("image/webp")
    )
    app.add_handler(MessageHandler(media_filter, handle_photo_or_document))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_unrecognized_text))
    app.add_handler(MessageHandler(filters.COMMAND, handle_unrecognized_text))

    print(f"Bot running... Database at: {DB_PATH}")
    app.run_polling(drop_pending_updates=True)