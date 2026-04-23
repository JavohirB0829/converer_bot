"""
Universal Converter Bot — aiogram 3.x | Windows optimized
Qo'llab-quvvatlanadigan konvertatsiyalar:
  • Bir nechta rasm  →  PDF  (Pillow)
  • DOCX / XLSX / PPTX  →  PDF  (LibreOffice)
  • PDF  →  DOCX  (pdf2docx)
"""

import asyncio
import logging
import os
import subprocess
import uuid
from pathlib import Path

from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.types import (
    BufferedInputFile,
    Document,
    FSInputFile,
    Message,
    PhotoSize,
)
from PIL import Image
from pdf2docx import Converter as PDF2DOCXConverter

# ─────────────────────────────────────────────
#  SOZLAMALAR
# ─────────────────────────────────────────────
BOT_TOKEN: str = "8724055816:AAHfaaSLyjCFWLsK9a1L_j6_BwRGFg5rxTM"          # <-- o'zgartiring

# Windows'dagi LibreOffice yo'li (kerak bo'lsa to'g'rilang)
SOFFICE_PATH: str = r"C:\Program Files\LibreOffice\program\soffice.exe"

DOWNLOADS_DIR = Path("downloads")
OUTPUTS_DIR   = Path("outputs")
DOWNLOADS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)

# Albom/ketma-ket rasm to'plash uchun bufer
# { user_id: {"images": [...paths], "task": asyncio.Task} }
IMAGE_BUFFER: dict[int, dict] = {}
COLLECT_TIMEOUT = 5          # soniya

OFFICE_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger(__name__)

bot = Bot(
    token=BOT_TOKEN,
    default=DefaultBotProperties(parse_mode=ParseMode.HTML),
)
dp = Dispatcher()


# ─────────────────────────────────────────────
#  YORDAMCHI FUNKSIYALAR
# ─────────────────────────────────────────────
def unique_path(directory: Path, suffix: str) -> Path:
    """Unikal fayl yo'lini qaytaradi."""
    return directory / f"{uuid.uuid4().hex}{suffix}"


async def download_telegram_file(
    file_id: str, dest_path: Path
) -> None:
    """Telegram faylini diskka yuklaydi."""
    file = await bot.get_file(file_id)
    await bot.download_file(file.file_path, destination=str(dest_path))  # type: ignore[arg-type]


def safe_remove(*paths: Path | str) -> None:
    """Mavjud bo'lsa faylni o'chiradi; xato bo'lsa log yozadi."""
    for p in paths:
        try:
            os.remove(p)
        except FileNotFoundError:
            pass
        except Exception as exc:
            logger.warning("O'chirib bo'lmadi: %s — %s", p, exc)


# ─────────────────────────────────────────────
#  KONVERTATSIYA FUNKSIYALARI
# ─────────────────────────────────────────────
async def images_to_pdf(image_paths: list[Path], output_path: Path) -> None:
    """Bir nechta rasm faylini bitta PDF ga birlashtiradi (Pillow)."""
    loop = asyncio.get_event_loop()

    def _convert() -> None:
        pil_images: list[Image.Image] = []
        for p in image_paths:
            img = Image.open(p).convert("RGB")
            pil_images.append(img)

        if not pil_images:
            raise ValueError("Rasmlar ro'yxati bo'sh.")

        first, rest = pil_images[0], pil_images[1:]
        first.save(
            str(output_path),
            save_all=True,
            append_images=rest,
        )
        for img in pil_images:
            img.close()

    await loop.run_in_executor(None, _convert)


async def office_to_pdf(input_path: Path, output_dir: Path) -> Path:
    """LibreOffice yordamida Office faylini PDF ga o'tkazadi (Windows)."""
    cmd = [
        SOFFICE_PATH,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(input_path),
    ]

    proc = await asyncio.create_subprocess_exec(
        *cmd,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    stdout, stderr = await proc.communicate()

    if proc.returncode != 0:
        raise RuntimeError(
            f"LibreOffice xatosi (code {proc.returncode}):\n"
            f"{stderr.decode(errors='replace')}"
        )

    # LibreOffice chiqaradigan fayl nomi: original_name.pdf
    pdf_name = input_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name

    if not pdf_path.exists():
        raise FileNotFoundError(
            f"Kutilgan PDF topilmadi: {pdf_path}"
        )
    return pdf_path


async def pdf_to_docx(input_path: Path, output_path: Path) -> None:
    """pdf2docx yordamida PDF ni DOCX ga o'tkazadi."""
    loop = asyncio.get_event_loop()

    def _convert() -> None:
        cv = PDF2DOCXConverter(str(input_path))
        cv.convert(str(output_path), start=0, end=None)
        cv.close()

    await loop.run_in_executor(None, _convert)


# ─────────────────────────────────────────────
#  RASMLARNI TO'PLASH VA PDF YUBORISH
# ─────────────────────────────────────────────
async def flush_images(user_id: int, chat_id: int) -> None:
    """
    5 soniya o'tgach ishga tushadi:
    to'plangan rasmlarni bitta PDF ga birlashtiradi va yuboradi.
    """
    await asyncio.sleep(COLLECT_TIMEOUT)

    entry = IMAGE_BUFFER.pop(user_id, None)
    if not entry:
        return

    image_paths: list[Path] = entry["images"]
    status_msg: Message | None = entry.get("status_msg")
    output_path: Path = unique_path(OUTPUTS_DIR, ".pdf")

    try:
        await images_to_pdf(image_paths, output_path)

        pdf_file = FSInputFile(str(output_path), filename="rasmlar.pdf")
        await bot.send_document(
            chat_id,
            pdf_file,
            caption=f"✅ {len(image_paths)} ta rasm PDF ga birlashtrildi.",
        )

    except Exception as exc:
        logger.exception("images_to_pdf xatosi: %s", exc)
        await bot.send_message(chat_id, f"❌ Xatolik: <code>{exc}</code>")

    finally:
        if status_msg:
            try:
                await status_msg.delete()
            except Exception:
                pass
        safe_remove(*image_paths, output_path)


async def schedule_flush(message: Message, img_path: Path) -> None:
    """Rasmni buferga qo'shadi va yoki taymerni boshlaydi/qayta boshlaydi."""
    user_id = message.from_user.id  # type: ignore[union-attr]
    chat_id = message.chat.id

    if user_id not in IMAGE_BUFFER:
        # Birinchi rasm: statusni chiqar va taymer qo'y
        status_msg = await message.answer("🖼 Rasmlar to'planmoqda, iltimos kuting…")
        task = asyncio.create_task(flush_images(user_id, chat_id))
        IMAGE_BUFFER[user_id] = {
            "images": [img_path],
            "task": task,
            "status_msg": status_msg,
        }
    else:
        # Keyingi rasmlar: eski taymerni bekor qil, yangi qo'y
        IMAGE_BUFFER[user_id]["task"].cancel()
        IMAGE_BUFFER[user_id]["images"].append(img_path)
        new_task = asyncio.create_task(flush_images(user_id, chat_id))
        IMAGE_BUFFER[user_id]["task"] = new_task


# ─────────────────────────────────────────────
#  HANDLERLАР
# ─────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: Message) -> None:
    text = (
        "👋 <b>Universal Converter Bot</b> ga xush kelibsiz!\n\n"
        "📌 <b>Qo'llab-quvvatlanadigan operatsiyalar:</b>\n"
        "• <b>Rasmlar → PDF</b>: bir yoki bir nechta rasm yuboring\n"
        "• <b>DOCX/XLSX/PPTX → PDF</b>: Office hujjatini yuboring\n"
        "• <b>PDF → DOCX</b>: PDF faylni yuboring\n\n"
        "⚡ Fayl yuboring va bot avtomatik aniqlaydi."
    )
    await message.answer(text)


# ── RASM HANDLERI ─────────────────────────────
@dp.message(F.photo)
async def handle_photo(message: Message) -> None:
    """Telegram photo (siqilgan)."""
    photo: PhotoSize = message.photo[-1]  # type: ignore[index]
    dest: Path = unique_path(DOWNLOADS_DIR, ".jpg")

    try:
        await download_telegram_file(photo.file_id, dest)
        await schedule_flush(message, dest)
    except Exception as exc:
        logger.exception("Photo yuklashda xato: %s", exc)
        safe_remove(dest)
        await message.answer(f"❌ Rasmni yuklab bo'lmadi: <code>{exc}</code>")


# ── HUJJAT HANDLERI ───────────────────────────
@dp.message(F.document)
async def handle_document(message: Message) -> None:
    doc: Document = message.document  # type: ignore[assignment]
    file_name: str = doc.file_name or "fayl"
    suffix: str = Path(file_name).suffix.lower()

    # ── Rasm hujjat sifatida kelgan (siqilmagan) ──────────────────
    image_suffixes = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".webp"}
    if suffix in image_suffixes:
        dest: Path = unique_path(DOWNLOADS_DIR, suffix)
        try:
            await download_telegram_file(doc.file_id, dest)
            await schedule_flush(message, dest)
        except Exception as exc:
            logger.exception("Rasm-hujjat yuklashda xato: %s", exc)
            safe_remove(dest)
            await message.answer(f"❌ Xatolik: <code>{exc}</code>")
        return

    # ── Office → PDF ──────────────────────────────────────────────
    if suffix in OFFICE_EXTENSIONS:
        input_path: Path = unique_path(DOWNLOADS_DIR, suffix)
        output_path: Path = Path("")          # keyin to'ldiriladi
        status_msg: Message | None = None

        try:
            status_msg = await message.answer("⚙️ Fayl ishlanmoqda…")
            await download_telegram_file(doc.file_id, input_path)

            output_path = await office_to_pdf(input_path, OUTPUTS_DIR)

            pdf_file = FSInputFile(
                str(output_path),
                filename=Path(file_name).stem + ".pdf",
            )
            await message.answer_document(
                pdf_file,
                caption="✅ PDF ga muvaffaqiyatli o'tkazildi.",
            )

        except Exception as exc:
            logger.exception("Office→PDF xatosi: %s", exc)
            await message.answer(f"❌ Konvertatsiya xatosi:\n<code>{exc}</code>")

        finally:
            if status_msg:
                try:
                    await status_msg.delete()
                except Exception:
                    pass
            safe_remove(input_path)
            if output_path != Path(""):
                safe_remove(output_path)
        return

    # ── PDF → DOCX ────────────────────────────────────────────────
    if suffix == ".pdf":
        input_path_pdf: Path = unique_path(DOWNLOADS_DIR, ".pdf")
        output_path_docx: Path = unique_path(OUTPUTS_DIR, ".docx")
        status_msg_pdf: Message | None = None

        try:
            status_msg_pdf = await message.answer("⚙️ Fayl ishlanmoqda…")
            await download_telegram_file(doc.file_id, input_path_pdf)
            await pdf_to_docx(input_path_pdf, output_path_docx)

            docx_file = FSInputFile(
                str(output_path_docx),
                filename=Path(file_name).stem + ".docx",
            )
            await message.answer_document(
                docx_file,
                caption="✅ DOCX ga muvaffaqiyatli o'tkazildi.",
            )

        except Exception as exc:
            logger.exception("PDF→DOCX xatosi: %s", exc)
            await message.answer(f"❌ Konvertatsiya xatosi:\n<code>{exc}</code>")

        finally:
            if status_msg_pdf:
                try:
                    await status_msg_pdf.delete()
                except Exception:
                    pass
            safe_remove(input_path_pdf, output_path_docx)
        return

    # ── Qo'llab-quvvatlanmagan format ────────────────────────────
    await message.answer(
        f"⚠️ <b>{suffix}</b> formati qo'llab-quvvatlanmaydi.\n"
        "Quyidagi formatlarni yuboring:\n"
        "• Rasmlar: JPG, PNG, BMP, TIFF, WEBP\n"
        "• Office: DOCX, XLSX, PPTX\n"
        "• PDF (DOCX ga o'tkazish uchun)"
    )


# ─────────────────────────────────────────────
#  BOTNI ISHGA TUSHIRISH
# ─────────────────────────────────────────────
async def main() -> None:
    me = await bot.get_me()
    print(f"Bot @{me.username} nomi bilan ishga tushdi ✅")
    logger.info("Bot @%s ishga tushdi.", me.username)

    await dp.start_polling(bot, skip_updates=True)


if __name__ == "__main__":
    asyncio.run(main())