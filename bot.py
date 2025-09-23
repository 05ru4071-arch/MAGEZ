# bot.py
# -*- coding: utf-8 -*-

import os
import io
import re
import asyncio
from typing import List, Dict, Optional

from aiogram import Bot, Dispatcher, F, Router
from aiogram.enums import ParseMode, ContentType
from aiogram.filters import CommandStart
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder
from aiogram.client.default import DefaultBotProperties

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# ---------- ТОКЕН ----------
BOT_TOKEN = "7741928533:AAFDsO77wqRsWLTR7cu39UQDvqMc5MsyEw4"

# ---------- Папки ----------
BASE_ARCHIVE_DIR = "archive"
BASE_TEMP_DIR = "temp_files"
os.makedirs(BASE_ARCHIVE_DIR, exist_ok=True)
os.makedirs(BASE_TEMP_DIR, exist_ok=True)

# ---------- Excel ----------
COLUMNS = ["Фото", "Ссылка", "Цвет", "Размер", "Количество", "Комментарий"]
HEADER_RED = "C00000"
HEADER_GREY = "808080"
WHITE = "FFFFFF"
DEFAULT_COL_WIDTHS = [18, 45, 18, 14, 12, 40]
ROW_HEIGHT_EXCEL = 80

# ---------- Хранилище ----------
user_items: Dict[int, List[Dict]] = {}
edit_context: Dict[int, Dict] = {}

# ---------- FSM ----------
class AddItemStates(StatesGroup):
    waiting_photo = State()
    waiting_link = State()
    waiting_color = State()
    waiting_size = State()
    waiting_qty = State()
    waiting_comment = State()

class SaveExcelStates(StatesGroup):
    waiting_cargo_code = State()

class EditStates(StatesGroup):
    waiting_field_value = State()

# ---------- Клавиатуры ----------
def kb_main_menu():
    kb = InlineKeyboardBuilder()
    kb.button(text="📂 Создать Excel", callback_data="create_excel")
    kb.button(text="📜 Посмотреть архив", callback_data="view_archive")
    kb.adjust(2)
    return kb.as_markup()

def kb_items_menu():
    kb = InlineKeyboardBuilder()
    kb.button(text="➕ Добавить товар", callback_data="add_item")
    kb.button(text="✏ Редактировать товар", callback_data="edit_item")
    kb.button(text="🗑 Удалить товар", callback_data="delete_item")
    kb.button(text="✅ Завершить Excel", callback_data="finish_excel")
    kb.adjust(2, 2)
    return kb.as_markup()

def kb_finish_menu():
    kb = InlineKeyboardBuilder()
    kb.button(text="💾 Сохранить Excel", callback_data="save_excel")
    kb.button(text="🔙 Назад", callback_data="back_to_edit_menu")
    kb.adjust(2)
    return kb.as_markup()

def kb_archive_files(user_id: int):
    kb = InlineKeyboardBuilder()
    user_dir = os.path.join(BASE_ARCHIVE_DIR, str(user_id))
    files = []
    if os.path.isdir(user_dir):
        for name in sorted(os.listdir(user_dir)):
            if name.lower().endswith(".xlsx"):
                files.append(name)
    if not files:
        kb.button(text="Архив пуст", callback_data="noop")
        kb.adjust(1)
        return kb.as_markup()
    for name in files:
        kb.button(text=name, callback_data=f"send_archive:{name}")
    kb.adjust(1)
    return kb.as_markup()

def kb_items_list(user_id: int, action: str):
    kb = InlineKeyboardBuilder()
    for i, it in enumerate(user_items.get(user_id, [])):
        kb.button(text=f"{i+1}. {short_item_title(it)}", callback_data=f"{action}:{i}")
    kb.adjust(1)
    return kb.as_markup()

def kb_edit_fields():
    kb = InlineKeyboardBuilder()
    kb.button(text="Фото", callback_data="field:photo")
    kb.button(text="Ссылка", callback_data="field:link")
    kb.button(text="Цвет", callback_data="field:color")
    kb.button(text="Размер", callback_data="field:size")
    kb.button(text="Количество", callback_data="field:qty")
    kb.button(text="Комментарий", callback_data="field:comment")
    kb.adjust(2, 2, 2)
    return kb.as_markup()

# ---------- Утилиты ----------
def short_item_title(it: Dict) -> str:
    base = it.get("file_name") or ("Фото" if it.get("photo_path") else "Без фото")
    return f"{base} | {it.get('color','-')} | {it.get('size','-')} | {it.get('qty','-')}"

def ensure_user_list(user_id: int):
    if user_id not in user_items:
        user_items[user_id] = []

def normalize_qty(text: str) -> Optional[int]:
    text = text.strip().replace(",", ".")
    if re.fullmatch(r"\d+", text):
        return int(text)
    return None

def is_image_file(path: str) -> bool:
    try:
        with PILImage.open(path) as im:
            im.verify()
        return True
    except:
        return False

def convert_to_png(path: str) -> Optional[str]:
    try:
        with PILImage.open(path) as im:
            im = im.convert("RGB")
            new_path = path.rsplit(".", 1)[0] + ".png"
            im.save(new_path, "PNG")
        return new_path
    except Exception:
        return None

def photo_cell_image(path: str, max_width_px=100, max_height_px=100):
    try:
        with PILImage.open(path) as im:
            im = im.convert("RGB")
            w, h = im.size
            scale = min(max_width_px / w, max_height_px / h, 1.0)
            im = im.resize((int(w*scale), int(h*scale)))
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
        return XLImage(buf)
    except:
        return None

def ensure_user_archive(user_id: int) -> str:
    path = os.path.join(BASE_ARCHIVE_DIR, str(user_id))
    os.makedirs(path, exist_ok=True)
    return path

def nice_items_text(items: List[Dict]) -> str:
    if not items:
        return "Список пуст."
    return "\n".join(
        f"{i+1}) {it.get('file_name') or 'Фото'} | {it.get('link','-')} | {it.get('color','-')} | "
        f"{it.get('size','-')} | {it.get('qty','-')} | {it.get('comment','-')}"
        for i, it in enumerate(items)
    )

def extract_url(text: str) -> str:
    match = re.search(r'(https?://\S+)', text)
    return match.group(1) if match else text.strip()

# ---------- Excel генерация ----------
def build_excel(user_id: int, cargo_code: str, items: List[Dict]) -> str:
    user_dir = ensure_user_archive(user_id)
    filename = f"{cargo_code}.xlsx"
    out_path = os.path.join(user_dir, filename)

    wb = Workbook()
    ws = wb.active

    ws.merge_cells("A1:F1")
    ws.merge_cells("A2:F2")
    head_font = Font(bold=True, color=WHITE, size=14)
    head_fill = PatternFill("solid", fgColor=HEADER_RED)
    center = Alignment(horizontal="center", vertical="center")

    ws["A1"].value = "MAGEZ"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = head_font, head_fill, center
    ws["A2"].value = "Торгово-Логистическая компания"
    ws["A2"].font, ws["A2"].fill, ws["A2"].alignment = head_font, head_fill, center

    header_font = Font(bold=True, color=WHITE)
    header_fill = PatternFill("solid", fgColor=HEADER_GREY)
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))
    row = 4
    for i, title in enumerate(COLUMNS, 1):
        c = ws.cell(row=row, column=i, value=title)
        c.font, c.fill, c.alignment, c.border = header_font, header_fill, center, border
        ws.column_dimensions[get_column_letter(i)].width = DEFAULT_COL_WIDTHS[i-1]

    row = 5
    for it in items:
        if it.get("photo_path") and it.get("photo_is_image"):
            img = photo_cell_image(it["photo_path"])
            if img:
                ws.row_dimensions[row].height = ROW_HEIGHT_EXCEL
                ws.add_image(img, f"A{row}")
        else:
            ws.cell(row=row, column=1, value=it.get("file_name") or "—")
        ws.cell(row=row, column=2, value=extract_url(it.get("link") or "—"))
        ws.cell(row=row, column=3, value=it.get("color") or "—")
        ws.cell(row=row, column=4, value=it.get("size") or "—")
        ws.cell(row=row, column=5, value=it.get("qty") or "—")
        ws.cell(row=row, column=6, value=it.get("comment") or "—")
        for c in range(1, 7):
            cell = ws.cell(row=row, column=c)
            cell.alignment, cell.border = center, border
        row += 1

    # --- Блок с контактами ---
    row += 2
    note = (
        "Для выкупа напишите одному из менеджеров:\n"
        "Мурад: WhatsApp +7 988 691-55-35\n"
        "Омаргаджи: WhatsApp +7 989 459-20-39"
    )
    ws.merge_cells(f"A{row}:F{row}")
    c = ws.cell(row=row, column=1, value=note)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    c.font = Font(bold=True, size=12, color="000000")

    wb.save(out_path)
    return out_path

# ---------- Роутер ----------
router = Router()

@router.message(CommandStart())
async def on_start(message: Message):
    ensure_user_list(message.from_user.id)
    await message.answer("Привет! Выберите действие:", reply_markup=kb_main_menu())

# ---------- Главное меню ----------
@router.callback_query(F.data == "create_excel")
async def cb_create_excel(call: CallbackQuery, state: FSMContext):
    user_items[call.from_user.id] = []
    await call.message.edit_text("Создаём Excel. Пока список пуст.", reply_markup=kb_items_menu())
    await call.answer()

@router.callback_query(F.data == "view_archive")
async def cb_view_archive(call: CallbackQuery):
    await call.message.edit_text("Ваши файлы архива:", reply_markup=kb_archive_files(call.from_user.id))
    await call.answer()

@router.callback_query(F.data.startswith("send_archive:"))
async def cb_send_archive(call: CallbackQuery):
    user_id = call.from_user.id
    filename = call.data.split(":", 1)[1]
    path = os.path.join(BASE_ARCHIVE_DIR, str(user_id), filename)
    if os.path.isfile(path):
        await call.message.answer_document(FSInputFile(path))
    await call.answer()

# ---------- FSM добавления ----------
@router.callback_query(F.data == "add_item")
async def cb_add_item(call: CallbackQuery, state: FSMContext):
    await state.set_state(AddItemStates.waiting_photo)
    await call.message.edit_text("Шаг 1/6. Отправьте фото или файл товара:")
    await call.answer()

@router.message(AddItemStates.waiting_photo, F.content_type.in_({ContentType.PHOTO, ContentType.DOCUMENT}))
async def step_photo(message: Message, state: FSMContext):
    file_name, saved_path, is_img = None, None, False
    if message.photo:
        tg_file = message.photo[-1]
        file_name = f"photo_{tg_file.file_unique_id}.jpg"
        saved_path = os.path.join(BASE_TEMP_DIR, file_name)
        await message.bot.download(tg_file, destination=saved_path)
        is_img = True
    elif message.document:
        doc = message.document
        file_name = doc.file_name
        saved_path = os.path.join(BASE_TEMP_DIR, file_name)
        await message.bot.download(doc, destination=saved_path)
        if is_image_file(saved_path):
            is_img = True
            new_path = convert_to_png(saved_path)
            if new_path:
                saved_path = new_path
    await state.update_data(photo_path=saved_path, file_name=file_name, photo_is_image=is_img)
    await state.set_state(AddItemStates.waiting_link)
    await message.answer("Шаг 2/6. Введите ссылку:")

@router.message(AddItemStates.waiting_link)
async def step_link(message: Message, state: FSMContext):
    await state.update_data(link=extract_url(message.text or ""))
    await state.set_state(AddItemStates.waiting_color)
    await message.answer("Шаг 3/6. Введите цвет:")

@router.message(AddItemStates.waiting_color)
async def step_color(message: Message, state: FSMContext):
    await state.update_data(color=message.text.strip())
    await state.set_state(AddItemStates.waiting_size)
    await message.answer("Шаг 4/6. Введите размер:")

@router.message(AddItemStates.waiting_size)
async def step_size(message: Message, state: FSMContext):
    await state.update_data(size=message.text.strip())
    await state.set_state(AddItemStates.waiting_qty)
    await message.answer("Шаг 5/6. Введите количество:")

@router.message(AddItemStates.waiting_qty)
async def step_qty(message: Message, state: FSMContext):
    qty = normalize_qty(message.text or "")
    if qty is None:
        await message.answer("Введите целое число:")
        return
    await state.update_data(qty=qty)
    await state.set_state(AddItemStates.waiting_comment)
    await message.answer("Шаг 6/6. Введите комментарий:")

@router.message(AddItemStates.waiting_comment)
async def step_comment(message: Message, state: FSMContext):
    user_id = message.from_user.id
    data = await state.get_data()
    item = {
        "photo_path": data.get("photo_path"),
        "photo_is_image": data.get("photo_is_image"),
        "file_name": data.get("file_name"),
        "link": data.get("link"),
        "color": data.get("color"),
        "size": data.get("size"),
        "qty": data.get("qty"),
        "comment": (message.text.strip() if message.text.strip() != "-" else "")
    }
    user_items[user_id].append(item)
    await state.clear()
    await message.answer("Товар добавлен ✅\n\n" + nice_items_text(user_items[user_id]), reply_markup=kb_items_menu())

# ---------- Редактирование ----------
@router.callback_query(F.data == "edit_item")
async def cb_edit_item(call: CallbackQuery):
    if not user_items.get(call.from_user.id):
        await call.message.edit_text("Список пуст.", reply_markup=kb_items_menu())
        return
    await call.message.edit_text("Выберите товар:", reply_markup=kb_items_list(call.from_user.id, "edit"))
    await call.answer()

@router.callback_query(F.data.startswith("edit:"))
async def cb_edit_select(call: CallbackQuery, state: FSMContext):
    idx = int(call.data.split(":")[1])
    edit_context[call.from_user.id] = {"index": idx}
    await call.message.edit_text("Выберите поле:", reply_markup=kb_edit_fields())
    await call.answer()

@router.callback_query(F.data.startswith("field:"))
async def cb_field(call: CallbackQuery, state: FSMContext):
    field = call.data.split(":")[1]
    edit_context[call.from_user.id]["field"] = field
    await state.set_state(EditStates.waiting_field_value)
    if field == "photo":
        await call.message.edit_text("Отправьте новое фото или файл товара:")
    else:
        await call.message.edit_text(f"Введите новое значение для {field}:")
    await call.answer()

@router.message(EditStates.waiting_field_value, F.content_type.in_({ContentType.PHOTO, ContentType.DOCUMENT}))
async def on_edit_photo(message: Message, state: FSMContext):
    ctx = edit_context.get(message.from_user.id)
    if not ctx or ctx.get("field") != "photo":
        return
    idx = ctx["index"]
    file_name, saved_path, is_img = None, None, False
    if message.photo:
        tg_file = message.photo[-1]
        file_name = f"photo_{tg_file.file_unique_id}.jpg"
        saved_path = os.path.join(BASE_TEMP_DIR, file_name)
        await message.bot.download(tg_file, destination=saved_path)
        is_img = True
    elif message.document:
        doc = message.document
        file_name = doc.file_name
        saved_path = os.path.join(BASE_TEMP_DIR, file_name)
        await message.bot.download(doc, destination=saved_path)
        if is_image_file(saved_path):
            is_img = True
            new_path = convert_to_png(saved_path)
            if new_path:
                saved_path = new_path
    user_items[message.from_user.id][idx].update({
        "photo_path": saved_path,
        "file_name": file_name,
        "photo_is_image": is_img
    })
    await state.clear()
    await message.answer("Фото изменено ✅\n\n" + nice_items_text(user_items[message.from_user.id]), reply_markup=kb_items_menu())

@router.message(EditStates.waiting_field_value, F.text)
async def on_edit_text(message: Message, state: FSMContext):
    ctx = edit_context.get(message.from_user.id)
    if not ctx:
        await state.clear()
        return
    idx, field = ctx["index"], ctx["field"]
    item = user_items[message.from_user.id][idx]
    if field == "qty":
        qty = normalize_qty(message.text or "")
        if qty:
            item["qty"] = qty
    elif field == "link":
        item["link"] = extract_url(message.text or "")
    else:
        item[field] = message.text.strip()
    await state.clear()
    await message.answer("Изменено ✅\n\n" + nice_items_text(user_items[message.from_user.id]), reply_markup=kb_items_menu())

# ---------- Удаление ----------
@router.callback_query(F.data == "delete_item")
async def cb_delete_item(call: CallbackQuery):
    if not user_items.get(call.from_user.id):
        await call.message.edit_text("Список пуст.", reply_markup=kb_items_menu())
        return
    await call.message.edit_text("Выберите товар для удаления:", reply_markup=kb_items_list(call.from_user.id, "del"))
    await call.answer()

@router.callback_query(F.data.startswith("del:"))
async def cb_del(call: CallbackQuery):
    idx = int(call.data.split(":")[1])
    del user_items[call.from_user.id][idx]
    await call.message.edit_text("Удалено ✅\n\n" + nice_items_text(user_items[call.from_user.id]), reply_markup=kb_items_menu())
    await call.answer()

# ---------- Завершение Excel ----------
@router.callback_query(F.data == "finish_excel")
async def cb_finish_excel(call: CallbackQuery):
    await call.message.edit_text("Список товаров:\n\n" + nice_items_text(user_items[call.from_user.id]), reply_markup=kb_finish_menu())
    await call.answer()

@router.callback_query(F.data == "save_excel")
async def cb_save_excel(call: CallbackQuery, state: FSMContext):
    await state.set_state(SaveExcelStates.waiting_cargo_code)
    await call.message.edit_text("Введите код груза (имя файла):")
    await call.answer()

@router.message(SaveExcelStates.waiting_cargo_code)
async def on_cargo_code(message: Message, state: FSMContext):
    path = build_excel(message.from_user.id, message.text.strip(), user_items.get(message.from_user.id, []))
    await message.answer_document(FSInputFile(path), caption="Файл сохранён ✅")
    user_items[message.from_user.id] = []
    await state.clear()
    await message.answer("Возврат в меню:", reply_markup=kb_main_menu())

# ---------- no-op ----------
@router.callback_query(F.data == "noop")
async def cb_noop(call: CallbackQuery):
    await call.answer()

# ---------- Запуск ----------
async def main():
    bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    print("Bot started.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        print("Bot stopped.")