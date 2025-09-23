import asyncio
import logging
import os
import uuid
from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from PIL import Image
import io

# ---------------- CONFIG ----------------
TOKEN = "7741928533:AAFDsO77wqRsWLTR7cu39UQDvqMc5MsyEw4"   # твой токен
ADMIN_ID = 1891138771                                      # твой Telegram ID
DATA_DIR = "archive"

# ---------------- LOGGER ----------------
logging.basicConfig(level=logging.INFO)

# ---------------- BOT INIT ----------------
bot = Bot(token=TOKEN)
dp = Dispatcher()

# ---------------- STORAGE ----------------
invites = set()  # одноразовые приглашения
allowed_users = set([ADMIN_ID])  # пользователи с доступом
user_products = {}  # временное хранилище товаров: user_id -> [dict]

# ---------------- FSM ----------------
class ProductForm(StatesGroup):
    photo = State()
    link = State()
    color = State()
    size = State()
    quantity = State()
    comment = State()
    editing_field = State()
    waiting_code = State()

# ---------------- KEYBOARDS ----------------
def main_menu():
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📂 Создать Excel", callback_data="create_excel")],
        [InlineKeyboardButton(text="📜 Посмотреть архив", callback_data="view_archive")]
    ])
    return kb

def after_product_menu():
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Добавить товар", callback_data="add_product")],
        [InlineKeyboardButton(text="✏ Редактировать товар", callback_data="edit_product")],
        [InlineKeyboardButton(text="🗑 Удалить товар", callback_data="delete_product")],
        [InlineKeyboardButton(text="✅ Завершить Excel", callback_data="finish_excel")]
    ])
    return kb

# ---------------- INVITES ----------------
@dp.message(CommandStart())
async def start_cmd(msg: types.Message):
    args = msg.text.split()
    if msg.from_user.id not in allowed_users:
        if len(args) == 2 and args[1] in invites:
            invites.remove(args[1])
            allowed_users.add(msg.from_user.id)
        else:
            await msg.answer("🚫 У вас нет доступа. Попросите приглашение у администратора.")
            return
    await msg.answer("Добро пожаловать в MAGEZ Bot!", reply_markup=main_menu())

@dp.message(F.text == "/invite")
async def create_invite(msg: types.Message):
    if msg.from_user.id != ADMIN_ID:
        return
    code = str(uuid.uuid4())
    invites.add(code)
    link = f"https://t.me/{(await bot.me()).username}?start={code}"
    await msg.answer(f"🎟 Приглашение создано:\n{link}")

# ---------------- FSM ADD PRODUCT ----------------
@dp.callback_query(F.data == "create_excel")
async def start_excel(cb: types.CallbackQuery, state: FSMContext):
    user_products[cb.from_user.id] = []
    await cb.message.edit_text("Добавляем первый товар...\nПришлите фото.", reply_markup=None)
    await state.set_state(ProductForm.photo)

@dp.callback_query(F.data == "add_product")
async def add_product(cb: types.CallbackQuery, state: FSMContext):
    await cb.message.answer("Пришлите фото товара:")
    await state.set_state(ProductForm.photo)

@dp.message(ProductForm.photo)
async def product_photo(msg: types.Message, state: FSMContext):
    if msg.photo:
        file = await bot.get_file(msg.photo[-1].file_id)
        file_path = f"{DATA_DIR}/{msg.from_user.id}/temp_{uuid.uuid4()}.jpg"
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        await bot.download_file(file.file_path, destination=file_path)
        await state.update_data(photo=file_path)
    elif msg.document:
        file = await bot.get_file(msg.document.file_id)
        file_path = f"{DATA_DIR}/{msg.from_user.id}/temp_{msg.document.file_name}"
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        await bot.download_file(file.file_path, destination=file_path)
        await state.update_data(photo=file_path)
    else:
        await msg.answer("Пришлите фото или файл.")
        return
    await msg.answer("Введите ссылку:")
    await state.set_state(ProductForm.link)

@dp.message(ProductForm.link)
async def product_link(msg: types.Message, state: FSMContext):
    await state.update_data(link=msg.text)
    await msg.answer("Введите цвет:")
    await state.set_state(ProductForm.color)

@dp.message(ProductForm.color)
async def product_color(msg: types.Message, state: FSMContext):
    await state.update_data(color=msg.text)
    await msg.answer("Введите размер:")
    await state.set_state(ProductForm.size)

@dp.message(ProductForm.size)
async def product_size(msg: types.Message, state: FSMContext):
    await state.update_data(size=msg.text)
    await msg.answer("Введите количество:")
    await state.set_state(ProductForm.quantity)

@dp.message(ProductForm.quantity)
async def product_quantity(msg: types.Message, state: FSMContext):
    await state.update_data(quantity=msg.text)
    await msg.answer("Введите комментарий:")
    await state.set_state(ProductForm.comment)

@dp.message(ProductForm.comment)
async def product_comment(msg: types.Message, state: FSMContext):
    data = await state.get_data()
    data["comment"] = msg.text
    user_products[msg.from_user.id].append(data)
    await state.clear()
    await msg.answer("Товар добавлен ✅", reply_markup=after_product_menu())

# ---------------- EDIT PRODUCT ----------------
@dp.callback_query(F.data == "edit_product")
async def choose_product_edit(cb: types.CallbackQuery, state: FSMContext):
    products = user_products.get(cb.from_user.id, [])
    if not products:
        await cb.message.answer("❌ Нет товаров для редактирования.")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"{i+1}. {p['link']}", callback_data=f"edit_{i}")]
        for i, p in enumerate(products)
    ])
    await cb.message.answer("Выберите товар для редактирования:", reply_markup=kb)

@dp.callback_query(F.data.startswith("edit_"))
async def edit_field(cb: types.CallbackQuery, state: FSMContext):
    idx = int(cb.data.split("_")[1])
    await state.update_data(edit_index=idx)
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Фото", callback_data="field_photo")],
        [InlineKeyboardButton(text="Ссылка", callback_data="field_link")],
        [InlineKeyboardButton(text="Цвет", callback_data="field_color")],
        [InlineKeyboardButton(text="Размер", callback_data="field_size")],
        [InlineKeyboardButton(text="Количество", callback_data="field_quantity")],
        [InlineKeyboardButton(text="Комментарий", callback_data="field_comment")]
    ])
    await cb.message.answer("Выберите поле для редактирования:", reply_markup=kb)

@dp.callback_query(F.data.startswith("field_"))
async def edit_selected_field(cb: types.CallbackQuery, state: FSMContext):
    field = cb.data.split("_")[1]
    await state.update_data(editing_field=field)
    await cb.message.answer(f"Введите новое значение для поля: {field}")
    await state.set_state(ProductForm.editing_field)

@dp.message(ProductForm.editing_field)
async def save_edit(msg: types.Message, state: FSMContext):
    data = await state.get_data()
    idx = data["edit_index"]
    field = data["editing_field"]
    if field == "photo":
        await msg.answer("Фото редактировать можно только через пересоздание товара.")
        return
    user_products[msg.from_user.id][idx][field] = msg.text
    await state.clear()
    await msg.answer("Изменения сохранены ✅", reply_markup=after_product_menu())

# ---------------- DELETE PRODUCT ----------------
@dp.callback_query(F.data == "delete_product")
async def choose_product_delete(cb: types.CallbackQuery):
    products = user_products.get(cb.from_user.id, [])
    if not products:
        await cb.message.answer("❌ Нет товаров для удаления.")
        return
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=f"{i+1}. {p['link']}", callback_data=f"del_{i}")]
        for i, p in enumerate(products)
    ])
    await cb.message.answer("Выберите товар для удаления:", reply_markup=kb)

@dp.callback_query(F.data.startswith("del_"))
async def delete_confirm(cb: types.CallbackQuery):
    idx = int(cb.data.split("_")[1])
    user_products[cb.from_user.id].pop(idx)
    await cb.message.answer("Товар удален 🗑", reply_markup=after_product_menu())

# ---------------- FINISH EXCEL ----------------
@dp.callback_query(F.data == "finish_excel")
async def finish_excel(cb: types.CallbackQuery):
    products = user_products.get(cb.from_user.id, [])
    if not products:
        await cb.message.answer("❌ Нет товаров для сохранения.")
        return
    text = "Список товаров:\n"
    for i, p in enumerate(products, 1):
        text += f"{i}. {p['link']} ({p['color']}, {p['size']}, {p['quantity']})\n"
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="💾 Сохранить Excel", callback_data="save_excel")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="add_product")]
    ])
    await cb.message.answer(text, reply_markup=kb)

@dp.callback_query(F.data == "save_excel")
async def ask_code(cb: types.CallbackQuery, state: FSMContext):
    await cb.message.answer("Введите код груза (имя файла):")
    await state.set_state(ProductForm.waiting_code)

@dp.message(ProductForm.waiting_code)
async def save_excel(msg: types.Message, state: FSMContext):
    code = msg.text.strip()
    products = user_products.get(msg.from_user.id, [])
    folder = f"{DATA_DIR}/{msg.from_user.id}/"
    os.makedirs(folder, exist_ok=True)
    filepath = f"{folder}{code}.xlsx"

    wb = Workbook()
    ws = wb.active

    # Header
    ws.merge_cells("A1:F1")
    ws["A1"] = "MAGEZ"
    ws["A1"].fill = PatternFill("solid", fgColor="FF0000")
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:F2")
    ws["A2"] = "Торгово-Логистическая компания"
    ws["A2"].fill = PatternFill("solid", fgColor="FF0000")
    ws["A2"].font = Font(bold=True, color="FFFFFF", size=12)
    ws["A2"].alignment = Alignment(horizontal="center")

    headers = ["Фото", "Ссылка", "Цвет", "Размер", "Количество", "Комментарий"]
    ws.append(headers)
    for col in range(1, 7):
        cell = ws.cell(row=3, column=col)
        cell.value = headers[col-1]
        cell.fill = PatternFill("solid", fgColor="808080")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data
    row = 4
    for product in products:
        ws.cell(row=row, column=2, value=product["link"])
        ws.cell(row=row, column=3, value=product["color"])
        ws.cell(row=row, column=4, value=product["size"])
        ws.cell(row=row, column=5, value=product["quantity"])
        ws.cell(row=row, column=6, value=product["comment"])

        try:
            img = Image.open(product["photo"])
            img.thumbnail((80, 80))
            bio = io.BytesIO()
            img.save(bio, format="PNG")
            bio.seek(0)
            xl_img = XLImage(bio)
            ws.add_image(xl_img, f"A{row}")
        except Exception:
            ws.cell(row=row, column=1, value=os.path.basename(product["photo"]))
        row += 1

    # Styling
    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 5
        for cell in col:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=Side(style="thin"),
                                 right=Side(style="thin"),
                                 top=Side(style="thin"),
                                 bottom=Side(style="thin"))

    wb.save(filepath)
    await msg.answer_document(InputFile(filepath))
    await msg.answer("Файл сохранен ✅", reply_markup=main_menu())
    await state.clear()

# ---------------- ARCHIVE ----------------
@dp.callback_query(F.data == "view_archive")
async def view_archive(cb: types.CallbackQuery):
    folder = f"{DATA_DIR}/{cb.from_user.id}/"
    if not os.path.exists(folder):
        await cb.message.answer("📂 Архив пуст.")
        return
    files = os.listdir(folder)
    if not files:
        await cb.message.answer("📂 Архив пуст.")
        return
    text = "Ваш архив:\n" + "\n".join(files)
    await cb.message.answer(text)

# ---------------- RUN ----------------
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
