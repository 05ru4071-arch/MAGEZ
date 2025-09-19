import logging
import os
from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage

# === НАСТРОЙКИ ===
API_TOKEN = "YOUR_BOT_TOKEN"  # вставь токен сюда
DATA_DIR = "data"

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# === ЛОГИ ===
logging.basicConfig(level=logging.INFO)

# === ИНИЦИАЛИЗАЦИЯ БОТА ===
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

# === ХРАНИЛИЩЕ В ПАМЯТИ ===
user_sessions = {}  # {user_id: [товары]}
file_counter = {}   # {user_id: номер_файла}
edit_mode = {}      # {user_id: индекс_товара}

# === КНОПКИ ===
def main_menu():
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton("📑 Создать Excel", callback_data="create_excel"))
    kb.add(InlineKeyboardButton("📂 Посмотреть архив", callback_data="view_archive"))
    return kb

def product_menu():
    kb = InlineKeyboardMarkup()
    kb.add(InlineKeyboardButton("➕ Добавить товар", callback_data="add_product"))
    kb.add(InlineKeyboardButton("👀 Список товаров", callback_data="list_products"))
    kb.add(InlineKeyboardButton("✏️ Редактировать", callback_data="choose_edit"))
    kb.add(InlineKeyboardButton("❌ Удалить", callback_data="choose_delete"))
    kb.add(InlineKeyboardButton("✅ Завершить Excel", callback_data="finish_excel"))
    return kb

# === СТАРТ ===
@dp.message_handler(commands=['start'])
async def start_cmd(message: types.Message):
    await message.answer("Добро пожаловать! Выберите действие:", reply_markup=main_menu())

# === ГЛАВНОЕ МЕНЮ ===
@dp.callback_query_handler(lambda c: c.data == "create_excel")
async def process_create_excel(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    user_sessions[user_id] = []
    await bot.send_message(user_id, "Начинаем создание Excel. Добавьте первый товар:", reply_markup=product_menu())

@dp.callback_query_handler(lambda c: c.data == "view_archive")
async def process_view_archive(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    user_dir = os.path.join(DATA_DIR, str(user_id))
    if not os.path.exists(user_dir) or not os.listdir(user_dir):
        await bot.send_message(user_id, "У вас пока нет архивов.", reply_markup=main_menu())
        return

    for f in os.listdir(user_dir):
        await bot.send_document(user_id, open(os.path.join(user_dir, f), "rb"))
    await bot.send_message(user_id, "Вот ваши файлы:", reply_markup=main_menu())

# === ДОБАВЛЕНИЕ ТОВАРА ===
@dp.callback_query_handler(lambda c: c.data == "add_product")
async def add_product(callback_query: types.CallbackQuery):
    await bot.send_message(callback_query.from_user.id, "Пришлите фото товара:")

@dp.message_handler(content_types=['photo'])
async def handle_photo(message: types.Message):
    user_id = message.from_user.id
    if user_id not in user_sessions:
        return

    file_info = await bot.get_file(message.photo[-1].file_id)
    photo_name = f"{message.photo[-1].file_id}.jpg"
    user_dir = os.path.join(DATA_DIR, str(user_id))
    os.makedirs(user_dir, exist_ok=True)
    photo_path = os.path.join(user_dir, photo_name)
    await bot.download_file(file_info.file_path, photo_path)

    user_sessions[user_id].append({
        "Фото": photo_path,
        "Ссылка": "",
        "Цвет": "",
        "Размер": "",
        "Количество": "",
        "Комментарий": ""
    })
    await message.answer("Фото получено ✅ Теперь введите ссылку:")

@dp.message_handler(lambda m: m.from_user.id in user_sessions and m.text and user_sessions[m.from_user.id][-1]["Ссылка"] == "")
async def handle_link(message: types.Message):
    user_sessions[message.from_user.id][-1]["Ссылка"] = message.text
    await message.answer("Введите цвет:")

@dp.message_handler(lambda m: m.from_user.id in user_sessions and m.text and user_sessions[m.from_user.id][-1]["Цвет"] == "")
async def handle_color(message: types.Message):
    user_sessions[message.from_user.id][-1]["Цвет"] = message.text
    await message.answer("Введите размер:")

@dp.message_handler(lambda m: m.from_user.id in user_sessions and m.text and user_sessions[m.from_user.id][-1]["Размер"] == "")
async def handle_size(message: types.Message):
    user_sessions[message.from_user.id][-1]["Размер"] = message.text
    await message.answer("Введите количество:")

@dp.message_handler(lambda m: m.from_user.id in user_sessions and m.text and user_sessions[m.from_user.id][-1]["Количество"] == "")
async def handle_qty(message: types.Message):
    user_sessions[message.from_user.id][-1]["Количество"] = message.text
    await message.answer("Введите комментарий:")

@dp.message_handler(lambda m: m.from_user.id in user_sessions and m.text and user_sessions[m.from_user.id][-1]["Комментарий"] == "")
async def handle_comment(message: types.Message):
    user_sessions[message.from_user.id][-1]["Комментарий"] = message.text
    await message.answer("Товар добавлен ✅", reply_markup=product_menu())

# === СПИСОК ТОВАРОВ ===
@dp.callback_query_handler(lambda c: c.data == "list_products")
async def list_products(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    products = user_sessions.get(user_id, [])
    if not products:
        await bot.send_message(user_id, "Список товаров пуст.", reply_markup=product_menu())
        return
    text = "Ваши товары:\n"
    for i, p in enumerate(products, 1):
        text += f"{i}. {p['Цвет']} | {p['Размер']} | {p['Количество']} шт.\n"
    await bot.send_message(user_id, text, reply_markup=product_menu())

# === ВЫБОР ТОВАРА ДЛЯ РЕДАКТИРОВАНИЯ ===
@dp.callback_query_handler(lambda c: c.data == "choose_edit")
async def choose_edit(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    products = user_sessions.get(user_id, [])
    if not products:
        await bot.send_message(user_id, "Список товаров пуст.", reply_markup=product_menu())
        return
    kb = InlineKeyboardMarkup()
    for i, _ in enumerate(products, 1):
        kb.add(InlineKeyboardButton(f"Товар {i}", callback_data=f"edit_{i-1}"))
    await bot.send_message(user_id, "Выберите товар для редактирования:", reply_markup=kb)

@dp.callback_query_handler(lambda c: c.data.startswith("edit_"))
async def edit_product(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    index = int(callback_query.data.split("_")[1])
    edit_mode[user_id] = index
    kb = InlineKeyboardMarkup()
    for field in ["Фото", "Ссылка", "Цвет", "Размер", "Количество", "Комментарий"]:
        kb.add(InlineKeyboardButton(field, callback_data=f"field_{field}"))
    await bot.send_message(user_id, "Что хотите изменить?", reply_markup=kb)

@dp.callback_query_handler(lambda c: c.data.startswith("field_"))
async def edit_field(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    field = callback_query.data.split("_")[1]
    if field == "Фото":
        await bot.send_message(user_id, "Пришлите новое фото:")
    else:
        await bot.send_message(user_id, f"Введите новое значение для поля {field}:")
    edit_mode[user_id] = (edit_mode[user_id], field)

@dp.message_handler(lambda m: m.from_user.id in edit_mode and isinstance(edit_mode[m.from_user.id], tuple))
async def save_edit(message: types.Message):
    user_id = message.from_user.id
    index, field = edit_mode[user_id]
    user_sessions[user_id][index][field] = message.text
    del edit_mode[user_id]
    await message.answer("Изменение сохранено ✅", reply_markup=product_menu())

@dp.message_handler(content_types=['photo'])
async def edit_photo(message: types.Message):
    user_id = message.from_user.id
    if user_id not in edit_mode or not isinstance(edit_mode[user_id], tuple):
        return
    index, field = edit_mode[user_id]
    if field != "Фото":
        return
    file_info = await bot.get_file(message.photo[-1].file_id)
    photo_name = f"{message.photo[-1].file_id}.jpg"
    user_dir = os.path.join(DATA_DIR, str(user_id))
    os.makedirs(user_dir, exist_ok=True)
    photo_path = os.path.join(user_dir, photo_name)
    await bot.download_file(file_info.file_path, photo_path)
    user_sessions[user_id][index]["Фото"] = photo_path
    del edit_mode[user_id]
    await message.answer("Фото обновлено ✅", reply_markup=product_menu())

# === УДАЛЕНИЕ ТОВАРА ===
@dp.callback_query_handler(lambda c: c.data == "choose_delete")
async def choose_delete(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    products = user_sessions.get(user_id, [])
    if not products:
        await bot.send_message(user_id, "Список товаров пуст.", reply_markup=product_menu())
        return
    kb = InlineKeyboardMarkup()
    for i, _ in enumerate(products, 1):
        kb.add(InlineKeyboardButton(f"Удалить товар {i}", callback_data=f"delete_{i-1}"))
    await bot.send_message(user_id, "Выберите товар для удаления:", reply_markup=kb)

@dp.callback_query_handler(lambda c: c.data.startswith("delete_"))
async def delete_product(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    index = int(callback_query.data.split("_")[1])
    products = user_sessions.get(user_id, [])
    if 0 <= index < len(products):
        products.pop(index)
        await bot.send_message(user_id, "Товар удалён ✅", reply_markup=product_menu())

# === ЗАВЕРШЕНИЕ ===
@dp.callback_query_handler(lambda c: c.data == "finish_excel")
async def finish_excel(callback_query: types.CallbackQuery):
    user_id = callback_query.from_user.id
    products = user_sessions.get(user_id, [])
    if not products:
        await bot.send_message(user_id, "Нельзя завершить — список товаров пуст.", reply_markup=product_menu())
        return

    wb = Workbook()
    ws = wb.active

    # === ШАПКА КОМПАНИИ ===
    ws.merge_cells("A1:F1")
    ws.merge_cells("A2:F2")
    red_fill = PatternFill("solid", fgColor="FF0000")
    white_font = Font(color="FFFFFF", bold=True, size=14)
    center_align = Alignment(horizontal="center", vertical="center")

    ws["A1"] = "MAGEZ"
    ws["A1"].fill = red_fill
    ws["A1"].font = white_font
    ws["A1"].alignment = center_align

    ws["A2"] = "Торгово-Логистическая компания"
    ws["A2"].fill = red_fill
    ws["A2"].font = white_font
    ws["A2"].alignment = center_align

    # === ЗАГОЛОВКИ ===
    headers = ["Фото", "Ссылка", "Цвет", "Размер", "Количество", "Комментарий"]
    ws.append(headers)
    for col in ws[3]:
        col.font = Font(bold=True)
        col.alignment = center_align
        col.fill = PatternFill("solid", fgColor="DDDDDD")

    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                         top=Side(style="thin"), bottom=Side(style="thin"))

    # === ТОВАРЫ ===
    row = 4
    for p in products:
        ws.append(["", p["Ссылка"], p["Цвет"], p["Размер"], p["Количество"], p["Комментарий"]])
        for col in ws[row]:
            col.border = thin_border
            col.alignment = Alignment(horizontal="center", vertical="center")
        if os.path.exists(p["Фото"]):
            img = XLImage(p["Фото"])
            img.width, img.height = 80, 80
            ws.add_image(img, f"A{row}")
        row += 1

    user_dir = os.path.join(DATA_DIR, str(user_id))
    os.makedirs(user_dir, exist_ok=True)
    file_counter[user_id] = file_counter.get(user_id, 0) + 1
    filename = f"Excel_{file_counter[user_id]}.xlsx"
    filepath = os.path.join(user_dir, filename)
    wb.save(filepath)

    await bot.send_document(user_id, open(filepath, "rb"))
    await bot.send_message(user_id, "Файл сохранён в архив ✅", reply_markup=main_menu())
    del user_sessions[user_id]

# === ЗАПУСК ===
if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
