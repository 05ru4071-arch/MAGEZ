import os
import asyncio
from io import BytesIO
from aiogram import Bot, Dispatcher, types
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup, InputFile
from aiogram.filters import Command
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

TOKEN = "ТОКЕН_ТВОЕГО_БОТА"

bot = Bot(token=TOKEN)
dp = Dispatcher()

# Хранилище товаров: user_id -> список товаров
user_data = {}

# Главное меню
def main_menu():
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📂 Создать Excel", callback_data="create_excel")],
        [InlineKeyboardButton(text="📜 Посмотреть архив", callback_data="view_archive")]
    ])
    return kb

# Меню работы с товарами
def product_menu():
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Добавить товар", callback_data="add_item")],
        [InlineKeyboardButton(text="✏ Редактировать товар", callback_data="edit_item")],
        [InlineKeyboardButton(text="🗑 Удалить товар", callback_data="delete_item")],
        [InlineKeyboardButton(text="✅ Завершить Excel", callback_data="finish_excel")]
    ])
    return kb

# Меню завершения
def finish_menu():
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="💾 Сохранить Excel", callback_data="save_excel")],
        [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_products")]
    ])
    return kb

# /start
@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    user_data[message.from_user.id] = []
    await message.answer("Привет! Выберите действие:", reply_markup=main_menu())

# Обработчик меню
@dp.callback_query()
async def menu_handler(callback: types.CallbackQuery):
    user_id = callback.from_user.id

    if callback.data == "create_excel":
        user_data[user_id] = []
        await callback.message.answer("Создание Excel. Добавьте товар:", reply_markup=product_menu())

    elif callback.data == "add_item":
        await callback.message.answer("Отправьте фото или файл товара:")
        dp.message.register(file_handler, user_id=user_id)

    elif callback.data == "finish_excel":
        if not user_data[user_id]:
            await callback.message.answer("Список товаров пуст!")
            return
        text = "📦 Товары:\n"
        for i, item in enumerate(user_data[user_id], start=1):
            text += f"{i}. {item['color']} | {item['size']} | {item['qty']} шт\n"
        await callback.message.answer(text, reply_markup=finish_menu())

    elif callback.data == "back_to_products":
        await callback.message.answer("Меню работы с товарами:", reply_markup=product_menu())

    elif callback.data == "save_excel":
        await callback.message.answer("Введите код груза:")
        dp.message.register(save_excel_name, user_id=user_id)

    elif callback.data == "view_archive":
        user_folder = f"archive/{user_id}"
        if not os.path.exists(user_folder) or not os.listdir(user_folder):
            await callback.message.answer("Ваш архив пуст 📂")
        else:
            files = os.listdir(user_folder)
            text = "📜 Ваши файлы:\n" + "\n".join(files)
            await callback.message.answer(text)

# Приём фото/файлов
async def file_handler(message: types.Message):
    user_id = message.from_user.id

    if message.photo:
        file_id = message.photo[-1].file_id
    elif message.document:
        file_id = message.document.file_id
    else:
        await message.answer("Пожалуйста, отправьте фото или файл 📷")
        return

    file = await bot.get_file(file_id)
    downloaded = await bot.download_file(file.file_path)
    file_bytes = downloaded.read()

    product = {
        "photo": file_bytes,
        "link": "https://пример.ссылка",
        "color": "красный",
        "size": "M",
        "qty": 1,
        "comment": "без комментариев"
    }

    user_data[user_id].append(product)
    await message.answer("✅ Товар добавлен!", reply_markup=product_menu())

# Сохраняем Excel
async def save_excel_name(message: types.Message):
    user_id = message.from_user.id
    code = message.text.strip()
    filename = f"{code}.xlsx"

    wb = Workbook()
    ws = wb.active

    # --- Шапка компании ---
    ws.merge_cells("A1:F1")
    ws.merge_cells("A2:F2")

    cell1 = ws["A1"]
    cell1.value = "MAGEZ"
    cell1.font = Font(color="FFFFFF", bold=True, size=16)
    cell1.fill = PatternFill("solid", fgColor="FF0000")
    cell1.alignment = Alignment(horizontal="center", vertical="center")

    cell2 = ws["A2"]
    cell2.value = "Торгово-Логистическая компания"
    cell2.font = Font(color="FFFFFF", bold=True, size=12)
    cell2.fill = PatternFill("solid", fgColor="FF0000")
    cell2.alignment = Alignment(horizontal="center", vertical="center")

    # --- Стили для таблицы ---
    header_fill = PatternFill("solid", fgColor="808080")
    header_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # --- Заголовки таблицы ---
    headers = ["Фото", "Ссылка", "Цвет", "Размер", "Количество", "Комментарий"]
    ws.append([])
    ws.append(headers)
    for col, name in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # --- Товары ---
    for i, item in enumerate(user_data[user_id], start=5):
        ws.cell(row=i, column=2).value = item["link"]
        ws.cell(row=i, column=3).value = item["color"]
        ws.cell(row=i, column=4).value = item["size"]
        ws.cell(row=i, column=5).value = item["qty"]
        ws.cell(row=i, column=6).value = item["comment"]

        for col in range(2, 7):
            cell = ws.cell(row=i, column=col)
            cell.alignment = center_align
            cell.border = thin_border

        if item["photo"]:
            img = XLImage(BytesIO(item["photo"]))
            img.width, img.height = 80, 80
            ws.add_image(img, f"A{i}")
            ws.row_dimensions[i].height = 60

    # Автоширина колонок
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 5

    # --- Сохраняем в архив пользователя ---
    user_folder = f"archive/{user_id}"
    os.makedirs(user_folder, exist_ok=True)
    filepath = f"{user_folder}/{filename}"
    wb.save(filepath)

    # Отправляем пользователю
    await message.answer_document(InputFile(filepath))
    await message.answer("✅ Файл сохранён в архив!", reply_markup=main_menu())

# Запуск
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
