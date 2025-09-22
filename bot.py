
# bot.py - Полный Telegram бот для создания Excel с товарами

import os
import io
import json
import secrets
import aiofiles
import aiosqlite
from datetime import datetime
from typing import List, Dict, Optional
from dotenv import load_dotenv

from aiogram import Bot, Dispatcher, Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, ContentType
from aiogram.filters import Command
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.utils.keyboard import InlineKeyboardBuilder

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image

# ============= КОНФИГУРАЦИЯ =============
load_dotenv()

BOT_TOKEN = os.getenv('BOT_TOKEN')
ADMIN_ID = int(os.getenv('ADMIN_ID'))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_DIR = os.path.join(BASE_DIR, 'archive')
TEMP_DIR = os.path.join(BASE_DIR, 'temp')
DB_PATH = os.path.join(BASE_DIR, 'bot.db')

os.makedirs(ARCHIVE_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)

# ============= СОСТОЯНИЯ FSM =============
class AuthStates(StatesGroup):
    waiting_for_invite = State()

class ProductStates(StatesGroup):
    waiting_for_photo = State()
    waiting_for_link = State()
    waiting_for_color = State()
    waiting_for_size = State()
    waiting_for_quantity = State()
    waiting_for_comment = State()

    selecting_product_to_edit = State()
    selecting_field_to_edit = State()
    editing_field = State()

    selecting_product_to_delete = State()
    confirming_deletion = State()

    waiting_for_cargo_code = State()

# ============= БАЗА ДАННЫХ =============
class Database:
    def __init__(self):
        self.db_path = DB_PATH

    async def create_tables(self):
        async with aiosqlite.connect(self.db_path) as db:
            await db.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    first_name TEXT,
                    joined_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            await db.execute('''
                CREATE TABLE IF NOT EXISTS invites (
                    code TEXT PRIMARY KEY,
                    created_by INTEGER,
                    used_by INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    used_at TIMESTAMP,
                    is_used BOOLEAN DEFAULT 0
                )
            ''')

            await db.execute('''
                CREATE TABLE IF NOT EXISTS sessions (
                    user_id INTEGER PRIMARY KEY,
                    products TEXT DEFAULT '[]',
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            await db.commit()

    async def add_user(self, user_id: int, username: str = None, first_name: str = None):
        async with aiosqlite.connect(self.db_path) as db:
            await db.execute(
                'INSERT OR IGNORE INTO users (user_id, username, first_name) VALUES (?, ?, ?)',
                (user_id, username, first_name)
            )
            await db.commit()

    async def is_user_authorized(self, user_id: int) -> bool:
        async with aiosqlite.connect(self.db_path) as db:
            cursor = await db.execute('SELECT user_id FROM users WHERE user_id = ?', (user_id,))
            result = await cursor.fetchone()
            return result is not None

    async def create_invite(self, created_by: int) -> str:
        code = secrets.token_urlsafe(16)
        async with aiosqlite.connect(self.db_path) as db:
            await db.execute(
                'INSERT INTO invites (code, created_by) VALUES (?, ?)',
                (code, created_by)
            )
            await db.commit()
        return code

    async def use_invite(self, code: str, user_id: int) -> bool:
        async with aiosqlite.connect(self.db_path) as db:
            cursor = await db.execute(
                'SELECT code FROM invites WHERE code = ? AND is_used = 0',
                (code,)
            )
            result = await cursor.fetchone()

            if result:
                await db.execute(
                    'UPDATE invites SET used_by = ?, used_at = CURRENT_TIMESTAMP, is_used = 1 WHERE code = ?',
                    (user_id, code)
                )
                await db.commit()
                return True
            return False

    async def get_session_products(self, user_id: int) -> List[Dict]:
        async with aiosqlite.connect(self.db_path) as db:
            cursor = await db.execute(
                'SELECT products FROM sessions WHERE user_id = ?',
                (user_id,)
            )
            result = await cursor.fetchone()
            if result:
                return json.loads(result[0])
            return []

    async def save_session_products(self, user_id: int, products: List[Dict]):
        async with aiosqlite.connect(self.db_path) as db:
            await db.execute(
                '''INSERT INTO sessions (user_id, products, updated_at) 
                   VALUES (?, ?, CURRENT_TIMESTAMP)
                   ON CONFLICT(user_id) DO UPDATE SET
                   products = excluded.products,
                   updated_at = CURRENT_TIMESTAMP''',
                (user_id, json.dumps(products, ensure_ascii=False))
            )
            await db.commit()

    async def clear_session(self, user_id: int):
        async with aiosqlite.connect(self.db_path) as db:
            await db.execute('DELETE FROM sessions WHERE user_id = ?', (user_id,))
            await db.commit()

# ============= КЛАВИАТУРЫ =============
def main_menu_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    kb.row(
        InlineKeyboardButton(text="📂 Создать Excel", callback_data="create_excel"),
        InlineKeyboardButton(text="📜 Посмотреть архив", callback_data="view_archive")
    )
    return kb.as_markup()

def product_actions_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    kb.row(
        InlineKeyboardButton(text="➕ Добавить товар", callback_data="add_product"),
        InlineKeyboardButton(text="✏️ Редактировать товар", callback_data="edit_product")
    )
    kb.row(
        InlineKeyboardButton(text="🗑 Удалить товар", callback_data="delete_product"),
        InlineKeyboardButton(text="✅ Завершить Excel", callback_data="finish_excel")
    )
    kb.row(InlineKeyboardButton(text="🏠 Главное меню", callback_data="main_menu"))
    return kb.as_markup()

def products_list_kb(products: List[Dict], action: str) -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    for i, product in enumerate(products):
        preview = f"Товар {i+1}: {product.get('color', 'Без цвета')}, {product.get('size', 'Без размера')}"
        kb.row(InlineKeyboardButton(text=preview[:50], callback_data=f"{action}_{i}"))
    kb.row(InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_actions"))
    return kb.as_markup()

def edit_fields_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    fields = [
        ("📷 Фото", "edit_photo"),
        ("🔗 Ссылка", "edit_link"),
        ("🎨 Цвет", "edit_color"),
        ("📏 Размер", "edit_size"),
        ("🔢 Количество", "edit_quantity"),
        ("💬 Комментарий", "edit_comment")
    ]
    for text, callback in fields:
        kb.row(InlineKeyboardButton(text=text, callback_data=callback))
    kb.row(InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_actions"))
    return kb.as_markup()

def confirm_delete_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    kb.row(
        InlineKeyboardButton(text="❌ Удалить", callback_data="confirm_delete"),
        InlineKeyboardButton(text="🔙 Отмена", callback_data="back_to_actions")
    )
    return kb.as_markup()

def save_excel_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    kb.row(
        InlineKeyboardButton(text="💾 Сохранить Excel", callback_data="save_excel"),
        InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_actions")
    )
    return kb.as_markup()

def cancel_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardBuilder()
    kb.row(InlineKeyboardButton(text="❌ Отмена", callback_data="cancel"))
    return kb.as_markup()

# ============= ГЕНЕРАТОР EXCEL =============
class ExcelGenerator:
    def __init__(self):
        self.red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        self.gray_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        self.white_bold_font = Font(bold=True, color="FFFFFF", size=14)
        self.header_font = Font(bold=True, color="FFFFFF", size=12)
        self.center_alignment = Alignment(horizontal="center", vertical="center")
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

    def create_excel(self, products: List[Dict], filename: str, file_path: str) -> str:
        wb = Workbook()
        ws = wb.active
        ws.title = "Products"

        # Company header
        ws.merge_cells('A1:F1')
        ws['A1'] = 'MAGEZ'
        ws['A1'].font = self.white_bold_font
        ws['A1'].fill = self.red_fill
        ws['A1'].alignment = self.center_alignment
        ws.row_dimensions[1].height = 30

        ws.merge_cells('A2:F2')
        ws['A2'] = 'Торгово-Логистическая компания'
        ws['A2'].font = Font(bold=True, color="FFFFFF", size=12)
        ws['A2'].fill = self.red_fill
        ws['A2'].alignment = self.center_alignment
        ws.row_dimensions[2].height = 25

        # Table headers
        headers = ['Фото', 'Ссылка', 'Цвет', 'Размер', 'Количество', 'Комментарий']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.font = self.header_font
            cell.fill = self.gray_fill
            cell.alignment = self.center_alignment
            cell.border = self.thin_border

        # Add products
        for row_idx, product in enumerate(products, start=5):
            # Photo handling
            photo_cell = ws.cell(row=row_idx, column=1)
            if product.get('photo_path') and os.path.exists(product['photo_path']):
                try:
                    img = Image.open(product['photo_path'])
                    img.thumbnail((100, 100))
                    img_buffer = io.BytesIO()
                    img.save(img_buffer, format='PNG')
                    img_buffer.seek(0)

                    xl_img = XLImage(img_buffer)
                    xl_img.width = 100
                    xl_img.height = 100

                    ws.add_image(xl_img, f'A{row_idx}')
                    ws.row_dimensions[row_idx].height = 80
                except Exception:
                    photo_cell.value = os.path.basename(product.get('photo_path', 'No photo'))
            else:
                photo_cell.value = product.get('photo_name', 'No photo')

            # Other fields
            ws.cell(row=row_idx, column=2, value=product.get('link', '')).alignment = self.center_alignment
            ws.cell(row=row_idx, column=3, value=product.get('color', '')).alignment = self.center_alignment
            ws.cell(row=row_idx, column=4, value=product.get('size', '')).alignment = self.center_alignment
            ws.cell(row=row_idx, column=5, value=product.get('quantity', '')).alignment = self.center_alignment
            ws.cell(row=row_idx, column=6, value=product.get('comment', '')).alignment = self.center_alignment

            # Apply borders
            for col in range(1, 7):
                ws.cell(row=row_idx, column=col).border = self.thin_border

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

            adjusted_width = min(max_length + 2, 40)
            ws.column_dimensions[column_letter].width = max(adjusted_width, 15)

        ws.column_dimensions['A'].width = 20

        full_path = os.path.join(file_path, f"{filename}.xlsx")
        wb.save(full_path)

        return full_path

# ============= ОБРАБОТЧИКИ =============
router = Router()
db = Database()
excel_gen = ExcelGenerator()

# --- Авторизация ---
@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    user_id = message.from_user.id

    if user_id == ADMIN_ID:
        await db.add_user(user_id, message.from_user.username, message.from_user.first_name)
        await message.answer(
            f"👋 Добро пожаловать, администратор!\n\n"
            f"Используйте /admin для доступа к админ-панели",
            reply_markup=main_menu_kb()
        )
        return

    if await db.is_user_authorized(user_id):
        await message.answer(
            f"👋 С возвращением, {message.from_user.first_name}!",
            reply_markup=main_menu_kb()
        )
    else:
        args = message.text.split()
        if len(args) > 1:
            invite_code = args[1]
            if await db.use_invite(invite_code, user_id):
                await db.add_user(user_id, message.from_user.username, message.from_user.first_name)
                await message.answer(
                    f"✅ Приглашение активировано!\n"
                    f"Добро пожаловать, {message.from_user.first_name}!",
                    reply_markup=main_menu_kb()
                )
            else:
                await message.answer(
                    "❌ Неверный или уже использованный код приглашения.\n"
                    "Введите действующий код:"
                )
                await state.set_state(AuthStates.waiting_for_invite)
        else:
            await message.answer(
                "🔒 Для доступа к боту необходим код приглашения.\n"
                "Введите код:"
            )
            await state.set_state(AuthStates.waiting_for_invite)

@router.message(AuthStates.waiting_for_invite)
async def process_invite(message: Message, state: FSMContext):
    invite_code = message.text.strip()
    user_id = message.from_user.id

    if await db.use_invite(invite_code, user_id):
        await db.add_user(user_id, message.from_user.username, message.from_user.first_name)
        await message.answer(
            f"✅ Приглашение активировано!\n"
            f"Добро пожаловать, {message.from_user.first_name}!",
            reply_markup=main_menu_kb()
        )
        await state.clear()
    else:
        await message.answer(
            "❌ Неверный или уже использованный код приглашения.\n"
            "Попробуйте ещё раз или обратитесь к администратору."
        )

# --- Админ команды ---
@router.message(Command("admin"))
async def admin_panel(message: Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer("❌ У вас нет доступа к админ-панели.")
        return

    await message.answer(
        "🔧 <b>Админ-панель</b>\n\n"
        "Доступные команды:\n"
        "/invite - создать новое приглашение\n"
        "/users - список пользователей",
        parse_mode="HTML"
    )

@router.message(Command("invite"))
async def create_invite(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    code = await db.create_invite(message.from_user.id)
    invite_link = f"https://t.me/{(await message.bot.me()).username}?start={code}"

    await message.answer(
        f"🎟 <b>Новое приглашение создано!</b>\n\n"
        f"Код: <code>{code}</code>\n"
        f"Ссылка: {invite_link}\n\n"
        f"<i>Приглашение одноразовое</i>",
        parse_mode="HTML"
    )

@router.message(Command("users"))
async def list_users(message: Message):
    if message.from_user.id != ADMIN_ID:
        return

    async with aiosqlite.connect(DB_PATH) as conn:
        cursor = await conn.execute(
            "SELECT user_id, username, first_name, joined_at FROM users ORDER BY joined_at DESC"
        )
        users = await cursor.fetchall()

    if not users:
        await message.answer("📭 Нет зарегистрированных пользователей")
        return

    text = "👥 <b>Список пользователей:</b>\n\n"
    for user in users:
        user_id, username, first_name, joined_at = user
        text += f"• {first_name or 'No name'} "
        if username:
            text += f"(@{username}) "
        text += f"[{user_id}]\n"
        text += f"  Присоединился: {joined_at[:10]}\n\n"

    await message.answer(text, parse_mode="HTML")

# --- Главное меню ---
@router.callback_query(F.data == "main_menu")
async def show_main_menu(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text(
        f"👋 Главное меню",
        reply_markup=main_menu_kb()
    )

@router.callback_query(F.data == "cancel")
async def cancel_action(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text(
        "❌ Действие отменено",
        reply_markup=product_actions_kb()
    )

# --- Создание Excel ---
@router.callback_query(F.data == "create_excel")
async def create_excel(callback: CallbackQuery, state: FSMContext):
    if not await db.is_user_authorized(callback.from_user.id):
        await callback.answer("❌ Необходима авторизация", show_alert=True)
        return

    await state.clear()
    await db.clear_session(callback.from_user.id)
    await callback.message.edit_text(
        "📂 <b>Создание нового Excel файла</b>\n\n"
        "Начните добавление товаров",
        parse_mode="HTML",
        reply_markup=product_actions_kb()
    )

@router.callback_query(F.data == "back_to_actions")
async def back_to_actions(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    products = await db.get_session_products(callback.from_user.id)
    await callback.message.edit_text(
        f"📦 Товаров добавлено: {len(products)}",
        reply_markup=product_actions_kb()
    )

# --- Добавление товара ---
@router.callback_query(F.data == "add_product")
async def start_add_product(callback: CallbackQuery, state: FSMContext):
    await callback.message.edit_text(
        "📷 Отправьте фото товара (или любой файл):",
        reply_markup=cancel_kb()
    )
    await state.set_state(ProductStates.waiting_for_photo)

@router.message(ProductStates.waiting_for_photo)
async def process_photo(message: Message, state: FSMContext):
    user_data = {}
    user_dir = os.path.join(TEMP_DIR, str(message.from_user.id))
    os.makedirs(user_dir, exist_ok=True)

    file_id = None
    file_name = None

    if message.photo:
        file_id = message.photo[-1].file_id
        file_name = f"photo_{datetime.now().timestamp()}.jpg"
    elif message.document:
        file_id = message.document.file_id
        file_name = message.document.file_name or f"doc_{datetime.now().timestamp()}"
    elif message.video:
        file_id = message.video.file_id
        file_name = f"video_{datetime.now().timestamp()}.mp4"
    else:
        await message.answer("❌ Пожалуйста, отправьте файл или фото")
        return

    file_path = os.path.join(user_dir, file_name)
    file = await message.bot.get_file(file_id)
    await message.bot.download_file(file.file_path, file_path)

    user_data['photo_path'] = file_path
    user_data['photo_name'] = file_name

    await state.update_data(**user_data)
    await message.answer("🔗 Введите ссылку на товар:")
    await state.set_state(ProductStates.waiting_for_link)

@router.message(ProductStates.waiting_for_link)
async def process_link(message: Message, state: FSMContext):
    await state.update_data(link=message.text)
    await message.answer("🎨 Введите цвет товара:")
    await state.set_state(ProductStates.waiting_for_color)

@router.message(ProductStates.waiting_for_color)
async def process_color(message: Message, state: FSMContext):
    await state.update_data(color=message.text)
    await message.answer("📏 Введите размер товара:")
    await state.set_state(ProductStates.waiting_for_size)

@router.message(ProductStates.waiting_for_size)
async def process_size(message: Message, state: FSMContext):
    await state.update_data(size=message.text)
    await message.answer("🔢 Введите количество:")
    await state.set_state(ProductStates.waiting_for_quantity)

@router.message(ProductStates.waiting_for_quantity)
async def process_quantity(message: Message, state: FSMContext):
    await state.update_data(quantity=message.text)
    await message.answer("💬 Введите комментарий:")
    await state.set_state(ProductStates.waiting_for_comment)

@router.message(ProductStates.waiting_for_comment)
async def process_comment(message: Message, state: FSMContext):
    await state.update_data(comment=message.text)

    product_data = await state.get_data()
    products = await db.get_session_products(message.from_user.id)
    products.append(product_data)

    await db.save_session_products(message.from_user.id, products)

    await message.answer(
        f"✅ Товар добавлен!\n\n"
        f"Всего товаров: {len(products)}",
        reply_markup=product_actions_kb()
    )
    await state.clear()

# --- Редактирование товара ---
@router.callback_query(F.data == "edit_product")
async def select_product_to_edit(callback: CallbackQuery, state: FSMContext):
    products = await db.get_session_products(callback.from_user.id)

    if not products:
        await callback.answer("❌ Нет товаров для редактирования", show_alert=True)
        return

    await callback.message.edit_text(
        "✏️ Выберите товар для редактирования:",
        reply_markup=products_list_kb(products, "edit_select")
    )
    await state.set_state(ProductStates.selecting_product_to_edit)

@router.callback_query(F.data.startswith("edit_select_"))
async def select_field_to_edit(callback: CallbackQuery, state: FSMContext):
    product_index = int(callback.data.split("_")[2])
    await state.update_data(editing_index=product_index)

    products = await db.get_session_products(callback.from_user.id)
    product = products[product_index]

    text = f"📦 <b>Товар {product_index + 1}</b>\n\n"
    text += f"🔗 Ссылка: {product.get('link', 'Не указано')}\n"
    text += f"🎨 Цвет: {product.get('color', 'Не указано')}\n"
    text += f"📏 Размер: {product.get('size', 'Не указано')}\n"
    text += f"🔢 Количество: {product.get('quantity', 'Не указано')}\n"
    text += f"💬 Комментарий: {product.get('comment', 'Не указано')}\n\n"
    text += "Выберите поле для редактирования:"

    await callback.message.edit_text(
        text,
        parse_mode="HTML",
        reply_markup=edit_fields_kb()
    )
    await state.set_state(ProductStates.selecting_field_to_edit)

@router.callback_query(ProductStates.selecting_field_to_edit)
async def edit_field(callback: CallbackQuery, state: FSMContext):
    field_map = {
        "edit_photo": ("photo", "📷 Отправьте новое фото:"),
        "edit_link": ("link", "🔗 Введите новую ссылку:"),
        "edit_color": ("color", "🎨 Введите новый цвет:"),
        "edit_size": ("size", "📏 Введите новый размер:"),
        "edit_quantity": ("quantity", "🔢 Введите новое количество:"),
        "edit_comment": ("comment", "💬 Введите новый комментарий:")
    }

    if callback.data in field_map:
        field, prompt = field_map[callback.data]
        await state.update_data(editing_field=field)
        await callback.message.edit_text(prompt)
        await state.set_state(ProductStates.editing_field)

@router.message(ProductStates.editing_field)
async def save_edited_field(message: Message, state: FSMContext):
    data = await state.get_data()
    field = data.get('editing_field')
    product_index = data.get('editing_index')

    products = await db.get_session_products(message.from_user.id)

    if field == 'photo':
        user_dir = os.path.join(TEMP_DIR, str(message.from_user.id))
        os.makedirs(user_dir, exist_ok=True)

        if message.photo:
            file_id = message.photo[-1].file_id
            file_name = f"photo_{datetime.now().timestamp()}.jpg"
        elif message.document:
            file_id = message.document.file_id
            file_name = message.document.file_name
        else:
            await message.answer("❌ Пожалуйста, отправьте файл или фото")
            return

        file_path = os.path.join(user_dir, file_name)
        file = await message.bot.get_file(file_id)
        await message.bot.download_file(file.file_path, file_path)

        products[product_index]['photo_path'] = file_path
        products[product_index]['photo_name'] = file_name
    else:
        products[product_index][field] = message.text

    await db.save_session_products(message.from_user.id, products)

    await message.answer(
        f"✅ Товар {product_index + 1} обновлен!",
        reply_markup=product_actions_kb()
    )
    await state.clear()

# --- Удаление товара ---
@router.callback_query(F.data == "delete_product")
async def select_product_to_delete(
