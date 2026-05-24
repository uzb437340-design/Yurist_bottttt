import asyncio
import logging
import os
import json
import aiosqlite
from datetime import datetime
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
import httpx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
ADMIN_IDS = list(map(int, os.getenv("ADMIN_IDS", "0").split(",")))
DB = "bot.db"
DOCS_DIR = "docs"
os.makedirs(DOCS_DIR, exist_ok=True)

# ─── KLAVIATURALAR ────────────────────────────────────────────────────────────

def main_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 Hujjat olish", callback_data="menu_docs")],
        [InlineKeyboardButton(text="🤖 AI Maslahat", callback_data="menu_ai")],
        [InlineKeyboardButton(text="ℹ️ Ma'lumot", callback_data="menu_info")],
    ])

def docs_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏠 Ijara shartnomasi - 10,000 so'm", callback_data="doc_ijara")],
        [InlineKeyboardButton(text="🤝 Oldi-sotdi shartnomasi - 10,000 so'm", callback_data="doc_oldi_sotdi")],
        [InlineKeyboardButton(text="✍️ Tilxat - 10,000 so'm", callback_data="doc_tilxat")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="back_main")],
    ])

def pay_menu(order_id):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="💳 To'lash (Demo)", callback_data=f"pay_{order_id}")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="back_main")],
    ])

def cancel_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="back_main")],
    ])

def admin_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Statistika", callback_data="admin_stats")],
        [InlineKeyboardButton(text="📢 Xabar tarqatish", callback_data="admin_broadcast")],
        [InlineKeyboardButton(text="⬅️ Bosh menyu", callback_data="back_main")],
    ])

# ─── MA'LUMOTLAR BAZASI ───────────────────────────────────────────────────────

async def init_db():
    async with aiosqlite.connect(DB) as db:
        await db.execute("""CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY, username TEXT, full_name TEXT)""")
        await db.execute("""CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
            doc_type TEXT, amount INTEGER, status TEXT DEFAULT 'pending',
            data TEXT, file_path TEXT)""")
        await db.commit()

async def save_user(user_id, username, full_name):
    async with aiosqlite.connect(DB) as db:
        await db.execute("INSERT OR IGNORE INTO users (user_id, username, full_name) VALUES (?,?,?)",
                         (user_id, username, full_name))
        await db.commit()

async def create_order(user_id, doc_type, amount, data):
    async with aiosqlite.connect(DB) as db:
        cur = await db.execute(
            "INSERT INTO orders (user_id, doc_type, amount, data) VALUES (?,?,?,?)",
            (user_id, doc_type, amount, json.dumps(data, ensure_ascii=False)))
        await db.commit()
        return cur.lastrowid

async def get_order(order_id):
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT * FROM orders WHERE id=?", (order_id,)) as cur:
            row = await cur.fetchone()
            if row:
                return {"id": row[0], "user_id": row[1], "doc_type": row[2],
                        "amount": row[3], "status": row[4], "data": json.loads(row[5] or "{}"), "file_path": row[6]}

async def update_order(order_id, status, file_path=None):
    async with aiosqlite.connect(DB) as db:
        await db.execute("UPDATE orders SET status=?, file_path=? WHERE id=?",
                         (status, file_path, order_id))
        await db.commit()

async def get_stats():
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT COUNT(*) FROM users") as c:
            users = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(*), SUM(amount) FROM orders WHERE status='paid'") as c:
            row = await c.fetchone()
            orders, revenue = row[0], row[1] or 0
    return users, orders, revenue

async def get_all_users():
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT user_id FROM users") as cur:
            return [r[0] for r in await cur.fetchall()]

# ─── HUJJAT GENERATSIYA ───────────────────────────────────────────────────────

def make_doc(doc_type, data, order_id):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    sana = datetime.now().strftime('%d.%m.%Y')

    if doc_type == "ijara":
        t = doc.add_paragraph()
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.add_run("IJARA SHARTNOMASI").bold = True
        doc.add_paragraph(f"Sana: {sana}  |  Shahar: {data.get('shahar','Toshkent')}")
        doc.add_paragraph(f"Ijara beruvchi: {data.get('beruvchi_fio','___')} (pasport: {data.get('beruvchi_pasport','___')})")
        doc.add_paragraph(f"Ijara oluvchi: {data.get('oluvchi_fio','___')} (pasport: {data.get('oluvchi_pasport','___')})")
        doc.add_paragraph(f"Manzil: {data.get('manzil','___')}")
        doc.add_paragraph(f"Oylik ijara haqi: {data.get('narx','___')} so'm")
        doc.add_paragraph(f"Muddat: {data.get('muddat','___')} oy")
        doc.add_paragraph("\nImzo: __________          Imzo: __________")

    elif doc_type == "oldi_sotdi":
        t = doc.add_paragraph()
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.add_run("OLDI-SOTDI SHARTNOMASI").bold = True
        doc.add_paragraph(f"Sana: {sana}  |  Shahar: {data.get('shahar','Toshkent')}")
        doc.add_paragraph(f"Sotuvchi: {data.get('sotuvchi_fio','___')} (pasport: {data.get('sotuvchi_pasport','___')})")
        doc.add_paragraph(f"Xaridor: {data.get('xaridor_fio','___')} (pasport: {data.get('xaridor_pasport','___')})")
        doc.add_paragraph(f"Mulk: {data.get('tovar','___')}")
        doc.add_paragraph(f"Narxi: {data.get('narx','___')} so'm")
        doc.add_paragraph("\nImzo: __________          Imzo: __________")

    elif doc_type == "tilxat":
        t = doc.add_paragraph()
        t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        t.add_run("TILXAT").bold = True
        doc.add_paragraph(f"Sana: {sana}  |  Shahar: {data.get('shahar','Toshkent')}")
        doc.add_paragraph(f"Men, {data.get('fio','___')}, pasport: {data.get('pasport','___')}, {data.get('manzil','___')} manzilida yashayman.")
        doc.add_paragraph(f"Tasdiqlayman: {data.get('mazmun','___')}")
        doc.add_paragraph(f"Miqdori: {data.get('miqdor','___')} so'm")
        doc.add_paragraph("\nImzo: __________")

    path = f"{DOCS_DIR}/{doc_type}_{order_id}.docx"
    doc.save(path)
    return path

# ─── FSM STATELARI ───────────────────────────────────────────────────────────

class Ijara(StatesGroup):
    beruvchi_fio = State()
    beruvchi_pasport = State()
    oluvchi_fio = State()
    oluvchi_pasport = State()
    manzil = State()
    narx = State()
    muddat = State()
    shahar = State()

class OldiSotdi(StatesGroup):
    sotuvchi_fio = State()
    sotuvchi_pasport = State()
    xaridor_fio = State()
    xaridor_pasport = State()
    tovar = State()
    narx = State()
    shahar = State()

class Tilxat(StatesGroup):
    fio = State()
    pasport = State()
    manzil = State()
    mazmun = State()
    miqdor = State()
    shahar = State()

class AI(StatesGroup):
    savol = State()

class Broadcast(StatesGroup):
    xabar = State()

# ─── BOT ─────────────────────────────────────────────────────────────────────

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

WELCOME = """⚖️ <b>Yuridik Yordamchi Botiga Xush Kelibsiz!</b>

📄 Hujjatlar konstruktori
🤖 AI huquqiy maslahat

Xizmatni tanlang 👇"""

@dp.message(CommandStart())
async def start(message: Message):
    await save_user(message.from_user.id, message.from_user.username or "", message.from_user.full_name or "")
    await message.answer(WELCOME, parse_mode="HTML", reply_markup=main_menu())

@dp.callback_query(F.data == "back_main")
async def back_main(cb: CallbackQuery, state: FSMContext):
    await state.clear()
    await cb.message.edit_text(WELCOME, parse_mode="HTML", reply_markup=main_menu())

@dp.callback_query(F.data == "menu_info")
async def info(cb: CallbackQuery):
    await cb.message.edit_text(
        "ℹ️ <b>Xizmatlar narxi:</b>\n\n• Ijara shartnomasi — 10,000 so'm\n• Oldi-sotdi shartnomasi — 10,000 so'm\n• Tilxat — 10,000 so'm\n• AI Maslahat — 10,000 so'm",
        parse_mode="HTML", reply_markup=main_menu())

@dp.callback_query(F.data == "menu_docs")
async def docs(cb: CallbackQuery):
    await cb.message.edit_text("📂 <b>Hujjat turini tanlang:</b>", parse_mode="HTML", reply_markup=docs_menu())

# ─── IJARA ───────────────────────────────────────────────────────────────────

@dp.callback_query(F.data == "doc_ijara")
async def ijara_start(cb: CallbackQuery, state: FSMContext):
    await state.set_state(Ijara.beruvchi_fio)
    await cb.message.edit_text("🏠 <b>Ijara Shartnomasi</b>\n\n1️⃣ Ijara beruvchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu())

@dp.message(Ijara.beruvchi_fio)
async def i1(m: Message, state: FSMContext):
    await state.update_data(beruvchi_fio=m.text)
    await state.set_state(Ijara.beruvchi_pasport)
    await m.answer("2️⃣ Ijara beruvchining pasporti (AB1234567):")

@dp.message(Ijara.beruvchi_pasport)
async def i2(m: Message, state: FSMContext):
    await state.update_data(beruvchi_pasport=m.text)
    await state.set_state(Ijara.oluvchi_fio)
    await m.answer("3️⃣ Ijara oluvchining to'liq ismi:")

@dp.message(Ijara.oluvchi_fio)
async def i3(m: Message, state: FSMContext):
    await state.update_data(oluvchi_fio=m.text)
    await state.set_state(Ijara.oluvchi_pasport)
    await m.answer("4️⃣ Ijara oluvchining pasporti:")

@dp.message(Ijara.oluvchi_pasport)
async def i4(m: Message, state: FSMContext):
    await state.update_data(oluvchi_pasport=m.text)
    await state.set_state(Ijara.manzil)
    await m.answer("5️⃣ Mulkning to'liq manzili:")

@dp.message(Ijara.manzil)
async def i5(m: Message, state: FSMContext):
    await state.update_data(manzil=m.text)
    await state.set_state(Ijara.narx)
    await m.answer("6️⃣ Oylik ijara haqi (so'mda):")

@dp.message(Ijara.narx)
async def i6(m: Message, state: FSMContext):
    await state.update_data(narx=m.text)
    await state.set_state(Ijara.muddat)
    await m.answer("7️⃣ Ijara muddati (oyda):")

@dp.message(Ijara.muddat)
async def i7(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text)
    await state.set_state(Ijara.shahar)
    await m.answer("8️⃣ Shahar:")

@dp.message(Ijara.shahar)
async def i8(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data()
    await state.clear()
    order_id = await create_order(m.from_user.id, "ijara", 10000, data)
    await m.answer(
        f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Beruvchi: {data['beruvchi_fio']}\n• Oluvchi: {data['oluvchi_fio']}\n• Manzil: {data['manzil']}\n• Narx: {data['narx']} so'm\n\n💳 To'lov qiling:",
        parse_mode="HTML", reply_markup=pay_menu(order_id))

# ─── OLDI-SOTDI ───────────────────────────────────────────────────────────────

@dp.callback_query(F.data == "doc_oldi_sotdi")
async def os_start(cb: CallbackQuery, state: FSMContext):
    await state.set_state(OldiSotdi.sotuvchi_fio)
    await cb.message.edit_text("🤝 <b>Oldi-Sotdi Shartnomasi</b>\n\n1️⃣ Sotuvchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu())

@dp.message(OldiSotdi.sotuvchi_fio)
async def os1(m: Message, state: FSMContext):
    await state.update_data(sotuvchi_fio=m.text)
    await state.set_state(OldiSotdi.sotuvchi_pasport)
    await m.answer("2️⃣ Sotuvchining pasporti:")

@dp.message(OldiSotdi.sotuvchi_pasport)
async def os2(m: Message, state: FSMContext):
    await state.update_data(sotuvchi_pasport=m.text)
    await state.set_state(OldiSotdi.xaridor_fio)
    await m.answer("3️⃣ Xaridorning to'liq ismi:")

@dp.message(OldiSotdi.xaridor_fio)
async def os3(m: Message, state: FSMContext):
    await state.update_data(xaridor_fio=m.text)
    await state.set_state(OldiSotdi.xaridor_pasport)
    await m.answer("4️⃣ Xaridorning pasporti:")

@dp.message(OldiSotdi.xaridor_pasport)
async def os4(m: Message, state: FSMContext):
    await state.update_data(xaridor_pasport=m.text)
    await state.set_state(OldiSotdi.tovar)
    await m.answer("5️⃣ Sotilayotgan mulk tavsifi:")

@dp.message(OldiSotdi.tovar)
async def os5(m: Message, state: FSMContext):
    await state.update_data(tovar=m.text)
    await state.set_state(OldiSotdi.narx)
    await m.answer("6️⃣ Mulk narxi (so'mda):")

@dp.message(OldiSotdi.narx)
async def os6(m: Message, state: FSMContext):
    await state.update_data(narx=m.text)
    await state.set_state(OldiSotdi.shahar)
    await m.answer("7️⃣ Shahar:")

@dp.message(OldiSotdi.shahar)
async def os7(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data()
    await state.clear()
    order_id = await create_order(m.from_user.id, "oldi_sotdi", 10000, data)
    await m.answer(
        f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Sotuvchi: {data['sotuvchi_fio']}\n• Xaridor: {data['xaridor_fio']}\n• Mulk: {data['tovar']}\n• Narx: {data['narx']} so'm\n\n💳 To'lov qiling:",
        parse_mode="HTML", reply_markup=pay_menu(order_id))

# ─── TILXAT ───────────────────────────────────────────────────────────────────

@dp.callback_query(F.data == "doc_tilxat")
async def tx_start(cb: CallbackQuery, state: FSMContext):
    await state.set_state(Tilxat.fio)
    await cb.message.edit_text("✍️ <b>Tilxat</b>\n\n1️⃣ To'liq ismi familiyangiz:", parse_mode="HTML", reply_markup=cancel_menu())

@dp.message(Tilxat.fio)
async def tx1(m: Message, state: FSMContext):
    await state.update_data(fio=m.text)
    await state.set_state(Tilxat.pasport)
    await m.answer("2️⃣ Pasport seriyasi va raqami:")

@dp.message(Tilxat.pasport)
async def tx2(m: Message, state: FSMContext):
    await state.update_data(pasport=m.text)
    await state.set_state(Tilxat.manzil)
    await m.answer("3️⃣ Yashash manzili:")

@dp.message(Tilxat.manzil)
async def tx3(m: Message, state: FSMContext):
    await state.update_data(manzil=m.text)
    await state.set_state(Tilxat.mazmun)
    await m.answer("4️⃣ Tilxat mazmuni (nima tasdiqlanmoqda):")

@dp.message(Tilxat.mazmun)
async def tx4(m: Message, state: FSMContext):
    await state.update_data(mazmun=m.text)
    await state.set_state(Tilxat.miqdor)
    await m.answer("5️⃣ Miqdori (so'mda):")

@dp.message(Tilxat.miqdor)
async def tx5(m: Message, state: FSMContext):
    await state.update_data(miqdor=m.text)
    await state.set_state(Tilxat.shahar)
    await m.answer("6️⃣ Shahar:")

@dp.message(Tilxat.shahar)
async def tx6(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data()
    await state.clear()
    order_id = await create_order(m.from_user.id, "tilxat", 10000, data)
    await m.answer(
        f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Muallif: {data['fio']}\n• Miqdor: {data['miqdor']} so'm\n\n💳 To'lov qiling:",
        parse_mode="HTML", reply_markup=pay_menu(order_id))

# ─── TO'LOV ───────────────────────────────────────────────────────────────────

@dp.callback_query(F.data.startswith("pay_"))
async def pay(cb: CallbackQuery):
    order_id = int(cb.data.split("_")[1])
    order = await get_order(order_id)
    if not order:
        await cb.answer("Buyurtma topilmadi!", show_alert=True)
        return
    await cb.message.edit_text(
        f"💳 <b>To'lov</b>\n\nBuyurtma #{order_id}\nMiqdor: {order['amount']:,} so'm\n\n<i>Demo rejim: Tasdiqlash tugmasini bosing</i>",
        parse_mode="HTML",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Tasdiqlash (Demo)", callback_data=f"confirm_{order_id}")],
            [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="back_main")],
        ]))

@dp.callback_query(F.data.startswith("confirm_"))
async def confirm(cb: CallbackQuery):
    order_id = int(cb.data.split("_")[1])
    order = await get_order(order_id)
    if not order:
        await cb.answer("Buyurtma topilmadi!", show_alert=True)
        return
    await cb.message.edit_text("⏳ Hujjat tayyorlanmoqda...")
    file_path = make_doc(order["doc_type"], order["data"], order_id)
    await update_order(order_id, "paid", file_path)
    doc_file = FSInputFile(file_path)
    await cb.message.answer_document(
        doc_file,
        caption="✅ <b>To'lov tasdiqlandi! Hujjatingiz tayyor.</b>\n\nWord (.docx) formatida — bosib chiqarib imzolashingiz mumkin.",
        parse_mode="HTML",
        reply_markup=main_menu())
    await cb.message.delete()

# ─── AI MASLAHAT ──────────────────────────────────────────────────────────────

@dp.callback_query(F.data == "menu_ai")
async def ai_start(cb: CallbackQuery, state: FSMContext):
    await state.set_state(AI.savol)
    await cb.message.edit_text(
        "🤖 <b>AI Huquqiy Maslahat</b>\n\nHuquqiy muammoingizni yozing:",
        parse_mode="HTML", reply_markup=cancel_menu())

@dp.message(AI.savol)
async def ai_answer(m: Message, state: FSMContext):
    await state.clear()
    msg = await m.answer("🤔 Tahlil qilinmoqda...")
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            r = await client.post(
                "https://api.openai.com/v1/chat/completions",
                headers={"Authorization": f"Bearer {OPENAI_API_KEY}"},
                json={"model": "gpt-4", "messages": [
                    {"role": "system", "content": "Siz O'zbekiston yuridik maslahatchi botisiz. Faqat O'zbekiston qonunchiligiga asoslanib javob bering. Javob oxirida: '⚠️ Bu boshlang'ich maslahat, to'liq yordam uchun advokatga murojaat qiling' deb yozing."},
                    {"role": "user", "content": m.text}
                ], "max_tokens": 800}
            )
            answer = r.json()["choices"][0]["message"]["content"]
    except:
        answer = "❌ AI xizmati hozircha mavjud emas. Keyinroq urinib ko'ring."
    await msg.delete()
    await m.answer(f"⚖️ <b>Huquqiy Maslahat:</b>\n\n{answer}", parse_mode="HTML", reply_markup=main_menu())

# ─── ADMIN ────────────────────────────────────────────────────────────────────

@dp.message(Command("admin"))
async def admin_cmd(m: Message):
    if m.from_user.id not in ADMIN_IDS:
        await m.answer("❌ Ruxsat yo'q.")
        return
    await m.answer("👨‍💼 <b>Admin Panel</b>", parse_mode="HTML", reply_markup=admin_menu())

@dp.callback_query(F.data == "admin_stats")
async def stats(cb: CallbackQuery):
    if cb.from_user.id not in ADMIN_IDS:
        return
    users, orders, revenue = await get_stats()
    await cb.message.edit_text(
        f"📊 <b>Statistika</b>\n\n👥 Foydalanuvchilar: {users}\n📄 Buyurtmalar: {orders}\n💰 Daromad: {revenue:,} so'm",
        parse_mode="HTML", reply_markup=admin_menu())

@dp.callback_query(F.data == "admin_broadcast")
async def broadcast_start(cb: CallbackQuery, state: FSMContext):
    if cb.from_user.id not in ADMIN_IDS:
        return
    await state.set_state(Broadcast.xabar)
    await cb.message.edit_text("📢 Barcha foydalanuvchilarga yuboriladigan xabarni yozing:")

@dp.message(Broadcast.xabar)
async def broadcast_send(m: Message, state: FSMContext):
    if m.from_user.id not in ADMIN_IDS:
        return
    await state.clear()
    users = await get_all_users()
    sent = 0
    for uid in users:
        try:
            await bot.send_message(uid, f"📢 {m.text}")
            sent += 1
        except:
            pass
    await m.answer(f"✅ {sent} ta foydalanuvchiga yuborildi.", reply_markup=admin_menu())

# ─── ISHGA TUSHIRISH ──────────────────────────────────────────────────────────

async def main():
    await init_db()
    logging.info("Bot ishga tushdi!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
