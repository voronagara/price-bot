from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, CallbackQueryHandler, filters
import pandas as pd
import json
import re
import os
import gdown

# 🔗 Google Drive файл (Excel)
GDRIVE_LINK = "https://drive.google.com/uc?id=1BVD0nAZoj5Ug2y3bytqfRwWRQp2P8hA2"
XLSX_FILE = "svodna_tablycya.xlsx"
excel_data = {}

def download_excel():
    if os.path.exists(XLSX_FILE):
        os.remove(XLSX_FILE)
    gdown.download(GDRIVE_LINK, XLSX_FILE, quiet=False)

def load_excel_to_memory():
    xls = pd.ExcelFile(XLSX_FILE)
    data = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [c.lower().strip() for c in df.columns]
        if "номенклатура товарів/послуг" in df.columns and "дата виписки" in df.columns:
            df["дата виписки"] = pd.to_datetime(df["дата виписки"], errors="coerce")
            data[sheet] = df
    return data

# 🔐 Адмін і користувачі
ADMIN_ID = 339950143
USERS_FILE = "allowed_users.json"

def load_users():
    if not os.path.exists(USERS_FILE):
        return []
    with open(USERS_FILE, "r") as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f)

allowed_users = load_users()

# 📍 /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    msg = f"👋 Привіт, {update.effective_user.first_name}!
Ваш Telegram ID: {user_id}"
    keyboard = [[InlineKeyboardButton("🔎 Зробити запит", callback_data="make_query")]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(msg, reply_markup=reply_markup)

# 📍 /id
async def get_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"Ваш Telegram ID: {update.effective_user.id}")

# 📍 /users
async def list_users(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ У вас немає прав на цю команду.")
        return
    await update.message.reply_text("👥 Список дозволених ID:\n" + "\n".join(str(uid) for uid in allowed_users))

# 📍 /admin add
async def admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("⛔ У вас немає прав на цю команду.")
        return

    args = context.args
    if len(args) != 2 or args[0] != "add":
        await update.message.reply_text("⚙️ Формат:\n/admin add 123456789")
        return

    try:
        new_id = int(args[1])
        if new_id not in allowed_users:
            allowed_users.append(new_id)
            save_users(allowed_users)
            await update.message.reply_text(f"✅ Користувача {new_id} додано.")
        else:
            await update.message.reply_text(f"ℹ️ Користувач {new_id} вже є.")
    except ValueError:
        await update.message.reply_text("❗ ID має бути числом.")

# 🔘 Кнопка
async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.message.reply_text("📌 Введіть запит:\n➔ Артикул (3364150) \u0447и VRP350/VRP 350/VRP-350, січень-грудень 2024")

# 📊 Аналіз
month_map = {
    "січень": "January", "лютий": "February", "березень": "March", "квітень": "April",
    "травень": "May", "червень": "June", "липень": "July", "серпень": "August",
    "вересень": "September", "жовтень": "October", "листопад": "November", "грудень": "December"
}

def normalize(text):
    return re.sub(r"[\s\-]", "", str(text)).lower()

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in allowed_users:
        await update.message.reply_text("⛔ У вас немає доступу до цього бота.")
        return

    text = update.message.text.lower().replace("–", "-")

    match = re.match(r"(.+?),\s*(.+?)\s*-\s*(.+?)\s*(\d{4})", text)
    if match:
        raw_skus, month_start, month_end, year = match.groups()
        sku_variants = [normalize(s) for s in raw_skus.split("/") if s.strip()]
        month_start_en = month_map.get(month_start.strip())
        month_end_en = month_map.get(month_end.strip())

        if not month_start_en or not month_end_en:
            await update.message.reply_text("Не вдалося розпізнати місяці.")
            return

        start_date = pd.to_datetime(f"1 {month_start_en} {year}", dayfirst=True)
        end_date = pd.to_datetime(f"1 {month_end_en} {year}", dayfirst=True) + pd.offsets.MonthEnd(0)
    else:
        sku_variants = [normalize(s) for s in text.split("/") if s.strip()]
        start_date, end_date = None, None

    rows = []
    for sheet, df in excel_data.items():
        df_filtered = df[df["номенклатура товарів/послуг"].apply(
            lambda x: any(variant in normalize(x) for variant in sku_variants)
        )]

        if start_date and end_date:
            df_filtered = df_filtered[
                (df_filtered["дата виписки"] >= start_date) &
                (df_filtered["дата виписки"] <= end_date)
            ]

        if not df_filtered.empty:
            qty = int(df_filtered["кількість (об’єм , обсяг)"].sum())
            avg = round(df_filtered["ціна з пдв"].mean(), 2)
            rows.append((sheet, qty, avg))

    if not rows:
        await update.message.reply_text("Продажів не знайдено.")
        return

    rows.sort(key=lambda x: x[1], reverse=True)
    table = "📊 <b>Аналіз продажів</b>\n\n"
    table += "<pre>{:<20} {:>10} {:>15}</pre>\n".format("Постачальник", "Кількість", "Середня ціна")
    for row in rows:
        name = row[0][:20]
        qty = f"{row[1]:,}".replace(",", " ")
        avg = f"{row[2]:,.2f}".replace(",", " ")
        table += "<pre>{:<20} {:>10} {:>15}</pre>\n".format(name, qty, avg)

    await update.message.reply_text(table, parse_mode="HTML")

# 🚀 Запуск

def main():
    print("☁️ Завантаження Excel з Google Drive...")
    download_excel()
    global excel_data
    excel_data = load_excel_to_memory()
    print("✅ Excel завантажено в пам'ять. Бот працює!")

    app = ApplicationBuilder().token("7762946339:AAGyH09Ty3eODo2SQTNW__ARZxshGo00NV8").build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("id", get_id))
    app.add_handler(CommandHandler("users", list_users))
    app.add_handler(CommandHandler("admin", admin_command))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.run_polling()

if __name__ == "__main__":
    main()
