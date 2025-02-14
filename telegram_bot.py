import logging
import os
import time
import schedule
import pandas as pd
import gspread
from aiogram import Bot, Dispatcher, types
from aiogram.types import InputFile
from aiogram.utils import executor
from oauth2client.service_account import ServiceAccountCredentials

# Настройки бота
TOKEN = "YOUR_BOT_TOKEN_HERE"  # Вставьте токен из BotFather
GROUP_CHAT_ID = -1001234567890  # Вставьте ID Telegram-группы
TOPIC_ID = 1263  # ID темы "Внешний вид ПКЮИ"

# Настройки Google Sheets
SPREADSHEET_NAME = "Report"
CREDENTIALS_FILE = "credentials.json"  # Файл JSON с ключами Google API

# Подключение к Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open(SPREADSHEET_NAME).sheet1

# Логирование
logging.basicConfig(level=logging.INFO)
bot = Bot(token=TOKEN)
dp = Dispatcher(bot)

# Хранилище данных
data = {}

@dp.message_handler(lambda message: message.is_topic_message and message.message_thread_id == TOPIC_ID)
async def handle_topic_message(message: types.Message):
    if message.text.isdigit():
        data[message.chat.id] = {'shop_number': message.text}
        await message.reply("Теперь отправьте фото ювелира.")
    elif message.photo:
        user_data = data.get(message.chat.id)
        if not user_data or 'shop_number' not in user_data:
            await message.reply("Сначала отправьте номер магазина.")
            return
        
        file_id = message.photo[-1].file_id
        file_path = f"photo_{user_data['shop_number']}.jpg"
        await bot.download_file_by_id(file_id, file_path)
        
        # Записываем данные в Google Sheets
        sheet.append_row([user_data['shop_number'], file_path, time.strftime("%Y-%m-%d %H:%M:%S")])
        
        await message.reply("Фото загружено. Спасибо!")
        del data[message.chat.id]

# Функция для создания отчета
async def send_report():
    records = sheet.get_all_records()
    df = pd.DataFrame(records)
    
    if df.empty:
        await bot.send_message(GROUP_CHAT_ID, "На сегодня данных нет.")
        return
    
    df['Фото загружено'] = df['Фото'].apply(lambda x: '✅' if x else '❌')
    report_path = "report.xlsx"
    df.to_excel(report_path, index=False)
    
    with open(report_path, "rb") as file:
        await bot.send_document(GROUP_CHAT_ID, InputFile(file, filename="Отчет.xlsx"))

# Планируем отправку отчета в 13:00
schedule.every().day.at("13:00").do(lambda: asyncio.run(send_report()))

async def scheduler():
    while True:
        schedule.run_pending()
        await asyncio.sleep(60)  # Проверять задачи каждую минуту

if __name__ == "__main__":
    loop = asyncio.get_event_loop()
    loop.create_task(scheduler())
    executor.start_polling(dp, skip_updates=True)
