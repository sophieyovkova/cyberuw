import os
import re
from google.cloud import vision
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep
import asyncio

# Подключение к боту Telegram
TOKEN = '8076865693:AAH7G8wDM-KI5iwDkOQyyrFXPxatapZcieg'
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "/Users/a2/Documents/arcuw-439010-d269eb0deee2.json"

# Подключение к Google Sheets
def connect_to_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1doCp5Bs2jaS3CBj28L07ija3GnMp3smb-V1OK3vwR1Q").sheet1
    return sheet

# Функция для проверки существующих заголовков таблицы
def check_headers():
    sheet = connect_to_google_sheets()
    city_column = sheet.col_values(1)
    product_row = sheet.row_values(2)
    
    if len(city_column) > 1 and len(product_row) > 1:
        return True
    else:
        return False

# Функция для создания заголовков таблицы
def create_table_headers():
    sheet = connect_to_google_sheets()
    products = ['Agate', 'Alcohol', 'Bananas', 'Carpets', 'Cloth', 'Diamonds', 'Dye', 'Firearms', 'Fish', 'Gold', 
                'Leather', 'Meat', 'Medicine', 'Paper', 'Peanuts', 'Pearls', 'Porcelain', 'Leaves', 'Tin', 'Tobacco']
    sheet.insert_row([''] + products, 2)

    cities = ['Aden', 'Alexandria', 'Amsterdam', 'Athens', 'Basrah', 'Boston', 'Brunei', 'Buenos', 'Calicut', 
              'Cape', 'Cayenne', 'Ceylon', 'Copenhagen', 'Darwin', 'Edo', 'Hamburg', 'Hangzhou', 'Istanbul', 
              'Jamaica', 'Kolkata', 'Lisbon', 'Las Palmas', 'London', 'Luanda', 'Malacca', 'Manila', 'Marseille', 
              'Mozambique', 'Nantes', 'Nassau', 'Panama City', 'Pinjarra', 'Quanzhou', 'Rio', 
              'Santo', 'Seville', 'St.', 'Stockholm', 'Tunis', 'Venice']

    for i, city in enumerate(cities, start=3):
        sheet.update_cell(i, 1, city)

# Функция для подсветки ячеек в зависимости от значения
def highlight_cells(sheet, row, col, value):
    if value.endswith('%'):
        percent = int(value[:-1])
        if percent >= 106:
            range_notation = gspread.utils.rowcol_to_a1(row, col)
            sheet.format(range_notation, {
                "backgroundColor": {"red": 0.56, "green": 0.93, "blue": 0.56}
            })

# Функция для экспорта данных в таблицу
def export_to_sheets(city, data):
    sheet = connect_to_google_sheets()
    cities = sheet.col_values(1)
    products = sheet.row_values(2)[1:]

    try:
        city_index = cities.index(city) + 1
    except ValueError:
        print(f"Город {city} не найден в таблице")
        return
    for product, percent in data:
        try:
            product_index = products.index(product) + 2
            sheet.update_cell(city_index, product_index, percent)
            highlight_cells(sheet, city_index, product_index, percent)
            sleep(0.5)
        except ValueError:
            print(f"Товар {product} не найден в таблице")

# Инициализация клиента Vision API
client = vision.ImageAnnotatorClient()

# Функция для извлечения города, товаров и процентов
def extract_relevant_data(text):
    city_pattern = r'\b[A-Za-zА-Яа-яёЁ\-]+\b'
    product_pattern = r'([A-Za-zА-Яа-яёЁ]+)\s*Price\s*\d+\s*\((\d+%)\)'

    city = re.search(city_pattern, text.split('\n')[0])
    products = re.findall(product_pattern, text)

    return city.group(0) if city else "Город неизвестен", products

# Обработчик команды /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('Hey! I\'m your smart bot. Send me an image in English, and watch the magic happen in the spreadsheet!')

# Обработчик фотографий
async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo_file = await context.bot.get_file(update.message.photo[-1].file_id)
    photo_path = 'photo.jpg'
    await photo_file.download_to_drive(photo_path)

    with open(photo_path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)
    response = client.text_detection(image=image)
    texts = response.text_annotations

    if response.error.message:
        await update.message.reply_text(f'API error: {response.error.message}')
    else:
        detected_text = texts[0].description if texts else ""
        city, products = extract_relevant_data(detected_text)

        product_info = "\n".join([f'{product}: {percent}' for product, percent in products])
        message = f'City: {city}\nRecognized data:\n{product_info}' if products else "No data recognized"

        await update.message.reply_text(message)

        if not check_headers():
            create_table_headers()

        if city and products:
            export_to_sheets(city, products)

# Основная функция для запуска бота
async def main():
    application = ApplicationBuilder().token(TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.PHOTO, handle_photo))

    await application.initialize()
    await application.start()
    await application.updater.start_polling()
    await asyncio.sleep(float('inf'))

if __name__ == '__main__':
    try:
        asyncio.run(main())
    except RuntimeError as e:
        if str(e) == 'This event loop is already running':
            print("Error: Event loop is already running. Please run the bot in a different environment.")
        else:
            raise
