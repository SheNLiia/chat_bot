import os
import requests
import telebot
from docx import Document
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

bot = telebot.TeleBot(os.getenv("BOT_TOKEN"))
YANDEX = os.getenv("YANDEX_TOKEN")

#JSON из Яндекс диска
def get_json():
    headers = {"Authorization": YANDEX}

    url = "https://cloud-api.yandex.net/v1/disk/resources/download"
    params = {"path": "Yandex.Forms/690df7f5068ff0fbd8626059/2025-11-16 КМПО.json"}

    r = requests.get(url, headers=headers, params=params)
    href = r.json()["href"]

    file = requests.get(href)
    return file.json()


#Преобразование строки в словарь
def parse_row(row):
    data = {}
    for item in row:
        key, value = item
        data[key] = value
    return data


#Создание DOCX
def generate_doc(fio, group):
    date_str = datetime.now().strftime("%d.%m.%Y")

    # Короткое ФИО
    parts = fio.split()
    if len(parts) >= 3:
        fio_short = f"{parts[0]} {parts[1][0]}. {parts[2][0]}."
    else:
        fio_short = fio

    doc = Document("template.docx")

    for p in doc.paragraphs:
        p.text = p.text.replace("{fio}", fio)
        p.text = p.text.replace("{fio_short}", fio_short)
        p.text = p.text.replace("{group}", group)
        p.text = p.text.replace("{date}", date_str)

    filename = "Заявление_студента.docx"
    doc.save(filename)
    return filename


# ---------------- TELEGRAM BOT ---------------- #

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Напиши /last чтобы получить заявление по последней записи формы.")


@bot.message_handler(commands=['last'])
def last(message):
    bot.send_message(message.chat.id, "Получаю данные...")

    raw = get_json()
    last_row = raw[-1]     # последняя запись
    data = parse_row(last_row)

    fio = data.get("ФИО студента", "-")
    group = data.get("Группа студента", "-")

    file_path = generate_doc(fio, group)

    with open(file_path, "rb") as f:
        bot.send_document(message.chat.id, f)

    bot.send_message(message.chat.id, "Готово!")


bot.polling(none_stop=True)
