import re

import pandas as pd
import telebot

bot = telebot.TeleBot("6038351834:AAHrRDBUteglTmvBGPEQjN1mor8_8Bm1Blk")
file_path = "excel_docs/Анкеты потенциальных Франчайзи.xlsx"


@bot.message_handler(content_types="text", commands=["send"])
def send(message):
    global file_path
    bot.send_document(chat_id=message.chat.id, document=open(file_path, "rb"))


@bot.message_handler(content_types=["document"])
def receive(message):
    global file_path
    file_id = message.document.file_id
    bot.download_file(file_id, file_path)


@bot.message_handler(content_types="text")
def handle_text_message(message):
    global file_path
    data_string = message.text
    text = data_string

    def update_excel_file(file_path):
        # Read the Excel file using pandas
        df = pd.read_excel(file_path)
        # Add a new column "Номер" with increasing numbers
        number = len(df) + 1
        # Update the columns based on the dictionary
        new_row = {
            "Номер": number,
            "Дата": date,
            "Имя": name,
            "Город": city,
            "Область": None,
            "Номер телефона": phone,
            "Почта": email,
            "Адрес": adress,
            "Дата первого звонка": None,
            "Комментарий": None,
            "Дата встречи": None,
            "Документы передала в СБ": None,
        }
        df = df.append(new_row, ignore_index=True)
        # Save the updated data back to the Excel file
        df.to_excel(file_path, index=False)

    date = re.search(r"Отправлено: (.+?)\n", text).group(1)
    date = date if date else None
    name = re.search(r"Ваше_імя_прізвище_по_батькові_: (.+?)\n", text).group(1)
    name = name if name else None
    city = re.search(
        r"В_якому_місті_Ви_плануєте_відкрити_діагностичне_відділення_МЛ_ДІЛА_: (.*)",
        text,
    )
    city = city.group(1) if city else None
    phone = re.search(r"Phone: (.*)", text)
    phone = phone.group(1) if phone else None
    email = re.search(r"Email: (.*)", text)
    email = email.group(1) if email else None
    adress = re.search(
        r"Якщо_у_Вас_є_приміщення_в_якому_ви_бажаєте_розмістити_франчайзингове_відділення_вкажіть_повну_адресу: (.*)",
        text,
    )
    adress = adress.group(1) if adress else None
    update_excel_file(file_path)
    bot.send_message(chat_id=message.chat.id, text="есть")


if __name__ == "__main__":
    bot.polling()
