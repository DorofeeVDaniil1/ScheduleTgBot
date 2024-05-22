# pip install --upgrade gspread
# pip install schedule
# pip install telebot
import os

import openpyxl
import telebot
from datetime import date
from datetime import datetime
from telebot import types
import time
import threading
import schedule
import gspread
from dateutil.relativedelta import relativedelta

token = "7192079410:AAHGHeUD1BXYDvzOiFTwI62HiMsHAQjJCxI"
bot = telebot.TeleBot(token)

# sa = gspread.service_account(filename="service_account.json")

from gspread import authorize
from oauth2client.service_account import ServiceAccountCredentials

scopes = ["https://spreadsheets.google.com/feeds",
          "https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive",
          "https://www.googleapis.com/auth/drive"]
cred = ServiceAccountCredentials.from_json_keyfile_name("D:/Irnitu/tg/telebot-423306-e9630311f084.json", scopes)
sa = authorize(cred)

sh = sa.open("tg_bot_dor")
wks = sh.worksheet("one")
headers = wks.row_values(1)


def days_between(d1, d2):
    return (d1 - d2).days


def draw_manager_menu():
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton("Отправить уведомление о сроках оплаты")
    btn2 = types.KeyboardButton("Показать информацию о клиентах")
    btn3 = types.KeyboardButton("Вывести данные в таблицу Excel")  # Новая кнопка
    btn4 = types.KeyboardButton("Проверить функциональность бота")
    markup.add(btn1, btn2, btn3, btn4)  # Добавляем новую кнопку
    return markup


def draw_user_menu():
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn1 = types.KeyboardButton("Узнать срок оплаты хостинга")
    btn2 = types.KeyboardButton("Связаться с администратором")
    btn3 = types.KeyboardButton("Узнать мой ID и ник")
    markup.add(btn1).add(btn2).add(btn3)
    return markup


@bot.message_handler(commands=['start'])
def start(message):
    if message.from_user.id == (523106911):
        markup = draw_manager_menu()
        bot.send_message(message.chat.id, text="Привет, {0.first_name}! Я твой бот помощник".format(message.from_user),
                         reply_markup=markup)
    else:
        markup = draw_user_menu()
        bot.send_message(message.chat.id,
                         text="Привет, {0.first_name}! Я бот, который напоминает об оплате хостинга".format(
                             message.from_user), reply_markup=markup)


day_token = lambda day_to: (
        (day_to in range(5, 20)) and 'дней' or
        (1 in (day_to, (diglast := day_to % 10))) and 'день' or
        ({day_to, diglast} & {2, 3, 4}) and 'дня' or 'дней')


def SendMsg(send=False):
    try:
        dataBase = wks.get_all_records()
        for user in dataBase:
            print(f"Processing user: {user}")  # Отладочный вывод

            dataOfUser = str(user['last_payment_date'])
            dataOfUser = datetime.strptime(dataOfUser, '%d.%m.%Y') + relativedelta(months=+1)
            today = datetime.strptime(str(date.today()), '%Y-%m-%d')
            betwens = days_between(dataOfUser, today)
            payment_sum = user['cost']

            print(f"dataOfUser: {dataOfUser}, today: {today}, betwens: {betwens}")  # Отладочный вывод

            if betwens >= 0:  # Убедитесь, что betwens не отрицательный
                if (betwens == 30 or (betwens <= 15 and betwens > 0 and betwens % 2 == 1)) or send:
                    if user['act/noact'] == 'act':
                        try:
                            bot.send_message(user['id_tg'],
                                             f'Через {betwens} {day_token(betwens)} вам необходимо оплатить хостинг на сумму: {payment_sum}')
                            print(f"Message sent to user {user['id_tg']}")  # Отладочный вывод
                        except Exception as e:
                            print(f"Failed to send message to user {user['id_tg']}: {e}")  # Обработка ошибки
                    else:
                        print("User is inactive, skipping.")  # Отладочный вывод
                        continue
                elif betwens > 30:
                    try:
                        bot.send_message(user['id_tg'], 'Вас скоро отключат от хостинга')
                        print(f"User {user['id_tg']} will be disconnected soon.")  # Отладочный вывод
                        rowToUpdate = wks.find(str(user['id_tg'])).row
                        colToUpdate = headers.index('act/noact') + 1
                        cellToUpdate = wks.cell(rowToUpdate, colToUpdate)
                        cellToUpdate.value = 'noact'
                        wks.update_cells([cellToUpdate])
                    except Exception as e:
                        print(f"Failed to send disconnection message to user {user['id_tg']}: {e}")  # Обработка ошибки
            else:
                print(f"Skipping user {user['id_tg']} because betwens is negative: {betwens}")  # Отладочный вывод

    except Exception as e:
        print(f"An error occurred: {e}")  # Отладочный вывод

# Обработчик нажатия на кнопку "Вывести данные в таблицу Excel"
@bot.message_handler(func=lambda message: message.text == "Вывести данные в таблицу Excel")
def export_to_excel(message):
    if message.from_user.id == 523106911:  # Проверяем, что пользователь - администратор
        ExportToExcel(message.chat.id)
    else:
        bot.send_message(message.chat.id, "У вас нет прав на выполнение этой команды.")


def GetInfo(admin):
    try:
        dataBase = wks.get_all_records()
        dataBase.sort(key=lambda k: datetime.strptime(k['last_payment_date'], '%d.%m.%Y'))

        for user in dataBase:
            dataOfUser = datetime.strptime(user['last_payment_date'], '%d.%m.%Y') + relativedelta(months=+1)
            today = datetime.strptime(str(date.today()), '%Y-%m-%d')
            betwens = days_between(dataOfUser, today)
            print(f"User: {user['id_client']}, betwens: {betwens}")  # Отладочный вывод
            if betwens <= 40 and betwens > 0:
                user_name = user['id_client']
                payment_sum = user['cost']
                phone = user['client_number']
                try:
                    bot.send_message(admin, f"У клиента {user_name} срок договора истекает через {betwens} {day_token(betwens)}\nНомер телефона клиента {phone}\nСумма оплаты: {payment_sum}")
                    print(f"Information sent for user {user_name}")  # Отладочный вывод
                except Exception as e:
                    print(f"Ошибка при отправке сообщения клиенту {user_name}: {e}")
                    continue
        bot.send_message(admin, "Информация о клиентах отправлена.")  # Сообщение администратору об успешной отправке
    except Exception as e:
        print(f"Произошла ошибка при получении информации о клиентах: {e}")

@bot.message_handler(func=lambda message: message.text == "Показать информацию о клиентах")
def show_client_info(message):
    if message.from_user.id == 523106911:
        print("Администратор запросил информацию о клиентах.")
        GetInfo(message.chat.id)
    else:
        bot.send_message(message.chat.id, "У вас нет прав на выполнение этой команды.")


def ExportToExcel(admin_id):
    try:
        # Получаем данные о клиентах
        dataBase = wks.get_all_records()
        dataBase.sort(key=lambda k: k['last_payment_date'])

        # Создаем новый файл Excel
        wb = openpyxl.Workbook()
        ws = wb.active

        # Добавляем заголовки
        headers = ['ID клиента', 'Последняя дата оплаты', 'Сумма оплаты', 'Номер телефона']
        ws.append(headers)

        # Записываем данные о клиентах в таблицу Excel
        for user in dataBase:
            user_data = [user['id_client'], user['last_payment_date'], user['cost'], user['client_number']]
            ws.append(user_data)

        # Сохраняем файл Excel
        excel_file_path = 'clients_info.xlsx'
        wb.save(excel_file_path)

        # Отправляем файл администратору
        with open(excel_file_path, 'rb') as file:
            bot.send_document(admin_id, file)

        # Удаляем временный файл Excel
        os.remove(excel_file_path)

        print("Excel файл успешно создан и отправлен.")

    except Exception as e:
        print("Произошла ошибка при создании и отправке Excel файла:", e)


# Обработчик нажатия на кнопку "Узнать мой ID и ник"
@bot.message_handler(func=lambda message: message.text == "Узнать мой ID и ник")
def get_user_id_and_username(message):
    user_id = message.from_user.id
    username = message.from_user.username if message.from_user.username else "ник не установлен"
    bot.send_message(message.chat.id, f"Ваш ID: {user_id}\nВаш ник: @{username}")


def GetMsg(user_id):
    # Ищем строку с данным user_id
    cell = wks.find(str(user_id))
    if cell is None:
        bot.send_message(user_id, "Ваш ID не найден в базе данных.")
        return

    # Получаем строку данных пользователя
    user = wks.row_values(cell.row)
    dataOfUser = str(user[3])
    dataOfUser = datetime.strptime(dataOfUser, '%d.%m.%Y') + relativedelta(months=+1)
    today = datetime.strptime(str(date.today()), '%Y-%m-%d')
    betwens = days_between(dataOfUser, today)
    payment_sum = user[4]
    bot.send_message(user_id,
                     f'Через {betwens} {day_token(betwens)} вам необходимо оплатить наш хостинг на сумму: {payment_sum}')


def check_bot_functionality(chat_id):
    try:
        wks.row_values(1)
        bot.send_message(chat_id, "Проверка Google Sheets: успешно.")
        bot.send_message(chat_id, "Тест отправки сообщения: успешно.")
        bot.send_message(chat_id, "Все системные проверки прошли успешно.")
    except Exception as e:
        bot.send_message(chat_id, f"Ошибка в системных проверках: {e}")


@bot.message_handler(content_types=['text'])
def mailing(message):
    if message.text == "Отправить уведомление о сроках оплаты":
        SendMsg(send=True)
    elif message.text == "Узнать срок оплаты хостинга":
        GetMsg(message.from_user.id)
    elif message.text == "Показать информацию о клиентах":
        show_client_info(message)  # Исправление вызова функции
    elif message.text == "Связаться с администратором":
        admin_username = 'dorofeev_daniil'
        bot.send_message(message.chat.id, f"Связаться с администратором можно по ссылке: t.me/{admin_username}")
    elif message.text == "Проверить функциональность бота":
        check_bot_functionality(message.chat.id)
    elif message.text == "Узнать мой ID и ник":
        get_user_id_and_username(message)
    else:
        bot.send_message(message.from_user.id, 'Неизвестная команда, попробуйте ещё раз.')


# bot.polling(none_stop=True)

def start_polling():
    bot.infinity_polling(none_stop=True)


polling_thread = threading.Thread(target=start_polling)
polling_thread.start()

schedule.every().day.at("23:54").do(SendMsg)

try:
    while True:
        schedule.run_pending()
        time.sleep(1)
except:
    pass
