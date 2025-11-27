import os
import requests  # Для выполнения HTTP-запросов к API Яндекс Форм
import telebot  # Для создания Telegram бота
from docx import Document
from datetime import datetime
from dotenv import load_dotenv

# Загружаем переменные окружения из файла .env
load_dotenv()

# Инициализируем бота с токеном из переменных окружения
bot = telebot.TeleBot(os.getenv("BOT_TOKEN"))
# Получаем токен доступа к API Яндекс Форм
YANDEX = os.getenv("YANDEX_TOKEN")
# Получаем ID формы из переменных окружения
SURVEY_ID = os.getenv("SURVEY_ID")


def get_all_form_answers():
    # Формируем URL для запроса к API Яндекс Форм
    url = f"https://api.forms.yandex.net/v1/surveys/{SURVEY_ID}/answers"
    # Заголовки запроса с авторизацией через OAuth-токен
    headers = {"Authorization": f"OAuth {YANDEX}"}
    # Параметры запроса: количество ответов на странице и сортировка по дате
    params = {
        "page_size": 50,  # Максимальное количество ответов (макс. 50 по API)
        "sort": "-submitted_at"  # Сортировка от новых к старым
    }

    # Выполняем GET-запрос к API
    r = requests.get(url, headers=headers, params=params)
    # Проверяем статус ответа (вызовет исключение при ошибке)
    r.raise_for_status()
    # Парсим JSON-ответ
    data = r.json()
    # Возвращаем список ответов или пустой список если их нет
    return data.get("answers", [])


def extract_form_data(answer, columns):
    # Создаем пустой словарь для данных
    data = {}

    # Проверяем наличие необходимых данных в ответе
    if not answer or "data" not in answer:
        return data

    # Проходим по всем колонкам и извлекаем данные
    for i, column in enumerate(columns):
        # Проверяем, что для текущей колонки есть данные
        if i < len(answer["data"]):
            # Обрабатываем случай с пустым значением
            if answer["data"][i] is None:
                data[column["text"]] = None
                continue

            # Получаем название колонки (текст вопроса)
            column_name = column["text"]
            # Получаем значение ответа
            value = answer["data"][i].get("value")

            # Обрабатываем разные форматы данных
            if value is None:
                data[column_name] = None
            elif isinstance(value, list):
                # Для списка берем первый элемент если он один
                if len(value) == 1:
                    data[column_name] = value[0]
                else:
                    data[column_name] = value
            else:
                data[column_name] = value

    return data


def get_gender_forms(gender, applicant_type):
    if applicant_type == "Студент(ка)":
        return {
            "my": "Я",
            "sex": "",
            "absence_verb": "буду отсутствовать",
            "responsibility": "за освоение учебного материала беру на себя"
        }
    else:
        # Для родителей в зависимости от пола ребенка
        if gender == "Женский":
            return {
                "my": "Моя",
                "sex": "дочь",
                "absence_verb": "будет отсутствовать",
                "responsibility": "за сохранность жизни и здоровья ребенка в указанный период, а также за освоение учебной программы, беру на себя"
            }
        else:
            return {
                "my": "Мой",
                "sex": "сын",
                "absence_verb": "будет отсутствовать",
                "responsibility": "за сохранность жизни и здоровья ребенка в указанный период, а также за освоение учебной программы, беру на себя"
            }


def format_period(period_start, period_end):
    # Конвертируем дату в формат ДД.ММ.ГГГГ
    start_date = datetime.strptime(str(period_start), "%Y-%m-%d").strftime("%d.%m.%Y")

    # Проверяем, указан ли период или одна дата
    if period_end and str(period_end) != str(period_start):
        end_date = datetime.strptime(str(period_end), "%Y-%m-%d").strftime("%d.%m.%Y")
        return f"с {start_date} по {end_date}"  # Период
    else:
        return start_date  # Одна дата


def format_surname_genitive(surname, gender):
    if not surname:
        return ""

    # Пример: surname = "Иванова", gender = "Женский" - "Ивановой"
    if gender == "Женский":
        if surname.endswith('ая'):  # Для фамилий типа "Градская"
            return surname[:-2] + 'ой'  # "Градской"
        elif surname.endswith('а'):  # Для фамилий типа "Иванова"
            return surname[:-1] + 'ой'  # "Ивановой"
        elif surname.endswith('я'):
            return surname[:-1] + 'ей'

    # Для мужских
    return surname + 'а'  # Пример: "Иванов" → "Иванова"


def generate_doc(fio_student, group, gender_student, applicant_type, period_start, period_end, fio_applicant=None):
    # Получаем текущую дату для подписи документа
    current_date = datetime.now().strftime("%d.%m.%Y")
    # Форматируем период отсутствия
    period_text = format_period(period_start, period_end)

    # Выбираем шаблон в зависимости от типа заявителя
    if applicant_type == "Студент(ка)":
        template_file = "template_student.docx"  # Шаблон для студентов
    else:
        template_file = "template_parent.docx"  # Шаблон для родителей

    # Разбиваем ФИО заявителя на составляющие
    parts_applicant = fio_applicant.split()
    if len(parts_applicant) >= 3:
        surname_applicant = parts_applicant[0]  # Фамилия
        name_applicant = parts_applicant[1]  # Имя
        patronymic_applicant = parts_applicant[2]  # Отчество

        # Склоняем фамилию для подписи
        applicant_gender = gender_student
        surname_genitive = format_surname_genitive(surname_applicant, applicant_gender)
        # Формируем краткое ФИО в родительном падеже (Пример: "Ивановой А. И.")
        fio_short = f"{surname_genitive} {name_applicant[0]}. {patronymic_applicant[0]}."
    else:
        # Если ФИО в неправильном формате, используем как есть
        fio_short = fio_applicant

    # Получаем правильные грамматические формы
    gender_forms = get_gender_forms(gender_student, applicant_type)

    # Открываем выбранный шаблон документа
    doc = Document(template_file)

    # Заменяем на реальные данные
    for p in doc.paragraphs:
        p.text = p.text.replace("{fio}", str(fio_student))
        p.text = p.text.replace("{fio_short}", str(fio_short))
        p.text = p.text.replace("{group}", str(group))
        p.text = p.text.replace("{date}", str(period_text))
        p.text = p.text.replace("{current_date}", str(current_date))
        p.text = p.text.replace("{My}", str(gender_forms["my"]))
        p.text = p.text.replace("{sex}", str(gender_forms["sex"]))

    # Создаем имя файла на основе ФИО студента
    filename = f"Заявление_{fio_student.replace(' ', '_')}.docx"
    # Сохраняем документ
    doc.save(filename)
    return filename


# Команда /start
@bot.message_handler(commands=['start'])
def start(message):
    # Отправляем сообщение с инструкцией
    bot.send_message(message.chat.id,
                     "Здравствуйте! Напишите /get и номер студенческого билета чтобы получить заявление. Пример: /get 000892")


# Команда /get
@bot.message_handler(commands=['get'])
def get_by_ticket(message):
    # Разбиваем сообщение на части: ['/get', '000892']
    command_parts = message.text.split()
    if len(command_parts) < 2:
        bot.send_message(message.chat.id, "Пожалуйста, укажите номер студенческого билета. Пример: /get 000892")
        return

    # Извлекаем номер студенческого билета
    ticket_number = command_parts[1]
    bot.send_message(message.chat.id, f"Ищу заявление для номера {ticket_number}...")

    # Формируем запрос к API Яндекс Форм
    url = f"https://api.forms.yandex.net/v1/surveys/{SURVEY_ID}/answers"
    headers = {"Authorization": f"OAuth {YANDEX}"}
    params = {"page_size": 50, "sort": "-submitted_at"}

    # Выполняем запрос
    r = requests.get(url, headers=headers, params=params)
    r.raise_for_status()
    full_response = r.json()

    # Извлекаем данные о колонках и ответы
    columns = full_response.get("columns", [])
    answers = full_response.get("answers", [])

    # Проверяем наличие ответов
    if not answers:
        bot.send_message(message.chat.id, "Нет данных в ответе от формы!")
        return

    # Ищем ответ по номеру студенческого билета
    found_answer = None
    for answer in answers:
        form_data = extract_form_data(answer, columns)
        student_ticket = form_data.get("Ведите номер студенческого билета (пример: 000893)", "")

        if student_ticket == ticket_number:
            found_answer = answer
            break

    # Если не нашли - сообщаем об ошибке
    if not found_answer:
        bot.send_message(message.chat.id, f"Заявление с номером студенческого билета {ticket_number} не найдено.")
        return

    # Извлекаем данные из найденного ответа
    form_data = extract_form_data(found_answer, columns)

    # Получаем необходимые поля из данных формы
    fio_student = form_data.get("Укажите ФИО студента", "-")
    group = form_data.get("Группа студента (пример: 403ИС-22)", "-")
    gender_student = form_data.get("Укажите пол студента", "-")
    applicant_type = form_data.get("Я", "-")
    fio_applicant = form_data.get("Укажите ФИО заявителя", None)

    # Обрабатываем период отсутствия (может быть списком дат)
    period_data = form_data.get("Укажите период отсутствия", [])

    period_start = None
    period_end = None

    # Извлекаем даты начала и окончания периода
    if isinstance(period_data, list) and len(period_data) >= 2:
        period_start = period_data[0]
        period_end = period_data[1]
    elif isinstance(period_data, list) and len(period_data) == 1:
        period_start = period_data[0]
        period_end = period_data[0]

    # Проверяем обязательные поля
    if fio_student == "-" or group == "-":
        bot.send_message(message.chat.id, "Не удалось извлечь необходимые данные")
        return

    # Генерируем документ
    file_path = generate_doc(fio_student, group, gender_student, applicant_type, period_start, period_end,
                             fio_applicant)

    # Отправляем документ пользователю
    with open(file_path, "rb") as f:
        bot.send_document(message.chat.id, f)

    # Отправляем подтверждение
    bot.send_message(message.chat.id, f"Готово! Создано заявление для: {applicant_type}")


# Запускаем бота
if __name__ == "__main__":
    bot.polling(none_stop=True)  # Бесконечный цикл опроса серверов Telegram
