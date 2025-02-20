import itertools
import random
import os
import time
from datetime import datetime
from string import digits, punctuation, ascii_letters
import win32com.client as client

# Словник для розпізнавання типу файлу
OFFICE_APPS = {
    "xlsx": "Excel.Application",
    "xls": "Excel.Application",
    "docx": "Word.Application",
    "doc": "Word.Application",
    "pptx": "PowerPoint.Application",
    "ppt": "PowerPoint.Application",
    "pst": "Outlook.Application",
    "ost": "Outlook.Application",
    "one": "OneNote.Application"
}

def detect_office_app(file_path):
    """Визначає, який додаток Office використовувати для файлу"""
    ext = file_path.split('.')[-1].lower()
    return OFFICE_APPS.get(ext)

def generate_passwords(symbols, min_length, max_length, random_mode=False):
    """Генерує паролі: за допомогою itertools або випадковим методом"""
    if random_mode:
        while True:
            length = random.randint(min_length, max_length)
            yield ''.join(random.choice(symbols) for _ in range(length))
    else:
        for length in range(min_length, max_length + 1):
            for password in itertools.product(symbols, repeat=length):
                yield "".join(password)

def brute_force_office():
    """Функція підбору паролю для офісних файлів"""
    try:
        file_path = input(r"Введіть шлях до файлу: ").strip()
        if not os.path.exists(file_path):
            print("Файл не знайдено! Переконайтеся, що шлях правильний.")
            return
    except Exception as e:
        print(f"Помилка: {e}")
        return

    office_app = detect_office_app(file_path)
    if not office_app:
        print("Непідтримуваний тип файлу.")
        return

    try:
        password_length = input("Введіть довжину паролю (наприклад 1-5): ").strip()
        min_length, max_length = map(int, password_length.split("-"))
        if min_length > max_length:
            print("Неправильний формат довжини паролю.")
            return
    except ValueError:
        print("Перевірте формат вводу (наприклад: 1-5).")
        return

    print(
        "Виберіть формат паролю:\n"
        "1 - тільки цифри\n"
        "2 - тільки букви\n"
        "3 - букви + цифри\n"
        "4 - букви + цифри + спец. символи\n"
        "5 - цифри + спец. символи\n"
        "6 - букви + спец. символи\n"
        "random - випадковий набір символів\n"
        "random_full - повністю випадкові паролі (різна довжина + змішані символи)\n"
    )

    choice = input(": ").strip()

    symbol_sets = {
        "1": digits,
        "2": ascii_letters,
        "3": digits + ascii_letters,
        "4": digits + ascii_letters + punctuation,
        "5": digits + punctuation,
        "6": ascii_letters + punctuation,
    }

    random_mode = False
    if choice in symbol_sets:
        possible_symbols = symbol_sets[choice]
    elif choice == "random":
        possible_symbols = random.choice([digits, ascii_letters, digits + ascii_letters, digits + punctuation])
    elif choice == "random_full":
        possible_symbols = digits + ascii_letters + punctuation
        random_mode = True
    else:
        print("Невідомий вибір!")
        return

    # Початок перебору
    start_timestamp = time.time()
    print(f"Початок роботи: {datetime.now().strftime('%H:%M:%S')}")

    count = 0
    office_instance = client.Dispatch(office_app)

    for password in generate_passwords(possible_symbols, min_length, max_length, random_mode):
        count += 1

        try:
            if "Excel" in office_app:
                office_instance.Workbooks.Open(file_path, False, True, None, password)
            elif "Word" in office_app:
                office_instance.Documents.Open(file_path, False, True, None, password)
            elif "PowerPoint" in office_app:
                office_instance.Presentations.Open(file_path, False, True, None, password)
            elif "Outlook" in office_app:
                office_instance.Session.Logon(file_path, password, False, True)
            elif "OneNote" in office_app:
                # OneNote не підтримує пряме відкриття за паролем, тут потрібно API Microsoft Graph
                print("OneNote захищені паролем файли потребують спеціального API.")
                return

            print(f"Пароль знайдено: {password}")
            print(f"Час: {time.time() - start_timestamp:.2f} секунд")
            return f"Спроба #{count}, пароль: {password}"

        except:
            print(f"Спроба #{count}, неправильний пароль: {password}")

        # Логування в файл
        with open("passwords_log.txt", "a", encoding="utf-8") as file:
            file.write(f'Спроба #{count} | Пароль: {password}\n')

    print("Перебір завершено. Пароль не знайдено.")

def main():
    print(brute_force_office())

if __name__ == '__main__':
    main()
