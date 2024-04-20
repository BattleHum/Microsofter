import itertools
from string import digits, punctuation, ascii_letters
import win32com.client as client
from datetime import datetime
import time


def brute_excel_doc():
    try:
        app = input(r"Введіть шлях де розташований файл: ")
    except:
        print("Введи правильний шлях!")

    try:
        password_length = input("Введіть довжину паролю, наприклад 1-5: ")
        password_length = [int(item) for item in password_length.split("-")]
    except:
        print("Перевірте данні")

    print("Якщо пароль містить тільки цифри введіть: 1\nЯкщо пароль иістить тільки букви введіть: 2\n"
          "Якщо пароль містить і букви і цифри введіть: 3\nЯкщо пароль містить і цифри, і букви і спец.символи введіть: 4\n"
          "Якщо пароь містить цифри і спец.символи введіть: 5\nЯкщо пароль містить букви і спец.символи введіть: 6\n"
          )

    try:
        choice = int(input(": "))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        elif choice == 5:
            possible_symbols = digits + punctuation
        elif choice == 6:
            possible_symbols = ascii_letters + punctuation
        else:
            possible_symbols = "???"
    except:
        print("???")
    # brute excel doc
    start_timestamp = time.time()
    print(f"Started at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

    count = 0
    for pass_length in range(password_length[0], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)


            opened_doc = client.Dispatch("Excel.Application")
            count += 1

            try:
                opened_doc.Workbooks.Open(
                    app,
                    False,
                    True,
                    None,
                    password
                )

                time.sleep(0.1)
                print(f"Finished at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
                print(f"Password cracking time - {time.time() - start_timestamp}")

                return f"Attempt #{count} Password is: {password}"

            except:
                print(f"Attempt #{count} Incorrect password: {password}")
                pass

            with open(file="Password (Excel)", mode="a", encoding="utf-8") as file:
                 file.write(f'Password: {password}\n{"#" * 20}\n')
def main():
    print(brute_excel_doc())


if __name__ == '__main__':
    main()