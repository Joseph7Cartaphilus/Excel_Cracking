import itertools
import win32com.client as client
import time
from datetime import datetime
from string import digits, punctuation, ascii_letters

symbols = digits + punctuation + ascii_letters
print(symbols)


def brute_excel_doc():
    print('***Hello Friend!***')

    try:
        password_length = input('Введите длину пароля, от скольки до скольки символом, например 3-7: ')
        password_length = [int(item) for item in password_length.split('-')]
    except:
        print('Проверьте введенные данные')

    print('Если пароль содержит только цифры, введите: 1\n'
          'Если пароль содержит только буквы, введите: 2\n'
          'Если пароль содержит цифры и буквы введите: 3\n'
          'Если пароль содержит цифры, буквы и спец символы введите: 4')

    try:
        choice = int(input(': '))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            possible_symbols = 'What`re you doing?'
    except:
        print('What`re you doing?')
    print(possible_symbols)

    # brute excel_doc
    start_timestamp = time.time()
    print(f"Started at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

    count = 0
    for pass_length in range(password_length[0], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = ''.join(password)
            print(password)

            opened_doc = client.Dispatch('Excel.Application')
            count += 1

            try:
                opened_doc.Workbooks.Open(
                    r'path to excel file',
                    False,
                    True,
                    None,
                    password
                )

                time.sleep(0.1)
                print(f"Finished at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
                print(f'Password cracking time  - {time.time() - start_timestamp}')

                return f'Attempt #{count} Password is {password}'
            except:
                print(f'Attempt #{count} Incorrect {password}')
                pass


def main():
    brute_excel_doc()


if __name__ == '__main__':
    main()
