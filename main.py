# Модуль автоматической работы со служебными записками по исправлению свода доходов в Самару
# Установить: pip install pywin32

# Импорты
import shutil
import glob
from ROI_base import *  # мой модуль для вывода времени

# Константы
NOW_DATE = datetime.now().strftime('%d.%m.%Y')  # текущая дата для работы в скрипте в формате 03.05.2021
NOTES_PATH = 'D:/Dropbox/Работа/Работа РОИ/+Выручка/!Корректировки в Самару/'
GENERATION_PATH = '+Генерация/'
SENDING_PATH = 'Отправленные/'


# Функции
def pause():
    a = input("\n Нажмите Enter для выхода...")
    print(a)


def search_file_attach() -> tuple:
    """Функция поиска файла для работы"""
    path0 = os.path.join(NOTES_PATH, GENERATION_PATH)
    files = glob.glob(path0 + f'*.pdf')  # f'{NOW_DATE}*.pdf'
    if not files:
        print_log('Нет файлов на оправку')
        pause()
        sys.exit()
    attach_file_pdf = files[0]
    _, file_name = attach_file_pdf.split('\\')
    return attach_file_pdf, file_name


def outlook_sending() -> None:
    """Функция отправки письма через MS Outlook"""
    print_log("Отправка файлов через MS Outlook")
    import win32com.client as win32  # импорт модуля для работы с Win32COM, pip install pywin32
    to_email = "slv1@rosinkas.ru; sev3@rosinkas.ru"  # основные получатели
    cc_email = "dsn2@rosinkas.ru; azd@rosinkas.ru; mev6@rosinkas.ru"  # получатели в копии
    attach_file_pdf = search_file_attach()[0]  # путь к файлу
    file_name = search_file_attach()[1]
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')  # вызов MS Outlook
    new_mail = outlook.CreateItem(0)  # создание письма в MS outlook
    new_mail.Subject = f"{file_name[:-4]}"
    new_mail.To = to_email  # обращение к списку получателей
    new_mail.CC = cc_email  # обращение к списку получателей в копии
    # new_mail.BodyFormat = 1  # формат PlainText
    new_mail.BodyFormat = 2  # формат HTML
    new_mail.Body = f"""Высылаю служебную записку '{file_name}'.
В САДД она уже прикреплена.
        
___________________
С уважением,
ведущий специалист РЦ "Ульяновск"
Митин Вячеслав Алексеевич, 8-902-004-27-98
"""
    print_log("Письмо для отправки через MS Outlook подготовлено", line_after=False)
    new_mail.Attachments.Add(Source=str(attach_file_pdf))  # присоединение вложения с файлом
    new_mail.Display(True)  # отображение подготовленного
    # new_mail.Send()  # немедленная отправка письма
    print_log(f"Письмо с файлом '{file_name}' отправлено", line_after=True)


def move_file_attach(file, name_file) -> None:
    """Функция перемещения отправленного файла"""
    path1 = os.path.join(NOTES_PATH, SENDING_PATH)
    print_log(f"Перемещение файла '{name_file}' в каталог '{path1}'")
    shutil.move(file, path1)


if __name__ == '__main__':
    print('модуль отправки служебных записок по корректировке свода доходов в самару'.upper() + '\n')
    search_file_attach()
    print_log(f"Файл для отправки '{search_file_attach()[1]}'", line_after=True)
    outlook_sending()
    move_file_attach(*search_file_attach())
    print_log(f"Работа модуля окончена", line_after=True)
    pause()
