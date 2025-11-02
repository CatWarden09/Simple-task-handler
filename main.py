import psutil
import os, sys
import subprocess
from dotenv import load_dotenv

# pywin32, subprocess


# script_dir = os.path.dirname(sys.executable) # for final version
script_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(script_dir, ".env")


def create_env_file():
    print("Укажите путь к папке с копиями программы")
    proc_folder = input()
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("FOLDER=" + str(proc_folder))


try:
    env_file = open(env_path, "r")
except FileNotFoundError:
    create_env_file()

load_dotenv()
FOLDER = os.getenv("FOLDER")


def print_stars():
    stars_count = 15
    print("*" * stars_count)


def terminate_process():
    process_found: bool = False
    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            if FOLDER in path:
                process_found = True
                if proc.is_running():
                    proc.terminate()
        except (psutil.NoSuchProcess, psutil.ZombieProcess, psutil.AccessDenied):
            continue
    if not process_found:
        print("Процессы не запущены.")
        print_stars()
        show_commands()
    else:
        print("Все процессы завершены.")
        print_stars()
        show_commands()


def list_process():
    process_counter: int = 0
    process_name_found: bool = False
    process_name = ""
    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            if FOLDER in path:
                print(path)
                process_counter += 1

                if not process_name_found:
                    process_name = proc.name()
                    process_name_found = True
                else:
                    continue

        except (psutil.NoSuchProcess, psutil.ZombieProcess, psutil.AccessDenied):
            continue
    if process_counter > 0:
        print_stars()
        print("Запущено ", process_counter, process_name)
        print_stars()

    else:
        print_stars()
        print("Процессы не запущены.")
        print_stars()
    process_counter = 0
    show_commands()


def select_range_start():
    print("Укажите начало диапазона")

    try:
        range_start = int(input())
        if range_start == 0 or range_start < 0:
            print("ОШИБКА: Неверное начало диапазона!")
            return select_range_start()
        return range_start
    except ValueError:
        print("ОШИБКА: Начало диапазона не является числом!")
        return select_range_start()


def select_range_end(range_start):
    print("Укажите конец диапазона")
    print("Начало диапазона:", range_start)
    try:
        range_end = int(input())
        if range_end <= 0:
            print("ОШИБКА: Неверный конец диапазона!")
            return select_range_end(range_start)
        elif range_end < range_start:
            print("ОШИБКА: Конец диапазона меньше начала диапазона!")
            return select_range_end(range_start)
        return range_end
    except ValueError:
        print("ОШИБКА: Конец диапазона не является числом!")
        return select_range_end(range_start)


def select_process_termination_range():
    range_start = select_range_start()
    range_end = select_range_end(range_start)
    terminate_selected_process(range_start, range_end)


def terminate_selected_process(range_start, range_end):
    process_found: bool = False
    terminate_range = []
    for i in range(range_start, range_end + 1):
        terminate_range.append(i)
    # print("Массив диапазона процессов:", terminate_range)
    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            folder_name = os.path.basename(os.path.dirname(path))
            if FOLDER in path and int(folder_name) in terminate_range:
                process_found = True
                if proc.is_running():
                    proc.terminate()

        except (psutil.NoSuchProcess, psutil.ZombieProcess, psutil.AccessDenied):
            continue
    if not process_found:
        print("Процессы в указанном диапазоне не запущены.")
        print_stars()
        show_commands()
    else:
        print("Процессы в выбранном диапазоне завершены.")
        print_stars()
        show_commands()


def select_process_start_range():
    range_start = select_range_start()
    range_end = select_range_end(range_start)
    start_selected_process(range_start, range_end)


def start_selected_process(range_start, range_end):
    path = FOLDER
    exe = "Telegram.exe"
    start_range = []
    folders_counter = len(next(os.walk(FOLDER))[1]) - 1  # - 1 for main account

    print("Folder counter", folders_counter)

    for i in range(range_start, range_end + 1):
        start_range.append(i)

    for k in range(1, folders_counter + 1):
        index = str(k)
        exe_path = os.path.join(path, index, exe)
        folder_name = os.path.basename(os.path.dirname(exe_path))
        try:
            if int(folder_name) in start_range:
                subprocess.Popen(exe_path)
        except FileNotFoundError:
            continue


def start_process():
    path = FOLDER
    exe = "Telegram.exe"

    folders_counter = len(next(os.walk(FOLDER))[1]) - 1  # - 1 for main account
    print("Folder counter", folders_counter)
    for i in range(1, folders_counter + 1):
        index = str(i)
        exe_path = os.path.join(path, index, exe)
        try:
            subprocess.Popen(exe_path)
        except FileNotFoundError:
            continue

    else:
        print("Запускаем все процессы...")
        print_stars()
        show_commands()


def show_commands():
    print("Введите команду:")
    print("1. Завершить все процессы")
    print("2. Показать список процессов")
    print("3. Завершить процессы выборочно")
    print("4. Запустить все процессы")
    print("5. Запустить процессы выборочно")
    print("6. Закрыть все активные окна TG🚧")
    print("0. Выход")
    input_command()


def input_command():
    cmd = input()
    match cmd:
        case "1":
            terminate_process()
        case "2":
            list_process()
        case "3":
            select_process_termination_range()
        case "4":
            start_process()
        case "5":
            select_process_start_range()
        case "0":
            exit
        case _:
            print("Неверная команда!")
            print_stars()
            show_commands()


show_commands()
