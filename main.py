import psutil
import os, sys
import subprocess
import pywinauto as pw

from dotenv import load_dotenv

if getattr(sys, "frozen", False):
    script_dir = os.path.dirname(sys.executable)  # for exe version
else:
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
        print_stars()
        print("Процессы не запущены.")
        print_stars()
        show_commands()
    else:
        print_stars()
        print("Все процессы завершены.")
        print_stars()
        show_commands()


def list_process():
    process_counter: int = 0
    process_name_found: bool = False
    process_name = "Telegram.exe"
    folders = []
    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            if FOLDER in path:
                # print(path)
                process_counter += 1
                folder_name = os.path.basename(os.path.dirname(path))
                folders.append(int(folder_name))
                if not process_name_found:
                    process_name = proc.name()
                    process_name_found = True
                else:
                    continue

        except (psutil.NoSuchProcess, psutil.ZombieProcess, psutil.AccessDenied):
            continue
    if process_counter > 0:
        print_stars()
        print("Список запущенных процессов:")
        folders.sort()
        for k in range(len(folders)):

            process_path = os.path.join(FOLDER, str(folders[k]), process_name)
            print(process_path)
        print_stars()
        print("Запущено ", process_counter, process_name)
        print_stars()

    else:
        print_stars()
        print("Процессы не запущены.")
        print_stars()
    process_counter = 0

    # print(folders)

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
        print_stars()
        print("Процессы в указанном диапазоне не запущены.")
        print_stars()
        show_commands()
    else:
        print_stars()
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
    folders_counter = len(next(os.walk(FOLDER))[1])

    # print("Folder counter", folders_counter)

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

    print_stars()
    print("Запускаем процессы в выбранном диапазоне...")
    print_stars()
    show_commands()


def start_process():
    path = FOLDER
    exe = "Telegram.exe"

    folders_counter = len(next(os.walk(FOLDER))[1])
    # print("Folder counter", folders_counter)
    for i in range(1, folders_counter + 1):
        index = str(i)
        exe_path = os.path.join(path, index, exe)
        try:
            subprocess.Popen(exe_path)
        except FileNotFoundError:
            continue

    else:
        print_stars()
        print("Запускаем все процессы...")
        print_stars()
        show_commands()


def close_all_windows():
    apps = pw.Desktop(backend="win32").windows(
        class_name="Qt51517QWindowIcon", visible_only=True, top_level_only=True
    )
    # print(apps)
    if len(apps) == 0:
        print_stars()
        print("Активные окна не найдены.")

    else:
        for app in apps:
            if app.is_visible():
                app.close()
            else:
                continue

        print_stars()
        print("Закрываем все активные окна...")
    print_stars()
    show_commands()


def start_single_process():
    path = FOLDER
    exe = "Telegram.exe"

    print("Укажите номер процесса для запуска")

    try:
        index = int(input())
        if index <= 0:
            print("ОШИБКА: неверный номер процесса!")
            start_single_process()
        else:
            exe_path = os.path.join(path, str(index), exe)
            # print(exe_path)
            subprocess.Popen(exe_path)
            print_stars()
            print("Запускаем процесс под номером ", index, "...", sep="")
            print_stars()
            show_commands()
    except ValueError:
        print("ОШИБКА: указанное значение не является числом!")
        start_single_process()
    except FileNotFoundError:
        print("ОШИБКА: процесс с указанным номером не найден!")
        start_single_process()


def terminate_single_process():
    process_found: bool = False
    path = FOLDER

    print("Укажите номер процесса для завершения")
    try:
        index = int(input())
        if index <= 0:
            print("ОШИБКА: Неверный номер процесса!")
            terminate_single_process()
        else:
            for proc in psutil.process_iter():
                try:
                    path = proc.exe()
                    folder_name = os.path.basename(os.path.dirname(path))
                    # print(folder_name)
                    if FOLDER in path and int(folder_name) == index:
                        if proc.is_running():
                            proc.terminate()
                            process_found = True
                except (
                    psutil.AccessDenied,
                    psutil.NoSuchProcess,
                    psutil.ZombieProcess,
                ):
                    continue
    except ValueError:
        print("ОШИБКА: указанное значение не является числом!")
        terminate_single_process()
    if not process_found:
        print("ОШИБКА! Процесс с указанным номером не найден!")
        terminate_single_process()
    else:
        print_stars()
        print("Процесс под номером", index, "завершен")
        print_stars()
        show_commands()


def show_commands():
    print("Введите команду:")
    print("1. Запустить все процессы. ВНИМАНИЕ! УЧИТЫВАЙТЕ ХАРАКТЕРИСТИКИ СВОЕГО ПК!")
    print("2. Завершить все процессы")
    print("3. Запустить процессы выборочно")
    print("4. Завершить процессы выборочно")
    print("5. Показать список процессов")
    print("6. Закрыть все активные окна")
    print("7. Запустить выбранный процесс")
    print("8. Завершить выбранный процесс")
    print("0. Выход")
    input_command()


def input_command():
    cmd = input()
    match cmd:
        case "1":
            start_process()
        case "2":
            terminate_process()
        case "3":
            select_process_start_range()
        case "4":
            select_process_termination_range()
        case "5":
            list_process()
        case "6":
            close_all_windows()
        case "7":
            start_single_process()
        case "8":
            terminate_single_process()
        case "0":
            exit
        case _:
            print("Неверная команда!")
            print_stars()
            show_commands()


show_commands()
