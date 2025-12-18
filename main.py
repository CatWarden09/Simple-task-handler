import psutil
import os, sys
import subprocess
import pywinauto as pw

from dotenv import load_dotenv
from pywinauto import timings

VERSION = "0.4.1"

if getattr(sys, "frozen", False):
    script_dir = os.path.dirname(sys.executable)  # for exe version
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

env_path = os.path.join(script_dir, ".env")


def create_env_file():
    print("Укажите путь к папке с копиями программы")
    proc_folder = input()
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("FOLDER=" + str(proc_folder) + "\n")
    print("Укажите метку для папок-исключений")
    skip_mark_input = input()
    with open(env_path, "a", encoding="utf-8") as f:
        f.write("SKIP_MARK=" + str(skip_mark_input) + "\n")
    print("Укажите название .exe-файла программы (например, Telegram.exe)")
    proc_exe = input()
    with open(env_path, "a", encoding="utf-8") as f:
        f.write("EXE=" + str(proc_exe) + "\n")


try:
    env_file = open(env_path, "r")
except FileNotFoundError:
    create_env_file()

load_dotenv()
FOLDER = os.getenv("FOLDER")
SKIP_MARK = os.getenv("SKIP_MARK")
EXE = os.getenv("EXE")


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
    else:
        print_stars()
        print("Все процессы завершены.")
        print_stars()


def list_process():
    process_counter: int = 0
    process_name_found: bool = False
    process_name = EXE
    folders = []
    skipped_folders = []
    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            if FOLDER in path:
                # print(path)
                process_counter += 1
                folder_name = os.path.basename(os.path.dirname(path))
                if SKIP_MARK in folder_name:
                    # создаем отдельный массив для исключаемых папок, т.к. нельзя смешивать сортировку str и int
                    skipped_folders.append(folder_name)
                    continue
                folders.append(int(folder_name))
                # перевод в int нужен для последующей сортировки массива и красивого отображения списка процессов по возрастанию номеров
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

        new_skipped_folders = []
        # создаем новый массив и добавляем туда индексы исключаемых папок для последующей сортировки
        for fold in skipped_folders:

            cleared_folder_name = fold.replace(str(SKIP_MARK), "")

            new_folder_name = int(cleared_folder_name)

            new_skipped_folders.append(new_folder_name)

        new_skipped_folders.sort()
        for k in range(len(folders)):
            process_path = os.path.join(FOLDER, str(folders[k]), process_name)
            print(process_path)
        for l in range(len(new_skipped_folders)):
            process_path = os.path.join(
                FOLDER, str(new_skipped_folders[l]) + SKIP_MARK, process_name
            )
            print(process_path)
        print_stars()
        print("Запущено ", process_counter, process_name)
        print_stars()

    else:
        print_stars()
        print("Процессы не запущены.")
        print_stars()


def select_range_start():
    while True:

        print("Укажите начало диапазона")

        try:
            range_start = int(input())
            if range_start == 0 or range_start < 0:
                print("ОШИБКА: Неверное начало диапазона!")
                continue
            return range_start
        except ValueError:
            print("ОШИБКА: Начало диапазона не является числом!")
            continue


def select_range_end(range_start):
    while True:
        print("Укажите конец диапазона")
        print("Начало диапазона:", range_start)
        try:
            range_end = int(input())
            if range_end <= 0:
                print("ОШИБКА: Неверный конец диапазона!")
                continue
            elif range_end < range_start:
                print("ОШИБКА: Конец диапазона меньше начала диапазона!")
                continue
            return range_end
        except ValueError:
            print("ОШИБКА: Конец диапазона не является числом!")
            continue


def select_process_termination_range():
    range_start = select_range_start()
    range_end = select_range_end(range_start)
    terminate_selected_process(range_start, range_end)


def terminate_selected_process(range_start, range_end):
    process_found: bool = False
    terminate_range = []
    for i in range(range_start, range_end + 1):
        terminate_range.append(i)

    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            folder_name = os.path.basename(os.path.dirname(path))

            folder_name = (
                folder_name.replace(SKIP_MARK, "")
                if SKIP_MARK in folder_name
                else folder_name
            )

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
    else:
        print_stars()
        print("Процессы в выбранном диапазоне завершены.")
        print_stars()


def select_process_start_range():
    range_start = select_range_start()
    range_end = select_range_end(range_start)
    start_selected_process(range_start, range_end)


def start_selected_process(range_start, range_end):
    path = FOLDER
    exe = EXE
    start_range = []
    folders = next(os.walk(path))[1]

    for i in range(range_start, range_end + 1):
        start_range.append(i)

    for folder in folders:
        exe_path = os.path.join(path, folder, exe)
        try:
            if not SKIP_MARK in folder and int(folder) in start_range:
                subprocess.Popen(exe_path)
        except FileNotFoundError:
            continue

    print_stars()
    print("Запускаем процессы в выбранном диапазоне...")
    print_stars()


def start_process():
    path = FOLDER
    exe = EXE

    folders = next(os.walk(FOLDER))[1]

    for folder in folders:
        exe_path = os.path.join(path, folder, exe)
        try:
            if not SKIP_MARK in folder:
                subprocess.Popen(exe_path)
        except FileNotFoundError:
            continue

    print_stars()
    print("Запускаем все процессы...")
    print_stars()


def close_all_windows():
    is_timeout: bool = False
    apps = pw.Desktop(backend="win32").windows(
        class_name="Qt51517QWindowIcon", visible_only=True, top_level_only=True
    )

    if len(apps) == 0:
        print_stars()
        print("Активные окна не найдены.")

    else:
        for app in apps:
            try:
                if app.is_visible():
                    app.close()
            except timings.TimeoutError:
                is_timeout = True
                continue
            else:
                continue

        print_stars()
        print("Закрываем все активные окна...")
        if is_timeout:
            print("______________________________________________________")
            print("")
            print("Внимание! Некоторые окна не были закрыты из-за ошибки!")
            print("Возможно, программа не может получить доступ к окну.")
            print("______________________________________________________")
    print_stars()
    print("Закрываем все активные окна...")
    print_stars()


def start_single_process():
    while True:
        path = FOLDER
        exe = EXE

        print("Укажите номер процесса для запуска")

        try:
            index = int(input())
            if index <= 0:
                print("ОШИБКА: неверный номер процесса!")
                continue
            else:
                exe_path = os.path.join(path, str(index), exe)

                subprocess.Popen(exe_path)
                print_stars()
                print("Запускаем процесс под номером ", index, "...", sep="")
                print_stars()
                break
        except ValueError:
            print("ОШИБКА: указанное значение не является числом!")
            continue
        except FileNotFoundError:
            print("ОШИБКА: процесс с указанным номером не найден!")
            continue


def terminate_single_process():
    while True:
        process_found: bool = False
        path = FOLDER

        print("Укажите номер процесса для завершения")
        try:

            index = int(input())
            if index <= 0:
                print("ОШИБКА: Неверный номер процесса!")
                continue
            else:
                for proc in psutil.process_iter():
                    try:
                        path = proc.exe()
                        folder_name = os.path.basename(os.path.dirname(path))
                        folder_name = (
                            folder_name.replace(SKIP_MARK, "")
                            if SKIP_MARK in folder_name
                            else folder_name
                        )

                        # перевод названия родительской папки в int здесь нужен для того, чтобы не закрывались все копии с определенным числом в названии (например, 1 для 11)
                        if FOLDER in path and int(folder_name) == index:
                            if proc.is_running():
                                proc.terminate()
                                process_found = True
                    except (
                        psutil.AccessDenied,
                        psutil.NoSuchProcess,
                        psutil.ZombieProcess,
                        ValueError,
                    ):
                        continue
        except ValueError:
            print("ОШИБКА: указанное значение не является числом!")
            continue
        if not process_found:
            print("ОШИБКА! Процесс с указанным номером не найден!")
            continue
        else:
            print_stars()
            print("Процесс под номером", index, "завершен")
            print_stars()
            break


def start_skipped_process():
    path = FOLDER
    exe = EXE
    skip_mark = SKIP_MARK.replace("(", "").replace(")", "")

    folders = next(os.walk(FOLDER))[1]

    for folder in folders:
        exe_path = os.path.join(path, folder, exe)
        if SKIP_MARK in exe_path:
            try:
                subprocess.Popen(exe_path)
            except FileNotFoundError:
                continue

    print_stars()
    print(f"Запускаем {skip_mark}-процессы ...")
    print_stars()


def terminate_skipped_process():
    skip_mark = SKIP_MARK.replace("(", "").replace(")", "")

    process_found: bool = False
    for proc in psutil.process_iter():
        try:
            path = proc.exe()
            if FOLDER in path and SKIP_MARK in path:
                process_found = True
                if proc.is_running():
                    proc.terminate()
        except (psutil.NoSuchProcess, psutil.ZombieProcess, psutil.AccessDenied):
            continue
    if not process_found:
        print_stars()
        print(f"{skip_mark}-процессы не запущены.")
        print_stars()
    else:
        print_stars()
        print(f"{skip_mark}-процессы завершены.")
        print_stars()


def count_skipped_folder():
    skip_mark = SKIP_MARK.replace("(", "").replace(")", "")

    skipped_counter = 0
    folders = next(os.walk(FOLDER))[1]
    for folder in folders:
        if SKIP_MARK in folder:
            skipped_counter += 1
        else:
            continue
    else:
        print_stars()
        print(f"Количество {skip_mark}-папок: {skipped_counter}")
        print_stars()


def show_commands():
    skip_mark = SKIP_MARK.replace("(", "").replace(")", "")

    print("Введите команду:")
    print("1. Запустить процессы. ВНИМАНИЕ! УЧИТЫВАЙТЕ ХАРАКТЕРИСТИКИ СВОЕГО ПК!")
    print("2. Завершить процессы")
    print("3. Запустить процессы выборочно")
    print("4. Завершить процессы выборочно")
    print("5. Показать список процессов")
    print("6. Закрыть все активные окна")
    print("7. Запустить выбранный процесс")
    print("8. Завершить выбранный процесс")
    print(f"9. Запустить {skip_mark}-процессы")
    print(f"10. Завершить {skip_mark}-процессы")
    print(f"11. Посчитать {skip_mark}-папки")
    print("0. Выход")


def input_command():
    while True:
        show_commands()
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

            case "9":
                start_skipped_process()

            case "10":
                terminate_skipped_process()

            case "11":
                count_skipped_folder()

            case "0":
                sys.exit(0)
            case _:
                print("Неверная команда!")
                print_stars()


print("|||||||||||||||||||||||||||||||||||||||")
print("Simple Task Handler v", VERSION, sep="")
input_command()
