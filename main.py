import configparser
import ctypes
import json
import os
import sys
import threading
import time
import tkinter as tk
import winreg
from datetime import datetime

import pystray
import schedule
from PIL import Image
from playwright.sync_api import sync_playwright
from pystray import MenuItem as item

APP_NAME = "IdleScholar"
APP_PATH = (
    f'"{sys.executable}"'
    if getattr(sys, "frozen", False)
    else f'"{os.path.abspath(__file__)}"'
)
MUTEX = "Global\\D1m7.IdleScholar"
SCHEDULE_FILE = "schedule.json"
WEEK_DAYS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
]
CONFIG_FILE = "settings.cfg"
global USER_NAME, CHROME_PATH

if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))


def check_single_instance():
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    kernel32.CreateMutexW(None, False, MUTEX)

    if kernel32.GetLastError() == 183:
        return False

    return True


def open_link(url):
    with sync_playwright() as p:
        user_data_dir = os.path.join(os.getcwd(), "chrome_user_data")
        browser = p.chromium.launch_persistent_context(
            user_data_dir,
            executable_path=CHROME_PATH,
            headless=False,
            args=[
                "--use-fake-ui-for-media-stream",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-extensions",
                "--disable-features=IsolateOrigins,site-per-process",
            ],
        )
        page = browser.new_page()
        page.add_init_script(
            """
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
        Object.defineProperty(navigator, 'platform', { get: () => 'Win32' });
        Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
        """
        )
        page.goto(url)
        page.wait_for_load_state("networkidle")

        input_name_selector = "#name"
        if page.locator(input_name_selector).is_visible():
            page.fill(input_name_selector, USER_NAME)
            page.keyboard.press("Enter")
        else:
            autoplay_button_selector = ".autoplay-video-allow-btn"
            if page.locator(autoplay_button_selector).is_visible():
                page.click(autoplay_button_selector)

        continue_button_selector = 'button:has-text("продолжить без микрофона")'
        if page.locator(continue_button_selector).is_visible():
            page.click(continue_button_selector)

        start_time = time.time()
        while time.time() - start_time < 2 * 60 * 60:
            page.mouse.move(100, 100)
            page.mouse.move(110, 110)
            time.sleep(300)

        browser.close()


def add_schedule(event=None):
    url = url_entry.get()
    time_str = time_entry.get().replace(" ", ":")
    day = day_var.get()

    if not url or not time_str or not day:
        return

    time_hours = time_str.split(":")[0]
    time_minutes = time_str.split(":")[1]

    if len(time_minutes) < 2:
        time_str = time_hours + ":0" + time_minutes

    if len(time_hours) < 2:
        time_str = "0" + time_str

    week_schedule = {
        "Monday": schedule.every().monday,
        "Tuesday": schedule.every().tuesday,
        "Wednesday": schedule.every().wednesday,
        "Thursday": schedule.every().thursday,
        "Friday": schedule.every().friday,
        "Saturday": schedule.every().saturday,
        "Sunday": schedule.every().sunday,
    }

    schedule_time = week_schedule.get(day).at(time_str)
    schedule_time.do(open_link, url)

    save_schedule()
    update_tasks_for_day()
    clear_entries()


def delete_schedule():
    selected_index = task_list.curselection()
    if not selected_index:
        return

    selected_task = task_list.get(selected_index)
    for job in schedule.get_jobs():
        task_time = job.at_time.strftime("%H:%M")
        task_url = job.job_func.args[0]
        task_str = f"{task_time} - {task_url}"
        if task_str == selected_task:
            schedule.cancel_job(job)
            break

    task_list.delete(selected_index)
    save_schedule()
    clear_entries()


def edit_schedule(event):
    if url_entry.get() or time_entry.get():
        return

    selected_index = task_list.curselection()
    if not selected_index:
        return

    selected_task = task_list.get(selected_index)
    time_str, url = selected_task.split(" - ")

    # Устанавливаем выбранные значения в поля ввода
    time_entry.delete(0, tk.END)
    time_entry.insert(0, time_str)

    url_entry.delete(0, tk.END)
    url_entry.insert(0, url)

    # Удаляем выбранную задачу из списка и планировщика
    task_list.delete(selected_index)
    for job in schedule.get_jobs():
        task_time = job.at_time.strftime("%H:%M")
        task_url = job.job_func.args[0]
        if task_time == time_str and task_url == url:
            schedule.cancel_job(job)
            break

    save_schedule()


def clear_entries():
    time_entry.delete(0, tk.END)
    url_entry.delete(0, tk.END)


def save_schedule():
    tasks = []
    for job in schedule.get_jobs():
        task = {
            "week_day": WEEK_DAYS[job.next_run.weekday()],
            "time": job.at_time.strftime("%H:%M"),
            "url": job.job_func.args[0],
        }
        tasks.append(task)

    if tasks:
        with open(SCHEDULE_FILE, "w") as f:
            json.dump(tasks, f, indent=4)


def load_schedule():
    schedule.clear()
    week_schedule = {
        "Monday": schedule.every().monday,
        "Tuesday": schedule.every().tuesday,
        "Wednesday": schedule.every().wednesday,
        "Thursday": schedule.every().thursday,
        "Friday": schedule.every().friday,
        "Saturday": schedule.every().saturday,
        "Sunday": schedule.every().sunday,
    }

    if os.path.exists(SCHEDULE_FILE):
        with open(SCHEDULE_FILE, "r") as f:
            tasks = json.load(f)
            for task in tasks:
                week_schedule = {
                    "Monday": schedule.every().monday,
                    "Tuesday": schedule.every().tuesday,
                    "Wednesday": schedule.every().wednesday,
                    "Thursday": schedule.every().thursday,
                    "Friday": schedule.every().friday,
                    "Saturday": schedule.every().saturday,
                    "Sunday": schedule.every().sunday,
                }

                day = task["week_day"]
                time_str = task["time"]
                url = task["url"]
                schedule_time = week_schedule[day].at(time_str)
                schedule_time.do(open_link, url)

    today = datetime.today().strftime("%A")
    day_var.set(today)
    update_tasks_for_day()


def update_tasks_for_day(*args):
    selected_day = day_var.get()
    day_tasks = [
        job
        for job in schedule.get_jobs()
        if WEEK_DAYS[job.next_run.weekday()] == selected_day
    ]
    task_list.delete(0, tk.END)
    for task in day_tasks:
        task_time = task.at_time.strftime("%H:%M")
        task_url = task.job_func.args[0]
        task_list.insert(tk.END, f"{task_time} - {task_url}")


def load_settings():
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
        user_name = config.get("Settings", "user_name", fallback="student")
        chrome_path = config.get(
            "Settings",
            "chrome_path",
            fallback="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        )
    else:
        user_name = "student"
        chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

    global USER_NAME, CHROME_PATH
    USER_NAME, CHROME_PATH = user_name, chrome_path


def save_settings(user_name, chrome_path):
    config = configparser.ConfigParser()
    config["Settings"] = {"user_name": user_name, "chrome_path": chrome_path}
    with open(CONFIG_FILE, "w") as cfg:
        config.write(cfg)

    load_settings()


def is_autostart_enabled():
    """Проверяет, включён ли автозапуск."""
    try:
        key = winreg.HKEY_CURRENT_USER
        subkey = r"Software\Microsoft\Windows\CurrentVersion\Run"

        with winreg.OpenKey(key, subkey, 0, winreg.KEY_READ) as reg_key:
            winreg.QueryValueEx(reg_key, APP_NAME)
        return True
    except FileNotFoundError:
        return False


def enable_autostart(is_on_startup):
    key = winreg.HKEY_CURRENT_USER
    subkey = r"Software\Microsoft\Windows\CurrentVersion\Run"

    with winreg.OpenKey(key, subkey, 0, winreg.KEY_SET_VALUE) as reg_key:
        if is_on_startup:
            winreg.SetValueEx(reg_key, APP_NAME, 0, winreg.REG_SZ, APP_PATH)
        else:
            try:
                winreg.DeleteValue(reg_key, APP_NAME)
            except FileNotFoundError:
                pass


def open_settings():
    def on_ok():
        user_name = entry1.get()
        chrome_path = entry2.get()
        save_settings(user_name, chrome_path)
        settings_window.destroy()

    def on_cancel():
        settings_window.destroy()

    settings_window = tk.Toplevel(root)
    settings_window.resizable(False, False)
    settings_window.title("Настройки")
    settings_window.transient(root)
    settings_window.grab_set()

    settings_window.geometry(f"+{root.winfo_rootx()}+{root.winfo_rooty()}")

    tk.Label(settings_window, text="Имя:").grid(row=0, column=0, padx=10, pady=5)
    entry1 = tk.Entry(settings_window)
    entry1.insert(0, USER_NAME)
    entry1.grid(row=0, column=1, padx=20, pady=5)

    tk.Label(settings_window, text="Путь к Chrome:").grid(
        row=1, column=0, padx=10, pady=5
    )
    entry2 = tk.Entry(settings_window)
    entry2.insert(0, CHROME_PATH)
    entry2.grid(row=1, column=1, padx=10, pady=5)

    is_on_startup = tk.BooleanVar()
    is_on_startup.set(is_autostart_enabled())
    startup_check = tk.Checkbutton(
        settings_window,
        text="Запуск вместе с Windows",
        variable=is_on_startup,
        command=enable_autostart(is_on_startup),
    )
    startup_check.grid(row=2, column=1, padx=10, pady=5)

    ok_button = tk.Button(settings_window, text="OK", command=on_ok)
    ok_button.grid(row=3, column=0, padx=10, pady=10)

    cancel_button = tk.Button(settings_window, text="Отмена", command=on_cancel)
    cancel_button.grid(row=3, column=1, padx=10, pady=10)

    settings_window.focus_set()
    settings_window.wait_window()


def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)


def hide_window():
    root.withdraw()
    create_tray_icon()


def show_window():
    app_icon.stop()
    root.deiconify()


def create_tray_icon():
    image = Image.open("icon.png")
    menu = (
        item("Show", lambda: show_window(), default=True),
        item("Exit", lambda: on_exit()),
    )
    global app_icon
    app_icon = pystray.Icon("schedule_app", image, "Schedule App", menu)
    threading.Thread(target=app_icon.run, daemon=True).start()


def on_exit():
    app_icon.stop()
    root.destroy()


if __name__ == "__main__":
    if not check_single_instance():
        print("Программа уже запущена.")
        sys.exit(0)

    print("Загрузка настроек..")
    load_settings()

    # GUI
    print("Загрузка вёрстки..")
    root = tk.Tk()
    root.title("Расписание занятий")
    root.resizable(False, False)

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    days = [
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
    ]
    day_var = tk.StringVar(value=days[0])

    settings_button = tk.Button(
        frame, text="⚙", command=open_settings, justify="left", anchor="w"
    )
    settings_button.grid(sticky="W", column=0, row=0)

    tk.Label(frame, text="Выберите день:").grid(row=0, column=0, padx=10, pady=5)
    day_menu = tk.OptionMenu(frame, day_var, *days)
    day_menu.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(frame, text="Время (HH:MM) / (HH MM):").grid(
        row=1, column=0, padx=5, pady=5
    )
    time_entry = tk.Entry(frame)
    time_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(frame, text="URL:").grid(row=2, column=0, padx=5, pady=5)
    url_entry = tk.Entry(frame)
    url_entry.grid(row=2, column=1, padx=5, pady=5)
    url_entry.bind("<Return>", add_schedule)

    add_button = tk.Button(frame, text="Добавить", command=add_schedule)
    add_button.grid(row=3, column=0, pady=10)

    delete_button = tk.Button(frame, text="Удалить", command=delete_schedule)
    delete_button.grid(row=3, column=1, pady=10)

    tk.Label(frame, text="Задачи:").grid(row=4, column=0, columnspan=2, pady=10)
    task_list = tk.Listbox(frame, width=50, height=10)
    task_list.grid(row=5, column=0, columnspan=2, pady=5)

    task_list.bind("<Double-1>", edit_schedule)

    day_var.trace_add("write", update_tasks_for_day)

    print("Загрузка расписания..")
    load_schedule()

    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()

    if is_autostart_enabled():
        print("Свёрнуто в трей")
        root.withdraw()

    create_tray_icon()

    root.protocol("WM_DELETE_WINDOW", hide_window)
    print("Успешный запуск!")
    root.mainloop()
