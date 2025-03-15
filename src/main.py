import json
import os
import sys
import threading
import time
import tkinter as tk
from datetime import datetime

import pystray
import schedule
from PIL import Image
from pystray import MenuItem as item

from browser_emulator import open_link
from src.settings import (
    WEEK_DAYS,
    SCHEDULE_FILE,
    save_settings,
    enable_autostart,
    check_single_instance,
    load_settings,
)

if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))


def get_task_week():
    return {
        "Monday": schedule.every().monday,
        "Tuesday": schedule.every().tuesday,
        "Wednesday": schedule.every().wednesday,
        "Thursday": schedule.every().thursday,
        "Friday": schedule.every().friday,
        "Saturday": schedule.every().saturday,
        "Sunday": schedule.every().sunday,
    }


class IdleScholar:
    def __init__(self, user_name, chrome_path, auto_start):
        self.USER_NAME: str = user_name
        self.CHROME_PATH: str = chrome_path
        self.AUTO_START: bool = auto_start
        self.APP_ICON: pystray.Icon = None

    def add_schedule(self, event=None):
        url = url_entry.get()
        time_str = time_entry.get().replace(" ", ":")
        day = day_var.get()
        msg = msg_entry.get()
        msg = msg if msg else ''

        if not url or not time_str or not day:
            return

        time_hours = time_str.split(":")[0]
        time_minutes = time_str.split(":")[1]

        if len(time_minutes) < 2:
            time_str = time_hours + ":0" + time_minutes

        if len(time_hours) < 2:
            time_str = "0" + time_str

        week_schedule = get_task_week()

        schedule_time = week_schedule.get(day).at(time_str)
        schedule_time.do(open_link, self, url=url, msg=msg)

        self.save_schedule()
        self.update_tasks_for_day()
        self.clear_entries()

    def delete_schedule(self):
        selected_index = task_list.curselection()
        if time_entry.get() or url_entry.get() or not selected_index:
            self.clear_entries()
            return

        selected_task = task_list.get(selected_index)
        for job in schedule.get_jobs():
            task_str = self.get_task_str(job)

            if task_str == selected_task:
                schedule.cancel_job(job)
                break

        task_list.delete(selected_index)
        self.save_schedule()
        self.clear_entries()

    @staticmethod
    def get_task_str(job):
        task_time = job.at_time.strftime("%H:%M")
        task_url = job.job_func.keywords["url"]
        task_msg = job.job_func.keywords["msg"]
        task_str = f"{task_time} | {task_msg} | {task_url}"
        return task_str

    def edit_schedule(self, event):
        if url_entry.get() or time_entry.get():
            return

        selected_index = task_list.curselection()
        if not selected_index:
            return

        selected_task = task_list.get(selected_index)
        time_str, msg, url = selected_task.split(" | ")

        # Устанавливаем выбранные значения в поля ввода
        self.clear_entries()
        self.set_entries(_time_entry=time_str, _url_entry=url, _msg_entry=msg)

        # Удаляем выбранную задачу из списка и планировщика
        task_list.delete(selected_index)
        for job in schedule.get_jobs():
            if self.get_task_str(job) == selected_task:
                schedule.cancel_job(job)
                break

        self.save_schedule()

    def clear_entries(self):
        time_entry.delete(0, tk.END)
        url_entry.delete(0, tk.END)
        msg_entry.delete(0, tk.END)

    def set_entries(self, _time_entry, _url_entry, _msg_entry):
        time_entry.insert(0, _time_entry)
        url_entry.insert(0, _url_entry)
        msg_entry.insert(0, _msg_entry)

    def save_schedule(self):
        tasks = []
        for job in schedule.get_jobs():
            task = {
                "week_day": WEEK_DAYS[job.next_run.weekday()],
                "time": job.at_time.strftime("%H:%M"),
                "url": job.job_func.keywords["url"],
                "msg": job.job_func.keywords["msg"],
            }
            tasks.append(task)

        if tasks:
            with open(SCHEDULE_FILE, "w") as f:
                json.dump(tasks, f, indent=4)

    def load_schedule(self):
        schedule.clear()

        if os.path.exists(SCHEDULE_FILE):
            with open(SCHEDULE_FILE, "r") as f:
                tasks = json.load(f)
                for task in tasks:
                    week_schedule = get_task_week()

                    day = task["week_day"]
                    time_str = task["time"]
                    url = task["url"]
                    msg = task.get("msg")
                    msg = msg if msg else ''

                    schedule_time = week_schedule[day].at(time_str)
                    schedule_time.do(open_link, self, url=url, msg=msg)

        today = datetime.today().strftime("%A")
        day_var.set(today)
        self.update_tasks_for_day()

    def update_tasks_for_day(self, *args):
        selected_day = day_var.get()
        day_tasks = [
            job
            for job in schedule.get_jobs()
            if WEEK_DAYS[job.next_run.weekday()] == selected_day
        ]
        task_list.delete(0, tk.END)
        for task in day_tasks:
            task_list.insert(tk.END, self.get_task_str(task))

    def open_settings(self):
        def on_ok():
            user_name = entry1.get()
            chrome_path = entry2.get()
            save_settings(user_name, chrome_path, self.AUTO_START)
            self.USER_NAME, self.CHROME_PATH, self.AUTO_START = load_settings()
            settings_window.destroy()

        def on_cancel():
            settings_window.destroy()

        def trigger_autostart():
            flag = is_on_startup.get()
            enable_autostart(flag)
            self.AUTO_START = flag

        settings_window = tk.Toplevel(root)
        settings_window.resizable(False, False)
        settings_window.title("Настройки")
        settings_window.transient(root)
        settings_window.grab_set()

        settings_window.geometry(f"+{root.winfo_rootx()}+{root.winfo_rooty()}")

        tk.Label(settings_window, text="Имя:").grid(row=0, column=0, padx=10, pady=5)
        entry1 = tk.Entry(settings_window)
        entry1.insert(0, self.USER_NAME)
        entry1.grid(row=0, column=1, padx=20, pady=5)

        tk.Label(settings_window, text="Путь к Chrome:").grid(
            row=1, column=0, padx=10, pady=5
        )
        entry2 = tk.Entry(settings_window)
        entry2.insert(0, self.CHROME_PATH)
        entry2.grid(row=1, column=1, padx=10, pady=5)

        is_on_startup = tk.BooleanVar()
        is_on_startup.set(self.AUTO_START)
        startup_check = tk.Checkbutton(
            settings_window,
            text="Запуск вместе с Windows",
            variable=is_on_startup,
            command=trigger_autostart,
        )
        startup_check.grid(row=2, column=1, padx=10, pady=5)

        ok_button = tk.Button(settings_window, text="OK", command=on_ok)
        ok_button.grid(row=3, column=0, padx=10, pady=10)

        cancel_button = tk.Button(settings_window, text="Отмена", command=on_cancel)
        cancel_button.grid(row=3, column=1, padx=10, pady=10)

        settings_window.focus_set()
        settings_window.wait_window()

    def run_scheduler(self):
        while True:
            schedule.run_pending()
            time.sleep(1)

    def hide_window(self):
        root.withdraw()
        self.create_tray_icon()

    def show_window(self):
        self.APP_ICON.stop()
        root.deiconify()

    def create_tray_icon(self):
        image = Image.open("icon.png")
        menu = (
            item("Show", lambda: self.show_window(), default=True),
            item("Exit", lambda: self.on_exit()),
        )
        self.APP_ICON = pystray.Icon("schedule_app", image, "Schedule App", menu)
        threading.Thread(target=self.APP_ICON.run, daemon=True).start()

    def on_exit(self):
        self.APP_ICON.stop()
        root.destroy()


if __name__ == "__main__":
    if not check_single_instance():
        print("Программа уже запущена.")
        sys.exit(0)

    print("Загрузка настроек..")
    scholar = IdleScholar(*load_settings())

    # GUI
    print("Загрузка вёрстки..")
    root = tk.Tk()
    root.title("Расписание занятий")
    root.resizable(False, False)

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    day_var = tk.StringVar(value=WEEK_DAYS[0])

    settings_button = tk.Button(
        frame, text="⚙", command=scholar.open_settings, justify="left", anchor="w"
    )
    settings_button.grid(sticky="W", column=0, row=0)

    tk.Label(frame, text="Выберите день:").grid(row=0, column=0, padx=10, pady=5)
    day_menu = tk.OptionMenu(frame, day_var, *WEEK_DAYS)
    day_menu.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(frame, text="Время (HH:MM) / (HH MM):").grid(
        row=1, column=0, padx=5, pady=5
    )
    time_entry = tk.Entry(frame)
    time_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(frame, text="URL:").grid(row=2, column=0, padx=5, pady=5)
    url_entry = tk.Entry(frame)
    url_entry.grid(row=2, column=1, padx=5, pady=5)
    url_entry.bind("<Return>", scholar.add_schedule)

    add_button = tk.Button(frame, text="Добавить", command=scholar.add_schedule)
    add_button.grid(row=3, column=0, pady=10)

    delete_button = tk.Button(frame, text="Удалить", command=scholar.delete_schedule)
    delete_button.grid(row=3, column=1, pady=10)

    msg_label = tk.Label(
        frame,
        text="Авто-сообщение в чат при заходе",
    )
    msg_label.grid(row=4, column=0, columnspan=1, padx=1, pady=10)

    msg_entry = tk.Entry(frame)
    msg_entry.grid(row=4, column=1)

    tk.Label(frame, text="Задачи:").grid(row=5, column=0, columnspan=2, pady=10)
    task_list = tk.Listbox(frame, width=50, height=10)
    task_list.grid(row=6, column=0, columnspan=2, pady=5)

    task_list.bind("<Double-1>", scholar.edit_schedule)

    day_var.trace_add("write", scholar.update_tasks_for_day)

    print("Загрузка расписания..")
    scholar.load_schedule()

    scheduler_thread = threading.Thread(target=scholar.run_scheduler, daemon=True)
    scheduler_thread.start()

    if scholar.AUTO_START:
        scholar.create_tray_icon()
        root.withdraw()
        print("Свёрнуто в трей")

    root.protocol("WM_DELETE_WINDOW", scholar.hide_window)
    print("Успешный запуск!")
    root.mainloop()
