import configparser
import winreg
import ctypes
import sys
import os


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


def check_single_instance() -> bool:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    kernel32.CreateMutexW(None, False, MUTEX)

    if kernel32.GetLastError() == 183:
        return False

    return True


def load_settings() -> (str, str, bool):
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
        user_name = config.get("Settings", "user_name", fallback="student")
        chrome_path = config.get(
            "Settings",
            "chrome_path",
            fallback="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        )
        auto_start = (
            True
            if config.get("Settings", "auto_start", fallback=False) == "True"
            else False
        )
    else:
        user_name = "student"
        chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        auto_start = False

    return user_name, chrome_path, auto_start


def save_settings(user_name, chrome_path, auto_start):
    config = configparser.ConfigParser()
    config["Settings"] = {
        "user_name": user_name,
        "chrome_path": chrome_path,
        "auto_start": auto_start,
    }
    with open(CONFIG_FILE, "w") as cfg:
        config.write(cfg)


def enable_autostart(enable: bool):
    key = winreg.HKEY_CURRENT_USER
    subkey = r"Software\Microsoft\Windows\CurrentVersion\Run"

    with winreg.OpenKey(key, subkey, 0, winreg.KEY_SET_VALUE) as reg_key:
        if enable:
            winreg.SetValueEx(reg_key, APP_NAME, 0, winreg.REG_SZ, APP_PATH)
            print("Created registry file at path", key, subkey)
        else:
            try:
                winreg.DeleteValue(reg_key, APP_NAME)
                print("Deleted registry file at path", key, subkey)
            except FileNotFoundError:
                print("Error deleting registry file at path", key, subkey)
