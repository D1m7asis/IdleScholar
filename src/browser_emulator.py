from playwright.sync_api import sync_playwright
import random
import time
import os


def login_scenario(self, page):
    input_name_selector = "#name"

    # Ввод имени пользователя
    if page.locator(input_name_selector).is_visible(timeout=0):
        page.fill(input_name_selector, self.USER_NAME)
        page.keyboard.press("Enter")
        return

    # Разрешение автозапуска видео
    autoplay_button_selector = ".autoplay-video-allow-btn"
    if page.locator(autoplay_button_selector).is_visible(timeout=0):
        page.click(autoplay_button_selector)
        page.wait_for_timeout(500)

    # Пропуск использования микрофона
    continue_button_selectors = [
        'button:has-text("продолжить без микрофона")',
        'button:has-text("Присоединиться без устройств")',
    ]

    for selector in continue_button_selectors * 4:
        if page.locator(selector).is_visible(timeout=5000):
            page.click(selector)
            page.wait_for_timeout(500)


def send_message_scenario(page, msg):
    input_message_selector = '[data-placeholder="Введите сообщение"]'

    try:
        if page.locator(input_message_selector).nth(0).is_visible(timeout=5000):
            page.locator(input_message_selector).nth(0).fill(msg)
            page.keyboard.press("Enter")
            print("Сообщение успешно отправлено.")
            return True
        else:
            print("Поле ввода сообщения не найдено.")
            return False
    except Exception as e:
        print(f"Ошибка при отправке сообщения: {e}")
        return False


def open_link(self, url, msg):
    with sync_playwright() as p:
        user_data_dir = os.path.join(os.getcwd(), "chrome_user_data")
        browser = p.chromium.launch_persistent_context(
            user_data_dir,
            executable_path=self.CHROME_PATH,
            headless=False,
            args=[
                "--use-fake-ui-for-media-stream",  # Отключает реальный доступ к камере/микрофону
                "--use-fake-device-for-media-stream",  # Подставляет фейковое устройство
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-extensions",
                "--disable-features=IsolateOrigins,site-per-process",
            ],
            permissions=[],  # Запрещаем все разрешения
        )

        # Запрещаем доступ к камере и микрофону
        browser.grant_permissions([], origin=url)

        page = browser.new_page()  # Используем context.new_page()

        try:
            page.goto(url, timeout=60_000)  # 60 секунд на загрузку
            page.wait_for_load_state("networkidle")

            login_scenario(self, page)  # Выполняем сценарий логина
            if msg:
                send_message_scenario(page, msg)

            start_time = time.time()
            while time.time() - start_time < 2 * 60 * 60:
                x, y = random.randint(100, 200), random.randint(100, 200)
                page.mouse.move(x, y)
                time.sleep(random.randint(240, 360))  # Ожидание от 4 до 6 минут

        except Exception as e:
            print(f"Ошибка: {e}")

        finally:
            page.close()
            browser.close()
