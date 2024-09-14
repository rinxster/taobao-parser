import os
import pandas as pd
from PIL import Image
from io import BytesIO
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent
import time
import random
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging

# Настройка логирования
logging.basicConfig(level=logging.INFO)

# Настройки веб-драйвера
options = Options()
options.add_argument("--headless")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("window-size=1920,1080")

def close_popup_if_present(driver):
    """Закрыть всплывающее окно, если оно появляется на странице."""
    try:
        wait = WebDriverWait(driver, 5)
        close_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '.login-box .fm-button.fm-submit.keep-login-btn.keep-login-confirm-btn.primary')))
        close_button.click()
        logging.info("Закрыто всплывающее окно")
    except Exception as e:
        # logging.warning(f'Всплывающее окно не найдено или не удалось закрыть: {e}')
        # logging.error(f'Всплывающее окно не найдено или не удалось закрыть: {e}')
        # logging.info(f'Всплывающее окно не найдено или не удалось закрыть: {e}')
        pass
def get_image_links(taobao_url):
    ua = UserAgent()
    user_agent = ua.random
    options.add_argument(f'user-agent={user_agent}')

    driver = webdriver.Chrome(options=options)
    driver.get(taobao_url)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    try:
        time.sleep(random.randint(5, 15))
        driver.implicitly_wait(10)

        # Закрытие всплывающего окна, если оно есть
        close_popup_if_present(driver)

        ul_element = driver.find_element(By.XPATH,
                                         '//*[@id="ice-container"]/div/div[2]/div[1]/div[2]/div/div[1]').find_element(
            By.CLASS_NAME, 'thumbnails--v976to2t')
        li_items = ul_element.find_elements(By.TAG_NAME, 'img')

        image_links = [img.get_attribute('src') for img in li_items]
        logging.info(f"Найденные ссылки на изображения: {image_links}")
        return image_links

    except Exception as e:
        logging.error(f"Ошибка: {e}")
        return []
    finally:
        driver.quit()

def create_folder_and_download_images(folder_name, image_urls):
    """Создать папку и загрузить изображения из списка URL."""
    formatted_folder_name = folder_name.replace(" ", "_").lower()
    folder_path = os.path.join('parseddata', folder_name)
    os.makedirs(folder_path, exist_ok=True)

    for index, url in enumerate(image_urls, start=1):
        download_image(url, folder_path, formatted_folder_name, index)

def download_image(url, folder_path, formatted_folder_name,  index):
    """Скачать изображение по URL и сохранить в указанной папке."""

    ua = UserAgent()
    headers = {
        'User-Agent': ua.random,
        'Referer': 'https://www.taobao.com/'
    }
    try:
        time.sleep(random.uniform(5, 10))

        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Проверка на ошибки

        image = Image.open(BytesIO(response.content))

        if image.format != 'JPEG':
            image = image.convert("RGB")

        file_name = os.path.join(folder_path, f"{formatted_folder_name}{index}_led7.jpg")
        image.save(file_name, format='JPEG')
        logging.info(f"Сохранено: {file_name}")
        time.sleep(random.uniform(5, 10))  # Задержка от 5 до 10 секунд

    except Exception as e:
        logging.error(f"Ошибка при загрузке {url}: {e}")

def main(excel_file):
    df = pd.read_excel(excel_file)

    for index, row in df.iterrows():
        if len(row) >= 3:
            position_name = row[4]  # Название позиции (пятая колонка)
            taobao_url = row[0]  # URL (первая колонка)
            logging.info(f"Обработка: {position_name}, URL: {taobao_url}")

            image_urls = get_image_links(taobao_url)  # Получаем ссылки на изображения
            if image_urls:
                create_folder_and_download_images(position_name, image_urls)  # Создаем папку и загружаем изображения

if __name__ == "__main__":
    main("links.xlsx")  # Замените на ваш файл Excel
