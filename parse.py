from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import openpyxl

BASE_URL = "https://karatltd.com.ua/"

# Настройки для Selenium
chrome_options = Options()
chrome_options.add_argument("--headless")  # Без отображения браузера
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")

# Запуск браузера
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Функция для парсинга товаров
def parse_products(driver, url):
    driver.get(url)
    time.sleep(3)  # Ждем загрузку страницы

    products = []

    # Получаем все карточки товаров
    product_cards = driver.find_elements(By.CLASS_NAME, "product-layout")

    if not product_cards:
        print("❗️ Товары не найдены на странице")

    for card in product_cards:
        id = card.find_element(By.CLASS_NAME, "sc-module-model").text.strip() if card.find_element(By.CLASS_NAME, "sc-module-model") else "No Id"
        title = card.find_element(By.CLASS_NAME, "sc-module-title").text.strip() if card.find_element(By.CLASS_NAME, "sc-module-title") else "No title"
        price = card.find_element(By.CLASS_NAME, "sc-module-price").text.strip() if card.find_element(By.CLASS_NAME, "sc-module-price") else "No price"
        link = card.find_element(By.CLASS_NAME, "sc-module-title").get_attribute("href") if card.find_element(By.CLASS_NAME, "sc-module-title") else "#"

        products.append({
            "id": id,
            "title": title,
            "price": price,
            "link": link
        })

    return products

# Функция для парсинга всех товаров с пагинацией
def parse_all_products_with_pagination(driver, category_url, max_pages=97):
    all_products = []
    
    for page_num in range(1, max_pages + 1):
        url = f"{category_url}?page={page_num}"
        print(f"📄 Парсинг страницы {page_num}...")
        
        # Получаем товары с текущей страницы
        products = parse_products(driver, url)
        
        if not products:
            print("❗️ Товары не найдены на странице, пропускаем.")
            continue
        
        all_products.extend(products)
    
    return all_products

# Функция для сохранения данных в Excel
def save_to_excel(products, filename="products.xlsx"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Products"

    # Заголовки
    sheet.append(["Id", "Название", "Цена", "Ссылка"])

    # Заполнение данными
    for product in products:
        sheet.append([
            product["id"],
            product["title"],
            product["price"],
            product["link"]
        ])

    workbook.save(filename)
    print(f"✅ Данные успешно сохранены в {filename}")

# Пример использования
if __name__ == "__main__":
    category_url = BASE_URL + "kabelno-providnykova-produktsiia"  # Укажите реальный путь к категории
    all_products = parse_all_products_with_pagination(driver, category_url)

    if all_products:
        save_to_excel(all_products)
    else:
        print("❗️ Нет данных для сохранения")

    driver.quit()  # Закрываем браузер после работы
