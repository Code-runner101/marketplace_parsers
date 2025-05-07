from selenium.common import TimeoutException, NoSuchElementException
from fake_useragent import UserAgent
import time
import random
import undetected_chromedriver as uc
from bs4 import BeautifulSoup
import pandas as pd
import re
from selenium.webdriver.common.by import By

ua = UserAgent()
options = uc.ChromeOptions()


# Функция для извлечения диагонали из названия
def extract_diagonal_from_title(title):
    # Поиск числа с возможными символами " или '
    diagonal_match = re.search(r'\b(\d{2})(["\'])?\b', title)
    if diagonal_match:
        diagonal = diagonal_match.group(1)
        return diagonal
    return ""


def replace_brand_in_title(title, brand_dict):
    title_lower = title.lower()
    for brand in brand_dict.values():  # Итерируемся по значениям словаря
        # Создаем шаблон регулярного выражения для точного совпадения бренда
        brand_pattern = rf'\b{brand.lower()}\b'
        if re.search(brand_pattern, title_lower):
            return brand
    return title


def random_delay(min_delay=1, max_delay=100):
    time.sleep(random.uniform(min_delay, max_delay))


def scroll_page(driver, times=2):
    #Прокрутка страницы вниз указанное количество раз.
    for _ in range(times):
        driver.execute_script("window.scrollBy(0, window.innerHeight);")
        random_delay(1, 2)  # Задержка между прокрутками


def get_backlight_type(product_url):
    options = uc.ChromeOptions()
    random_user_agent = ua.random
    options.add_argument(f"user-agent={random_user_agent}")

    driver = uc.Chrome(options=options)

    try:

        driver.get(product_url)
        random_delay(1, 3)  # Ожидание для имитации загрузки контента

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Ищем элемент "Все характеристики" через BeautifulSoup
        specs_button = soup.find('span', class_="_1_47u _2SUA6 _33utW IFARr _1A5yJ", string="Все характеристики")
        if specs_button:
            # Если найден, получаем XPATH его родительского элемента, чтобы кликнуть через Selenium
            specs_button_xpath = '//*[@id="cardContent"]/div[1]/div/div[1]/div[3]/div[2]/div/div[9]/div/div/div/div[3]/span'
            try:
                button = driver.find_element(By.XPATH, specs_button_xpath)
                button.click()
                random_delay(2, 3)  # Ожидание появления всплывающего окна

                # Пробуем искать информацию во всплывающем окне
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                containers = soup.find_all('div', class_='_1j1RQ _1MOwX _2eMnU i1Eun')
                for container in containers:
                    label = container.find('span', class_='_1EbOn')
                    if label and 'Тип подсветки' in label.get_text(strip=True):
                        # Ищем элемент с типом подсветки
                        backlight_value = container.find('span', class_='YwVL7')
                        if backlight_value:
                            value_text = backlight_value.get_text(strip=True)
                            return value_text
                return "Не указана"
            except NoSuchElementException:
                print("Элемент 'Все характеристики' не найден.")

        scroll_page(driver, times=2)

        button_xpath = '// *[ @ id = "specs-list"] / div[2] / div / div / div / div / div / div[1] / button'
        try:
            button = driver.find_element(By.XPATH, button_xpath)
        except NoSuchElementException:
            print("Кнопка не найдена.")
            return "Не указана"

        # Получаем HTML после прокрутки
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Ищем элементы с классом "_3rW2x _1MOwX _2eMnU" с помощью BeautifulSoup
        containers = soup.find_all('div', class_='_3rW2x _1MOwX _2eMnU')

        for container in containers:
            label = container.find('span', class_='ds-text')
            if label and 'Тип подсветки' in label.get_text(strip=True):
                # Ищем следующий элемент, где содержится значение
                backlight_value = container.find('div', class_='b2ZT4')
                if backlight_value:
                    value_text = backlight_value.get_text(strip=True)
                    return value_text
        else:
            try:
                button.click()
                random_delay(2, 4)
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                containers = soup.find_all('div', class_='_3rW2x _1MOwX _2eMnU')

                for container in containers:
                    label = container.find('span', class_='ds-text')
                    if label and 'Тип подсветки' in label.get_text(strip=True):
                        # Ищем следующий элемент, где содержится значение
                        backlight_value = container.find('div', class_='b2ZT4')
                        if backlight_value:
                            value_text = backlight_value.get_text(strip=True)
                            return value_text
            except Exception as e:
                print(f"Ошибка при нажатии кнопки: {e}")
                # Пробуем снова искать информацию после нажатия кнопки

    except TimeoutException:
        print(f"Ошибка: Превышен таймаут загрузки страницы или элемента для {product_url}")
        return "Не указана"
    except Exception as e:
        print(f"Ошибка при получении типа подсветки: {e}")
        return "Не указана"
    finally:
        driver.quit()

    return "Не указана"


# Загрузка HTML-контента из файла
with open('/Users/macbookpro/Downloads/tv_ya3.htm', 'r', encoding='utf-8') as file:
    soup = BeautifulSoup(file, 'html.parser')

# Инициализация списков для хранения извлечённых данных
titles = []
diagonals = []
resolutions = []
smart_platforms = []
prices = []
backlight_types = []
links = []

# Список бредов ТВ
brand_dict = {
    0: "Витязь",
    1: "Триколор",
    2: "Asano",
    3: "artel",
    4: "AIWA",
    5: "Aceline",
    6: "Accesstyle",
    7: "AVEL",
    8: "AquaView",
    9: "ARAL",
    10: "BBK",
    11: "BQ",
    12: "Blackton",
    13: "Blaupunkt",
    14: "BAFF",
    15: "Candy",
    16: "CENTEK",
    17: "Carrera",
    18: "Cameron",
    19: "DEXP",
    20: "DIGMA",
    21: "Digma PRO",
    22: "Doffler",
    23: "Daewoo Electronics",
    24: "DENN",
    25: "DLED",
    26: "Erisson",
    27: "ECON",
    28: "Eplutus",
    29: "Ecotronic",
    30: "ExeGate",
    31: "Grundig",
    32: "GoldStar",
    33: "GARLYN",
    34: "GEFEST",
    35: "Goodview",
    36: "Gemy",
    37: "Hisense",
    38: "Haier",
    39: "HYUNDAI",
    40: "HARPER",
    41: "Hi",
    42: "Horizont",
    43: "HIPER",
    44: "HARTENS",
    45: "HIBERG",
    46: "Holleberg",
    47: "Horion",
    48: "Hikvision",
    49: "Hyundai/KIA",
    50: "Irbis",
    51: "iFFALCON",
    52: "Philips",
    53: "Panasonic",
    54: "Polarline",
    55: "Polar",
    56: "Prestigio",
    57: "Pioneer",
    58: "PRINX",
    59: "PROLISS",
    60: "realme",
    61: "Rombica",
    62: "RENOVA",
    63: "RAZZ",
    64: "Samsung",
    65: "Sony",
    66: "SBER",
    67: "Sharp",
    68: "SUPRA",
    69: "Smart TV",
    70: "STARWIND",
    71: "SkyLine",
    72: "Skyworth",
    73: "Sunwind",
    74: "Shivaki",
    75: "SoundMAX",
    76: "Scoole",
    77: "Schaub Lorenz",
    78: "SSMART",
    79: "TCL",
    80: "Toshiba",
    81: "Tuvio",
    82: "Thomson",
    83: "TELEFUNKEN",
    84: "TopDevice",
    85: "TESLA",
    86: "Viomi",
    87: "VEKTA",
    88: "VESTA",
    89: "Vestel",
    90: "VR",
    91: "Xiaomi",
    92: "Yuno",
    93: "Yasin",
    94: "ZeepDeep"
}

# Поиск всех контейнеров товаров
product_containers = soup.find_all('div', class_='_2rw4E _2g7lE')

# Обработка каждого контейнера товара
for product in product_containers:
    link = product.find('a', class_="EQlfk")
    if link:
        product_link = link.get('href')
        print(product_link)
        backlight_type = get_backlight_type(product_link)
        print(backlight_type)
        backlight_types.append(backlight_type)
        links.append(product_link)

    title_tag = product.find('span', class_='ds-text ds-text_lineClamp_2 ds-text_weight_med ds-text_color_text-primary ds-text_typography_lead-text ds-text_lead-text_normal ds-text_lead-text_med ds-text_lineClamp')
    if title_tag:
        title_text = title_tag.get_text(strip=True)
        title = replace_brand_in_title(title_text, brand_dict)
        titles.append(title)
    else:
        titles.append("")

    # Извлечение диагонали экрана
    diagonal = ""
    diagonal_tag = product.find('span', string=re.compile('Диагональ:'))
    if diagonal_tag:
        diagonal = diagonal_tag.find_next_sibling('span').get_text(strip=True).replace('"', '')
    else:
        # Если диагональ не найдена в описании, пытаемся найти её в названии товара
        diagonal = extract_diagonal_from_title(title_text)
    diagonals.append(diagonal)

    # Извлечение разрешения экрана
    resolution = ""
    resolution_tag = product.find('span', string=re.compile('Разрешение HD:'))
    if resolution_tag:
        resolution_text = resolution_tag.find_next_sibling('span').get_text(strip=True)
        resolution = re.search(r'(8K|4K|Full HD|HD)', resolution_text)
        resolution = resolution.group(0) if resolution else resolution_text
    resolutions.append(resolution)

    # Извлечение платформы Smart TV
    platform = ""
    platform_tag = product.find('span', string=re.compile('Платформа Smart TV:'))
    if platform_tag:
        platform = platform_tag.find_next_sibling('span').get_text(strip=True)
    smart_platforms.append(platform)

    # Извлечение цены товара
    price = ""
    price_classes = [
        'ds-text ds-text_weight_bold ds-text_color_price-term ds-text_typography_headline-5 ds-text_headline-5_tight ds-text_headline-5_bold',
        'ds-text ds-text_weight_bold ds-text_color_text-primary ds-text_typography_headline-5 ds-text_headline-5_tight ds-text_headline-5_bold',
        'ds-text ds-text_weight_bold ds-text_color_price-sale ds-text_typography_headline-5 ds-text_headline-5_tight ds-text_headline-5_bold',
    ]

for price_class in price_classes:
    price_tag = product.find('span', class_=price_class)
    if price_tag:
        price = re.sub(r'[^\d]', '', price_tag.get_text(strip=True))
        break

prices.append(price)

# Создание DataFrame для хранения извлечённых данных
data = {
    'Название товара': titles,
    'Диагональ': diagonals,
    'Разрешение экрана': resolutions,
    'OS': smart_platforms,
    'Цена (₽)': prices,
    'Тип подсветки': backlight_types,
    'links (to delete)': links
}

min_length = min(len(titles), len(diagonals), len(resolutions), len(smart_platforms), len(prices), len(backlight_types),
                 len(links))
titles = titles[:min_length]
diagonals = diagonals[:min_length]
resolutions = resolutions[:min_length]
smart_platforms = smart_platforms[:min_length]
prices = prices[:min_length]
backlight_types = backlight_types[:min_length]
links = links[:min_length]

df = pd.DataFrame(data)

# Сохранение DataFrame в файл Excel с гиперссылками
output_path = '/Users/macbookpro/Downloads/ya_tv_data2.xlsx'
writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)

worksheet = writer.sheets['Sheet1']

# Форматирование ссылок
link_format = writer.book.add_format({'font_color': 'blue', 'underline': 1})

# Добавление гиперссылок в столбец 'Ссылка'
for row_num, link in enumerate(links, start=1):  # start=1 чтобы пропустить заголовок
    worksheet.write_url(row_num, len(data) - 1, link, link_format)

writer.close()
