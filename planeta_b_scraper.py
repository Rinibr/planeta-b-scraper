# -*- coding: utf-8 -*-

import sys
import os
import time
import re
import csv
import json
import logging
import asyncio
import webbrowser # Для открытия ссылок
import html # Для HTML экранирования
# Импорты из typing для совместимости < Python 3.9
from typing import Dict, List, Any, Optional

# --- Определение пути к ресурсам (для PyInstaller) ---
def get_resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу, работает для разработки и PyInstaller """
    try:
        # PyInstaller создает временную папку и сохраняет путь в sys._MEIPASS
        base_path = sys._MEIPASS
        # logger.debug(f"Running bundled. Base path: {base_path}") # Отладка
    except AttributeError:
        # sys._MEIPASS не определен, значит запущено как обычный скрипт
        base_path = os.path.dirname(os.path.abspath(__file__))
        # logger.debug(f"Running as script. Base path: {base_path}") # Отладка
    except Exception as e:
        # На всякий случай ловим другие ошибки и используем CWD
        # Логгер может быть еще не настроен здесь, используем print
        print(f"Error getting base_path: {e}. Using CWD.", file=sys.stderr)
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- Определение путей для логов и категорий ---
if getattr(sys, 'frozen', False):
    # Если запущено из упакованного .exe
    application_path = os.path.dirname(sys.executable)
else:
    # Если запущено как обычный скрипт .py
    application_path = os.path.dirname(os.path.abspath(__file__))

LOG_FILENAME_ONLY = 'scraper_app.log' # Только имя файла
LOG_FILE_PATH = os.path.join(application_path, LOG_FILENAME_ONLY) # Полный путь к логу
CATEGORIES_FILENAME_ONLY = 'categories.json' # Только имя файла

# --- Настройка логирования (ЕДИНСТВЕННЫЙ БЛОК) ---
from logging.handlers import RotatingFileHandler

log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(name)s] %(message)s')
log_handler = RotatingFileHandler(LOG_FILE_PATH, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
log_handler.setFormatter(log_formatter)

logger = logging.getLogger() # Настраиваем корневой, чтобы ловить все
logger.setLevel(logging.DEBUG) # Установите DEBUG для отладки

# Удаляем существующие обработчики (если были добавлены ранее)
for hdlr in logger.handlers[:]:
    logger.removeHandler(hdlr)

logger.addHandler(log_handler) # Добавляем наш файловый обработчик

# Добавляем вывод в консоль
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)
console_handler.setLevel(logging.INFO) # INFO для консоли
logger.addHandler(console_handler)
logger.info(f"Логгер настроен. Лог-файл: {LOG_FILE_PATH}")
# --- Конец настройки логирования ---

# --- Загрузка категорий (ЕДИНСТВЕННЫЙ БЛОК) ---
# Используем get_resource_path для файла категорий
CATEGORIES_FILE_PATH = get_resource_path(CATEGORIES_FILENAME_ONLY)
logger.debug(f"Ожидаемый путь к файлу категорий: {CATEGORIES_FILE_PATH}")

def load_categories_from_file(full_filepath: str) -> Dict[str, str]:
    """Загружает категории и URL из JSON файла по полному пути."""
    logger.debug(f"Попытка загрузки категорий из файла: {full_filepath}")
    default_categories = {"--- Выберите категорию ---": ""}
    categories_to_return = default_categories.copy()

    if not os.path.exists(full_filepath):
        logger.warning(f"Файл категорий '{full_filepath}' не найден.")
        try:
            with open(full_filepath, 'w', encoding='utf-8') as f:
                example_categories = {
                    "--- Выберите категорию ---": "",
                    "Пример Категории (отредактируйте categories.json)": "https://example.com"
                }
                json.dump(example_categories, f, ensure_ascii=False, indent=4)
            logger.info(f"Создан пример файла категорий '{full_filepath}'. Пожалуйста, отредактируйте его.")
            categories_to_return = example_categories
        except IOError as e:
             logger.error(f"Не удалось создать файл категорий '{full_filepath}': {e}")
             categories_to_return = default_categories.copy()
        return categories_to_return

    try:
        with open(full_filepath, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if not content:
                logger.warning(f"Файл категорий '{full_filepath}' пуст. Используются дефолтные значения.")
                categories_to_return = default_categories.copy()
            else:
                f.seek(0)
                data = json.load(f)
                if not isinstance(data, dict):
                    logger.error(f"Ошибка в файле категорий '{full_filepath}': ожидался JSON объект (словарь). Используются дефолтные значения.")
                    categories_to_return = default_categories.copy()
                else:
                    if "--- Выберите категорию ---" not in data:
                        categories_to_return = {**default_categories, **data}
                    else:
                        categories_to_return = data
    except json.JSONDecodeError as e:
        logger.error(f"Ошибка декодирования JSON в файле категорий '{full_filepath}': {e}. Используются дефолтные значения.")
        categories_to_return = default_categories.copy()
    except IOError as e:
        logger.error(f"Ошибка чтения файла категорий '{full_filepath}': {e}. Используются дефолтные значения.")
        categories_to_return = default_categories.copy()
    except Exception as e:
        logger.error(f"Неожиданная ошибка при загрузке категорий из '{full_filepath}': {e}", exc_info=True)
        categories_to_return = default_categories.copy()

    logger.info(f"Загружено {len(categories_to_return)} категорий из '{full_filepath}'.")
    logger.debug(f"Возвращаемые категории: {categories_to_return}")
    return categories_to_return
# --- Конец загрузки категорий ---

# --- Проверка и импорт зависимостей ---
try:
    import qasync
except ImportError:
     print("ОШИБКА: Библиотека qasync не найдена. Установите ее: pip install qasync", file=sys.stderr)
     sys.exit(1)

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    logger.info("Библиотека openpyxl не найдена. Сохранение в Excel будет недоступно.")

try:
    import aiohttp # Для асинхронной загрузки изображений
    AIOHTTP_AVAILABLE = True
except ImportError:
    AIOHTTP_AVAILABLE = False
    logger.info("Библиотека aiohttp не найдена. Превью изображений будет недоступно.")

# --- PyQt5 Импорты ---
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel, QLineEdit,
                             QPushButton, QTextEdit, QVBoxLayout, QHBoxLayout, QSplitter,
                             QGridLayout, QFileDialog, QMessageBox, QSpinBox, QCheckBox,
                             QProgressBar, QComboBox, QTableWidget, QTableWidgetItem,
                             QHeaderView, QTabWidget) # <-- Добавлен QTabWidget
from PyQt5.QtGui import QPixmap, QResizeEvent, QCloseEvent
from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot, QSettings, QTimer, Qt, QSize

# --- Импорт Playwright и заглушки ---
class MockAsyncContextManager:
    async def __aenter__(self): return self
    async def __aexit__(self, exc_type, exc_val, exc_tb): pass
    def chromium(self): return MockAsyncBrowserLauncher()
class MockAsyncBrowserLauncher:
    async def launch(self, headless): return MockAsyncBrowser()
class MockAsyncBrowser:
    async def new_page(self): return MockAsyncPage()
    async def close(self): pass
    def is_connected(self): return False
    async def new_context(self, **kwargs): return MockAsyncContext()
class MockAsyncContext:
     async def new_page(self): return MockAsyncPage()
     async def close(self): pass
class MockAsyncPage:
    async def goto(self, url, timeout=None, wait_until=None): await asyncio.sleep(0); logger.warning("Playwright заглушка: goto")
    async def content(self): await asyncio.sleep(0); logger.warning("Playwright заглушка: content"); return ""
    def locator(self, selector): return MockAsyncLocator()
    async def wait_for_selector(self, selector, timeout=None, state=None): await asyncio.sleep(0); logger.warning("Playwright заглушка: wait_for_selector")
    def set_default_navigation_timeout(self, timeout): pass
    def set_default_timeout(self, timeout): pass
    async def close(self): await asyncio.sleep(0); logger.warning("Playwright заглушка: close page")
    @property
    def url(self): return "http://mock.url"
class MockAsyncLocator:
    async def count(self): await asyncio.sleep(0); return 0
    def nth(self, i): return self
    async def is_visible(self, timeout=None): await asyncio.sleep(0); return False
    async def inner_text(self, timeout=None): await asyncio.sleep(0); return "Mock Text"
    async def all_inner_texts(self): await asyncio.sleep(0); return []
    async def get_attribute(self, name, timeout=None): await asyncio.sleep(0); return None
    async def click(self, timeout=None, **kwargs): await asyncio.sleep(0); logger.warning("Playwright заглушка: click")

PlaywrightTimeoutErrorMock = type('PlaywrightTimeoutErrorMock', (Exception,), {})
PlaywrightErrorMock = type('PlaywrightErrorMock', (Exception,), {})

async_playwright = MockAsyncContextManager()
PlaywrightTimeoutError = PlaywrightTimeoutErrorMock
PlaywrightError = PlaywrightErrorMock
PLAYWRIGHT_AVAILABLE = False

try:
    from playwright.async_api import async_playwright as async_playwright_real
    from playwright.async_api import TimeoutError as PlaywrightTimeoutErrorReal
    from playwright.async_api import Error as PlaywrightErrorReal

    async_playwright = async_playwright_real
    PlaywrightTimeoutError = PlaywrightTimeoutErrorReal
    PlaywrightError = PlaywrightErrorReal
    PLAYWRIGHT_AVAILABLE = True
    logger.info("Playwright успешно импортирован.")
except ImportError:
    logger.critical("Playwright не найден! Установите его (`pip install playwright && playwright install chromium`)")
# --- Конец импорта Playwright ---

# --- Вспомогательная функция ---
def try_convert_to_number(value):
    if isinstance(value, (int, float)): return value
    try:
        cleaned_value = str(value).replace('\xa0', '').replace(' ', '').replace(',', '.')
        cleaned_value = re.sub(r'[^\d.\-].*$', '', cleaned_value) # Убираем все после не-цифр/точки/минуса
        if not cleaned_value or cleaned_value == '-': return value # Если ничего не осталось или только минус
        if '.' in cleaned_value:
             # Проверка на несколько точек
             if cleaned_value.count('.') > 1: return value
             # Проверка на точку в начале/конце или '--'
             if cleaned_value.startswith('.') or cleaned_value.endswith('.') or '--' in cleaned_value: return value
             # Проверка на '-' не в начале
             if '-' in cleaned_value and not cleaned_value.startswith('-'): return value
             return float(cleaned_value)
        else:
             # Проверка на '-' не в начале
             if '-' in cleaned_value and not cleaned_value.startswith('-'): return value
             return int(cleaned_value)
    except (ValueError, TypeError):
        return value

# ==============================================================================
# Рабочий класс для парсинга (ScraperWorker)
# ==============================================================================
class ScraperWorker(QObject):
    log_message = pyqtSignal(str)
    error = pyqtSignal(str)
    finished = pyqtSignal(list) # List[Dict[str, Any]]
    progress = pyqtSignal(int, int)

    def __init__(self, start_url: str, max_pages: int, scrape_details: bool = True):
        super().__init__()
        self.start_url = start_url
        self.max_pages = max_pages
        self.scrape_details = scrape_details
        self.is_running = True
        # --- Селекторы (можно вынести в конфиг/json при желании) ---
        # Каталог
        self.CARD_LOCATOR_SELECTOR = 'div.panel.panel-default' # Контейнер карточки
        self.NAME_LINK_SELECTOR = 'h3.media-heading a'          # Ссылка с названием
        self.PRICE_SELECTOR = 'h3.wholesale span.summ'          # Элемент с ценой
        self.IMAGE_SELECTOR = 'div.media a.media-left img'      # Картинка
        self.NEXT_PAGE_SELECTOR = 'ul.pagination a[aria-label="Next"], ul.pagination a[rel="next"]' # Кнопка "след."
        self.PAGINATION_BLOCK_SELECTOR = 'ul.pagination'        # Блок пагинации (для проверки наличия)
        # Страница товара
        self.TAB1_SELECTOR = 'div#home'                         # Первая вкладка с характеристиками
        self.CHAR_TABLE1_SELECTOR = 'table.table-bordered'      # Таблица в первой вкладке
        self.CHAR_ROW1_SELECTOR = 'tr'                          # Строка в таблице 1
        self.CHAR_NAME1_SELECTOR = 'td:nth-of-type(1)'          # Ячейка с названием характеристики (табл.1)
        self.CHAR_VALUE1_SELECTOR = 'td:nth-of-type(2)'         # Ячейка со значением характеристики (табл.1)
        self.TAB2_SELECTOR = 'div#settings'                     # Вторая вкладка с характеристиками
        self.CHAR_TABLE2_SELECTOR = 'table.vendorenabled'       # Таблица во второй вкладке
        self.CHAR_ROW2_SELECTOR = 'tr'                          # Строка в таблице 2
        self.CHAR_NAME2_SELECTOR = 'td:nth-of-type(1) b'        # Ячейка с названием характеристики (табл.2)
        self.CHAR_VALUE2_SELECTOR = 'td:nth-of-type(2) span'    # Ячейка со значением характеристики (табл.2)

    def _log(self, message: str, level: int = logging.INFO):
        # Используем QTimer для безопасной отправки сигнала из async потока в GUI
        QTimer.singleShot(0, lambda msg=message: self.log_message.emit(msg))
        # Пишем в файл/консоль через основной логгер
        log_func = logger.info
        if level == logging.ERROR: log_func = logger.error
        elif level == logging.WARNING: log_func = logger.warning
        elif level == logging.DEBUG: log_func = logger.debug
        log_func(f"[Парсер] {message}")

    async def run(self) -> None:
        if not PLAYWRIGHT_AVAILABLE:
            err_msg = "Playwright не найден. Установите его: pip install playwright && playwright install chromium"
            QTimer.singleShot(0, lambda: self.error.emit(err_msg))
            QTimer.singleShot(0, lambda: self.finished.emit([]))
            return

        all_results: List[Dict[str, Any]] = []
        page_count = 0
        browser = None
        context = None
        page = None

        try:
            self._log(f"Запуск парсинга... URL: {self.start_url}")
            if self.max_pages > 0: self._log(f"Макс. страниц: {self.max_pages}")
            else: self._log(f"Макс. страниц: Все")
            self._log(f"Сбор характеристик: {'Включен' if self.scrape_details else 'Выключен'}")

            async with async_playwright() as p:
                self._log(f"Запуск браузера (headless=True)...")
                try:
                    # Убрали channel="chrome", используем стандартный chromium
                    browser = await p.chromium.launch(headless=True)
                except Exception as launch_err:
                    self._log(f"ОШИБКА запуска браузера: {launch_err}", level=logging.ERROR)
                    raise RuntimeError(f"Не удалось запустить браузер Playwright. Убедитесь, что он установлен (`playwright install chromium`). Ошибка: {launch_err}")

                self._log(f"Создание контекста браузера...")
                context = await browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
                    java_script_enabled=True,
                    accept_downloads=False,
                    # Можно добавить блокировку ресурсов для ускорения (опционально)
                    # route = lambda route: route.abort() if route.request.resource_type in {"image", "stylesheet", "font"} else route.continue_()
                )

                page = await context.new_page()
                self._log(f"Установка таймаутов и переход на стартовую страницу...")
                page.set_default_navigation_timeout(90000) # 90 секунд на навигацию
                page.set_default_timeout(60000)           # 60 секунд на действия/ожидания

                # Переход на начальный URL с retry
                try:
                    retries = 2
                    for attempt in range(retries + 1):
                        try:
                            await page.goto(self.start_url, wait_until='domcontentloaded', timeout=90000)
                            self._log(f"Успешный переход на {self.start_url}")
                            break # Выходим из цикла retry
                        except PlaywrightTimeoutError as goto_timeout:
                            if attempt < retries:
                                self._log(f"Таймаут при переходе на {self.start_url} (попытка {attempt+1}/{retries+1}). Повтор через 5 сек...", level=logging.WARNING)
                                await asyncio.sleep(5)
                            else:
                                self._log(f"ОШИБКА: Таймаут при переходе на {self.start_url} после {retries+1} попыток.", level=logging.ERROR)
                                raise # Перебрасываем ошибку таймаута
                        except Exception as goto_err:
                            # Ловим другие ошибки goto (например, network error)
                             self._log(f"ОШИБКА при переходе на {self.start_url}: {goto_err}", level=logging.ERROR)
                             raise # Перебрасываем ошибку

                except Exception as initial_goto_err:
                     # Если даже с retry не получилось, выходим
                     raise initial_goto_err

                # --- Основной цикл парсинга по страницам ---
                while self.is_running:
                    page_count += 1
                    QTimer.singleShot(0, lambda pc=page_count, mp=self.max_pages: self.progress.emit(pc, mp if mp > 0 else -1))

                    self._log(f"--- Обработка страницы #{page_count} ---")
                    current_url = page.url
                    self._log(f"Текущий URL: {current_url}")

                    # Ожидание появления карточек
                    self._log(f"Ожидание карточек товара ({self.CARD_LOCATOR_SELECTOR})...")
                    try:
                        # Ждем первую карточку + небольшая пауза на всякий случай
                        await page.locator(self.CARD_LOCATOR_SELECTOR).first.wait_for(state='visible', timeout=60000)
                        self._log(f"Карточки найдены. Ожидание стабилизации...")
                        await asyncio.sleep(1.5) # Небольшая пауза для динамического контента
                    except PlaywrightTimeoutError:
                        self._log(f"Таймаут ожидания карточек на странице {page_count}.", level=logging.ERROR)
                        # Проверим, есть ли пагинация. Если нет, возможно, просто пустая категория или конец
                        pagination_exists = await page.locator(self.PAGINATION_BLOCK_SELECTOR).count() > 0
                        if not pagination_exists and page_count > 1:
                             self._log(f"Карточек нет, пагинации тоже нет. Вероятно, достигнут конец каталога.", level=logging.INFO)
                        elif not pagination_exists and page_count == 1:
                             self._log(f"Карточек нет на первой странице и нет пагинации. Категория пуста?", level=logging.WARNING)
                        break # Прерываем цикл по страницам
                    except Exception as wait_error:
                        self._log(f"Ошибка ожидания контента: {wait_error}", level=logging.ERROR)
                        break

                    if not self.is_running: self._log(f"Остановка после ожидания контента."); break

                    # --- Извлечение данных с карточек ---
                    self._log(f"Извлечение данных с карточек...")
                    card_locators = page.locator(self.CARD_LOCATOR_SELECTOR)
                    count = await card_locators.count()
                    self._log(f"Найдено {count} карточек.")
                    page_results_count = 0

                    for i in range(count):
                        if not self.is_running: self._log(f"Остановка во время обработки карточки #{i+1}."); break
                        card = card_locators.nth(i)
                        name = "N/A"; link = "N/A"; price = "N/A"; chars: Dict[str, str] = {}; image_url: Optional[str] = None

                        try:
                            # Имя и ссылка
                            name_loc = card.locator(self.NAME_LINK_SELECTOR)
                            if await name_loc.count() > 0 and await name_loc.is_visible(timeout=5000):
                                name = (await name_loc.inner_text(timeout=5000)).strip()
                                href = await name_loc.get_attribute('href', timeout=5000)
                                if href:
                                    link = href
                                    # Сборка абсолютного URL
                                    if link.startswith('/'): link = 'https://planeta-b.ru' + link
                                    elif not link.startswith('http'):
                                         try: from urllib.parse import urljoin; link = urljoin(current_url, link)
                                         except ImportError: pass # Маловероятно

                            # Цена
                            price_loc = card.locator(self.PRICE_SELECTOR)
                            if await price_loc.count() > 0 and await price_loc.is_visible(timeout=5000):
                                raw_price = await price_loc.inner_text(timeout=5000)
                                price = raw_price.strip() # Сохраняем как есть, конвертация позже
                            else: price = "Цена не найдена" # Явно указываем

                            # URL изображения
                            image_loc = card.locator(self.IMAGE_SELECTOR)
                            if await image_loc.count() > 0:
                                img_src = await image_loc.first.get_attribute('src', timeout=5000)
                                if img_src:
                                    image_url = img_src
                                    # Сборка абсолютного URL
                                    if image_url.startswith('//'): image_url = 'https:' + image_url
                                    elif image_url.startswith('/'): image_url = 'https://planeta-b.ru' + image_url
                                    elif not image_url.startswith('http'):
                                        try: from urllib.parse import urljoin; image_url = urljoin(current_url, image_url)
                                        except ImportError: pass

                            # --- Парсинг характеристик (если включен) ---
                            if self.scrape_details and link != "N/A" and link.startswith('http') and self.is_running:
                                self._log(f"Запрос деталей: {name[:40]}... ({link})", level=logging.DEBUG)
                                prod_page = None
                                try:
                                    prod_page = await context.new_page()
                                    prod_page.set_default_navigation_timeout(45000)
                                    prod_page.set_default_timeout(30000)
                                    await prod_page.goto(link, wait_until='domcontentloaded', timeout=45000)
                                    if not self.is_running: raise asyncio.CancelledError("Остановка во время загрузки деталей")

                                    # Функция для парсинга таблицы
                                    async def parse_char_table(tab_locator, table_sel, row_sel, name_sel, value_sel) -> Dict[str, str]:
                                        parsed_chars = {}
                                        if not self.is_running: return parsed_chars
                                        try:
                                            if await tab_locator.is_visible(timeout=10000):
                                                table = tab_locator.locator(table_sel).first
                                                if await table.is_visible(timeout=5000):
                                                    rows = table.locator(row_sel)
                                                    row_count = await rows.count()
                                                    for j in range(row_count):
                                                        if not self.is_running: break
                                                        row = rows.nth(j)
                                                        try:
                                                            n_loc = row.locator(name_sel).first
                                                            v_loc = row.locator(value_sel).first
                                                            if await n_loc.is_visible(timeout=1000) and await v_loc.is_visible(timeout=1000):
                                                                k = (await n_loc.inner_text(timeout=1000)).strip().replace(':','').strip()
                                                                v = (await v_loc.inner_text(timeout=1000)).strip()
                                                                if k and v and v != '-': # Добавляем, если ключ/значение непустые и значение не '-'
                                                                    # Добавляем только если ключа еще нет (избегаем дублей между табл 1 и 2)
                                                                    if k not in chars and k not in parsed_chars:
                                                                        parsed_chars[k] = v
                                                        except PlaywrightTimeoutError: continue # Игнор таймаутов чтения строки
                                                        except Exception as row_err:
                                                            self._log(f"Ошибка в строке {j+1} таблицы ({table_sel}): {row_err}", level=logging.DEBUG)
                                        except PlaywrightTimeoutError:
                                            self._log(f"Таймаут ожидания вкладки/таблицы ({tab_locator} / {table_sel})", level=logging.DEBUG)
                                        return parsed_chars

                                    # Парсинг Таблицы 1 (div#home)
                                    tab1_chars = await parse_char_table(prod_page.locator(self.TAB1_SELECTOR).first,
                                                                        self.CHAR_TABLE1_SELECTOR, self.CHAR_ROW1_SELECTOR,
                                                                        self.CHAR_NAME1_SELECTOR, self.CHAR_VALUE1_SELECTOR)
                                    chars.update(tab1_chars)
                                    if not self.is_running: break # Проверка после первой таблицы

                                    # Парсинг Таблицы 2 (div#settings)
                                    tab2_chars = await parse_char_table(prod_page.locator(self.TAB2_SELECTOR).first,
                                                                        self.CHAR_TABLE2_SELECTOR, self.CHAR_ROW2_SELECTOR,
                                                                        self.CHAR_NAME2_SELECTOR, self.CHAR_VALUE2_SELECTOR)
                                    chars.update(tab2_chars) # Добавляем уникальные ключи из второй таблицы

                                except asyncio.CancelledError:
                                    self._log(f"Парсинг деталей прерван для: {link}", level=logging.WARNING)
                                except PlaywrightTimeoutError:
                                    self._log(f"Таймаут при загрузке/парсинге страницы деталей: {link}", level=logging.WARNING)
                                except Exception as detail_err:
                                    self._log(f"Ошибка при парсинге деталей {link}: {detail_err}", level=logging.ERROR)
                                    logger.error(f"Ошибка деталей {link}", exc_info=True) # Полный traceback в лог-файл
                                finally:
                                    if prod_page and not prod_page.is_closed():
                                        await prod_page.close()
                                self._log(f"Найдено характеристик: {len(chars)} для {name[:20]}...", level=logging.DEBUG)
                            # --- Конец парсинга характеристик ---

                        except PlaywrightTimeoutError as card_timeout_err:
                             self._log(f"Таймаут при извлечении данных из карточки #{i+1}: {card_timeout_err}", level=logging.WARNING)
                        except Exception as extract_error:
                             self._log(f"Ошибка извлечения данных из карточки #{i+1}: {extract_error}", level=logging.ERROR)
                             logger.error(f"Ошибка извлечения из карточки #{i+1} на URL {current_url}", exc_info=True)

                        # Добавление результата (только если парсер работает и есть имя)
                        if self.is_running and name != "N/A":
                            item_data: Dict[str, Any] = {
                                'name': name,
                                'price': price, # Сохраняем исходную строку цены
                                'link': link,
                                'image_url': image_url,
                                'characteristics': chars # Словарь характеристик
                            }
                            all_results.append(item_data)
                            page_results_count += 1
                    # --- Конец цикла по карточкам ---

                    if not self.is_running: self._log(f"Остановка после обработки карточек."); break
                    self._log(f"Добавлено {page_results_count} товаров со страницы {page_count}.")

                    # Проверка лимита страниц
                    if self.max_pages > 0 and page_count >= self.max_pages:
                        self._log(f"Достигнут лимит страниц ({self.max_pages}). Завершение.")
                        break

                    # --- Пагинация ---
                    self._log(f"Поиск кнопки 'Следующая страница' ({self.NEXT_PAGE_SELECTOR})...")
                    next_btn = page.locator(self.NEXT_PAGE_SELECTOR).first

                    try:
                        is_visible = await next_btn.is_visible(timeout=10000)
                        is_enabled = await next_btn.is_enabled(timeout=5000) if is_visible else False

                        if is_visible and is_enabled:
                            self._log(f"Клик по кнопке 'Следующая страница'...")
                            await next_btn.click(timeout=15000)
                            self._log(f"Ожидание загрузки следующей страницы...")
                            # Ждем загрузки DOM и снова появления первой карточки
                            await page.wait_for_load_state('domcontentloaded', timeout=60000)
                            # Доп. ожидание видимости карточки для надежности
                            await page.locator(self.CARD_LOCATOR_SELECTOR).first.wait_for(state='visible', timeout=60000)
                            self._log(f"Следующая страница загружена.")
                            await asyncio.sleep(1) # Небольшая пауза
                        else:
                            self._log(f"Кнопка 'Следующая страница' не найдена/не активна. Конец пагинации.")
                            break # Выходим из цикла while
                    except PlaywrightTimeoutError:
                         self._log(f"Таймаут при поиске/клике/ожидании 'Следующая страница'. Завершение пагинации.", level=logging.WARNING)
                         break
                    except Exception as next_page_err:
                         # Проверяем, не связана ли ошибка с остановкой
                         if not self.is_running and ("Target page" in str(next_page_err) or "closed" in str(next_page_err)):
                             self._log(f"Ошибка пагинации после сигнала стоп (ожидаемо): {next_page_err}", level=logging.INFO)
                         else:
                             self._log(f"Ошибка при переходе на следующую страницу: {next_page_err}", level=logging.ERROR)
                             logger.error("Ошибка пагинации", exc_info=True)
                         break
                # --- Конец цикла while ---
                self._log(f"Основной цикл парсинга завершен.")

        except RuntimeError as rt_err: # Ошибка запуска браузера
             err_msg = f"Критическая ошибка Playwright: {rt_err}"
             QTimer.singleShot(0, lambda: self.error.emit(err_msg))
             logger.critical(f"Критическая ошибка Playwright в ScraperWorker: {rt_err}", exc_info=True)
        except asyncio.CancelledError:
             self._log(f"Задача парсинга была отменена.", level=logging.WARNING)
        except Exception as e:
             # Глобальная ошибка в воркере
             if isinstance(e, PlaywrightError) and "Target page" in str(e) and "closed" in str(e) and not self.is_running:
                 self._log(f"Ошибка Playwright после остановки (ожидаемо): {e}", level=logging.INFO)
             else:
                 err_msg = f"Глобальная ошибка в парсере: {e}"
                 QTimer.singleShot(0, lambda: self.error.emit(err_msg))
                 logger.error(f"Глобальная ошибка в ScraperWorker", exc_info=True)
        finally:
            self._log(f"Блок finally воркера.")
            self.is_running = False
            # Аккуратное закрытие ресурсов
            try:
                if page and not page.is_closed():
                    self._log(f"Закрытие страницы...", level=logging.DEBUG)
                    await page.close()
            except Exception as page_close_err:
                 self._log(f"Ошибка при закрытии страницы: {page_close_err}", level=logging.WARNING)
            try:
                if context:
                    self._log(f"Закрытие контекста...", level=logging.DEBUG)
                    await context.close()
            except Exception as context_close_err:
                 self._log(f"Ошибка при закрытии контекста: {context_close_err}", level=logging.WARNING)
            try:
                if browser and browser.is_connected():
                    self._log(f"Закрытие браузера...", level=logging.DEBUG)
                    await browser.close()
                    self._log(f"Браузер закрыт.", level=logging.DEBUG)
                elif browser:
                    self._log(f"Браузер уже был отключен.", level=logging.DEBUG)
            except Exception as browser_close_err:
                 self._log(f"Ошибка при закрытии браузера: {browser_close_err}", level=logging.WARNING)

            # Отправляем результаты в GUI
            self._log(f"Отправка сигнала finished с {len(all_results)} результатами.")
            QTimer.singleShot(0, lambda res=all_results: self.finished.emit(res))

    def stop(self):
        if self.is_running:
            self._log(f"Получен сигнал stop. Установка is_running = False.")
            self.is_running = False
        else:
             self._log(f"Сигнал stop получен, но парсер уже не активен.", level=logging.DEBUG)

# ==============================================================================
# Главный класс GUI
# ==============================================================================
class ScraperApp(QMainWindow):
    # Версия для вкладки "О программе"
    APP_VERSION = "1.9.5 (Tabs+Fixes)" # Обновленная версия

    def __init__(self):
        super().__init__()
        self.settings = QSettings("Videokot", "Парсер сайта Планета безопасности")
        self.setWindowTitle(f'Парсер сайта Планета безопасности v{self.APP_VERSION}') # Используем версию
        self.setGeometry(100, 100, 1100, 850)
        self.scraping_task: Optional[asyncio.Task] = None
        self.worker: Optional[ScraperWorker] = None
        self.is_scraper_running: bool = False
        self.all_scraper_results: List[Dict[str, Any]] = []
        self.http_session: Optional[aiohttp.ClientSession] = None
        self.current_pixmap: Optional[QPixmap] = None

        # Загрузка категорий (использует функцию и путь, определенные ВЫШЕ)
        self.loaded_categories: Dict[str, str] = load_categories_from_file(CATEGORIES_FILE_PATH)

        self.load_settings()
        self.initUI() # Вызываем обновленный initUI с вкладками
        self.apply_settings_to_widgets()
        self._log_file_path: str = LOG_FILE_PATH # Используем путь, определенный ВЫШЕ
        logger.info(f"Путь к лог-файлу: {self._log_file_path}")

        logger.info("[GUI] Приложение инициализировано.")
        if not PLAYWRIGHT_AVAILABLE:
             self.handle_error("Playwright не найден! Парсинг невозможен.")
        if not qasync:
             self.handle_error("qasync не найден! Приложение не может работать.")

    def _open_log_file(self) -> None:
        if os.path.exists(self._log_file_path):
            try:
                webbrowser.open(f"file:///{os.path.abspath(self._log_file_path)}")
                logger.info(f"Попытка открыть лог-файл: {self._log_file_path}")
            except Exception as e:
                logger.error(f"Не удалось открыть лог-файл '{self._log_file_path}': {e}")
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть лог-файл:\n{self._log_file_path}\n\nОшибка: {e}")
        else:
            logger.warning(f"Лог-файл не найден: {self._log_file_path}")
            QMessageBox.warning(self, "Файл не найден", f"Лог-файл не найден:\n{self._log_file_path}")

    def initUI(self) -> None:
        # --- Создаем QTabWidget ---
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget) # Устанавливаем TabWidget как центральный

        # --- 1. Создаем виджет и компоновку для первой вкладки ("Парсер") ---
        parser_tab = QWidget()
        main_layout = QVBoxLayout(parser_tab) # Теперь это компоновка первой вкладки

        # --- Переносим сюда все элементы интерфейса ---

        # --- 1.1 Блок Настроек ---
        settings_layout = QGridLayout()
        settings_layout.addWidget(QLabel('Категория:'), 0, 0)
        self.category_combo = QComboBox()
        if self.loaded_categories:
            category_keys = list(self.loaded_categories.keys())
            self.category_combo.addItems(category_keys)
        else:
            logger.error("Словарь категорий пуст перед добавлением в QComboBox!")
            self.category_combo.addItem("--- Ошибка загрузки категорий ---")
        settings_layout.addWidget(self.category_combo, 0, 1, 1, 3)

        settings_layout.addWidget(QLabel('Макс. страниц (0=все):'), 1, 0)
        self.max_pages_input = QSpinBox()
        self.max_pages_input.setRange(0, 9999)
        settings_layout.addWidget(self.max_pages_input, 1, 1)

        self.scrape_details_checkbox = QCheckBox("Собирать характеристики")
        self.scrape_details_checkbox.setToolTip("Сбор детальной информации со страницы каждого товара (замедляет парсинг)")
        settings_layout.addWidget(self.scrape_details_checkbox, 1, 2)

        settings_layout.addWidget(QLabel('Формат файла:'), 2, 0)
        self.format_combo = QComboBox()
        available_formats = ['CSV', 'JSON']
        if OPENPYXL_AVAILABLE:
            available_formats.append('Excel (.xlsx)')
        self.format_combo.addItems(available_formats)
        self.format_combo.currentTextChanged.connect(self._update_save_dialog_filter)
        settings_layout.addWidget(self.format_combo, 2, 1)

        settings_layout.addWidget(QLabel('Сохранить в:'), 3, 0)
        self.outfile_input = QLineEdit()
        self.browse_button = QPushButton('Обзор...')
        self.browse_button.clicked.connect(self.browse_file)
        settings_layout.addWidget(self.outfile_input, 3, 1, 1, 2)
        settings_layout.addWidget(self.browse_button, 3, 3)
        main_layout.addLayout(settings_layout) # Добавляем в компоновку вкладки

        # --- 1.2 Блок Управления ---
        control_layout = QHBoxLayout()
        self.start_button = QPushButton('Начать парсинг')
        self.start_button.setStyleSheet("background-color: lightgreen; font-weight: bold;")
        self.start_button.clicked.connect(self.start_scraping)
        self.stop_button = QPushButton('Остановить')
        self.stop_button.setStyleSheet("background-color: lightcoral; font-weight: bold;")
        self.stop_button.clicked.connect(self.stop_scraping)
        self.stop_button.setEnabled(False)

        self.clear_log_button = QPushButton('Очистить лог')
        self.clear_log_button.clicked.connect(self._clear_log_output)

        self.open_log_button = QPushButton("Открыть лог")
        self.open_log_button.clicked.connect(self._open_log_file)

        control_layout.addWidget(self.start_button)
        control_layout.addWidget(self.stop_button)
        control_layout.addStretch(1)
        control_layout.addWidget(self.clear_log_button)
        control_layout.addWidget(self.open_log_button)
        main_layout.addLayout(control_layout) # Добавляем в компоновку вкладки

        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar) # Добавляем в компоновку вкладки

        # --- 1.3 Блок Фильтров ---
        filter_group_box = QWidget()
        filter_main_layout = QVBoxLayout(filter_group_box)
        filter_main_layout.setContentsMargins(5, 5, 5, 5)

        filter_row1_layout = QHBoxLayout()
        filter_row1_layout.addWidget(QLabel("Фильтры: Название содерж.:"))
        self.filter_name_input = QLineEdit()
        self.filter_name_input.setPlaceholderText("Часть названия (регистр не важен)")
        filter_row1_layout.addWidget(self.filter_name_input)
        filter_row1_layout.addWidget(QLabel("Цена от:"))
        self.filter_price_min_input = QSpinBox()
        self.filter_price_min_input.setRange(0, 9999999)
        self.filter_price_min_input.setValue(0)
        self.filter_price_min_input.setSuffix(" руб.")
        self.filter_price_min_input.setFixedWidth(100)
        filter_row1_layout.addWidget(self.filter_price_min_input)
        filter_row1_layout.addWidget(QLabel("до:"))
        self.filter_price_max_input = QSpinBox()
        self.filter_price_max_input.setRange(0, 9999999)
        self.filter_price_max_input.setValue(0)
        self.filter_price_max_input.setSpecialValueText("Любой")
        self.filter_price_max_input.setSuffix(" руб.")
        self.filter_price_max_input.setFixedWidth(100)
        filter_row1_layout.addWidget(self.filter_price_max_input)
        filter_main_layout.addLayout(filter_row1_layout)

        filter_row2_layout = QHBoxLayout()
        filter_row2_layout.addWidget(QLabel("Характеристика:"))
        self.filter_char_name_combo = QComboBox()
        self.filter_char_name_combo.addItem("--- Любая ---")
        self.filter_char_name_combo.setMinimumWidth(180)
        self.filter_char_name_combo.setToolTip("Выберите характеристику из списка (заполняется после парсинга)")
        filter_row2_layout.addWidget(self.filter_char_name_combo)
        filter_row2_layout.addWidget(QLabel("Значение содержит:"))
        self.filter_char_value_input = QLineEdit()
        self.filter_char_value_input.setPlaceholderText("Часть значения (регистр не важен)")
        filter_row2_layout.addWidget(self.filter_char_value_input)
        filter_row2_layout.addStretch(1)
        self.apply_filter_button = QPushButton("Применить фильтр")
        self.apply_filter_button.clicked.connect(self.apply_filter)
        filter_row2_layout.addWidget(self.apply_filter_button)
        self.reset_filter_button = QPushButton("Сбросить фильтр")
        self.reset_filter_button.clicked.connect(self.reset_filter)
        filter_row2_layout.addWidget(self.reset_filter_button)
        filter_main_layout.addLayout(filter_row2_layout)

        main_layout.addWidget(filter_group_box) # Добавляем в компоновку вкладки

        # --- 1.4 Блок Результатов, Деталей, Изображения и Лога ---
        output_splitter = QSplitter(Qt.Horizontal)

        # Левая часть - Таблица
        table_widget = QWidget()
        table_layout = QVBoxLayout(table_widget)
        table_layout.setContentsMargins(0,0,0,0)
        self.results_label = QLabel("Результаты: 0")
        table_layout.addWidget(self.results_label)
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(3)
        self.results_table.setHorizontalHeaderLabels(['Название', 'Цена', 'Ссылка'])
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setSelectionMode(QTableWidget.SingleSelection)
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.results_table.setSortingEnabled(True)
        self.results_table.itemSelectionChanged.connect(self.show_details)
        self.results_table.itemDoubleClicked.connect(self.open_product_link)
        table_layout.addWidget(self.results_table)
        output_splitter.addWidget(table_widget)

        # Правая часть - Вертикальный сплиттер для деталей/картинки и лога
        right_splitter = QSplitter(Qt.Vertical)

        # Верхняя правая часть - Детали и Картинка
        details_image_widget = QWidget()
        details_image_layout = QHBoxLayout(details_image_widget)
        details_image_layout.setContentsMargins(5,0,0,0)
        # Панель картинки
        image_panel_layout = QVBoxLayout()
        image_panel_layout.addWidget(QLabel("Изображение:"))
        self.image_preview_label = QLabel("Выберите товар в таблице")
        self.image_preview_label.setAlignment(Qt.AlignCenter)
        self.image_preview_label.setMinimumSize(200, 200)
        self.image_preview_label.setScaledContents(False) # Масштабируем вручную
        self.image_preview_label.setStyleSheet("QLabel { background-color : #f0f0f0; border: 1px solid grey; color: grey; }")
        image_panel_layout.addWidget(self.image_preview_label)
        image_panel_layout.addStretch(1)
        details_image_layout.addLayout(image_panel_layout)
        # Панель характеристик
        details_panel_layout = QVBoxLayout()
        details_panel_layout.addWidget(QLabel("Характеристики:"))
        self.details_output = QTextEdit()
        self.details_output.setReadOnly(True)
        details_panel_layout.addWidget(self.details_output)
        details_image_layout.addLayout(details_panel_layout, 1)
        right_splitter.addWidget(details_image_widget)

        # Нижняя правая часть - Лог
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(5,0,0,0)
        log_layout.addWidget(QLabel("Лог:"))
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        log_layout.addWidget(self.log_output)
        right_splitter.addWidget(log_widget)

        right_splitter.setSizes([350, 250])
        output_splitter.addWidget(right_splitter)
        output_splitter.setSizes([650, 450])

        main_layout.addWidget(output_splitter) # Добавляем сплиттер в компоновку вкладки

        # --- Добавляем первую вкладку в QTabWidget ---
        self.tab_widget.addTab(parser_tab, "Парсер")

        # --- 2. Создаем виджет и компоновку для второй вкладки ("О программе") ---
        about_tab = QWidget()
        about_layout = QVBoxLayout(about_tab)
        about_layout.setContentsMargins(15, 15, 15, 15) # Добавим отступы

        # Заполняем вторую вкладку
        title_label = QLabel(f"<b>Парсер Planeta-B</b><br>Версия: {self.APP_VERSION}")
        title_label.setAlignment(Qt.AlignCenter)
        font = title_label.font()
        font.setPointSize(14)
        title_label.setFont(font)

        description_label = QLabel(
            "Это приложение предназначено для сбора информации о товарах "
            "(название, цена, характеристики, изображения) с сайта planeta-b.ru.<br><br>"
            "<b>Используемые технологии:</b><br>"
            "- Python<br>"
            "- PyQt5 (для графического интерфейса)<br>"
            "- Playwright (для взаимодействия с веб-страницей)<br>"
            "- asyncio / qasync (для асинхронной работы)<br>"
            "- aiohttp (для загрузки изображений, опционально)<br>"
            "- openpyxl (для сохранения в Excel, опционально)"
        )
        description_label.setWordWrap(True)
        description_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)

        # Замените URL на актуальный, если есть репозиторий
        repo_link_label = QLabel(
            'Исходный код: <a href="https://github.com/ваша_учетная_запись/ваш_репозиторий">GitHub</a> (если доступен)'
        )
        repo_link_label.setOpenExternalLinks(True)
        repo_link_label.setAlignment(Qt.AlignCenter)

        # Замените на свое имя/ник
        author_label = QLabel("Разработчик:Ринат Ибрагимов/Rinibr")
        author_label.setAlignment(Qt.AlignCenter)

        about_layout.addWidget(title_label)
        about_layout.addSpacing(20)
        about_layout.addWidget(description_label)
        about_layout.addStretch(1)
        about_layout.addWidget(repo_link_label)
        about_layout.addWidget(author_label)
        about_layout.addSpacing(10)

        # --- Добавляем вторую вкладку в QTabWidget ---
        self.tab_widget.addTab(about_tab, "О программе")

        # --- Строка состояния ---
        self.statusBar().showMessage('Готов')
        logger.debug("Инициализация UI с вкладками завершена.")


    def load_settings(self) -> None:
        logger.debug("Загрузка настроек...")
        saved_category = self.settings.value("category", "--- Выберите категорию ---")
        if saved_category in self.loaded_categories:
            self.default_category_text = saved_category
        else:
            self.default_category_text = "--- Выберите категорию ---"
            if saved_category != "--- Выберите категорию ---":
                logger.warning(f"Сохраненная категория '{saved_category}' не найдена. Сброс на дефолтную.")

        self.default_output_file = self.settings.value("outputFile", os.path.join(os.getcwd(), 'planeta_b_data.csv'))
        self.default_format = self.settings.value("outputFormat", "CSV")
        try:
            self.default_max_pages = int(self.settings.value("maxPages", 0))
        except (ValueError, TypeError): self.default_max_pages = 0
        self.default_scrape_details = self.settings.value("scrapeDetails", True, type=bool)
        logger.debug(f"Загруженные настройки: категория='{self.default_category_text}', файл='{self.default_output_file}', формат='{self.default_format}', макс.стр={self.default_max_pages}, детали={self.default_scrape_details}")


    def apply_settings_to_widgets(self) -> None:
         logger.debug("Применение настроек к виджетам...")
         # Проверяем наличие виджетов перед обращением (на всякий случай)
         if hasattr(self, 'category_combo'):
             index = self.category_combo.findText(self.default_category_text)
             if index != -1: self.category_combo.setCurrentIndex(index)
             elif self.category_combo.count() > 0: self.category_combo.setCurrentIndex(0)
         if hasattr(self, 'outfile_input'): self.outfile_input.setText(self.default_output_file)
         if hasattr(self, 'format_combo'):
             current_format = self.default_format
             if current_format == 'Excel (.xlsx)' and not OPENPYXL_AVAILABLE:
                 logger.warning("Формат Excel выбран, но openpyxl недоступен. Сброс на CSV.")
                 current_format = 'CSV'
             index = self.format_combo.findText(current_format)
             if index != -1: self.format_combo.setCurrentIndex(index)
             elif self.format_combo.count() > 0: self.format_combo.setCurrentIndex(0)
             self._update_save_dialog_filter() # Обновить расширение файла
         if hasattr(self, 'max_pages_input'): self.max_pages_input.setValue(self.default_max_pages)
         if hasattr(self, 'scrape_details_checkbox'): self.scrape_details_checkbox.setChecked(self.default_scrape_details)
         logger.debug("Настройки применены к виджетам.")


    def save_settings(self) -> None:
        logger.debug("Сохранение настроек...")
        if hasattr(self, 'category_combo'): self.settings.setValue("category", self.category_combo.currentText())
        if hasattr(self, 'outfile_input'): self.settings.setValue("outputFile", self.outfile_input.text())
        if hasattr(self, 'format_combo'): self.settings.setValue("outputFormat", self.format_combo.currentText())
        if hasattr(self, 'max_pages_input'): self.settings.setValue("maxPages", self.max_pages_input.value())
        if hasattr(self, 'scrape_details_checkbox'): self.settings.setValue("scrapeDetails", self.scrape_details_checkbox.isChecked())
        self.settings.sync()
        logger.debug("Настройки сохранены.")


    def _update_save_dialog_filter(self) -> None:
        selected_format = self.format_combo.currentText()
        current_file = self.outfile_input.text()
        base, current_ext = os.path.splitext(current_file)
        new_ext = ""
        self.file_dialog_filter = ""

        if "JSON" in selected_format:
            new_ext = ".json"; self.file_dialog_filter = "JSON (*.json);;Все файлы (*)"
        elif "Excel" in selected_format:
            new_ext = ".xlsx"; self.file_dialog_filter = "Excel (*.xlsx);;Все файлы (*)"
        else: # CSV по умолчанию
            new_ext = ".csv"; self.file_dialog_filter = "CSV (*.csv);;Все файлы (*)"

        # Обновляем расширение, если оно не соответствует выбранному формату
        if new_ext and (not current_ext or current_ext.lower() != new_ext):
             # Проверяем, не содержит ли базовое имя файла уже новое расширение
             if not base.lower().endswith(new_ext):
                 self.outfile_input.setText(base + new_ext)
             else: # Имя уже содержит расширение (например, file.csv), но оно было без точки
                 # Проверим, нужно ли заменить, если пользователь ввел filecsv
                 if not current_file.endswith(new_ext):
                     self.outfile_input.setText(base + new_ext)
        logger.debug(f"Фильтр диалога обновлен: {self.file_dialog_filter}, имя файла: {self.outfile_input.text()}")


    def browse_file(self) -> None:
        options = QFileDialog.Options()
        initial_file = self.outfile_input.text()
        initial_dir = os.path.dirname(initial_file) if os.path.dirname(initial_file) else os.getcwd()
        if not os.path.isdir(initial_dir): initial_dir = os.getcwd()
        initial_file = os.path.join(initial_dir, os.path.basename(initial_file))

        # Убедимся, что фильтр установлен перед открытием диалога
        if not hasattr(self, 'file_dialog_filter') or not self.file_dialog_filter:
            self._update_save_dialog_filter()

        fileName, selected_filter = QFileDialog.getSaveFileName(self,
                                                  "Сохранить файл как",
                                                  initial_file,
                                                  self.file_dialog_filter,
                                                  options=options)
        if fileName:
            logger.debug(f"Выбран файл: {fileName}, Фильтр: {selected_filter}")
            # Добавляем/исправляем расширение на основе фильтра
            required_ext = ""
            if "(*.json)" in selected_filter: required_ext = ".json"
            elif "(*.xlsx)" in selected_filter: required_ext = ".xlsx"
            elif "(*.csv)" in selected_filter: required_ext = ".csv"
            else: # Если фильтр "Все файлы (*)" или что-то неожиданное, берем из комбобокса
                 selected_format = self.format_combo.currentText()
                 if "JSON" in selected_format: required_ext = ".json"
                 elif "Excel" in selected_format: required_ext = ".xlsx"
                 else: required_ext = ".csv"

            base, ext = os.path.splitext(fileName)
            if required_ext and (not ext or ext.lower() != required_ext):
                fileName = base + required_ext
            self.outfile_input.setText(fileName)
            logger.info(f"Установлен файл для сохранения: {fileName}")


    def update_log(self, message: str) -> None:
        try:
            # Используем форматирование из файлового логгера для консистентности
            log_record = logging.LogRecord(name='GUI', level=logging.INFO, pathname="", lineno=0, msg=message, args=[], exc_info=None)
            formatted_message = log_handler.formatter.format(log_record) if log_handler.formatter else message

            self.log_output.append(formatted_message)
            # Автопрокрутка, только если ползунок внизу
            scrollbar = self.log_output.verticalScrollBar()
            if not scrollbar.isVisible() or scrollbar.value() >= scrollbar.maximum() - 15:
                scrollbar.setValue(scrollbar.maximum())
        except Exception as e:
             print(f"Ошибка при обновлении лога GUI: {e}", file=sys.stderr)


    def _clear_log_output(self) -> None:
        self.log_output.clear()
        self.details_output.clear()
        self.image_preview_label.clear()
        self.image_preview_label.setText("Выберите товар в таблице")
        self.image_preview_label.setStyleSheet("QLabel { background-color : #f0f0f0; border: 1px solid grey; color: grey; }")
        self.current_pixmap = None
        self.statusBar().showMessage("Лог, детали и превью очищены", 3000)
        logger.info("Лог GUI, детали и превью очищены.")


    def update_progress(self, current_page: int, max_pages: int) -> None:
        if self.is_scraper_running:
            if max_pages > 0: # Если известно макс. количество страниц
                self.progress_bar.setMaximum(max_pages)
                self.progress_bar.setValue(current_page)
                self.progress_bar.setFormat(f"Стр. {current_page} / {max_pages}")
                self.statusBar().showMessage(f"Обработка страницы {current_page} из {max_pages}...")
            else: # Если макс. количество страниц не задано (0)
                self.progress_bar.setMaximum(0) # Режим "бесконечного" прогресса
                self.progress_bar.setValue(0)
                self.progress_bar.setFormat(f"Стр. {current_page}")
                self.statusBar().showMessage(f"Обработка страницы {current_page}...")


    def handle_error(self, error_message: str) -> None:
        log_msg = f"--- ОШИБКА ---\n{error_message}\n-----------------"
        logger.error(log_msg) # В основной лог
        # В лог GUI красным цветом
        self.log_output.append(f"<font color='red'>{html.escape(log_msg)}</font>")
        scrollbar = self.log_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum()) # Прокрутка вниз

        if self.isVisible(): # Показываем QMessageBox только если окно видимо
             QMessageBox.critical(self, "Ошибка", error_message)
        else:
             print(f"Критическая ошибка до показа окна: {error_message}", file=sys.stderr)

        # Сброс состояния, если парсер был активен или заблокирован при инициализации
        if self.is_scraper_running or not self.start_button.isEnabled():
             self._reset_gui_state("Ошибка!")


    def _display_results(self, data_list: List[Dict[str, Any]]) -> None:
        self.results_table.setSortingEnabled(False) # Отключаем сортировку для скорости
        self.results_table.clearContents()
        self.results_table.setRowCount(0)

        count = len(data_list)
        self.results_label.setText(f"Результаты: {count}")
        logger.info(f"Отображение {count} результатов...")

        if not data_list:
            self.details_output.clear()
            self.image_preview_label.clear(); self.image_preview_label.setText("Нет данных")
            self.current_pixmap = None
            return

        self.results_table.setRowCount(count)
        for row_idx, item in enumerate(data_list):
            try:
                # Столбец 0: Название (храним весь dict в UserRole)
                name_item = QTableWidgetItem(item.get('name', ''))
                name_item.setData(Qt.UserRole, item) # <-- Важно для деталей
                self.results_table.setItem(row_idx, 0, name_item)

                # Столбец 1: Цена (пытаемся конвертировать для сортировки)
                price_item = QTableWidgetItem()
                raw_price = item.get('price', '')
                numeric_price = try_convert_to_number(raw_price)
                if isinstance(numeric_price, (int, float)):
                    price_item.setData(Qt.EditRole, numeric_price) # Число для сортировки
                    price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                # Отображаем всегда исходную строку
                price_item.setData(Qt.DisplayRole, str(raw_price))
                self.results_table.setItem(row_idx, 1, price_item)

                # Столбец 2: Ссылка
                link_text = item.get('link', '')
                link_item = QTableWidgetItem(link_text)
                link_item.setToolTip(link_text) # Подсказка с полным URL
                self.results_table.setItem(row_idx, 2, link_item)

            except Exception as e:
                 logger.error(f"Ошибка при создании ячейки для строки {row_idx}, данных: {item}: {e}", exc_info=True)
                 # Добавляем строку с ошибкой, чтобы не сдвигать нумерацию
                 error_item = QTableWidgetItem(f"Ошибка отображения строки {row_idx}")
                 error_item.setData(Qt.UserRole, None)
                 self.results_table.setItem(row_idx, 0, error_item)

        self.results_table.setSortingEnabled(True) # Включаем сортировку обратно
        logger.info("Таблица результатов обновлена.")
        if count > 0:
             self.results_table.selectRow(0) # Выделяем первую строку для показа деталей
             self.show_details()


    def show_details(self) -> None:
        selected_items = self.results_table.selectedItems()
        self.details_output.clear()
        self.image_preview_label.clear()
        self.image_preview_label.setText("Загрузка...")
        self.image_preview_label.setStyleSheet("QLabel { background-color : #f0f0f0; border: 1px solid grey; color: grey; }")
        self.current_pixmap = None # Сбрасываем текущее изображение

        if not selected_items:
            self.image_preview_label.setText("Выберите товар в таблице")
            return

        selected_row = selected_items[0].row()
        name_cell = self.results_table.item(selected_row, 0)
        if not name_cell:
            logger.warning(f"Не удалось получить ячейку (0) для строки {selected_row} в show_details")
            self.details_output.setText("Ошибка: Нет данных для строки.")
            self.image_preview_label.setText("Ошибка данных")
            return

        item_data = name_cell.data(Qt.UserRole)

        if not isinstance(item_data, dict):
             logger.warning(f"Некорректные данные в UserRole для строки {selected_row}: {type(item_data)}")
             self.details_output.setText("Ошибка: Некорректные данные.")
             self.image_preview_label.setText("Ошибка данных")
             return

        # Отображение характеристик
        characteristics = item_data.get('characteristics', {})
        if isinstance(characteristics, dict) and characteristics:
            details_html = ""
            for key in sorted(characteristics.keys()): # Сортируем для консистентности
                value = characteristics[key]
                details_html += f"<b>{html.escape(str(key))}:</b> {html.escape(str(value))}<br>"
            self.details_output.setHtml(details_html)
        else:
            self.details_output.setText("Характеристики не найдены или не были собраны.")

        # Загрузка изображения
        image_url = item_data.get('image_url')
        if image_url and AIOHTTP_AVAILABLE:
             logger.debug(f"Запуск загрузки изображения: {image_url}")
             try:
                 # Получаем текущий event loop (qasync должен его предоставить)
                 try: loop = asyncio.get_running_loop()
                 except RuntimeError: loop = asyncio.get_event_loop()

                 if loop.is_running():
                    # Создаем задачу для асинхронной загрузки
                    loop.create_task(self.fetch_and_display_image(image_url))
                 else:
                    logger.warning("Цикл asyncio не запущен, не могу загрузить изображение.")
                    self.image_preview_label.setText("Ошибка\nцикла")
             except Exception as e:
                 logger.error(f"Ошибка при получении event loop для загрузки изображения: {e}", exc_info=True)
                 self.image_preview_label.setText("Ошибка\nцикла")
        elif not AIOHTTP_AVAILABLE:
             self.image_preview_label.setText("Библиотека\naiohttp\nне найдена")
             self.image_preview_label.setStyleSheet("QLabel { background-color : #f0f0f0; border: 1px solid grey; color: red; }")
        else:
             self.image_preview_label.setText("URL изображения\nотсутствует")


    async def fetch_and_display_image(self, url: str) -> None:
        if not AIOHTTP_AVAILABLE: return # Двойная проверка

        # Создаем сессию, если ее нет или она закрыта
        if not self.http_session or self.http_session.closed:
            try:
                loop = asyncio.get_event_loop()
                timeout = aiohttp.ClientTimeout(total=20, connect=10) # Таймаут 20 сек
                self.http_session = aiohttp.ClientSession(loop=loop, timeout=timeout)
                logger.info("Создана новая сессия aiohttp.")
            except Exception as session_err:
                logger.error(f"Не удалось создать сессию aiohttp: {session_err}", exc_info=True)
                QTimer.singleShot(0, lambda: self.image_preview_label.setText("Ошибка\nсессии") if self.image_preview_label else None)
                self.current_pixmap = None
                return

        if not self.http_session: # Если сессия все еще не создана
             QTimer.singleShot(0, lambda: self.image_preview_label.setText("Ошибка\nсессии") if self.image_preview_label else None)
             self.current_pixmap = None
             return

        # Асинхронная загрузка
        try:
            logger.debug(f"Загрузка изображения: {url}")
            async with self.http_session.get(url) as response:
                response.raise_for_status() # Проверка на HTTP ошибки (4xx, 5xx)
                image_data = await response.read()
                logger.debug(f"Изображение загружено ({len(image_data)} байт).")

                pixmap = QPixmap()
                loaded = pixmap.loadFromData(image_data)

                if loaded:
                    self.current_pixmap = pixmap # Сохраняем оригинал
                    # Обновляем QLabel в GUI потоке через QTimer
                    QTimer.singleShot(0, self.rescale_preview_image)
                else:
                    logger.warning(f"Ошибка декодирования изображения: {url}")
                    self.current_pixmap = None
                    QTimer.singleShot(0, lambda: self.image_preview_label.setText("Ошибка\nформата") if self.image_preview_label else None)

        except asyncio.TimeoutError:
             logger.warning(f"Таймаут при загрузке изображения {url}")
             self.current_pixmap = None
             QTimer.singleShot(0, lambda: self.image_preview_label.setText("Таймаут\nзагрузки") if self.image_preview_label else None)
        except aiohttp.ClientResponseError as http_err:
             logger.warning(f"Ошибка HTTP {http_err.status} при загрузке {url}: {http_err.message}")
             self.current_pixmap = None
             QTimer.singleShot(0, lambda code=http_err.status: self.image_preview_label.setText(f"Ошибка {code}") if self.image_preview_label else None)
        except aiohttp.ClientConnectionError as conn_err:
             logger.warning(f"Ошибка соединения при загрузке {url}: {conn_err}")
             self.current_pixmap = None
             QTimer.singleShot(0, lambda: self.image_preview_label.setText("Ошибка\nсоединения") if self.image_preview_label else None)
        except Exception as e:
            logger.error(f"Неожиданная ошибка при загрузке изображения {url}", exc_info=True)
            self.current_pixmap = None
            QTimer.singleShot(0, lambda: self.image_preview_label.setText("Ошибка\nзагрузки") if self.image_preview_label else None)


    def rescale_preview_image(self) -> None:
        """Масштабирует self.current_pixmap под размер self.image_preview_label."""
        if self.current_pixmap and self.image_preview_label and self.image_preview_label.isVisible():
            label_size = self.image_preview_label.size()
            # Масштабируем, только если размер виджета > 1x1
            if label_size.width() > 1 and label_size.height() > 1:
                scaled_pixmap = self.current_pixmap.scaled(
                    label_size,
                    Qt.KeepAspectRatio,     # Сохранять пропорции
                    Qt.SmoothTransformation # Плавное масштабирование
                )
                self.image_preview_label.setPixmap(scaled_pixmap)
                self.image_preview_label.setStyleSheet("") # Убираем стиль ошибки/загрузки
            # else: # Не масштабируем, если виджет слишком мал или скрыт
            #     logger.debug("Размер QLabel для превью некорректен, масштабирование пропущено.")


    def resizeEvent(self, event: QResizeEvent) -> None:
        """Перемасштабирует изображение при изменении размера окна."""
        super().resizeEvent(event)
        # Вызываем перерисовку с небольшой задержкой, чтобы избежать лишних вызовов
        QTimer.singleShot(50, self.rescale_preview_image)


    def open_product_link(self, item: QTableWidgetItem) -> None:
        if not item: return
        selected_row = item.row()
        name_cell = self.results_table.item(selected_row, 0)
        if not name_cell: return

        data_item = name_cell.data(Qt.UserRole)
        if isinstance(data_item, dict):
            url = data_item.get('link')
            if url and url != "N/A" and url.startswith('http'):
                try:
                    logger.info(f"Открытие ссылки в браузере: {url}")
                    webbrowser.open(url)
                except Exception as e:
                    err_msg = f"Не удалось открыть ссылку {url}: {e}"
                    logger.error(err_msg)
                    QMessageBox.warning(self, "Ошибка открытия ссылки", err_msg)
            else:
                logger.warning(f"Двойной клик: Некорректная ссылка для строки {selected_row}. URL: {url}")
                QMessageBox.information(self, "Нет ссылки", "Ссылка на товар отсутствует или некорректна.")


    def save_results(self, results: List[Dict[str, Any]]) -> None:
        output_file = self.outfile_input.text().strip()
        selected_format = self.format_combo.currentText()

        if not results:
            logger.info("Нет данных для сохранения.")
            return

        if not output_file:
            QMessageBox.warning(self, "Файл не указан", "Укажите файл для сохранения.")
            self.browse_file() # Предлагаем выбрать
            output_file = self.outfile_input.text().strip()
            if not output_file: return # Если все равно не выбрали

        logger.info(f"Сохранение {len(results)} результатов в: {output_file} (Формат: {selected_format})...")
        self.statusBar().showMessage('Сохранение результатов...')
        QApplication.processEvents() # Обновляем GUI

        try:
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
                logger.info(f"Создана директория: {output_dir}")

            if "JSON" in selected_format:
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(results, f, ensure_ascii=False, indent=4)

            elif "Excel" in selected_format:
                if not OPENPYXL_AVAILABLE: raise ImportError("Библиотека openpyxl не найдена.")
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Результаты парсинга"
                # Добавляем числовую цену и характеристики отдельными колонками
                headers = ['Название', 'Цена (строка)', 'Цена (число)', 'Ссылка', 'URL Изображения', 'Характеристики (JSON)']
                ws.append(headers)

                for item in results:
                    char_str = ""
                    char_data = item.get('characteristics')
                    if isinstance(char_data, dict) and char_data:
                        try: char_str = json.dumps(char_data, ensure_ascii=False, sort_keys=True)
                        except Exception: char_str = "Ошибка JSON"

                    raw_price = item.get('price', '')
                    numeric_price = try_convert_to_number(raw_price)
                    if not isinstance(numeric_price, (int, float)): numeric_price = None # Пустое значение для Excel

                    ws.append([
                        item.get('name', ''),
                        raw_price,
                        numeric_price,
                        item.get('link', ''),
                        item.get('image_url', ''),
                        char_str
                    ])
                # Автоподбор ширины (опционально, может замедлять на больших файлах)
                # for col_idx, column_cells in enumerate(ws.columns): ... (код автоподбора)
                wb.save(output_file)

            else: # CSV
                with open(output_file, 'w', newline='', encoding='utf-8-sig') as file: # utf-8-sig для Excel
                    writer = csv.writer(file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                    headers = ['Название', 'Цена', 'Ссылка', 'URL Изображения', 'Характеристики (JSON)']
                    writer.writerow(headers)
                    for item in results:
                        char_str = ""
                        char_data = item.get('characteristics')
                        if isinstance(char_data, dict) and char_data:
                             try: char_str = json.dumps(char_data, ensure_ascii=False, sort_keys=True)
                             except Exception: char_str = "Ошибка JSON"
                        writer.writerow([
                            item.get('name', ''), item.get('price', ''),
                            item.get('link', ''), item.get('image_url', ''),
                            char_str
                        ])

            logger.info(f"Результаты ({len(results)} шт.) успешно сохранены.")
            self.statusBar().showMessage('Результаты сохранены.', 5000)
            QMessageBox.information(self, "Сохранение успешно", f"{len(results)} записей сохранено в:\n{output_file}")

        except ImportError as import_err:
             error_msg = f"Ошибка сохранения: {import_err}"
             logger.error(error_msg)
             self.statusBar().showMessage('Ошибка сохранения!', 0)
             QMessageBox.critical(self, "Ошибка библиотеки", f"Не удалось сохранить файл: {import_err}\nУстановите библиотеку (например, `pip install openpyxl` для Excel).")
        except IOError as io_err:
             error_msg = f"Ошибка записи файла {output_file}: {io_err}"
             logger.error(error_msg, exc_info=True)
             self.statusBar().showMessage('Ошибка записи файла!', 0)
             QMessageBox.critical(self, "Ошибка записи", f"Не удалось записать в файл:\n{output_file}\nПроверьте права доступа.\nОшибка: {io_err}")
        except Exception as save_error:
             error_msg = f"Неожиданная ошибка при сохранении: {save_error}"
             logger.error(error_msg, exc_info=True)
             self.statusBar().showMessage('Ошибка сохранения!', 0)
             QMessageBox.critical(self, "Ошибка сохранения", f"Ошибка при сохранении:\n{save_error}")


    def _reset_gui_state(self, status_message: str = "Готов") -> None:
        """Сбрасывает состояние GUI к начальному."""
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.category_combo.setEnabled(True)
        self.max_pages_input.setEnabled(True)
        self.scrape_details_checkbox.setEnabled(True)
        self.outfile_input.setEnabled(True)
        self.browse_button.setEnabled(True)
        self.format_combo.setEnabled(True)

        self.is_scraper_running = False
        self.worker = None
        self.scraping_task = None

        self.progress_bar.setMaximum(100) # Сброс прогресс-бара
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(status_message.replace("!", ""))

        self.statusBar().showMessage(status_message, 5000)
        logger.info(f"Состояние GUI сброшено. Статус: {status_message}")


    def scraping_finished(self, results: List[Dict[str, Any]]) -> None:
        """Обработчик завершения работы ScraperWorker."""
        logger.info("Сигнал `finished` получен от воркера.")
        # Определяем, была ли остановка ручной (кнопка Stop была нажата и стала неактивной)
        was_stopped_manually = not self.stop_button.isEnabled() and not self.start_button.isEnabled()

        self.statusBar().showMessage('Парсинг завершен. Обработка результатов...', 0)
        self.all_scraper_results = results # Сохраняем все результаты
        logger.info(f"Получено {len(self.all_scraper_results)} результатов.")

        # --- Обновление списка характеристик для фильтра ---
        all_char_keys = set()
        if self.all_scraper_results:
            for item in self.all_scraper_results:
                if isinstance(item.get('characteristics'), dict):
                    all_char_keys.update(k for k in item['characteristics'].keys() if k) # Собираем все непустые ключи

        sorted_char_keys = sorted(list(all_char_keys))
        # Сохраняем текущий выбор фильтра
        current_filter_key = self.filter_char_name_combo.currentText()
        current_filter_value = self.filter_char_value_input.text()
        # Обновляем комбобокс
        self.filter_char_name_combo.clear()
        self.filter_char_name_combo.addItem("--- Любая ---")
        self.filter_char_name_combo.addItems(sorted_char_keys)
        # Восстанавливаем выбор, если возможно
        index = self.filter_char_name_combo.findText(current_filter_key)
        if index != -1:
             self.filter_char_name_combo.setCurrentIndex(index)
             self.filter_char_value_input.setText(current_filter_value) # Восстанавливаем и значение
        else: # Если старого ключа нет, сбрасываем
             self.filter_char_name_combo.setCurrentIndex(0)
             self.filter_char_value_input.clear()
        logger.info(f"Обновлен список характеристик для фильтра ({len(sorted_char_keys)} уникальных).")
        # --- Конец обновления фильтра ---

        # Сбрасываем фильтры и отображаем ВСЕ результаты
        self.reset_filter(display_results=False) # Сначала сброс полей
        self._display_results(self.all_scraper_results) # Отображение всех данных

        # Предлагаем сохранить, если не было ручной остановки И есть результаты
        if not was_stopped_manually and results:
            self.save_results(results)
        elif was_stopped_manually and results:
             logger.info("Парсинг остановлен вручную, автосохранение пропущено.")
             reply = QMessageBox.question(self, 'Сохранить?',
                                      f"Парсинг был остановлен. Собрано {len(results)} записей.\nСохранить их?",
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
             if reply == QMessageBox.Yes: self.save_results(results)
        elif not results:
             logger.info("Результатов нет, сохранение не требуется.")

        final_status = "Остановлено" if was_stopped_manually else "Завершено"
        self._reset_gui_state(final_status) # Сброс GUI в финальное состояние


    def apply_filter(self) -> None:
        if self.is_scraper_running:
            QMessageBox.warning(self, "Парсинг активен", "Дождитесь завершения или остановите парсинг.")
            return
        if not self.all_scraper_results:
            logger.info("Нет данных для фильтрации.")
            return

        filter_name = self.filter_name_input.text().strip().lower()
        filter_price_min = self.filter_price_min_input.value()
        filter_price_max = self.filter_price_max_input.value()
        filter_char_name = self.filter_char_name_combo.currentText()
        filter_char_value = self.filter_char_value_input.text().strip().lower()

        apply_char_filter = (filter_char_name != "--- Любая ---" and filter_char_value)
        apply_price_max_filter = (filter_price_max > 0) # Фильтр по макс. цене активен, если > 0

        conditions = []
        if filter_name: conditions.append(f"Название~'{filter_name}'")
        if filter_price_min > 0: conditions.append(f"Цена>={filter_price_min}")
        if apply_price_max_filter: conditions.append(f"Цена<={filter_price_max}")
        if apply_char_filter: conditions.append(f"Характеристика['{filter_char_name}']~'{filter_char_value}'")

        if not conditions:
            logger.info("Нет активных фильтров.")
            self.reset_filter() # Сбрасываем и показываем все
            return

        logger.info(f"[GUI] Применение фильтра: {', '.join(conditions)}")
        self.statusBar().showMessage('Применение фильтра...')
        QApplication.processEvents()

        filtered_results: List[Dict[str, Any]] = []
        for item in self.all_scraper_results:
            # Фильтр по названию
            if filter_name and filter_name not in item.get('name', '').lower(): continue

            # Фильтр по цене
            price_match = True
            if filter_price_min > 0 or apply_price_max_filter:
                price_val = try_convert_to_number(item.get('price', ''))
                if isinstance(price_val, (int, float)):
                    if filter_price_min > 0 and price_val < filter_price_min: price_match = False
                    if price_match and apply_price_max_filter and price_val > filter_price_max: price_match = False
                else: price_match = False # Нечисловая цена не проходит фильтр диапазона
                if not price_match: continue

            # Фильтр по характеристикам
            if apply_char_filter:
                item_chars = item.get('characteristics', {})
                char_actual_value = str(item_chars.get(filter_char_name, '')).lower() if isinstance(item_chars, dict) else ""
                if filter_char_value not in char_actual_value: continue

            # Если все фильтры пройдены
            filtered_results.append(item)

        count = len(filtered_results)
        logger.info(f"Найдено по фильтру: {count} элементов.")
        self._display_results(filtered_results) # Отображаем отфильтрованные
        self.statusBar().showMessage(f"Фильтр применен. Найдено: {count}", 5000)


    def reset_filter(self, display_results: bool = True) -> None:
        logger.info("Сброс фильтров...")
        self.filter_name_input.clear()
        self.filter_price_min_input.setValue(0)
        self.filter_price_max_input.setValue(0)
        if self.filter_char_name_combo.count() > 0: self.filter_char_name_combo.setCurrentIndex(0)
        self.filter_char_value_input.clear()

        if display_results:
            self._display_results(self.all_scraper_results) # Показываем все результаты
            self.statusBar().showMessage("Фильтры сброшены.", 3000)
        else:
            self.statusBar().showMessage("Поля фильтров очищены.", 3000)


    def start_scraping(self) -> None:
        if self.is_scraper_running:
            QMessageBox.warning(self, "Парсер запущен", "Процесс парсинга уже выполняется.")
            return

        if not PLAYWRIGHT_AVAILABLE:
            self.handle_error("Playwright не найден! Установите его (`pip install playwright && playwright install chromium`) и перезапустите.")
            return

        selected_category_name = self.category_combo.currentText()
        start_url = self.loaded_categories.get(selected_category_name)
        max_pages = self.max_pages_input.value()
        scrape_details_flag = self.scrape_details_checkbox.isChecked()

        if not start_url or selected_category_name.startswith("---"):
             QMessageBox.warning(self, "Категория не выбрана", "Пожалуйста, выберите категорию.")
             return

        # Подготовка GUI к запуску
        logger.info("="*20 + f" Новый запуск: {selected_category_name} " + "="*20)
        logger.info(f"Параметры: Макс.стр: {max_pages if max_pages > 0 else 'Все'}, Детали: {scrape_details_flag}, URL: {start_url}")
        self.update_log("="*10 + " Запуск парсинга " + "="*10) # Сообщение в лог GUI

        self.results_table.setRowCount(0) # Очистка таблицы
        self.results_label.setText("Результаты: 0")
        self.details_output.clear()
        self.image_preview_label.clear(); self.image_preview_label.setText("Ожидание данных...")
        self.current_pixmap = None
        self.reset_filter(display_results=False) # Сброс полей фильтра
        # Очищаем и комбобокс характеристик
        self.filter_char_name_combo.clear(); self.filter_char_name_combo.addItem("--- Любая ---")
        self.all_scraper_results = [] # Очищаем предыдущие результаты

        # Блокируем элементы управления
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.category_combo.setEnabled(False)
        self.max_pages_input.setEnabled(False)
        self.scrape_details_checkbox.setEnabled(False)
        self.outfile_input.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.format_combo.setEnabled(False)

        self.is_scraper_running = True
        self.statusBar().showMessage('Запуск парсера...')
        self.progress_bar.setMaximum(0) # Начальное состояние - "бегунок"
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Запуск...")

        # Создаем и запускаем воркер
        self.worker = ScraperWorker(start_url, max_pages, scrape_details=scrape_details_flag)
        self.worker.log_message.connect(self.update_log)
        self.worker.error.connect(self.handle_error)
        self.worker.finished.connect(self.scraping_finished)
        self.worker.progress.connect(self.update_progress)

        try:
            loop = asyncio.get_event_loop()
            self.scraping_task = loop.create_task(self.worker.run())
            logger.info("Асинхронная задача парсинга создана.")
        except Exception as task_err:
             self.handle_error(f"Критическая ошибка создания задачи: {task_err}")


    def stop_scraping(self) -> None:
        if self.worker and self.is_scraper_running:
            logger.warning("Получен сигнал остановки от пользователя.")
            self.statusBar().showMessage('Остановка парсера...')
            self.stop_button.setEnabled(False) # Блокируем кнопку "Стоп" сразу

            self.worker.stop() # Устанавливаем флаг is_running = False в воркере

            if self.scraping_task and not self.scraping_task.done():
                 # Отменяем задачу asyncio (это вызовет CancelledError внутри run)
                 cancelled = self.scraping_task.cancel()
                 logger.info(f"Попытка отмены задачи asyncio (успешно: {cancelled}). Ожидание завершения...")
                 # Обработка результатов и сброс GUI произойдет в scraping_finished
            else:
                 logger.info("Задача парсинга не найдена или уже завершена при попытке остановки.")
                 # Если задача уже завершилась сама по себе, но GUI не сбросился
                 if self.is_scraper_running:
                     self._reset_gui_state("Остановлено (задача не найдена)")
        else:
            logger.info("Нет активного процесса парсинга для остановки.")


    def closeEvent(self, event: QCloseEvent) -> None:
        logger.info("Получен сигнал закрытия окна.")
        self.save_settings() # Сохраняем настройки при выходе

        close_accepted = False
        if self.is_scraper_running:
            reply = QMessageBox.question(self, 'Подтверждение выхода',
                                         "Парсинг активен. Выйти?\n(Процесс будет остановлен)",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                logger.info("Выход подтвержден во время парсинга.")
                self.stop_scraping() # Инициируем остановку
                # Не ждем здесь завершения, просто выходим
                close_accepted = True
            else:
                logger.info("Выход отменен.")
                event.ignore() # Отменяем закрытие окна
                return
        else:
            close_accepted = True

        if close_accepted:
            logger.info("Закрытие приложения...")
            # Пытаемся закрыть сессию aiohttp
            if self.http_session and not self.http_session.closed:
                logger.info("Попытка закрытия сессии aiohttp...")
                try:
                    loop = asyncio.get_event_loop()
                    if loop.is_running():
                        # Планируем закрытие в цикле событий
                        asyncio.ensure_future(self.http_session.close(), loop=loop)
                        logger.info("Запланировано асинхронное закрытие сессии aiohttp.")
                        # Даем циклу немного времени на обработку
                        #loop.call_later(0.1, lambda: loop.stop() if loop.is_running() else None)
                    else:
                         logger.warning("Цикл asyncio остановлен, сессия aiohttp не закроется асинхронно.")
                except RuntimeError:
                     logger.warning("Цикл asyncio не найден/остановлен, сессия aiohttp не закроется асинхронно.")
                except Exception as e:
                     logger.error(f"Ошибка при планировании закрытия сессии aiohttp: {e}", exc_info=True)

            event.accept() # Разрешаем закрытие окна

# --- Запуск приложения с qasync ---
if __name__ == '__main__':
    # Проверка критических зависимостей перед запуском GUI
    print("--- Проверка зависимостей ---")
    critical_missing = False
    if not PLAYWRIGHT_AVAILABLE:
        print("\nКРИТИЧЕСКАЯ ОШИБКА: Playwright не найден.", file=sys.stderr)
        print("Пожалуйста, установите его:", file=sys.stderr)
        print("1. pip install playwright", file=sys.stderr)
        print("2. playwright install chromium", file=sys.stderr)
        critical_missing = True

    if 'qasync' not in sys.modules:
        print("\nКРИТИЧЕСКАЯ ОШИБКА: qasync не найден.", file=sys.stderr)
        print("Пожалуйста, установите его: pip install qasync", file=sys.stderr)
        critical_missing = True

    if critical_missing:
        # Можно показать MessageBox перед выходом, если QApplication уже можно создать
        try:
            temp_app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Критическая ошибка", "Отсутствуют необходимые библиотеки (Playwright или qasync).\nСмотрите вывод в консоли. Приложение будет закрыто.")
        except Exception as mb_err:
             print(f"Не удалось показать QMessageBox: {mb_err}", file=sys.stderr)
        sys.exit(1)

    # Информационные сообщения об опциональных библиотеках (уже логируются при импорте)
    # if not AIOHTTP_AVAILABLE: print("\nПРЕДУПРЕЖДЕНИЕ: aiohttp не найден...")
    # if not OPENPYXL_AVAILABLE: print("\nПРЕДУПРЕЖДЕНИЕ: openpyxl не найден...")
    print("--- Проверка завершена ---\n")

    # Инициализация приложения и event loop
    app = QApplication(sys.argv)
    try:
        loop = qasync.QEventLoop(app)
        asyncio.set_event_loop(loop)
    except Exception as loop_init_err:
         logger.critical(f"Критическая ошибка инициализации event loop qasync: {loop_init_err}", exc_info=True)
         QMessageBox.critical(None, "Ошибка запуска", f"Не удалось инициализировать event loop qasync:\n{loop_init_err}")
         sys.exit(1)

    logger.info("--- Запуск приложения Planeta-B Scraper ---")

    mainWin = ScraperApp()
    mainWin.show()

    exit_code = 0
    try:
        with loop: # Запускаем цикл событий
            logger.info("Запуск главного цикла событий...")
            loop.run_forever()
        logger.info("Главный цикл событий завершен.")
    except KeyboardInterrupt:
        logger.warning("Приложение прервано пользователем (Ctrl+C).")
        if mainWin.isVisible(): mainWin.close() # Пытаемся закрыть окно штатно
        exit_code = 1
    except Exception as main_loop_err:
         logger.critical(f"КРИТИЧЕСКАЯ ОШИБКА в главном цикле: {main_loop_err}", exc_info=True)
         try:
            QMessageBox.critical(mainWin, "Критическая ошибка", f"Ошибка в главном цикле приложения:\n{main_loop_err}\n\nСм. лог-файл {LOG_FILENAME_ONLY} для деталей.")
         except: pass
         exit_code = 1
    finally:
        logger.info("Завершение работы приложения...")
        # Блок очистки asyncio (закрытие оставшихся задач, закрытие цикла)
        if loop.is_running():
             logger.debug("Остановка цикла asyncio...")
             loop.stop()
        if not loop.is_closed():
            logger.info("Ожидание завершения фоновых задач asyncio...")
            try:
                # Собираем все задачи, кроме текущей
                tasks = asyncio.all_tasks(loop=loop)
                current = asyncio.current_task(loop=loop)
                tasks.discard(current) # Удаляем текущую задачу из набора
                if tasks:
                    logger.debug(f"Ожидающие задачи: {len(tasks)}")
                    # Даем им немного времени на завершение
                    loop.run_until_complete(asyncio.wait(tasks, timeout=1.5)) # 1.5 секунды
                    logger.debug("Ожидание задач завершено.")
                else: logger.debug("Нет ожидающих задач.")
            except RuntimeError as e: logger.warning(f"Ошибка при ожидании задач (цикл закрыт?): {e}")
            except Exception as e: logger.error(f"Неожиданная ошибка при ожидании задач: {e}", exc_info=True)
            finally:
                 if not loop.is_closed():
                     loop.close()
                     logger.info("Цикл asyncio закрыт.")
        else: logger.info("Цикл asyncio уже был закрыт.")

    logger.info(f"--- Приложение завершено с кодом {exit_code} ---")
    sys.exit(exit_code)