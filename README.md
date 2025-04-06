# Парсер Planeta-B

Графическое приложение на Python и PyQt5 для парсинга данных о товарах с сайта planeta-b.ru.

## Возможности

*   Выбор категории из `categories.json`.
*   Настройка максимального количества страниц.
*   Опциональный сбор детальных характеристик.
*   Отображение результатов в таблице.
*   Фильтрация результатов.
*   Предпросмотр изображений (требует `aiohttp`).
*   Сохранение в CSV, JSON, Excel (требует `openpyxl`).

## Установка зависимостей

```bash
pip install PyQt5 qasync playwright aiohttp openpyxl
playwright install chromium

##Запуск
```bash
python planeta_b_scraper.py

##Файл категорий
Отредактируйте файл categories.json, чтобы добавить нужные вам категории и ссылки на них.
```bash
{
  "--- Выберите категорию ---": "",
  "Уличные IP камеры": "https://planeta-b.ru/ulichnye-ip-kamery.html",
  "Внутренние IP камеры": "https://planeta-b.ru/vnutrennie-ip-kamery.html",
  "Видеорегистраторы IP (NVR)": "https://planeta-b.ru/videoregistratori-ip-nvr.html",
  "СКУД (Общая)": "https://planeta-b.ru/catalog/sistemy-kontrolya-dostupa/",
  "Биометрические терминалы": "https://planeta-b.ru/biometricheskie-terminaly.html",
  "Жесткие диски для видеонаблюдения": "https://planeta-b.ru/zhestkie-diski-dlya-videonabludeniya.html"
}
