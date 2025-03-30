# Автоматизация визовых заявлений 🛂

**Проект для автоматического формирования документов для рабочих виз в Россию**  
Решение генерирует полный пакет документов (заявление, анкету, ходатайство) по шаблонам компаний.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

## 🌟 Особенности
- Поддержка 4 компаний с уникальными шаблонами
- Автоматическая валидация вводимых данных
- Генерация документов в формате DOCX
- Интеллектуальная обработка дат и персональных данных
- Автоматическое создание структуры папок

## 🛠 Технологии
- **Core**: Python 3.10
- **DOCX Processing**: `python-docx`, `docxtpl`
- **Data Validation**: Регулярные выражения
- **Date Handling**: `datetime`, `calendar`

## 📁 Структура проекта

Visa_application_automatization/
├── templates/ # Шаблоны документов по компаниям

├── OUTPUT/ # Сгенерированные документы 

├── FUNCTIONS_AND_CLASSES/ # Основная логика 

│ ├── FUNCTIONS.py # Утилиты обработки данных 

│ ├── request_factory.py # Генератор заявлений 

│ ├── anketa_factory.py # Генератор анкет 

│ └── hodataistvo_factory.py # Генератор ходатайств 

└── REQUEST_ANKETA_BLANK.py # Точка входа 

Copy


## ⚙️ Установка
1. Клонируйте репозиторий:
```bash
git clone https://github.com/teamdlok/Visa_application_automatization.git
cd Visa_application_automatization

    Установите зависимости:

bash
Copy

pip install -r requirements.txt

🚀 Использование

    Запустите REQUEST_ANKETA_BLANK.py:

bash
Copy

python main.py

    Следуйте интерактивным подсказкам:

Copy

ВЫБЕРИТЕ КОМПАНИЮ 
1 - АВАНТА 
2 - ПРОФЕССИОНАЛ 
3 - ВИЗАР ВОСТОК 
4 - ТЭНФЭЙ 

Нужно ли заявление? 
(ДА)  (НЕТ)

    Вводите данные по запросам системы. Пример:

Copy

ВВЕДИТЕ ФАМИЛИЮ: IVANOV
ВВЕДИТЕ ИМЯ: ALEXEY
ВВЕДИТЕ ДЕНЬ РОЖДЕНИЯ: 15
...

    Результаты сохраняются в папке OUTPUT/

🔧 Основные функции
Валидация данных

    Автоматическая проверка формата дат

    Контроль длины полей (паспорт, РНП)

    Преобразование месяцев (AUG → 08)

Генерация документов

    Динамическое заполнение шаблонов

    Автоматическое форматирование:

        Выравнивание текста

        Настройки шрифтов (Arial Narrow, 11pt)

        Обработка длинных адресов
