#!/bin/bash
# Скрипт для запуска проекта сверки оборудования на MacOS

cd "$(dirname "$0")"

# Проверка наличия виртуального окружения
if [ ! -d ".venv" ]; then
  echo "Создаю виртуальное окружение..."
  python3 -m venv .venv
fi

# Активация виртуального окружения
source .venv/bin/activate

# Установка зависимостей
pip install --upgrade pip
pip install -r requirements.txt

# Освобождаем порт, если занят
lsof -ti:8001 | xargs kill -9 2>/dev/null || true

echo ""
echo "  Сервер запускается: http://127.0.0.1:8001"
echo "  Для остановки нажмите Ctrl+C"
echo ""

# Запуск приложения
.venv/bin/uvicorn app.main:app --reload --port 8001
