#!/usr/bin/env python3
"""
Тестовый скрипт для проверки чтения листов Excel
"""
import sys
import os

sys.path.insert(0, '/Users/mirslava/Desktop/myapps')

from app.excel_logic import get_engine, get_sheet_names, get_columns, auto_detect_columns

# Путь к тестовому файлу
test_file = "/Users/mirslava/Desktop/myapps/ФОП_приоритетная поддержка 30-01.xlsx"

print("=" * 60)
print("ТЕСТИРОВАНИЕ ЧТЕНИЯ EXCEL ФАЙЛА")
print("=" * 60)
print(f"\nФайл: {test_file}")
print(f"Существует: {os.path.exists(test_file)}")

if not os.path.exists(test_file):
    print("\n❌ ФАЙЛ НЕ НАЙДЕН!")
    print("Пожалуйста, укажите корректный путь к файлу")
    sys.exit(1)

print(f"Размер: {os.path.getsize(test_file) / 1024:.2f} KB")

# Определяем engine
engine = get_engine(test_file)
print(f"Engine: {engine}")

# Пробуем прочитать листы
print("\n" + "-" * 60)
print("ЧТЕНИЕ ЛИСТОВ:")
print("-" * 60)

try:
    sheets = get_sheet_names(test_file, engine)
    print(f"✓ Найдено листов: {len(sheets)}")
    for i, sheet in enumerate(sheets, 1):
        print(f"  {i}. {sheet}")
    
    # Пробуем прочитать столбцы первого листа
    if sheets:
        print("\n" + "-" * 60)
        print(f"ЧТЕНИЕ СТОЛБЦОВ ПЕРВОГО ЛИСТА: '{sheets[0]}'")
        print("-" * 60)
        try:
            cols = get_columns(test_file, engine, sheets[0])
            print(f"✓ Найдено столбцов: {len(cols)}")
            for i, col in enumerate(cols[:10], 1):  # Первые 10
                print(f"  {i}. {col}")
            if len(cols) > 10:
                print(f"  ... и ещё {len(cols) - 10} столбцов")
            
            # Автодетект
            print("\n" + "-" * 60)
            print("АВТОДЕТЕКТ СТОЛБЦОВ:")
            print("-" * 60)
            detected = auto_detect_columns(cols)
            print(f"Серийный номер: {detected['serial'] or '❌ не найден'}")
            print(f"Дата: {detected['date'] or '❌ не найден'}")
            
        except Exception as e:
            print(f"❌ Ошибка чтения столбцов: {e}")
            import traceback
            traceback.print_exc()
    
except Exception as e:
    print(f"❌ ОШИБКА: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)
print("ТЕСТ ЗАВЕРШЁН")
print("=" * 60)
