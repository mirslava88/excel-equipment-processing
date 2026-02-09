"""Создание тестового Excel файла с разными датами для проверки логики устаревания"""
import pandas as pd
from datetime import datetime

# Тестовые данные с разными годами
test_data = {
    "Дата отражения проводки (когда вносились данные в этот файл)": [
        datetime(2025, 1, 15),  # 1 год - не устарело
        datetime(2023, 6, 20),  # 3 года - не устарело (граничное)
        datetime(2020, 3, 10),  # 6 лет - устарело
        datetime(2019, 8, 1),   # 7 лет - устарело
        datetime(2018, 12, 25), # 8 лет - устарело
        datetime(2016, 4, 5),   # 10 лет - КРИТИЧНО
        datetime(2015, 7, 30),  # 11 лет - КРИТИЧНО
        datetime(2012, 2, 14),  # 14 лет - КРИТИЧНО
        "invalid_date",         # Невалидная дата
        None,                   # Пустая дата
    ],
    "Серийный номер": [
        "SN001", "SN002", "SN003", "SN004", "SN005",
        "SN006", "SN007", "SN008", "SN009", "SN010"
    ],
    "Модель": [
        "Laptop", "Desktop", "Monitor", "Printer", "Router",
        "Switch", "Server", "Tablet", "Phone", "Camera"
    ]
}

# Создаем DataFrame
df = pd.DataFrame(test_data)

# Сохраняем в Excel
output_path = "test_equipment.xlsx"
df.to_excel(output_path, index=False, sheet_name="Заполнить")

print(f"✓ Создан тестовый файл: {output_path}")
print(f"  Строк: {len(df)}")
print("\nДаты в файле:")
for i, date in enumerate(df["Дата отражения проводки (когда вносились данные в этот файл)"], 1):
    print(f"  {i}. {date} → {df['Серийный номер'][i-1]}")
