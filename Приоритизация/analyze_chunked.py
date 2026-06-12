#!/usr/bin/env python3
"""
ПРИОРИТИЗАЦИЯ АВТОМОБИЛЕЙ - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ
Читает большой файл по частям для экономии памяти!

drom_archive_2007-2025_full_27-07-2025.csv
"""

import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sys
import gc

print("""
╔════════════════════════════════════════════════════════════════════════════╗
║                                                                            ║
║   🚗 ПРИОРИТИЗАЦИЯ АВТОМОБИЛЕЙ - ОПТИМИЗИРОВАННАЯ ВЕРСИЯ                 ║
║                                                                            ║
║   Файл читается по частям (chunks) для экономии памяти!                  ║
║                                                                            ║
╚════════════════════════════════════════════════════════════════════════════╝
""")

# Найти файл
possible_paths = [
    Path("drom_archive_2007-2025_full_27-07-2025.csv"),
    Path.home() / "Desktop" / "drom_archive_2007-2025_full_27-07-2025.csv",
]

for p in Path(".").glob("drom_archive*.csv"):
    possible_paths.insert(0, p)

DATA_FILE = None
for path in possible_paths:
    if path.exists():
        DATA_FILE = path
        break

if DATA_FILE is None:
    print("❌ Файл drom_archive_2007-2025_full_27-07-2025.csv не найден!")
    print("\nРекомендация: скопируйте этот скрипт в папку с файлом CSV")
    sys.exit(1)

print(f"✓ Файл: {DATA_FILE}")
print(f"✓ Размер: {DATA_FILE.stat().st_size / (1024**3):.2f} ГБ")

OUTPUT_FILE = DATA_FILE.parent / "prioritized_cars_export.xlsx"

# ============================================================================
# НАСТРОЙКИ
# ============================================================================

INFLATION_FACTORS = {
    2007: 0.57, 2008: 0.64, 2009: 0.68, 2010: 0.70, 2011: 0.73, 2012: 0.76,
    2013: 0.79, 2014: 0.83, 2015: 0.87, 2016: 0.88, 2017: 0.91, 2018: 1.0,
    2019: 1.065, 2020: 1.098, 2021: 1.188, 2022: 1.582, 2023: 1.754, 2024: 1.85, 2025: 1.90,
}

DOMESTIC_BRANDS = {
    'лада', 'вав', 'уаз', 'камаз', 'волга', 'ока', 'москвич', 'иж',
    'газ', 'азлк', 'нива', 'таврия', 'богдан', 'чери', 'ваз'
}

# ============================================================================
# ФУНКЦИИ
# ============================================================================

def extract_brand(car_name):
    """Извлечь марку из названия машины"""
    if pd.isna(car_name):
        return 'Неизвестно'
    
    name_lower = str(car_name).lower().split()
    
    for brand in DOMESTIC_BRANDS:
        if brand in name_lower:
            return brand.upper()
    
    if name_lower:
        return name_lower[0].upper()
    
    return 'Неизвестно'

def is_domestic(brand):
    """Определить отечественная ли машина"""
    if pd.isna(brand):
        return 'неизвестно'
    
    brand_lower = str(brand).lower()
    
    if any(domestic in brand_lower for domestic in DOMESTIC_BRANDS):
        return 'да'
    
    return 'нет'

def determine_car_type(body_type):
    """Определить тип машины"""
    if pd.isna(body_type):
        return 'Неизвестно'
    
    body_str = str(body_type).lower()
    
    if any(x in body_str for x in ['седан', 'хэтчбэк', 'универсал', 'купе', 'кабриолет', 'liftback']):
        return 'Легковые'
    
    if any(x in body_str for x in ['автобус', 'микроавтобус', 'минивэн', 'van', 'bus']):
        return 'Пассажирские'
    
    if any(x in body_str for x in ['грузовик', 'пикап', 'truck', 'фургон', 'самосвал']):
        return 'Грузовые'
    
    return 'Другие'

def adjust_price_for_inflation(price, year):
    """Привести цену к уровню 2025 года"""
    if pd.isna(price) or pd.isna(year):
        return np.nan
    
    try:
        year = int(year)
        price = float(price)
    except:
        return np.nan
    
    factor = INFLATION_FACTORS.get(year, 1.0)
    return price * factor

def calculate_age(year):
    """Рассчитать возраст машины"""
    if pd.isna(year):
        return np.nan
    
    try:
        year = int(year)
    except:
        return np.nan
    
    return 2025 - year

# ============================================================================
# ЗАГРУЗКА И ОБРАБОТКА ДАННЫХ ПО ЧАСТЯМ
# ============================================================================

print("\n" + "="*70)
print("📊 Загрузка данных (по частям)...")
print("="*70)

# Аккумулятор для результатов
all_results = {}
chunk_size = 50000  # Читаем по 50 тыс строк за раз
total_rows = 0
processed_rows = 0

try:
    # Читаем файл по частям
    for chunk_num, chunk in enumerate(pd.read_csv(
        DATA_FILE,
        usecols=[0, 1, 4, 20],
        names=['car_name', 'year', 'price', 'body_type'],
        skiprows=1,
        chunksize=chunk_size,
        dtype={'year': 'Int64', 'price': 'float32'},
        on_bad_lines='skip',
        engine='python'
    )):
        total_rows += len(chunk)
        
        print(f"  Обработка части {chunk_num + 1} ({len(chunk):,} строк)...")
        
        # Очистка данных для этой части
        chunk = chunk.dropna(subset=['year', 'price'])
        chunk = chunk[(chunk['price'] > 50000) & (chunk['price'] < 100000000)]
        chunk = chunk[(chunk['year'] >= 2000) & (chunk['year'] <= 2025)]
        
        if len(chunk) == 0:
            print("    → нет подходящих данных в этой части")
            continue
        
        # Анализ
        chunk['brand'] = chunk['car_name'].apply(extract_brand)
        chunk['domestic'] = chunk['brand'].apply(is_domestic)
        chunk['type'] = chunk['body_type'].apply(determine_car_type)
        chunk['price_adjusted'] = chunk.apply(
            lambda row: adjust_price_for_inflation(row['price'], row['year']),
            axis=1
        )
        chunk['age'] = chunk['year'].apply(calculate_age)
        
        # Аккумулирование результатов
        for idx, row in chunk.iterrows():
            key = (row['domestic'], row['type'])
            
            if key not in all_results:
                all_results[key] = {
                    'count': 0,
                    'prices': [],
                    'ages': [],
                    'min_price': float('inf'),
                    'max_price': float('-inf')
                }
            
            all_results[key]['count'] += 1
            all_results[key]['prices'].append(row['price_adjusted'])
            all_results[key]['ages'].append(row['age'])
            all_results[key]['min_price'] = min(all_results[key]['min_price'], row['price'])
            all_results[key]['max_price'] = max(all_results[key]['max_price'], row['price'])
        
        processed_rows += len(chunk)
        
        # Очистка памяти
        del chunk
        gc.collect()
        
        print(f"    ✓ Всего обработано: {processed_rows:,} строк")

except Exception as e:
    print(f"❌ Ошибка при чтении: {e}")
    print("\nПопытка использовать другой движок...")
    
    # Пробуем с другим движком
    try:
        for chunk_num, chunk in enumerate(pd.read_csv(
            DATA_FILE,
            usecols=[0, 1, 4, 20],
            names=['car_name', 'year', 'price', 'body_type'],
            skiprows=1,
            chunksize=chunk_size,
            dtype={'year': 'Int64', 'price': 'float32'},
            on_bad_lines='skip'
        )):
            total_rows += len(chunk)
            
            print(f"  Обработка части {chunk_num + 1} ({len(chunk):,} строк)...")
            
            chunk = chunk.dropna(subset=['year', 'price'])
            chunk = chunk[(chunk['price'] > 50000) & (chunk['price'] < 100000000)]
            chunk = chunk[(chunk['year'] >= 2000) & (chunk['year'] <= 2025)]
            
            if len(chunk) == 0:
                continue
            
            chunk['brand'] = chunk['car_name'].apply(extract_brand)
            chunk['domestic'] = chunk['brand'].apply(is_domestic)
            chunk['type'] = chunk['body_type'].apply(determine_car_type)
            chunk['price_adjusted'] = chunk.apply(
                lambda row: adjust_price_for_inflation(row['price'], row['year']),
                axis=1
            )
            chunk['age'] = chunk['year'].apply(calculate_age)
            
            for idx, row in chunk.iterrows():
                key = (row['domestic'], row['type'])
                
                if key not in all_results:
                    all_results[key] = {
                        'count': 0,
                        'prices': [],
                        'ages': [],
                        'min_price': float('inf'),
                        'max_price': float('-inf')
                    }
                
                all_results[key]['count'] += 1
                all_results[key]['prices'].append(row['price_adjusted'])
                all_results[key]['ages'].append(row['age'])
                all_results[key]['min_price'] = min(all_results[key]['min_price'], row['price'])
                all_results[key]['max_price'] = max(all_results[key]['max_price'], row['price'])
            
            processed_rows += len(chunk)
            del chunk
            gc.collect()
            
            print(f"    ✓ Всего обработано: {processed_rows:,} строк")
    
    except Exception as e2:
        print(f"❌ Ошибка и при втором способе: {e2}")
        sys.exit(1)

if not all_results:
    print("❌ После обработки данных не осталось!")
    sys.exit(1)

print(f"\n✓ Всего загружено: {total_rows:,} строк")
print(f"✓ Обработано: {processed_rows:,} строк")
print(f"✓ Групп для анализа: {len(all_results)}")

# ============================================================================
# ПОДГОТОВКА РЕЗУЛЬТАТОВ
# ============================================================================

print("\n" + "="*70)
print("📊 Подготовка результатов...")
print("="*70)

result_data = []

for (domestic, car_type), data in all_results.items():
    if data['count'] < 10:  # Мінімум 10 объявлений
        continue
    
    prices = np.array(data['prices'])
    ages = np.array(data['ages'])
    
    prices_clean = prices[~np.isnan(prices)]
    ages_clean = ages[~np.isnan(ages)]
    
    if len(prices_clean) == 0:
        continue
    
    result_data.append({
        'domestic': domestic,
        'type': car_type,
        'count': data['count'],
        'avg_price_adjusted': np.mean(prices_clean),
        'median_price_adjusted': np.median(prices_clean),
        'avg_age': np.mean(ages_clean) if len(ages_clean) > 0 else 0,
        'min_price': data['min_price'],
        'max_price': data['max_price']
    })

result_df = pd.DataFrame(result_data)

if len(result_df) == 0:
    print("❌ Нет данных для анализа!")
    sys.exit(1)

# Расчёт приоритета
domestic_score = result_df['domestic'].map({'да': 3, 'нет': 1, 'неизвестно': 0.5})
type_score = result_df['type'].map({'Легковые': 3, 'Пассажирские': 2, 'Грузовые': 1, 'Другие': 0})

count_norm = (result_df['count'] - result_df['count'].min()) / (result_df['count'].max() - result_df['count'].min() + 1)
price_norm = 1 - np.abs(np.log(result_df['avg_price_adjusted'] / result_df['avg_price_adjusted'].median())) / 4
price_norm = np.clip(price_norm, 0, 1)
age_norm = 1 - (result_df['avg_age'] - result_df['avg_age'].min()) / (result_df['avg_age'].max() - result_df['avg_age'].min() + 1)

result_df['priority_score'] = (
    domestic_score * 0.30 + type_score * 0.25 + count_norm * 0.25 +
    price_norm * 0.10 + age_norm * 0.10
)

result_df = result_df.sort_values('priority_score', ascending=False).reset_index(drop=True)
result_df['rank'] = range(1, len(result_df) + 1)

result = result_df[['rank', 'domestic', 'type', 'count', 'avg_price_adjusted', 
                    'median_price_adjusted', 'avg_age', 'min_price', 'max_price', 'priority_score']]

# ============================================================================
# ЭКСПОРТ В EXCEL
# ============================================================================

print("\n" + "="*70)
print("📥 Экспорт в Excel...")
print("="*70)

wb = Workbook()
ws = wb.active
ws.title = "Приоритизация"

headers = ['Ранг', 'Отечественные', 'Тип', 'Кол-во объявлений', 
           'Средняя цена (₽, инфляция 2025)', 'Медиана цены (₽)', 
           'Средний возраст (лет)', 'Мин. цена (₽)', 'Макс. цена (₽)', 'Приоритет']

ws.append(headers)

header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
header_font = Font(bold=True, color='FFFFFF', size=11)
border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                top=Side(style='thin'), bottom=Side(style='thin'))

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border

for idx, row in result.iterrows():
    ws.append([
        int(row['rank']),
        row['domestic'],
        row['type'],
        int(row['count']),
        f"{row['avg_price_adjusted']:,.0f}",
        f"{row['median_price_adjusted']:,.0f}",
        f"{row['avg_age']:.1f}",
        f"{row['min_price']:,.0f}",
        f"{row['max_price']:,.0f}",
        f"{row['priority_score']:.2f}"
    ])

for row_num in range(2, len(result) + 2):
    for col_num in range(1, 11):
        cell = ws.cell(row=row_num, column=col_num)
        cell.border = border
        if col_num > 3:
            cell.alignment = Alignment(horizontal='right')
        else:
            cell.alignment = Alignment(horizontal='center')
        
        if col_num == 10:
            priority = float(result.iloc[row_num-2]['priority_score'])
            if priority > 0.75:
                cell.fill = PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
            elif priority > 0.50:
                cell.fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
            else:
                cell.fill = PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')

widths = [6, 15, 15, 16, 28, 20, 18, 18, 18, 12]
for i, width in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = width

ws.row_dimensions[1].height = 30

wb.save(OUTPUT_FILE)

# РЕЗУЛЬТАТЫ
print(f"""
{'='*70}
✅ АНАЛИЗ ЗАВЕРШЁН!
{'='*70}

📊 РЕЗУЛЬТАТЫ:
  Всего категорий: {len(result)}
  Всего объявлений: {result['count'].sum():,}
  Среднее на категорию: {result['count'].mean():.0f}
  Ценовой диапазон: ₽{result['avg_price_adjusted'].min():,.0f} - ₽{result['avg_price_adjusted'].max():,.0f}

📁 Результат сохранён:
   {OUTPUT_FILE}

🎉 Откройте этот файл в Excel!
{'='*70}
""")

print("\nТОП-15 КАТЕГОРИЙ ДЛЯ ВЫГРУЗКИ:")
print("="*70)
print(result.head(15).to_string(index=False))

print("\n✅ Готово!")
