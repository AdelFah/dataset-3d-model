#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ПРИОРИТИЗАЦИЯ АВТОМОБИЛЕЙ v3 — на DuckDB.
Считает агрегаты по 5-ГБ CSV напрямую (без pandas-loop), затем
формирует красивый Excel на русском.
"""
import sys
import math
from pathlib import Path

import duckdb
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

DATA_DIR = Path(__file__).parent
CSV_FILE = DATA_DIR / "drom_archive_2007-2025_full_27-07-2025.csv"
OUTPUT_FILE = DATA_DIR / "Приоритизация_автомобилей.xlsx"

# Коэффициенты инфляции к уровню 2025
INFLATION_FACTORS = {
    2007: 3.55, 2008: 3.05, 2009: 2.84, 2010: 2.62, 2011: 2.47, 2012: 2.32,
    2013: 2.20, 2014: 2.04, 2015: 1.78, 2016: 1.69, 2017: 1.65, 2018: 1.59,
    2019: 1.53, 2020: 1.47, 2021: 1.36, 2022: 1.18, 2023: 1.10, 2024: 1.04,
    2025: 1.00,
}

RUSSIAN_BRANDS = {
    'lada', 'vaz', 'ваз', 'лада', 'uaz', 'уаз', 'gaz', 'газ', 'kamaz', 'камаз',
    'volga', 'волга', 'moskvich', 'москвич', 'niva', 'нива', 'oka', 'ока',
    'izh', 'иж', 'azlk', 'азлк', 'taz', 'paz', 'паз', 'maz', 'маз', 'liaz',
    'лиаз', 'zil', 'зил', 'tagaz', 'тагаз', 'bogdan', 'богдан', 'aurus',
    'аурус', 'evolute',
}

LOCALIZED_BRANDS = {
    'hyundai', 'kia', 'renault', 'volkswagen', 'skoda', 'nissan', 'toyota',
    'ford', 'chevrolet', 'datsun', 'mazda', 'mitsubishi', 'peugeot', 'citroen',
    'haval', 'geely', 'chery', 'great_wall', 'lifan', 'ssangyong', 'fiat',
    'audi', 'bmw', 'volvo', 'isuzu', 'iveco', 'man', 'scania', 'mercedes-benz',
    'opel', 'omoda', 'exeed', 'jaecoo', 'changan', 'jac', 'gac', 'faw',
    'foton', 'dongfeng', 'baw', 'tank', 'voyah',
}


def sql_list(items):
    return ', '.join(f"'{x}'" for x in items)


def build_aggregate_sql(parquet_path: Path, with_brand: bool):
    keys = ['origin', 'ctype']
    if with_brand:
        keys.insert(1, 'brand')
    return f"""
    SELECT
        {', '.join(keys)},
        count(*)          AS count,
        avg(price_adj)    AS avg_price_adj,
        avg(price)        AS avg_price,
        min(price)        AS min_price,
        max(price)        AS max_price,
        avg(age)          AS avg_age,
        avg(yr)           AS avg_year
    FROM read_parquet('{parquet_path}')
    GROUP BY {', '.join(keys)}
    """


def build_extract_sql(out_parquet: Path):
    """Один проход CSV → лёгкий parquet с уже посчитанными классификациями."""
    origin_case = (
        f"CASE "
        f"  WHEN lower(trim(\"Метка\")) IN ({sql_list(RUSSIAN_BRANDS)}) THEN 'Российские' "
        f"  WHEN lower(trim(\"Метка\")) IN ({sql_list(LOCALIZED_BRANDS)}) THEN 'Локализованные' "
        f"  WHEN \"Метка\" IS NULL OR trim(\"Метка\") = '' THEN 'Неизвестно' "
        f"  ELSE 'Импорт' END"
    )
    type_case = (
        "CASE "
        "  WHEN regexp_matches(lower(coalesce(\"Тип кузова\", '')), "
        "       '(грузов|самосвал|фургон|тягач|truck|рефрижератор|бортов|эвакуатор|манипулятор|автокран)') "
        "    THEN 'Грузовые' "
        "  WHEN regexp_matches(lower(coalesce(\"Тип кузова\", '')), '(автобус|микроавтобус|bus)') "
        "    THEN 'Пассажирские' "
        "  WHEN regexp_matches(lower(coalesce(\"Тип кузова\", '')), "
        "       '(седан|хэтчб|хетчб|универсал|купе|кабриолет|liftback|лифтбек|минивэн|внедорожник|кроссовер|пикап|родстер|тарга)') "
        "    THEN 'Легковые' "
        "  ELSE 'Другое' END"
    )
    infl_when = ' '.join(
        f"WHEN cast(\"Год\" as INT) = {y} THEN {f}" for y, f in INFLATION_FACTORS.items()
    )
    infl_case = f"CASE {infl_when} ELSE 1.0 END"

    return f"""
    COPY (
        SELECT
            {origin_case} AS origin,
            lower(trim("Метка")) AS brand,
            {type_case} AS ctype,
            cast("Год" as INT) AS yr,
            cast("Цена" AS DOUBLE) AS price,
            cast("Цена" AS DOUBLE) * {infl_case} AS price_adj,
            (2025 - cast("Год" as INT)) AS age
        FROM read_csv('{CSV_FILE}',
                      header=True, sample_size=2048,
                      ignore_errors=True, parallel=True,
                      types={{'Год': 'DOUBLE', 'Цена': 'DOUBLE',
                              'Метка': 'VARCHAR', 'Тип кузова': 'VARCHAR'}})
        WHERE "Цена" IS NOT NULL
          AND "Цена" BETWEEN 30000 AND 200000000
          AND "Год" IS NOT NULL
          AND cast("Год" as INT) BETWEEN 2000 AND 2025
    ) TO '{out_parquet}' (FORMAT PARQUET, COMPRESSION ZSTD)
    """


def add_priority(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        df['Балл приоритета'] = []
        return df

    origin_w = df['Отечественность'].map(
        {'Российские': 3.0, 'Локализованные': 2.0, 'Импорт': 1.0, 'Неизвестно': 0.5}
    ).fillna(0.5)
    type_w = df['Тип ТС'].map(
        {'Легковые': 3.0, 'Пассажирские': 2.0, 'Грузовые': 1.5, 'Другое': 0.5}
    ).fillna(0.5)

    cnt = df['Кол-во объявлений'].astype(float)
    cnt_lo, cnt_hi = np.log1p(cnt.min()), np.log1p(cnt.max())
    cnt_norm = (np.log1p(cnt) - cnt_lo) / max(cnt_hi - cnt_lo, 1e-9)

    price = df['Средняя цена, ₽ (с инфл.)'].astype(float)
    med = price.median()
    price_norm = 1.0 - (
        np.log(price.replace(0, np.nan) / med).abs() / 3.0
    ).clip(0, 1).fillna(0.5)

    age = df['Средний возраст, лет'].astype(float)
    age = age.fillna(age.max())
    age_min, age_max = age.min(), age.max()
    age_norm = 1.0 - (age - age_min) / max(age_max - age_min, 1e-9)

    score = (
        origin_w / 3.0 * 0.25 +
        type_w / 3.0 * 0.20 +
        cnt_norm * 0.30 +
        price_norm * 0.10 +
        age_norm * 0.15
    )
    df = df.copy()
    df['Балл приоритета'] = (score * 100).round(1)
    df = df.sort_values('Балл приоритета', ascending=False).reset_index(drop=True)
    df.insert(0, 'Ранг', range(1, len(df) + 1))
    return df


def rename_cols(df: pd.DataFrame, with_brand: bool) -> pd.DataFrame:
    mapping = {
        'origin': 'Отечественность',
        'brand': 'Бренд',
        'ctype': 'Тип ТС',
        'count': 'Кол-во объявлений',
        'avg_price_adj': 'Средняя цена, ₽ (с инфл.)',
        'avg_price': 'Средняя цена, ₽ (как есть)',
        'min_price': 'Мин. цена, ₽',
        'max_price': 'Макс. цена, ₽',
        'avg_age': 'Средний возраст, лет',
        'avg_year': 'Средний год выпуска',
    }
    df = df.rename(columns=mapping)
    order = ['Отечественность']
    if with_brand:
        order.append('Бренд')
    order += [
        'Тип ТС', 'Кол-во объявлений',
        'Средняя цена, ₽ (с инфл.)',
        'Средняя цена, ₽ (как есть)',
        'Мин. цена, ₽', 'Макс. цена, ₽',
        'Средний возраст, лет', 'Средний год выпуска',
    ]
    return df[order]


def write_excel(df_summary, df_brand, total_rows):
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill('solid', fgColor='1F4E79')
    header_font = Font(bold=True, color='FFFFFF', size=11)
    border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF'),
    )
    widths = {
        'Ранг': 6, 'Отечественность': 16, 'Бренд': 18, 'Тип ТС': 14,
        'Кол-во объявлений': 16,
        'Средняя цена, ₽ (с инфл.)': 22,
        'Медиана цены, ₽ (с инфл.)': 22,
        'Средняя цена, ₽ (как есть)': 22,
        'Мин. цена, ₽': 14, 'Макс. цена, ₽': 16,
        'Средний возраст, лет': 18, 'Средний год выпуска': 18,
        'Балл приоритета': 14,
    }

    def write_sheet(name, df):
        ws = wb.create_sheet(name)
        if df.empty:
            ws.append(['нет данных'])
            return

        cols = list(df.columns)
        ws.append(cols)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
        ws.row_dimensions[1].height = 38
        ws.freeze_panes = 'A2'

        money_cols = {c for c in cols if 'цена' in c.lower()}
        int_cols = {'Ранг', 'Кол-во объявлений'}
        float1_cols = {'Балл приоритета', 'Средний возраст, лет', 'Средний год выпуска'}

        for _, row in df.iterrows():
            excel_row = []
            for c in cols:
                v = row[c]
                if pd.isna(v):
                    excel_row.append(None)
                elif c in int_cols:
                    excel_row.append(int(v))
                else:
                    excel_row.append(float(v) if c in money_cols or c in float1_cols else v)
            ws.append(excel_row)

        for col_idx, c in enumerate(cols, start=1):
            if c in money_cols:
                fmt = '#,##0\\ ₽'
            elif c in int_cols:
                fmt = '#,##0'
            elif c in float1_cols:
                fmt = '0.0'
            else:
                fmt = None
            for row_num in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.border = border
                if fmt:
                    cell.number_format = fmt
                if c in ('Отечественность', 'Тип ТС', 'Бренд'):
                    cell.alignment = Alignment(horizontal='left')
                else:
                    cell.alignment = Alignment(horizontal='right')

        if 'Балл приоритета' in cols:
            score_col_idx = cols.index('Балл приоритета') + 1
            score_letter = get_column_letter(score_col_idx)
            rng = f'{score_letter}2:{score_letter}{ws.max_row}'
            rule = ColorScaleRule(
                start_type='min', start_color='F8696B',
                mid_type='percentile', mid_value=50, mid_color='FFEB84',
                end_type='max', end_color='63BE7B',
            )
            ws.conditional_formatting.add(rng, rule)

        for col_idx, c in enumerate(cols, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(c, 14)
        ws.auto_filter.ref = ws.dimensions

    # Лист «Описание»
    ws_info = wb.create_sheet('Описание')
    info_rows = [
        ['Файл', 'Приоритизация автомобилей по архиву DROM (2007–2025)'],
        ['Источник', CSV_FILE.name],
        ['Размер источника', f'{CSV_FILE.stat().st_size / (1024**3):.2f} ГБ'],
        ['Строк после фильтра', f'{int(total_rows):,}'.replace(',', ' ')],
        ['', ''],
        ['Листы', ''],
        ['  Сводка', 'Категории по уровню локализации × типу ТС'],
        ['  Бренды', 'Главный список для выгрузки: бренд × тип, отсортирован по приоритету'],
        ['  Топ-100', 'Топ-100 строк из листа «Бренды»'],
        ['', ''],
        ['Категории отечественности', ''],
        ['  Российские', 'Lada, UAZ, ГАЗ, КамАЗ, Москвич, Aurus и т. п.'],
        ['  Локализованные', 'Бренды с производством в РФ: Hyundai, Kia, VW, Renault, Toyota, Haval, Geely и т. п.'],
        ['  Импорт', 'Все остальные бренды'],
        ['', ''],
        ['Веса в формуле приоритета', ''],
        ['  Отечественность', '25 %'],
        ['  Тип ТС', '20 %'],
        ['  Кол-во объявлений (log-нормировка)', '30 %'],
        ['  Средняя цена (близость к медиане)', '10 %'],
        ['  Средний возраст (моложе → выше)', '15 %'],
        ['', ''],
        ['Цены', 'приведены к уровню 2025 года через коэффициенты ИПЦ'],
        ['Балл приоритета', 'от 0 до 100; чем выше — тем выше приоритет для выгрузки'],
    ]
    for r in info_rows:
        ws_info.append(r)
    ws_info.column_dimensions['A'].width = 36
    ws_info.column_dimensions['B'].width = 80
    for r in range(1, ws_info.max_row + 1):
        ws_info.cell(row=r, column=1).font = Font(bold=True, color='1F4E79')
        for c in (1, 2):
            ws_info.cell(row=r, column=c).alignment = Alignment(vertical='top', wrap_text=True)

    write_sheet('Сводка', df_summary)
    write_sheet('Бренды', df_brand)
    write_sheet('Топ-100', df_brand.head(100))

    wb.save(OUTPUT_FILE)


def main():
    if not CSV_FILE.exists():
        print(f"❌ Не найден: {CSV_FILE}")
        sys.exit(1)

    print(f"📂 {CSV_FILE.name}  ({CSV_FILE.stat().st_size/1024**3:.2f} ГБ)")

    con = duckdb.connect()
    con.execute("PRAGMA threads=4")
    con.execute("PRAGMA memory_limit='4GB'")

    parquet_path = DATA_DIR / "_filtered.parquet"
    if not parquet_path.exists():
        print("⚙️  Шаг 1/2: CSV → parquet (один проход)...")
        con.execute(build_extract_sql(parquet_path))
        print(f"   parquet: {parquet_path.stat().st_size/1024/1024:.1f} МБ")
    else:
        print(f"⚙️  Использую готовый parquet ({parquet_path.stat().st_size/1024/1024:.1f} МБ)")

    print("📊 Шаг 2/2: агрегаты по parquet...")
    df_summary_raw = con.execute(build_aggregate_sql(parquet_path, with_brand=False)).fetch_df()
    df_brand_raw = con.execute(build_aggregate_sql(parquet_path, with_brand=True)).fetch_df()

    total_rows = df_summary_raw['count'].sum()
    print(f"\n✓ строк в выборке: {int(total_rows):,}")
    print(f"✓ категорий-сводки: {len(df_summary_raw)}")
    print(f"✓ бренд-категорий: {len(df_brand_raw)}")

    # фильтр редких групп
    df_summary = df_summary_raw[df_summary_raw['count'] >= 50].copy()
    df_brand = df_brand_raw[df_brand_raw['count'] >= 100].copy()

    df_summary = rename_cols(df_summary, with_brand=False)
    df_brand = rename_cols(df_brand, with_brand=True)

    df_summary = add_priority(df_summary)
    df_brand = add_priority(df_brand)

    print("\n📥 Сохраняю Excel...")
    write_excel(df_summary, df_brand, total_rows)
    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"✅ {OUTPUT_FILE} ({size_kb:.1f} КБ)")
    print(f"   Сводка: {len(df_summary)} строк | Бренды: {len(df_brand)} строк")


if __name__ == '__main__':
    main()
