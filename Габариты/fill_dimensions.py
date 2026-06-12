#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Заполнение габаритов для MVP-1600 из Wikipedia — версия 2.

Что нового по сравнению с v1:
  * Парсер переписан: сначала пробует чистый wikitext через MediaWiki API
    (action=parse&prop=wikitext), оттуда тащит параметры |length= |width=
    |height= |wheelbase= |ground_clearance= шаблона Infobox automobile.
  * Fallback на HTML с BeautifulSoup: ищет в infobox <th>Label</th> и
    соседний <td> с мм-значением.
  * Поддержка ru-wiki параметров: длина/ширина/высота/колёсная база/клиренс.
  * Авто-сброс прогресса, если в нём 0 попаданий (значит был сломан).
  * Подробнее печатает диагностику при «не найдено».

Установка и запуск (Windows):
    pip install requests openpyxl beautifulsoup4
    python fill_dimensions.py
"""
import json
import re
import sys
import time
from pathlib import Path

import requests
from openpyxl import load_workbook

try:
    from bs4 import BeautifulSoup
except ImportError:
    print('❌ Нет beautifulsoup4. Поставь: pip install beautifulsoup4')
    sys.exit(1)

# =========================================================================
# КОНФИГУРАЦИЯ
# =========================================================================

SCRIPT_DIR = Path(__file__).parent
XLSX = SCRIPT_DIR / 'MVP_dataset_restored.xlsx'
PROGRESS = SCRIPT_DIR / '_dimensions_progress.json'
PARSER_VERSION = 2          # при изменении автоматически сбрасывает прогресс

SHEET = 'MVP top-1600'
HEADER_ROW = 2
DATA_START = 3
COL_MARK = 'Марка'
COL_MODEL = 'Модель'
COL_GEN = 'Поколение'
COL_LENGTH = 'Длина, мм'
COL_WIDTH = 'Ширина, мм'
COL_HEIGHT = 'Высота, мм'
COL_WHEELBASE = 'Колёсная база, мм'
COL_CLEARANCE = 'Клиренс, мм'
COL_SOURCE = 'Источник габаритов'
COL_NOTE = 'Комментарий'

RATE_LIMIT_SEC = 0.5
TIMEOUT = 25
HEADERS = {
    'User-Agent': 'cars_analysis/2.0 (educational; adelf5556@gmail.com)'
}
SAVE_EVERY = 25


# =========================================================================
# ВЫТАСКИВАНИЕ МЁМ-ЗНАЧЕНИЯ
# =========================================================================

MM_RE = re.compile(
    r'(\d{1,2}[\s,\.]?\d{3}|\d{3,5})\s*(?:mm|мм)\b',
    re.IGNORECASE,
)


def first_mm(value: str):
    """'4,290 mm (168.9 in)' / '{{convert|4290|mm|...}}' → 4290"""
    if not value:
        return None

    # Вариант 1: шаблон convert — {{convert|4290|mm|...}}
    m = re.search(r'\{\{\s*convert\s*\|\s*(\d{3,5})\s*\|\s*mm\b', value, re.IGNORECASE)
    if m:
        return int(m.group(1))

    # Вариант 2: явный поиск «N mm» (любой регистр, ru/en)
    matches = MM_RE.findall(value)
    if matches:
        first = matches[0].replace(' ', '').replace(',', '').replace('.', '')
        try:
            n = int(first)
            if 100 <= n <= 25000:
                return n
        except ValueError:
            pass
    return None


# =========================================================================
# СТРАТЕГИЯ 1: WIKITEXT ЧЕРЕЗ MediaWiki API
# =========================================================================

WIKITEXT_PARAM_RE = re.compile(
    r'^\s*\|\s*([a-zA-Zа-яА-ЯёЁ_\- ]+?)\s*=\s*(.+?)\s*$',
    re.MULTILINE,
)

PARAM_KEYS = {
    'length': 'length', 'длина': 'length',
    'width': 'width', 'ширина': 'width',
    'height': 'height', 'высота': 'height',
    'wheelbase': 'wheelbase',
    'колёсная база': 'wheelbase', 'колесная база': 'wheelbase', 'база': 'wheelbase',
    'ground clearance': 'clearance', 'ground_clearance': 'clearance',
    'дорожный просвет': 'clearance', 'клиренс': 'clearance',
}


def parse_wikitext(wikitext: str):
    """Ищет параметры |length= |width= ... в шаблоне Infobox automobile."""
    out = {}
    if not wikitext:
        return out

    # сузим до первого Infobox automobile / Авто (если есть)
    m = re.search(
        r'\{\{\s*Infobox\s*(?:automobile|vehicle|car|Автомобиль|Авто)\b',
        wikitext, re.IGNORECASE,
    )
    if m:
        # отрежем от начала infobox до 8000 символов (хватит на всю карточку)
        chunk = wikitext[m.start(): m.start() + 8000]
    else:
        chunk = wikitext

    for pm in WIKITEXT_PARAM_RE.finditer(chunk):
        raw_key = pm.group(1).strip().lower()
        raw_val = pm.group(2)
        key = PARAM_KEYS.get(raw_key)
        if key and key not in out:
            n = first_mm(raw_val)
            if n:
                out[key] = n
    return out


def fetch_wikitext(title: str, lang: str):
    url = f'https://{lang}.wikipedia.org/w/api.php'
    params = {
        'action': 'parse', 'page': title, 'prop': 'wikitext',
        'section': 0, 'format': 'json', 'redirects': 1,
    }
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        data = r.json()
        return data.get('parse', {}).get('wikitext', {}).get('*', '')
    except Exception:
        return ''


# =========================================================================
# СТРАТЕГИЯ 2: HTML ЧЕРЕЗ BeautifulSoup
# =========================================================================

HTML_LABELS = {
    'length': 'length', 'длина': 'length',
    'width': 'width', 'ширина': 'width',
    'height': 'height', 'высота': 'height',
    'wheelbase': 'wheelbase',
    'колёсная база': 'wheelbase', 'колесная база': 'wheelbase',
    'ground clearance': 'clearance',
    'дорожный просвет': 'clearance', 'клиренс': 'clearance',
}


def parse_html(html: str):
    """Парсит infobox у Wikipedia: ищет <tr> с <th>label</th><td>...mm...</td>."""
    out = {}
    if not html:
        return out
    try:
        soup = BeautifulSoup(html, 'html.parser')
    except Exception:
        return out

    # ищем сначала infobox (на ru — также может быть «вертикальная карточка»)
    infobox = soup.find('table', class_=re.compile(r'\binfobox\b', re.IGNORECASE))
    scope = infobox or soup

    for tr in scope.find_all('tr'):
        th = tr.find('th')
        td = tr.find('td')
        if not th or not td:
            continue
        label = th.get_text(' ', strip=True).lower().strip(':')
        key = HTML_LABELS.get(label)
        if not key or key in out:
            continue
        text = td.get_text(' ', strip=True)
        n = first_mm(text)
        if n:
            out[key] = n
    return out


def fetch_html(title: str, lang: str):
    url = f'https://{lang}.wikipedia.org/wiki/{title.replace(" ", "_")}'
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        if r.status_code == 200:
            return r.text, url
    except Exception:
        pass
    return '', ''


# =========================================================================
# ПОИСК СТРАНИЦЫ
# =========================================================================

def wiki_search(query: str, lang: str):
    url = f'https://{lang}.wikipedia.org/w/api.php'
    params = {
        'action': 'opensearch', 'search': query, 'limit': 3,
        'format': 'json', 'redirects': 'resolve',
    }
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        data = r.json()
        return data[1] if len(data) > 1 and data[1] else []
    except Exception:
        return []


def fetch_dimensions(mark, model, generation):
    """Главная функция: пробует разные запросы и стратегии."""
    mark_str = str(mark).strip()
    model_str = str(model).strip()
    gen_str = str(generation).strip()

    queries = []
    # для отечественных — VAZ/ВАЗ короткое имя
    if mark_str.lower() in ('lada', 'ваз', 'vaz'):
        queries.append((f'VAZ-{model_str}', 'en'))
        queries.append((f'ВАЗ-{model_str}', 'ru'))
        queries.append((f'Lada {model_str}', 'en'))
        queries.append((f'Лада {model_str}', 'ru'))

    queries.extend([
        (f'{mark_str} {model_str} {gen_str}th generation', 'en'),
        (f'{mark_str} {model_str} ({gen_str})', 'en'),
        (f'{mark_str} {model_str}', 'en'),
        (f'{mark_str} {model_str}', 'ru'),
    ])

    tried_titles = set()

    for q, lang in queries:
        titles = wiki_search(q, lang)
        time.sleep(RATE_LIMIT_SEC / 2)
        for title in titles[:2]:
            if (title, lang) in tried_titles:
                continue
            tried_titles.add((title, lang))

            # 1. wikitext
            wt = fetch_wikitext(title, lang)
            time.sleep(RATE_LIMIT_SEC / 2)
            dims = parse_wikitext(wt)
            if dims:
                dims['_source'] = f'https://{lang}.wikipedia.org/wiki/{title.replace(" ", "_")}'
                dims['_query'] = f'wikitext: {q}'
                return dims

            # 2. HTML fallback
            html, url = fetch_html(title, lang)
            time.sleep(RATE_LIMIT_SEC / 2)
            dims = parse_html(html)
            if dims:
                dims['_source'] = url
                dims['_query'] = f'html: {q}'
                return dims
    return {}


# =========================================================================
# СВЯЗКА С XLSX
# =========================================================================

def get_col_indices(ws):
    headers = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if isinstance(v, str):
            headers[v.strip()] = c
    return headers


def load_progress():
    if PROGRESS.exists():
        try:
            state = json.loads(PROGRESS.read_text(encoding='utf-8'))
            if state.get('parser_version') != PARSER_VERSION:
                print('   парсер обновился → сбрасываю прогресс')
                return new_state()
            if state.get('hits', 0) == 0 and state.get('misses', 0) > 50:
                print(f'   предыдущий прогон с {state["misses"]} промахами → сбрасываю')
                return new_state()
            return state
        except Exception:
            pass
    return new_state()


def new_state():
    return {'last_row': DATA_START - 1, 'hits': 0, 'misses': 0,
            'parser_version': PARSER_VERSION}


def save_progress(state):
    PROGRESS.write_text(json.dumps(state, ensure_ascii=False, indent=2),
                        encoding='utf-8')


def main():
    if not XLSX.exists():
        print(f'❌ Не найден: {XLSX}')
        sys.exit(1)

    print(f'📂 {XLSX.name}')
    try:
        wb = load_workbook(XLSX)
    except PermissionError:
        print('❌ Файл открыт в Excel. Закрой его и запусти снова.')
        sys.exit(1)

    if SHEET not in wb.sheetnames:
        print(f'❌ Лист «{SHEET}» не найден. Есть: {wb.sheetnames}')
        sys.exit(1)
    ws = wb[SHEET]
    cols = get_col_indices(ws)

    required = [COL_MARK, COL_MODEL, COL_GEN, COL_LENGTH, COL_WIDTH,
                COL_HEIGHT, COL_WHEELBASE, COL_CLEARANCE, COL_SOURCE, COL_NOTE]
    missing = [c for c in required if c not in cols]
    if missing:
        print(f'❌ Нет колонок: {missing}')
        sys.exit(1)

    state = load_progress()
    start_row = max(state['last_row'] + 1, DATA_START)
    total = ws.max_row - DATA_START + 1
    print(f'   строк: {total}, начинаю с {start_row}')
    print(f'   попадание: {state["hits"]}, мимо: {state["misses"]}')
    print()

    for r in range(start_row, ws.max_row + 1):
        mark = ws.cell(row=r, column=cols[COL_MARK]).value
        model = ws.cell(row=r, column=cols[COL_MODEL]).value
        gen = ws.cell(row=r, column=cols[COL_GEN]).value
        if not mark or not model:
            state['last_row'] = r
            continue

        try:
            dims = fetch_dimensions(mark, model, gen)
        except KeyboardInterrupt:
            print('\n⏸  прервано пользователем — сохраняю...')
            wb.save(XLSX); save_progress(state)
            sys.exit(0)
        except Exception as e:
            dims = {}
            print(f'  [{r:>4}] {mark} {model} g{gen}  ! {type(e).__name__}: {e}')

        if dims:
            state['hits'] += 1
            ws.cell(row=r, column=cols[COL_LENGTH]).value = dims.get('length')
            ws.cell(row=r, column=cols[COL_WIDTH]).value = dims.get('width')
            ws.cell(row=r, column=cols[COL_HEIGHT]).value = dims.get('height')
            ws.cell(row=r, column=cols[COL_WHEELBASE]).value = dims.get('wheelbase')
            ws.cell(row=r, column=cols[COL_CLEARANCE]).value = dims.get('clearance')
            ws.cell(row=r, column=cols[COL_SOURCE]).value = dims.get('_source')

            old = ws.cell(row=r, column=cols[COL_NOTE]).value or ''
            old = re.sub(r'\s*\|\s*wiki:[^|]+', '', old).strip(' |')
            new = (old + (' | ' if old else '') + dims.get('_query', '')).strip(' |')
            ws.cell(row=r, column=cols[COL_NOTE]).value = new

            L = dims.get('length') or '—'
            W = dims.get('width') or '—'
            H = dims.get('height') or '—'
            WB = dims.get('wheelbase') or '—'
            print(f'  [{r:>4}] {mark} {model} g{gen}  ✓ L={L} W={W} H={H} WB={WB}')
        else:
            state['misses'] += 1
            old = ws.cell(row=r, column=cols[COL_NOTE]).value or ''
            old = re.sub(r'\s*\|\s*wiki:[^|]+', '', old).strip(' |')
            new = (old + (' | ' if old else '') + 'wiki: not found').strip(' |')
            ws.cell(row=r, column=cols[COL_NOTE]).value = new
            print(f'  [{r:>4}] {mark} {model} g{gen}  — не найдено')

        state['last_row'] = r

        if (r - start_row + 1) % SAVE_EVERY == 0:
            try:
                wb.save(XLSX)
            except PermissionError:
                print('   ⚠ xlsx занят — продолжаю в памяти')
            save_progress(state)
            print(f'   ── промежуточное (hits: {state["hits"]}, miss: {state["misses"]})')

    try:
        wb.save(XLSX)
    except PermissionError:
        print('❌ Финал: xlsx занят. Закрой Excel и запусти ещё раз.')
        sys.exit(1)
    save_progress(state)
    rate = state['hits'] / max(state['hits'] + state['misses'], 1) * 100
    print(f'\n✅ Готово. hits: {state["hits"]}, miss: {state["misses"]}')
    print(f'   процент попадания: {rate:.1f}%')


if __name__ == '__main__':
    main()
