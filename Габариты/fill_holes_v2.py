#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Дополняет пропущенные габариты в датасете.
Если для машины не нашлась длина/ширина/высота в Wikipedia,
этот скрипт пробует Wikidata и японскую Wikipedia для JDM-моделей.
Заполняет только пустые ячейки, уже найденные данные не трогает.

Запуск:
    python fill_holes_v2.py
"""
import json, re, sys, time
from pathlib import Path

import requests
from openpyxl import load_workbook

try:
    from bs4 import BeautifulSoup
except ImportError:
    print('pip install beautifulsoup4'); sys.exit(1)

SCRIPT_DIR   = Path(__file__).parent
XLSX         = SCRIPT_DIR / 'MVP_dataset_restored.xlsx'
PROGRESS     = SCRIPT_DIR / '_holes_v2_progress.json'
SHEET        = 'MVP top-1600'
HEADER_ROW   = 2
DATA_START   = 3
RATE         = 0.4
TIMEOUT      = 20
SAVE_EVERY   = 20
HEADERS_HTTP = {'User-Agent': 'cars_analysis/2.1 (educational; adelf5556@gmail.com)'}

COL_ID        = 'ID'
COL_MARK      = 'Марка'
COL_MODEL     = 'Модель'
COL_GEN       = 'Поколение'
COL_LENGTH    = 'Длина, мм'
COL_WIDTH     = 'Ширина, мм'
COL_HEIGHT    = 'Высота, мм'
COL_WHEELBASE = 'Колёсная база, мм'
COL_CLEARANCE = 'Клиренс, мм'
COL_SOURCE    = 'Источник габаритов'

JDM_MARKS = {
    'toyota': 'トヨタ',
    'honda': 'ホンダ',
    'nissan': 'ニッサン',
    'mazda': 'マツダ',
    'subaru': 'スバル',
    'mitsubishi': 'ミツビシ',
    'suzuki': 'スズキ',
    'daihatsu': 'ダイハツ',
    'isuzu': 'いすず',
    'lexus': 'レクサス',
    'infiniti': 'インフィニティ',
}

MM_RE = re.compile(r'(\d{1,2}[\s,.]?\d{3}|\d{3,5})\s*(?:mm|мм)\b', re.I)

def first_mm(val):
    if not val: return None
    m = re.search(r'\{\{\s*convert\s*\|\s*(\d{3,5})\s*\|\s*mm\b', str(val), re.I)
    if m: return int(m.group(1))
    for s in MM_RE.findall(str(val)):
        n = s.replace(' ','').replace(',','').replace('.','')
        try:
            n = int(n)
            if 800 <= n <= 25000: return n
        except: pass
    return None

PARAM_MAP = {
    'length':'L', 'длина':'L',
    'width':'W', 'ширина':'W',
    'height':'H', 'высота':'H',
    'wheelbase':'WB', 'колёсная база':'WB', 'колесная база':'WB', 'база':'WB',
    'ground clearance':'C', 'ground_clearance':'C',
    'дорожный просвет':'C', 'клиренс':'C',
}

def parse_wikitext(wt):
    out = {}
    if not wt: return out
    m = re.search(r'\{\{\s*[Ii]nfobox\s*(?:automobile|vehicle|car|Автомобиль|Авто)\b', wt, re.I)
    chunk = wt[m.start():m.start()+8000] if m else wt
    for pm in re.finditer(r'^\s*\|\s*([a-zA-Zа-яА-ЯёЁ_\- ]+?)\s*=\s*(.+?)\s*$', chunk, re.M):
        k = pm.group(1).strip().lower()
        mapped = PARAM_MAP.get(k)
        if mapped and mapped not in out:
            n = first_mm(pm.group(2))
            if n: out[mapped] = n
    return out

HTML_LABELS = {
    'length':'L','длина':'L','width':'W','ширина':'W',
    'height':'H','высота':'H',
    'wheelbase':'WB','колёсная база':'WB','колесная база':'WB',
    'ground clearance':'C','дорожный просвет':'C','клиренс':'C',
    '全長':'L','全幅':'W','全高':'H','ホイールベース':'WB','最低地上高':'C',
}

def parse_html(html):
    out = {}
    if not html: return out
    try: soup = BeautifulSoup(html, 'html.parser')
    except: return out
    infobox = soup.find('table', class_=re.compile(r'\binfobox\b', re.I))
    scope = infobox or soup
    for tr in scope.find_all('tr'):
        th = tr.find('th')
        td = tr.find('td')
        if not th or not td: continue
        label = th.get_text(' ', strip=True).lower().strip(':').strip()
        mapped = HTML_LABELS.get(label)
        if mapped and mapped not in out:
            n = first_mm(td.get_text(' ', strip=True))
            if n: out[mapped] = n
    return out

def wiki_fetch_wikitext(title, lang):
    try:
        r = requests.get(
            f'https://{lang}.wikipedia.org/w/api.php',
            params={'action':'parse','page':title,'prop':'wikitext',
                    'section':0,'format':'json','redirects':1},
            headers=HEADERS_HTTP, timeout=TIMEOUT)
        return r.json().get('parse',{}).get('wikitext',{}).get('*','')
    except: return ''

def wiki_fetch_html(title, lang):
    try:
        r = requests.get(
            f'https://{lang}.wikipedia.org/wiki/{title.replace(" ","_")}',
            headers=HEADERS_HTTP, timeout=TIMEOUT)
        if r.status_code == 200:
            return r.text, r.url
    except: pass
    return '', ''

def wiki_search(query, lang):
    try:
        r = requests.get(
            f'https://{lang}.wikipedia.org/w/api.php',
            params={'action':'opensearch','search':query,'limit':3,
                    'format':'json','redirects':'resolve'},
            headers=HEADERS_HTTP, timeout=TIMEOUT)
        data = r.json()
        return data[1] if len(data)>1 else []
    except: return []

PROP_MAP = {'P2043':'L','P2049':'W','P2048':'H','P2783':'WB','P2926':'C'}

def wikidata_search(mark, model):
    try:
        r = requests.get(
            'https://www.wikidata.org/w/api.php',
            params={'action':'wbsearchentities','search':f'{mark} {model}',
                    'language':'en','format':'json','limit':3,'type':'item'},
            headers=HEADERS_HTTP, timeout=TIMEOUT)
        results = r.json().get('search', [])
        for res in results:
            desc = (res.get('description') or '').lower()
            if any(w in desc for w in ('automobile','car','vehicle','sedan','suv')):
                return res['id']
        return results[0]['id'] if results else None
    except: return None

def wikidata_dims(qid):
    if not qid: return {}
    sparql = f"""
    SELECT ?prop ?val WHERE {{
      VALUES ?prop {{ wd:P2043 wd:P2049 wd:P2048 wd:P2783 wd:P2926 }}
      wd:{qid} ?prop [ wikibase:quantityAmount ?val ] .
    }}"""
    try:
        r = requests.get(
            'https://query.wikidata.org/sparql',
            params={'query': sparql, 'format': 'json'},
            headers={**HEADERS_HTTP, 'Accept': 'application/sparql-results+json'},
            timeout=TIMEOUT)
        out = {}
        for row in r.json().get('results',{}).get('bindings',[]):
            prop_uri = row['prop']['value'].split('/')[-1]
            val = float(row['val']['value'])
            mm = int(round(val * 1000)) if val < 100 else int(round(val))
            mapped = PROP_MAP.get(prop_uri)
            if mapped and mapped not in out and 800 <= mm <= 25000:
                out[mapped] = mm
        return out
    except: return {}

def fetch_dimensions(mark, model, gen):
    m, mo, g = str(mark).strip(), str(model).strip(), str(gen).strip()
    ml = m.lower()
    queries = []
    if ml in ('lada','ваз','vaz'):
        queries += [(f'VAZ-{mo}','en'),(f'ВАЗ-{mo}','ru'),(f'Lada {mo}','en')]
    if ml in ('газ','gaz'):
        queries += [(f'GAZ-{mo}','en'),(f'ГАЗ-{mo}','ru')]
    queries += [
        (f'{m} {mo} {g}th generation', 'en'),
        (f'{m} {mo} ({g})', 'en'),
        (f'{m} {mo}', 'en'),
        (f'{m} {mo}', 'ru'),
    ]
    if ml in JDM_MARKS:
        queries.append((f'{m} {mo}', 'ja'))
    tried = set()
    for q, lang in queries:
        titles = wiki_search(q, lang)
        time.sleep(RATE / 2)
        for title in titles[:2]:
            key = (title, lang)
            if key in tried: continue
            tried.add(key)
            wt = wiki_fetch_wikitext(title, lang)
            time.sleep(RATE / 2)
            dims = parse_wikitext(wt)
            if dims:
                dims['_src'] = f'https://{lang}.wikipedia.org/wiki/{title.replace(" ","_")}'
                return dims
            html, url = wiki_fetch_html(title, lang)
            time.sleep(RATE / 2)
            dims = parse_html(html)
            if dims:
                dims['_src'] = url
                return dims
    qid = wikidata_search(m, mo)
    time.sleep(RATE)
    if qid:
        dims = wikidata_dims(qid)
        if dims:
            dims['_src'] = f'https://www.wikidata.org/wiki/{qid}'
            return dims
    return {}

def get_cols(ws):
    cols = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if isinstance(v, str):
            cols[v.strip()] = c
    return cols

def load_progress():
    if PROGRESS.exists():
        try: return set(json.loads(PROGRESS.read_text('utf-8')).get('done', []))
        except: pass
    return set()

def save_progress(done):
    PROGRESS.write_text(json.dumps({'done': list(done)}, ensure_ascii=False), 'utf-8')

def main():
    if not XLSX.exists():
        print(f'Net fayla: {XLSX}'); sys.exit(1)
    try:
        wb = load_workbook(XLSX)
    except PermissionError:
        print('Zakroi fail v Excel!'); sys.exit(1)

    ws = wb[SHEET]
    cols = get_cols(ws)
    for c in [COL_MARK, COL_MODEL, COL_GEN, COL_LENGTH]:
        if c not in cols:
            print(f'Net kolonki: {c}'); sys.exit(1)

    done = load_progress()
    todo = []
    for r in range(DATA_START, ws.max_row + 1):
        vid = ws.cell(row=r, column=cols[COL_ID]).value or f'row{r}'
        if vid in done: continue
        L = ws.cell(row=r, column=cols[COL_LENGTH]).value
        W = ws.cell(row=r, column=cols[COL_WIDTH]).value
        H = ws.cell(row=r, column=cols[COL_HEIGHT]).value
        if not L or not W or not H:
            todo.append((r, vid))

    print(f'Strok k obrabotke: {len(todo)}')
    if not todo:
        print('Vse zapolneno!'); return

    hits = miss = 0
    for i, (r, vid) in enumerate(todo, 1):
        mark  = ws.cell(row=r, column=cols[COL_MARK]).value
        model = ws.cell(row=r, column=cols[COL_MODEL]).value
        gen   = ws.cell(row=r, column=cols[COL_GEN]).value

        try:
            dims = fetch_dimensions(mark, model, gen)
        except KeyboardInterrupt:
            print('\nPrervanno — sohranyayu...')
            wb.save(XLSX); save_progress(done); sys.exit(0)
        except Exception as e:
            dims = {}
            print(f'  [{i}/{len(todo)}] {mark} {model} g{gen} ! {e}')

        if dims:
            hits += 1
            def fill(col_name, key):
                if col_name not in cols: return
                if not ws.cell(row=r, column=cols[col_name]).value and key in dims:
                    ws.cell(row=r, column=cols[col_name]).value = dims[key]
            fill(COL_LENGTH,    'L')
            fill(COL_WIDTH,     'W')
            fill(COL_HEIGHT,    'H')
            fill(COL_WHEELBASE, 'WB')
            fill(COL_CLEARANCE, 'C')
            if COL_SOURCE in cols and not ws.cell(row=r, column=cols[COL_SOURCE]).value:
                ws.cell(row=r, column=cols[COL_SOURCE]).value = dims.get('_src','')
            print(f'  [{i}/{len(todo)}] OK {mark} {model} g{gen}  L={dims.get("L")} W={dims.get("W")} H={dims.get("H")}')
        else:
            miss += 1
            print(f'  [{i}/{len(todo)}] -- {mark} {model} g{gen}')

        done.add(vid)

        if i % SAVE_EVERY == 0 or i == len(todo):
            try: wb.save(XLSX)
            except PermissionError: print('  xlsx zanyat')
            save_progress(done)
            print(f'  -- sohraneno  hits={hits}  miss={miss}')

    print(f'\nGotovo: hits={hits}, miss={miss}')
    if hits+miss:
        print(f'Procent: {hits/(hits+miss)*100:.1f}%')

if __name__ == '__main__':
    main()
