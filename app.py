"""Bábolna Sped: Agroorg létszám, kézi besorolás és arculatos exportok."""
from datetime import date
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import re
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd

BLUE, YELLOW, GRAY, WHITE = '#2A3756', '#F2E47D', '#D1D5DB', '#FFFFFF'
CATEGORIES = ['Nemzetközis', 'Belföldes', 'Hibrid', 'Ónódi', '4-kezes', 'Kihagyás']
FACTORS = {'Nemzetközis': (1., 0.), 'Belföldes': (0., 1.),
           'Hibrid': (.8, .2), 'Ónódi': (.625, 0.), '4-kezes': (.75, 0.), 'Kihagyás': (0., 0.), '0 Ft kereset': (0., 0.)}
BASE_COLUMNS = ['Törzsszám', 'Név', 'FEOR', 'Alaplétszám', 'Besorolás']
EXTRA_COLUMNS = ['Kereset', 'Kihagyás oka', 'Megjegyzés']
EXCLUDED = ['Kihagyás', '0 Ft kereset']
ZERO_REASON = '0 kereset miatt kihagyva'
GROUPS = ['Belföldes', 'Nemzetközis', 'Hibrid', 'Ónódi', '4-kezes', 'Nem számolt']


CONFIG_COLUMNS = ['Kategória', 'Nemzetközi %', 'Belföldi %']
RESERVED_SHEETS = {'összesítés', 'részletezés', 'kihagyottak', 'besorolások', 'ellenőrzés', 'havi belépők', 'kategóriák', 'nem számolt', '0 ft kereset', 'history'}


def default_config():
    return pd.DataFrame([[name, FACTORS[name][0] * 100, FACTORS[name][1] * 100]
                         for name in CATEGORIES], columns=CONFIG_COLUMNS)


def validate_config(config, used=()):
    result = config[CONFIG_COLUMNS].copy()
    result['Kategória'] = result['Kategória'].fillna('').astype(str).str.strip()
    names = result['Kategória']
    if names.eq('').any() or names.str.casefold().duplicated().any():
        raise ValueError('A kategórianevek legyenek kitöltve és egyediek.')
    for name in names:
        if len(name) > 31 or re.search(r"[\[\]:*?/\\]", name) or name.startswith("'") or name.endswith("'") or name.casefold() in RESERVED_SHEETS:
            raise ValueError(f'Nem használható kategórianév: {name}. Legfeljebb 31 karakter, Excelben is érvényes név szükséges.')
    if not {'Nemzetközis', 'Kihagyás'}.issubset(names):
        raise ValueError('A Nemzetközis és Kihagyás alapértelmezett kategória nem törölhető vagy nevezhető át.')
    missing = set(used) - set(names) - {'0 Ft kereset'}
    if missing:
        raise ValueError('Előbb sorold át az érintett dolgozókat, majd töröld vagy nevezd át a kategóriát: ' + ', '.join(sorted(missing)))
    for col in CONFIG_COLUMNS[1:]:
        result[col] = pd.to_numeric(result[col], errors='coerce')
        if result[col].isna().any() or not np.isfinite(result[col]).all() or result[col].lt(0).any():
            raise ValueError('A százalékok nemnegatív, véges számok legyenek.')
    if result.loc[names.eq('Kihagyás'), CONFIG_COLUMNS[1:]].ne(0).any().any():
        raise ValueError('A Kihagyás szorzói kötelezően 0%-osak.')
    return result.reset_index(drop=True)


def factors_for(config=None):
    rows = validate_config(default_config() if config is None else config)
    return {**{row['Kategória']: (row['Nemzetközi %']/100, row['Belföldi %']/100)
               for _, row in rows.iterrows()}, '0 Ft kereset': (0., 0.)}


def restore_config(data):
    book = pd.ExcelFile(BytesIO(data))
    return validate_config(pd.read_excel(book, sheet_name='Kategóriák')) if 'Kategóriák' in book.sheet_names else default_config()


def grouped(frame, config=None):
    counted = ~frame['Besorolás'].isin(EXCLUDED) & frame['Alaplétszám'].gt(0)
    return [(name, frame.loc[~counted if name == 'Nem számolt' else counted & frame['Besorolás'].eq(name)])
            for name in ([n for n in GROUPS[:-1] if n in factors_for(config)] +
                         [n for n in factors_for(config) if n not in GROUPS and n not in EXCLUDED] + ['Nem számolt'])]


def empty_employment_date(value):
    """Az Agroorg üres dátuma kötőjelekkel is érkezhet; ez nem hibás dátum."""
    return pd.isna(value) or str(value).strip() in ('', '0', '0.0') or bool(re.fullmatch(r'[\s\-–—]+', str(value)))


def employment_date(value):
    if empty_employment_date(value):
        return pd.NaT
    if isinstance(value, (int, float, np.number)):
        return pd.to_datetime(value, unit='D', origin='1899-12-30', errors='coerce')
    text = str(value).strip()
    return pd.to_datetime(text, dayfirst=not bool(re.match(r'^\d{4}[-./]', text)), errors='coerce')


def monthly_entrants(frame, month):
    """Kiegészítő nézet, nem külön számítási kategória; nem duplázza az összesítést."""
    dates = frame.get('Belépés', pd.Series('', index=frame.index)).map(employment_date)
    start = pd.Timestamp(month + '-01')
    end = start + pd.offsets.MonthEnd(0)
    return frame.loc[dates.ge(start) & dates.le(end)].sort_values(['Belépés', 'Név'])


def audit(frame, month):
    """Dátumokból kizárólag összehasonlítás: a forráslétszámot nem arányosítjuk újra."""
    start = pd.Timestamp(month + '-01')
    end = start + pd.offsets.MonthEnd(0)
    issues = []
    for _, row in frame.iterrows():
        begin = employment_date(row.get('Belépés', ''))
        finish = employment_date(row.get('Kilépés', ''))
        notes = []
        ratio = None
        if row['Alaplétszám'] > 1:
            notes.append('1 fő feletti alaplétszám – ellenőrizendő')
        if row['Kereset'] < 0:
            notes.append('Negatív kereset – korrekció lehet, nem automatikus kihagyás')
        if row['Kereset'] != 0 and row['Alaplétszám'] == 0:
            notes.append('Van kereset, de nulla alaplétszám; a hozzájárulás nulla')
        if row.get('Dátumhiba', ''):
            notes.append(row['Dátumhiba'])
        if pd.notna(begin) and pd.notna(finish) and finish < begin:
            notes.append('Kilépés a belépés előtt')
        elif pd.notna(begin):
            overlap = max(0, (min(end, finish) - max(start, begin)).days + 1) if pd.notna(finish) else max(0, (end - max(start, begin)).days + 1)
            ratio = overlap / end.day
            if overlap == 0 and row['Alaplétszám'] > 0:
                notes.append('A jogviszony nem érinti ezt a hónapot, mégis pozitív az alaplétszám')
            elif (start < begin <= end or pd.notna(finish) and start <= finish < end):
                if abs(row['Alaplétszám'] - ratio) > 0.00011:
                    notes.append('Hóközi be-/kilépés: a naptári arány eltér az Agroorg-értéktől; nem feltétlenül hiba')
                else:
                    notes.append('Hóközi be-/kilépés: az alaplétszám egyezik a naptári aránnyal')
        elif pd.notna(finish) and finish < start and row['Alaplétszám'] > 0:
            notes.append('Hónap előtti kilépés mellett pozitív alaplétszám')
        if notes:
            issues.append({'Törzsszám': row['Törzsszám'], 'Név': row['Név'],
                           'Belépés': '' if pd.isna(begin) else begin.strftime('%Y-%m-%d'),
                           'Kilépés': '' if pd.isna(finish) else finish.strftime('%Y-%m-%d'),
                           'Alaplétszám': row['Alaplétszám'], 'Naptári arány': ratio,
                           'Ellenőrzés': '; '.join(notes)})
    return pd.DataFrame(issues)


def identifier(value):
    if pd.isna(value):
        return ''
    return re.sub(r'\.0$', '', str(value).strip())


def read_source(data, filename):
    """A név szerinti oszlopazonosítás a régi, kézzel bővített exportot is kezeli."""
    suffix = Path(filename).suffix.lower()
    if suffix in ('.xls', '.xlsx'):
        raw = pd.read_excel(BytesIO(data), engine='xlrd' if suffix == '.xls' else 'openpyxl')
    elif suffix == '.csv':
        decoded = None
        for encoding in ('utf-8-sig', 'cp1250'):
            try:
                decoded = data.decode(encoding)
                break
            except UnicodeDecodeError:
                pass
        if decoded is None:
            raise ValueError('A CSV karakterkódolása nem olvasható.')
        from io import StringIO
        raw = pd.read_csv(StringIO(decoded), sep=None, engine='python')
    else:
        raise ValueError('XLS, XLSX vagy CSV fájlt tölts fel.')
    raw.columns = [str(c).strip().lower() for c in raw.columns]
    required = ['torzsszam', 'nev', 'letsz', 'feorkód', 'kereset']
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError('Hiányzó Agroorg-oszlop: ' + ', '.join(missing))
    raw = raw.dropna(how='all').reset_index(drop=True)
    if raw.empty:
        raise ValueError('A fájl nem tartalmaz személyeket.')
    amounts = pd.to_numeric(raw['letsz'].astype(str).str.replace('\u00a0', '', regex=False)
                            .str.replace(' ', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
    invalid = amounts.isna() | ~np.isfinite(amounts) | (amounts < 0)
    if invalid.any():
        rows = ', '.join(str(i + 2) for i in raw.index[invalid])
        raise ValueError(f'Hiányzó, hibás vagy negatív létszámérték. Excel-sorok: {rows}. Javítsd a forrást.')
    earnings = pd.to_numeric(raw['kereset'].astype(str).str.replace('\u00a0', '', regex=False)
                             .str.replace(' ', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
    invalid_earnings = earnings.isna() | ~np.isfinite(earnings)
    if invalid_earnings.any():
        rows = ', '.join(str(i + 2) for i in raw.index[invalid_earnings])
        raise ValueError(f'Hiányzó vagy hibás kereset. Excel-sorok: {rows}. A hiányzó adat nem tekinthető nullának.')
    ids = raw['torzsszam'].map(identifier)
    names = raw['nev'].fillna('').astype(str).str.strip().str.title()
    if ids.eq('').any() or names.eq('').any():
        raise ValueError('Hiányzó törzsszám vagy név. Javítsd a forrást.')
    if ids.duplicated().any():
        raise ValueError('Ismétlődő törzsszám: ' + ', '.join(ids[ids.duplicated()].unique()) +
                         '. Az eltérő jogviszonyok összesítését ellenőrizni kell.')
    feor = raw['feorkód'].map(identifier)
    result = pd.DataFrame({'Törzsszám': ids, 'Név': names, 'FEOR': feor,
                           'Alaplétszám': amounts, 'Besorolás': np.where(feor.eq('8417'), 'Nemzetközis', 'Kihagyás'),
                           'Kereset': earnings, 'Kihagyás oka': '', 'Megjegyzés': ''})
    result['Dátumhiba'] = ''
    for source, target in [('jogvkezd', 'Belépés'), ('jogvvég', 'Kilépés')]:
        result[target] = ''
        if source in raw:
            for i, value in raw[source].items():
                parsed = employment_date(value)
                if pd.notna(parsed):
                    result.loc[i, target] = parsed.strftime('%Y-%m-%d')
                elif not empty_employment_date(value):
                    result.loc[i, 'Dátumhiba'] += f'Hibás {target.lower()} dátum; '
    result = result.sort_values('Név').reset_index(drop=True)
    return enforce_rules(result)


def enforce_rules(frame):
    """A kötelező kizárás import, visszatöltés, számítás és export során is érvényes."""
    result = frame.copy()
    result['Név'] = result['Név'].astype(str).str.strip().str.title()
    if result['Kereset'].isna().any() or not np.isfinite(result['Kereset']).all():
        raise ValueError('Hiányzó vagy hibás kereset.')
    zero = result['Kereset'].eq(0)
    result.loc[zero, 'Besorolás'] = '0 Ft kereset'
    if 'Megjegyzés' not in result:
        result['Megjegyzés'] = ''
    result['Megjegyzés'] = result['Megjegyzés'].fillna('').astype(str)
    if (result['Kereset'].ne(0) & result['Besorolás'].eq('0 Ft kereset')).any():
        raise ValueError('A 0 Ft kereset besorolás csak nulla keresethez tartozhat.')
    result['Kihagyás oka'] = np.where(zero, ZERO_REASON,
                                    np.where(result['Besorolás'].eq('Kihagyás'), 'Kézi kihagyás', ''))
    return result


def suggested_month(filename):
    matches = re.findall(r'(20\d{2})(0[1-9]|1[0-2])(?!\d)', Path(filename).stem)
    return date(int(matches[-1][0]), int(matches[-1][1]), 1) if matches else date.today().replace(day=1)


def restore_categories(data, current, month=None, config=None):
    """A besorolás átvihető másik hónapra, a megjegyzés csak ugyanarra a hónapra."""
    saved = pd.read_excel(BytesIO(data), sheet_name='Besorolások', dtype={'Törzsszám': str})
    if not {'Törzsszám', 'Besorolás'}.issubset(saved.columns):
        raise ValueError('A kiválasztott fájl nem az alkalmazás Excel-exportja.')
    saved['Törzsszám'] = saved['Törzsszám'].map(identifier)
    if saved['Törzsszám'].eq('').any() or saved['Törzsszám'].duplicated().any():
        raise ValueError('A mentett besorolások törzsszámai hiányosak vagy ismétlődnek.')
    saved['Besorolás'] = saved['Besorolás'].replace({'Váltó / Ónódi': 'Ónódi'})
    if not saved['Besorolás'].isin(factors_for(config)).all():
        raise ValueError('Ismeretlen kategória található a mentett besorolásokban.')
    automatic_zero = saved['Besorolás'].eq('0 Ft kereset')
    if 'Kihagyás oka' in saved:
        automatic_zero |= saved['Kihagyás oka'].eq(ZERO_REASON)
    saved.loc[automatic_zero, 'Besorolás'] = None
    mapping = saved.set_index('Törzsszám')['Besorolás']
    result = current.copy()
    result['Megjegyzés'] = ''
    defaults = pd.Series(np.where(result['FEOR'].eq('8417'), 'Nemzetközis', 'Kihagyás'), index=result.index)
    result['Besorolás'] = result['Törzsszám'].map(mapping).fillna(defaults)
    if month and 'Megjegyzés' in saved:
        summary = pd.read_excel(BytesIO(data), sheet_name='Összesítés', header=None)
        if str(summary.iloc[3, 1]) == month:
            notes = saved.set_index('Törzsszám')['Megjegyzés']
            result['Megjegyzés'] = result['Törzsszám'].map(notes).fillna('')
    result = enforce_rules(result)
    return result, int(result['Törzsszám'].isin(mapping.index).sum())


def calculate(frame, config=None):
    factors = factors_for(config)
    frame = enforce_rules(frame)
    if not frame['Besorolás'].isin(factors).all():
        raise ValueError('Minden személynél válassz érvényes besorolást.')
    result = frame.copy()
    result['N szorzó'] = result['Besorolás'].map(lambda c: factors[c][0])
    result['B szorzó'] = result['Besorolás'].map(lambda c: factors[c][1])
    result['Nemzetközi'] = result['Alaplétszám'] * result['N szorzó']
    result['Belföldi'] = result['Alaplétszám'] * result['B szorzó']
    return result


def excel_export(frame, month, source, config=None):
    """Az alkalmazás futás közben készített Excel-exportja képletekkel és gyorsítótárazott eredménnyel."""
    import xlsxwriter
    frame = calculate(frame, config)
    output = BytesIO()
    book = xlsxwriter.Workbook(output, {'in_memory': True, 'strings_to_formulas': False,
                                      'strings_to_urls': False})
    normal = book.add_format({'font_name': 'Calibri', 'font_size': 11, 'font_color': BLUE})
    number = book.add_format({'font_name': 'Calibri', 'num_format': '0.0000', 'font_color': BLUE})
    note_format = book.add_format({'font_name': 'Calibri', 'font_size': 11, 'font_color': BLUE, 'text_wrap': True, 'valign': 'top'})
    money = book.add_format({'font_name': 'Calibri', 'num_format': '#,##0.00', 'font_color': BLUE})
    head = book.add_format({'font_name': 'Calibri', 'bold': True, 'bg_color': BLUE,
                           'font_color': WHITE, 'text_wrap': True, 'valign': 'vcenter'})
    total = book.add_format({'font_name': 'Calibri', 'bold': True, 'bg_color': YELLOW,
                            'font_color': BLUE, 'num_format': '0.0000'})
    summary = book.add_worksheet('Összesítés')
    detail = book.add_worksheet('Részletezés')
    skipped = book.add_worksheet('Kihagyottak')
    settings = book.add_worksheet('Besorolások')
    for sheet in (summary, detail, skipped, settings):
        sheet.hide_gridlines(2)
        sheet.set_tab_color(BLUE)
        sheet.set_default_row(21)
        sheet.set_column('A:A', 16, normal)
        sheet.set_column('B:B', 34, normal)
        sheet.set_column('C:L', 18, normal)
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
    # A besorolások munkalap az exportból történő visszatöltéshez is használható.
    settings.write_row(0, 0, BASE_COLUMNS + EXTRA_COLUMNS, head)
    settings.set_column('G:H', 30, normal)
    settings.set_row(0, 32)
    for i, row in enumerate(frame[BASE_COLUMNS + EXTRA_COLUMNS].itertuples(index=False, name=None), 1):
        settings.write_row(i, 0, row, normal)
        settings.write_number(i, 3, row[3], number)
        settings.write_number(i, 5, row[5], money)
        settings.write_string(i, 7, row[7], note_format)
        settings.set_row(i, min(409, max(21, (len(row[7]) // 28 + 1) * 15)))
    config = validate_config(default_config() if config is None else config)
    category_sheet = book.add_worksheet('Kategóriák')
    category_sheet.hide_gridlines(2)
    category_sheet.set_column('A:A', 32, normal)
    category_sheet.set_column('B:C', 20, number)
    category_sheet.write_row(0, 0, CONFIG_COLUMNS, head)
    for i, row in enumerate(config.itertuples(index=False, name=None), 1):
        category_sheet.write_row(i, 0, row, normal)
    category_sheet.freeze_panes(1, 1)
    category_end = len(config) + 1
    book.define_name('Kategoriak', f"='Kategóriák'!$A$2:$A${category_end}")
    settings.data_validation(1, 4, len(frame), 4, {'validate': 'list', 'source': '=Kategoriak'})
    settings.freeze_panes(1, 2)
    settings.autofilter(0, 0, len(frame), 7)
    headers = BASE_COLUMNS + ['N szorzó', 'B szorzó', 'Nemzetközi', 'Belföldi'] + EXTRA_COLUMNS
    detail.set_column('K:L', 30, normal)
    detail.write_row(0, 0, headers, head)
    detail.set_row(0, 32)
    for i, (_, row) in enumerate(frame.iterrows(), 1):
        r = i + 1
        for j, col in enumerate(BASE_COLUMNS):
            detail.write_formula(i, j, f"='Besorolások'!{chr(65+j)}{r}",
                                 number if j == 3 else normal, row[col])
        # Kereset=0 esetén az Excelben átírt besorolás sem változtathatja meg a számítást.
        detail.write_formula(i, 4, f'IF(J{r}=0,"0 Ft kereset",\'Besorolások\'!E{r})', normal, row['Besorolás'])
        detail.write_formula(i, 9, f"='Besorolások'!F{r}", money, row['Kereset'])
        detail.write_formula(i, 10, f'IF(J{r}=0,"{ZERO_REASON}",IF(E{r}="Kihagyás","Kézi kihagyás",""))', normal, row['Kihagyás oka'])
        detail.write_formula(i, 11, f"=IF('Besorolások'!H{r}=\"\",\"\",'Besorolások'!H{r})", note_format, row['Megjegyzés'])
        detail.set_row(i, min(409, max(21, (len(row["Megjegyzés"]) // 28 + 1) * 15)))
        for j, factor_col in [(5, 2), (6, 3)]:
            formula = f'IF(OR(E{r}="0 Ft kereset",E{r}="Kihagyás"),0,IFERROR(VLOOKUP(E{r},\'Kategóriák\'!$A$2:$C${category_end},{factor_col},FALSE)/100,0))'
            detail.write_formula(i, j, formula, number, row['N szorzó' if j == 5 else 'B szorzó'])
        detail.write_formula(i, 7, f'D{r}*F{r}', number, row['Nemzetközi'])
        detail.write_formula(i, 8, f'D{r}*G{r}', number, row['Belföldi'])
    detail.freeze_panes(1, 2)
    detail.autofilter(0, 0, len(frame), 11)
    detail.repeat_rows(0)
    end = len(frame) + 1
    summary.set_column('A:A', 44, normal)
    summary.merge_range('A1:D2', 'BÁBOLNA SPED | SOFŐRLÉTSZÁM', head)
    summary.write_row('A4', ['Elszámolási hónap', month], normal)
    summary.write_row('A5', ['Forrásfájl', source], normal)
    for i, (label, col) in enumerate([('Nemzetközi súlyozott létszám', 'H'), ('Belföldi súlyozott létszám', 'I')], 7):
        summary.write(i, 0, label, total)
        value = frame['Nemzetközi' if col == 'H' else 'Belföldi'].sum()
        summary.write_formula(i, 1, f"SUM('Részletezés'!{col}2:{col}{end})", total, value)
    summary.write('A10', 'Összes súlyozott létszám', total)
    summary.write_formula('B10', 'SUM(B8:B9)', total, frame[['Nemzetközi', 'Belföldi']].sum().sum())
    summary.write_row('A12', ['Forrásban szereplő személyek', len(frame)], normal)
    summary.write('A13', 'Besorolt személyek (nullás értékkel is)', normal)
    summary.write_formula('B13', f'COUNTIFS(\'Részletezés\'!E2:E{end},"<>Kihagyás",\'Részletezés\'!E2:E{end},"<>0 Ft kereset")', normal, int((~frame['Besorolás'].isin(EXCLUDED)).sum()))
    summary.write('A14', 'Pozitív alaplétszámú besorolt személyek', normal)
    summary.write_formula('B14', f'COUNTIFS(\'Részletezés\'!E2:E{end},"<>Kihagyás",\'Részletezés\'!E2:E{end},"<>0 Ft kereset",\'Részletezés\'!D2:D{end},">0")', normal,
                          int((~frame['Besorolás'].isin(EXCLUDED) & frame['Alaplétszám'].gt(0)).sum()))
    summary.write('A15', 'Kihagyott személyek', normal)
    summary.write_formula('B15', f'COUNTIF(\'Részletezés\'!E2:E{end},"Kihagyás")+COUNTIF(\'Részletezés\'!E2:E{end},"0 Ft kereset")', normal, int(frame['Besorolás'].isin(EXCLUDED).sum()))
    summary.write('A16', 'Ebből 0 kereset miatt kihagyva', normal)
    summary.write_formula('B16', f'COUNTIF(\'Részletezés\'!J2:J{end},0)', normal, int(frame['Kereset'].eq(0).sum()))
    for i, line in enumerate(['Mértékegység: fő. A súlyozott létszám nem személyek darabszáma.',
                             'Az alkalmazott százalékokat a Kategóriák munkalap tartalmazza.',
                             'A Besorolások lap szerkesztése frissíti a Részletezést és az Összesítést.',
                             'A csoportlapok az exportáláskori állapotot rögzítik; frissítéshez új export kell.'], 17):
        summary.merge_range(i, 0, i, 4, line, normal)
    summary.merge_range('A22:E22', 'Más FEOR-kódú dolgozó kézzel bevonható. 0 Ft kereset: kötelező kizárás.', normal)
    skipped.write_row(0, 0, BASE_COLUMNS + EXTRA_COLUMNS, head)
    skipped.set_column('G:H', 30, normal)
    for i, row in enumerate(frame.loc[frame['Besorolás'].isin(EXCLUDED), BASE_COLUMNS + EXTRA_COLUMNS].itertuples(index=False, name=None), 1):
        skipped.write_row(i, 0, row, normal)
        skipped.write_number(i, 3, row[3], number)
        skipped.write_number(i, 5, row[5], money)
        skipped.write_string(i, 7, row[7], note_format)
        skipped.set_row(i, min(409, max(21, (len(row[7]) // 28 + 1) * 15)))
    for group_name, subset in grouped(frame, config):
        sheet = book.add_worksheet(group_name)
        sheet.set_tab_color(BLUE)
        sheet.hide_gridlines(2)
        sheet.set_column('A:A', 15, normal)
        sheet.set_column('B:B', 32, normal)
        sheet.set_column('C:G', 19, normal)
        sheet.set_column('H:H', 40, note_format)
        sheet.merge_range('A1:H1', f'{group_name} – {month} – exportáláskori állapot', head)
        cols = ['Törzsszám', 'Név', 'FEOR', 'Besorolás', 'Alaplétszám', 'Nemzetközi', 'Belföldi', 'Megjegyzés']
        sheet.write_row(2, 0, cols, head)
        for i, (_, row) in enumerate(subset.iterrows(), 3):
            for j, col in enumerate(cols):
                sheet.write(i, j, row[col], number if 4 <= j <= 6 else note_format if j == 7 else normal)
            sheet.set_row(i, min(409, max(21, (len(row['Megjegyzés']) // 35 + 1) * 15)))
        sheet.freeze_panes(3, 2)
        if len(subset):
            sheet.autofilter(2, 0, len(subset) + 2, 7)
    check = book.add_worksheet('Ellenőrzés')
    check.set_column('A:F', 20, normal)
    check.set_column('G:G', 90, note_format)
    check.write('A1', 'Az ellenőrzés nem írja át az Agroorg alaplétszámát.', normal)
    issues = audit(frame, month)
    if not issues.empty:
        check.write_row(2, 0, list(issues.columns), head)
        for i, row in enumerate(issues.fillna('').itertuples(index=False, name=None), 3):
            check.write_row(i, 0, row, note_format)
            check.set_row(i, 60)
    else:
        check.write('A3', 'Nincs jelzett eltérés. Hiányzó dátumokból nem végezhető teljes időszak-ellenőrzés.', normal)
    entrants = monthly_entrants(frame, month)
    entrant_sheet = book.add_worksheet('Havi belépők')
    entrant_sheet.hide_gridlines(2)
    entrant_sheet.set_column('A:A', 15, normal)
    entrant_sheet.set_column('B:B', 32, normal)
    entrant_sheet.set_column('C:H', 19, normal)
    entrant_sheet.merge_range('A1:H1', f'{month} – Havi belépők – kiegészítő nézet', head)
    entrant_sheet.merge_range('A2:H2', 'Ezek a dolgozók a kategórialapokon is szerepelnek. Az összesítés csak egyszer számolja őket.', normal)
    entrant_columns = ['Törzsszám', 'Név', 'Belépés', 'Kilépés', 'Besorolás', 'Alaplétszám', 'Nemzetközi', 'Belföldi']
    entrant_sheet.write_row(3, 0, entrant_columns, head)
    for i, (_, row) in enumerate(entrants.iterrows(), 4):
        for j, col in enumerate(entrant_columns):
            entrant_sheet.write(i, j, row.get(col, ''), number if j >= 5 else normal)
    entrant_sheet.freeze_panes(4, 2)
    book.close()
    return output.getvalue()


def pdf_export(frame, month, source, config=None):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    fonts = Path(__file__).parent / 'assets'
    for name, filename in [('Report', 'DejaVuSans.ttf'), ('ReportBold', 'DejaVuSans-Bold.ttf')]:
        if name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(name, str(fonts / filename)))
    frame = calculate(frame, config)
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, leftMargin=32, rightMargin=32,
                            topMargin=32, bottomMargin=32, title=f'Sofőrlétszám – {month}', author='Bábolna Sped Kft.')
    body = ParagraphStyle('body', fontName='Report', fontSize=8, leading=11, textColor=colors.HexColor(BLUE))
    title = ParagraphStyle('title', parent=body, fontName='ReportBold', fontSize=17, leading=22, spaceAfter=10)
    heading = ParagraphStyle('heading', parent=body, fontName='ReportBold', fontSize=11, leading=15, spaceBefore=14, spaceAfter=7, keepWithNext=True)
    white = ParagraphStyle('white', parent=body, fontName='ReportBold', textColor=colors.white)
    p = lambda text: Paragraph(escape(str(text)), body)
    fmt = lambda value: f'{value:.4f}'.replace('.', ',')
    story = [Paragraph('BÁBOLNA SPED', title), Paragraph('Sofőrlétszám-kimutatás', heading),
             p(f'Elszámolási hónap: {month}'), p(f'Forrás: {source}'), Spacer(1, 12)]
    values = [['Nemzetközi súlyozott létszám', fmt(frame['Nemzetközi'].sum()) + ' fő'],
              ['Belföldi súlyozott létszám', fmt(frame['Belföldi'].sum()) + ' fő'],
              ['Összes súlyozott létszám', fmt(frame[['Nemzetközi', 'Belföldi']].sum().sum()) + ' fő']]
    table = Table([[p(a), p(b)] for a, b in values], colWidths=[371, 160])
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(YELLOW)),
                               ('TOPPADDING', (0, 0), (-1, -1), 9), ('BOTTOMPADDING', (0, 0), (-1, -1), 9)]))
    story += [table, Spacer(1, 10)]
    included = ~frame['Besorolás'].isin(EXCLUDED)
    story += [p(f'Személyek a forrásban: {len(frame)} | Besorolt: {included.sum()} | '
                f'Pozitív alaplétszámú besorolt: {(included & frame["Alaplétszám"].gt(0)).sum()} | Kihagyott: {(~included).sum()}'),
              p(f'0 kereset miatt kötelezően kihagyva: {frame["Kereset"].eq(0).sum()} személy. Más FEOR-kódú dolgozó is beszámítható kézi besorolással.'),
              p('A súlyozott létszám nem azonos a személyek számával. A számítás a forrás létszámértékeit és a kimutatásban rögzített kategóriaarányokat használja; kerekítés csak a megjelenítésnél.')]
    entrants = monthly_entrants(frame, month)
    story.append(Paragraph(f'Havi belépők ({len(entrants)} fő)', heading))
    story.append(p('Kiegészítő nézet: a kategóriákban is szerepelnek, az összesítésben egyszer számítanak.'))
    if entrants.empty:
        story.append(p('Nincs belépő ebben a hónapban.'))
    else:
        for _, row in entrants.iterrows():
            story.append(p(f'{row["Név"]} ({row["Törzsszám"]}) – belépés: {row["Belépés"]}; besorolás: {row["Besorolás"]}; alaplétszám: {fmt(row["Alaplétszám"])} fő.'))
    for label, subset in grouped(frame, config):
        story.append(Paragraph(label, heading))
        if subset.empty:
            story.append(p('Nincs.'))
            continue
        columns = ['Törzsszám', 'Név', 'Besorolás', 'Alaplétszám', 'Nemzetközi', 'Belföldi']
        rows = [[Paragraph(c, white) for c in columns]]
        note_rows = []
        for _, row in subset.iterrows():
            cells = [p(fmt(row[c]) if c in columns[3:] else row[c]) for c in columns]
            rows.append(cells)
            if row['Megjegyzés']:
                note_rows.append(len(rows))
                rows.append([p('Megjegyzés'), p(row['Megjegyzés']), '', '', '', ''])
        t = Table(rows, colWidths=[66, 149, 97, 73, 73, 73], repeatRows=1, hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(BLUE)),
                              ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                              ('LINEBELOW', (0, 1), (-1, -1), .3, colors.HexColor(GRAY)),
                              ('TOPPADDING', (0, 0), (-1, -1), 6),
                              ('BOTTOMPADDING', (0, 0), (-1, -1), 6)]))
        for note_row in note_rows:
            t.setStyle(TableStyle([('SPAN', (1, note_row), (-1, note_row))]))
        story.append(t)
    story.append(Paragraph('Kategóriák és számítási arányok', heading))
    for name, (n, b) in factors_for(config).items():
        story.append(p(f'{name}: {n*100:g}% nemzetközi, {b*100:g}% belföldi.'))
    checks = audit(frame, month)
    story.append(Paragraph('Időszak- és adatellenőrzés', heading))
    story.append(p('A belépés és kilépés dátuma csak ellenőrzési adat. A számítás az Agroorg alaplétszámát használja, újabb időarányosítás nélkül.'))
    if checks.empty:
        story.append(p('Nincs jelzett eltérés. Hiányzó dátumok esetén a jogviszony időszaka nem ellenőrizhető teljesen.'))
    else:
        for _, check in checks.iterrows():
            story.append(p(f'{check["Név"]} ({check["Törzsszám"]}): {check["Ellenőrzés"]}'))
    doc.build(story)
    return output.getvalue()


def main():
    import streamlit as st
    st.set_page_config(page_title='Bábolna Sped | Sofőrlétszám', layout='wide')
    st.title('Sofőrlétszám')
    st.caption('Bábolna Sped Kft. • Agroorg havi létszámkimutatás')
    if 'category_config' not in st.session_state:
        st.session_state['category_config'] = default_config()
        st.session_state['config_revision'] = 0
    with st.expander('Beállítások – kategóriák és arányok', expanded=False):
        st.caption('100% = 1 fő, 75% = 0,75 fő az alaplétszám minden 1 főjére. A két arány összege nem kötelezően 100%. Új kategóriát az üres alsó sorban adhatsz hozzá. Használt kategória törlése vagy átnevezése előtt sorold át a dolgozókat. A Nemzetközis név és a Kihagyás szabály védett.')
        with st.form('settings_form'):
            settings_config = st.data_editor(st.session_state['category_config'], hide_index=True, width='stretch', num_rows='dynamic',
                key=f'category_settings_{st.session_state["config_revision"]}',
                column_config={col: st.column_config.NumberColumn(min_value=0., step=0.1, format='%.2f', required=True) for col in CONFIG_COLUMNS[1:]})
            save_settings = st.form_submit_button('Kategóriák mentése')
        if save_settings:
            try:
                used = st.session_state['base']['Besorolás'] if 'base' in st.session_state else []
                st.session_state['category_config'] = validate_config(settings_config, used)
                st.session_state['config_revision'] += 1
                st.session_state['revision'] = st.session_state.get('revision', 0) + 1
                st.rerun()
            except ValueError as error:
                st.error(str(error))
    config = st.session_state['category_config']
    categories = config['Kategória'].tolist()
    uploaded = st.file_uploader('Agroorg-export feltöltése', type=['xls', 'xlsx', 'csv'])
    if uploaded is None:
        st.info('Töltsd fel a havi Agroorg-exportot. Minden dolgozó megjelenik. A 8417 FEOR alapbesorolása Nemzetközis, más FEOR esetén Kihagyás. A 0 keresetűek láthatók maradnak, de kimaradnak a számításból.')
        return
    data = uploaded.getvalue()
    fingerprint = sha256(b'v6-entrants-empty-exit' + data + uploaded.name.encode()).hexdigest()
    try:
        if st.session_state.get('source_id') != fingerprint:
            st.session_state['base'] = read_source(data, uploaded.name)
            st.session_state['source_id'] = fingerprint
            st.session_state['revision'] = 0
            st.session_state['period'] = suggested_month(uploaded.name)
        if st.session_state['base'].empty:
            st.info('A fájlban nincs dolgozó.')
            return
        period = st.date_input('Elszámolási hónap – a kiválasztott dátum hónapja számít', key='period')
        month = period.strftime('%Y-%m')
        with st.expander('Korábbi besorolások visszatöltése'):
            st.write('A besorolásokat törzsszám alapján vesszük át. Az aktuális havi 0 Ft kereset mindig kötelező kizárás. A megjegyzések csak azonos elszámolási hónapból tölthetők vissza.')
            previous = st.file_uploader('Korábbi Excel-export', type=['xlsx'], key='previous')
            if st.button('Besorolások visszatöltése', disabled=previous is None):
                loaded_config = restore_config(previous.getvalue())
                restored, count = restore_categories(previous.getvalue(), st.session_state['base'], month, loaded_config)
                st.session_state['category_config'] = loaded_config
                st.session_state['config_revision'] += 1
                config = loaded_config
                categories = config['Kategória'].tolist()
                st.session_state['base'] = restored
                st.session_state['notes_month'] = month
                st.session_state['revision'] += 1
                st.success(f'{count} személy besorolása visszatöltve.')
        st.subheader('Dolgozók besorolás szerint')
        st.caption('A módosítások után nyomd meg a Mentés és újraszámítás gombot. Az eredmények és exportok a mentett állapotot mutatják.')
        base = enforce_rules(st.session_state['base'])
        if st.session_state.get('notes_month') != month:
            base['Megjegyzés'] = ''
        parts = []
        with st.form('classification_form', enter_to_submit=False):
            groups = grouped(base, config)
            entrants = monthly_entrants(base, month)
            tabs = st.tabs([f'Havi belépők ({len(entrants)})'] + [f'{label} ({len(subset)})' for label, subset in groups])
            with tabs[0]:
                st.caption('Itt is módosíthatod a besorolást, majd nyomd meg a Mentés és újraszámítás gombot. Az összesítés egyszer számolja a belépőket.')
                if entrants.empty:
                    st.info('Ebben a hónapban nincs belépő.')
                else:
                    active_entrants = entrants.loc[entrants['Kereset'].ne(0)]
                    zero_entrants = entrants.loc[entrants['Kereset'].eq(0)]
                    columns = ['Törzsszám', 'Név', 'Belépés', 'Kilépés', 'FEOR', 'Alaplétszám', 'Kereset', 'Besorolás']
                    if not active_entrants.empty:
                        parts.append(st.data_editor(active_entrants[columns], hide_index=True, width='stretch',
                            key=f'entrants_{fingerprint}_{month}_{st.session_state["revision"]}',
                            disabled=[column for column in columns if column != 'Besorolás'],
                            column_config={'Besorolás': st.column_config.SelectboxColumn(options=categories, required=True),
                                           'Alaplétszám': st.column_config.NumberColumn(format='%.4f')}))
                    if not zero_entrants.empty:
                        st.caption('0 Ft kereset: kötelező kihagyás. A megjegyzés szerkeszthető.')
                        parts.append(st.data_editor(zero_entrants[columns + ['Megjegyzés']], hide_index=True, width='stretch',
                            key=f'entrants_zero_{fingerprint}_{month}_{st.session_state["revision"]}',
                            disabled=columns,
                            column_config={'Megjegyzés': st.column_config.TextColumn(width='large', max_chars=500)}))
            for tab, (label, subset) in zip(tabs[1:], groups):
                with tab:
                    if subset.empty:
                        st.caption('Nincs dolgozó ebben a csoportban.')
                        continue
                    active = subset.loc[subset['Kereset'].ne(0)]
                    zero = subset.loc[subset['Kereset'].eq(0)]
                    if not active.empty:
                        parts.append(st.data_editor(active[BASE_COLUMNS + ['Kereset']], hide_index=True, width='stretch',
                            key=f'group_{fingerprint}_{month}_{label}_{st.session_state["revision"]}',
                            disabled=['Törzsszám', 'Név', 'FEOR', 'Alaplétszám', 'Kereset'],
                            column_config={'Besorolás': st.column_config.SelectboxColumn(options=categories, required=True),
                                           'Alaplétszám': st.column_config.NumberColumn(format='%.4f')}))
                    if not zero.empty:
                        st.caption('0 Ft kereset: kötelező kihagyás. A megjegyzés szerkeszthető.')
                        parts.append(st.data_editor(zero[BASE_COLUMNS + ['Kereset', 'Megjegyzés']], hide_index=True, width='stretch',
                            key=f'zero_{fingerprint}_{month}_{st.session_state["revision"]}',
                            disabled=BASE_COLUMNS + ['Kereset'],
                            column_config={'Megjegyzés': st.column_config.TextColumn(width='large', max_chars=500)}))
            submitted = st.form_submit_button('Mentés és újraszámítás')
        if submitted:
            updated = base.set_index('Törzsszám').copy()
            changes = {}
            for part in parts:
                patch = part.set_index('Törzsszám')
                for column in ['Besorolás', 'Megjegyzés']:
                    if column in patch:
                        # Only changed cells: an unchanged duplicate view must not
                        # overwrite an edit made on the monthly entrants tab.
                        for employee, value in patch[column].items():
                            if value == updated.loc[employee, column]:
                                continue
                            cell = (employee, column)
                            if cell in changes and changes[cell] != value:
                                raise ValueError('Ugyanazt a mezőt két fülön eltérően módosítottad. Állítsd azonosra a Havi belépők és a kategóriafül értékét, majd ments újra.')
                            changes[cell] = value
            for (employee, column), value in changes.items():
                updated.loc[employee, column] = value
            st.session_state['base'] = enforce_rules(updated.reset_index())
            st.session_state['notes_month'] = month
            st.session_state['revision'] += 1
            st.rerun()
        edited = base
        if suggested_month(uploaded.name).strftime('%Y-%m') != month:
            st.warning('A kiválasztott hónap eltér a fájlnévből felajánlott hónaptól. A hónap átállítása nem számítja újra az Agroorg alaplétszámát.')
        with st.expander('Hóközi be-/kilépés és lehetséges adathibák', expanded=False):
            st.write('A számítás az Agroorg letsz mezőjét használja. Nem szorozzuk meg újra a hónapban töltött napok arányával. A naptári arány összehasonlító adat, nem automatikus korrekció.')
            issues = audit(edited, month)
            if issues.empty:
                st.info('Nincs jelzett eltérés a rendelkezésre álló adatokban.')
            else:
                st.dataframe(issues, hide_index=True, width='stretch')
            if edited['Belépés'].eq('').any():
                st.caption('Egyes dolgozóknál hiányzik a belépési dátum; náluk az időszak ellenőrzése nem teljes.')
        calculated = calculate(edited, config)
        st.subheader('Eredmény')
        n, b, total = st.columns(3)
        n.metric('Nemzetközi súlyozott létszám', f'{calculated["Nemzetközi"].sum():.4f}'.replace('.', ',') + ' fő')
        b.metric('Belföldi súlyozott létszám', f'{calculated["Belföldi"].sum():.4f}'.replace('.', ',') + ' fő')
        total.metric('Összes súlyozott létszám', f'{calculated[["Nemzetközi", "Belföldi"]].sum().sum():.4f}'.replace('.', ',') + ' fő')
        included = ~edited['Besorolás'].isin(EXCLUDED)
        st.caption(f'Forrás: {len(edited)} személy • Besorolt: {included.sum()} • Pozitív alaplétszámú besorolt: '
                   f'{(included & edited["Alaplétszám"].gt(0)).sum()} • Kihagyott: {(~included).sum()}')
        st.caption('A számítás a mentett kategóriaarányokat használja. A köztes értékeket nem kerekítjük.')
        with st.expander('Számítás részletei'):
            st.dataframe(calculated, hide_index=True, width='stretch')
        st.subheader('Export')
        # Mindkét export az aktuális felületi állapotból készül.
        left, right = st.columns(2)
        left.download_button('Excel letöltése', excel_export(edited, month, uploaded.name, config),
                             file_name=f'letszam_{month}.xlsx',
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        right.download_button('PDF letöltése', pdf_export(edited, month, uploaded.name, config),
                              file_name=f'letszam_{month}.pdf', mime='application/pdf')
        st.caption('A besorolásokat az Excel-export őrzi meg. A böngésző bezárása után ebből tölthetők vissza.')
    except Exception as error:
        st.error(f'A feldolgozás nem fejezhető be: {error}')


if __name__ == '__main__':
    main()
