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
CATEGORIES = ['Nemzetközis', 'Belföldes', 'Hibrid', 'Váltó / Ónódi', 'Kihagyás']
FACTORS = {'Nemzetközis': (1., 0.), 'Belföldes': (0., 1.),
           'Hibrid': (.8, .2), 'Váltó / Ónódi': (.625, 0.), 'Kihagyás': (0., 0.)}
BASE_COLUMNS = ['Törzsszám', 'Név', 'FEOR', 'Alaplétszám', 'Besorolás']
EXTRA_COLUMNS = ['Kereset', 'Kihagyás oka']
ZERO_REASON = '0 kereset miatt kihagyva'


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
    # Más munkakörök személyei egyetlen listába vagy exportba sem kerülnek.
    raw = raw.loc[raw['feorkód'].map(identifier).eq('8417')].copy()
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
    names = raw['nev'].fillna('').astype(str).str.strip()
    if ids.eq('').any() or names.eq('').any():
        raise ValueError('Hiányzó törzsszám vagy név. Javítsd a forrást.')
    if ids.duplicated().any():
        raise ValueError('Ismétlődő törzsszám: ' + ', '.join(ids[ids.duplicated()].unique()) +
                         '. Az eltérő jogviszonyok összesítését ellenőrizni kell.')
    feor = raw['feorkód'].map(identifier)
    result = pd.DataFrame({'Törzsszám': ids, 'Név': names, 'FEOR': feor,
                           'Alaplétszám': amounts, 'Besorolás': 'Nemzetközis',
                           'Kereset': earnings, 'Kihagyás oka': ''}).sort_values('Név').reset_index(drop=True)
    return enforce_rules(result)


def enforce_rules(frame):
    """A kötelező kizárás import, visszatöltés, számítás és export során is érvényes."""
    result = frame.loc[frame['FEOR'].map(identifier).eq('8417')].copy()
    if result['Kereset'].isna().any() or not np.isfinite(result['Kereset']).all():
        raise ValueError('Hiányzó vagy hibás kereset.')
    zero = result['Kereset'].eq(0)
    result.loc[zero, 'Besorolás'] = 'Kihagyás'
    result['Kihagyás oka'] = np.where(zero, ZERO_REASON,
                                    np.where(result['Besorolás'].eq('Kihagyás'), 'Kézi kihagyás', ''))
    return result


def suggested_month(filename):
    matches = re.findall(r'(20\d{2})(0[1-9]|1[0-2])(?!\d)', Path(filename).stem)
    return date(int(matches[-1][0]), int(matches[-1][1]), 1) if matches else date.today().replace(day=1)


def restore_categories(data, current):
    """Csak a besorolás kerül vissza, az alaplétszám és a név nem."""
    saved = pd.read_excel(BytesIO(data), sheet_name='Besorolások', dtype={'Törzsszám': str})
    if not {'Törzsszám', 'Besorolás'}.issubset(saved.columns):
        raise ValueError('A kiválasztott fájl nem az alkalmazás Excel-exportja.')
    saved['Törzsszám'] = saved['Törzsszám'].map(identifier)
    if saved['Törzsszám'].eq('').any() or saved['Törzsszám'].duplicated().any():
        raise ValueError('A mentett besorolások törzsszámai hiányosak vagy ismétlődnek.')
    if not saved['Besorolás'].isin(CATEGORIES).all():
        raise ValueError('Ismeretlen kategória található a mentett besorolásokban.')
    # Az előző hónap automatikus nullás kizárása nem válik tartós kézi kizárássá.
    if 'Kihagyás oka' in saved:
        saved.loc[saved['Kihagyás oka'].eq(ZERO_REASON), 'Besorolás'] = 'Nemzetközis'
    mapping = saved.set_index('Törzsszám')['Besorolás']
    result = current.copy()
    result['Besorolás'] = result['Törzsszám'].map(mapping).fillna('Nemzetközis')
    result = enforce_rules(result)
    return result, int(result['Törzsszám'].isin(mapping.index).sum())


def calculate(frame):
    frame = enforce_rules(frame)
    if not frame['Besorolás'].isin(CATEGORIES).all():
        raise ValueError('Minden személynél válassz érvényes besorolást.')
    result = frame.copy()
    result['N szorzó'] = result['Besorolás'].map(lambda c: FACTORS[c][0])
    result['B szorzó'] = result['Besorolás'].map(lambda c: FACTORS[c][1])
    result['Nemzetközi'] = result['Alaplétszám'] * result['N szorzó']
    result['Belföldi'] = result['Alaplétszám'] * result['B szorzó']
    return result


def excel_export(frame, month, source):
    """Az alkalmazás futás közben készített Excel-exportja képletekkel és gyorsítótárazott eredménnyel."""
    import xlsxwriter
    frame = calculate(frame)
    output = BytesIO()
    book = xlsxwriter.Workbook(output, {'in_memory': True, 'strings_to_formulas': False,
                                      'strings_to_urls': False})
    normal = book.add_format({'font_name': 'Calibri', 'font_size': 11, 'font_color': BLUE})
    number = book.add_format({'font_name': 'Calibri', 'num_format': '0.0000', 'font_color': BLUE})
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
        sheet.set_column('C:K', 18, normal)
        sheet.set_landscape()
        sheet.fit_to_pages(1, 0)
    # A besorolások munkalap az exportból történő visszatöltéshez is használható.
    settings.write_row(0, 0, BASE_COLUMNS + EXTRA_COLUMNS, head)
    settings.set_column('G:G', 30, normal)
    settings.set_row(0, 32)
    for i, row in enumerate(frame[BASE_COLUMNS + EXTRA_COLUMNS].itertuples(index=False, name=None), 1):
        settings.write_row(i, 0, row, normal)
        settings.write_number(i, 3, row[3], number)
        settings.write_number(i, 5, row[5], money)
    settings.data_validation(1, 4, len(frame), 4, {'validate': 'list', 'source': CATEGORIES})
    settings.freeze_panes(1, 2)
    settings.autofilter(0, 0, len(frame), 6)
    headers = BASE_COLUMNS + ['N szorzó', 'B szorzó', 'Nemzetközi', 'Belföldi'] + EXTRA_COLUMNS
    detail.set_column('K:K', 30, normal)
    detail.write_row(0, 0, headers, head)
    detail.set_row(0, 32)
    for i, (_, row) in enumerate(frame.iterrows(), 1):
        r = i + 1
        for j, col in enumerate(BASE_COLUMNS):
            detail.write_formula(i, j, f"='Besorolások'!{chr(65+j)}{r}",
                                 number if j == 3 else normal, row[col])
        # Kereset=0 esetén az Excelben átírt besorolás sem változtathatja meg a számítást.
        detail.write_formula(i, 4, f'IF(J{r}=0,"Kihagyás",\'Besorolások\'!E{r})', normal, row['Besorolás'])
        detail.write_formula(i, 9, f"='Besorolások'!F{r}", money, row['Kereset'])
        detail.write_formula(i, 10, f'IF(J{r}=0,"{ZERO_REASON}",IF(E{r}="Kihagyás","Kézi kihagyás",""))', normal, row['Kihagyás oka'])
        detail.write_formula(i, 5, f'IF(E{r}="Nemzetközis",1,IF(E{r}="Hibrid",0.8,IF(E{r}="Váltó / Ónódi",0.625,0)))', number, row['N szorzó'])
        detail.write_formula(i, 6, f'IF(E{r}="Belföldes",1,IF(E{r}="Hibrid",0.2,0))', number, row['B szorzó'])
        detail.write_formula(i, 7, f'D{r}*F{r}', number, row['Nemzetközi'])
        detail.write_formula(i, 8, f'D{r}*G{r}', number, row['Belföldi'])
    detail.freeze_panes(1, 2)
    detail.autofilter(0, 0, len(frame), 10)
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
    summary.write_row('A12', ['8417 FEOR-kódú személyek', len(frame)], normal)
    summary.write('A13', 'Besorolt személyek (nullás értékkel is)', normal)
    summary.write_formula('B13', f'COUNTIF(\'Részletezés\'!E2:E{end},"<>Kihagyás")', normal, int(frame['Besorolás'].ne('Kihagyás').sum()))
    summary.write('A14', 'Pozitív alaplétszámú besorolt személyek', normal)
    summary.write_formula('B14', f'COUNTIFS(\'Részletezés\'!E2:E{end},"<>Kihagyás",\'Részletezés\'!D2:D{end},">0")', normal,
                          int((frame['Besorolás'].ne('Kihagyás') & frame['Alaplétszám'].gt(0)).sum()))
    summary.write('A15', 'Kihagyott személyek', normal)
    summary.write_formula('B15', f'COUNTIF(\'Részletezés\'!E2:E{end},"Kihagyás")', normal, int(frame['Besorolás'].eq('Kihagyás').sum()))
    summary.write('A16', 'Ebből 0 kereset miatt kihagyva', normal)
    summary.write_formula('B16', f'COUNTIF(\'Részletezés\'!J2:J{end},0)', normal, int(frame['Kereset'].eq(0).sum()))
    for i, line in enumerate(['Mértékegység: fő. A súlyozott létszám nem személyek darabszáma.',
                             'Hibrid: 80% nemzetközi, 20% belföldi. Váltó: 62,5% nemzetközi.',
                             'A Besorolások lap szerkesztése frissíti a Részletezést és az Összesítést.',
                             'A Kihagyottak lap az exportáláskori állapotot rögzíti.'], 17):
        summary.merge_range(i, 0, i, 4, line, normal)
    summary.merge_range('A22:E22', 'Csak 8417 FEOR. Kereset = 0: kötelező kihagyás, látható személy.', normal)
    skipped.write_row(0, 0, BASE_COLUMNS + EXTRA_COLUMNS, head)
    skipped.set_column('G:G', 30, normal)
    for i, row in enumerate(frame.loc[frame['Besorolás'].eq('Kihagyás'), BASE_COLUMNS + EXTRA_COLUMNS].itertuples(index=False, name=None), 1):
        skipped.write_row(i, 0, row, normal)
        skipped.write_number(i, 3, row[3], number)
        skipped.write_number(i, 5, row[5], money)
    book.close()
    return output.getvalue()


def pdf_export(frame, month, source):
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
    frame = calculate(frame)
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, leftMargin=32, rightMargin=32,
                            topMargin=32, bottomMargin=32, title=f'Sofőrlétszám – {month}', author='Bábolna Sped Kft.')
    body = ParagraphStyle('body', fontName='Report', fontSize=8, leading=11, textColor=colors.HexColor(BLUE))
    title = ParagraphStyle('title', parent=body, fontName='ReportBold', fontSize=17, leading=22, spaceAfter=10)
    heading = ParagraphStyle('heading', parent=body, fontName='ReportBold', fontSize=11, leading=15, spaceBefore=14, spaceAfter=7)
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
    included = frame['Besorolás'].ne('Kihagyás')
    story += [p(f'8417 FEOR-kódú személyek: {len(frame)} | Besorolt: {included.sum()} | '
                f'Pozitív alaplétszámú besorolt: {(included & frame["Alaplétszám"].gt(0)).sum()} | Kihagyott: {(~included).sum()}'),
              p(f'0 kereset miatt kötelezően kihagyva: {frame["Kereset"].eq(0).sum()} személy. Más FEOR-kódú dolgozók nem szerepelnek a kimutatásban.'),
              p('A súlyozott létszám nem azonos a személyek számával. Hibrid: 80% nemzetközi, 20% belföldi. Váltó / Ónódi: 62,5% nemzetközi. A számítás a forrás létszámértékeiből történik, kerekítés csak a megjelenítésnél.')]
    for label, subset in [('Besorolt személyek', frame[included]), ('Kihagyott személyek', frame[~included])]:
        story.append(Paragraph(label, heading))
        if subset.empty:
            story.append(p('Nincs.'))
            continue
        columns = ['Törzsszám', 'Név', 'Besorolás', 'Alaplétszám', 'Nemzetközi', 'Belföldi']
        rows = [[Paragraph(c, white) for c in columns]]
        for _, row in subset.iterrows():
            cells = [p(fmt(row[c]) if c in columns[3:] else row[c]) for c in columns]
            if row['Kihagyás oka']:
                cells[2] = p(row['Kihagyás oka'])
            rows.append(cells)
        t = Table(rows, colWidths=[66, 149, 97, 73, 73, 73], repeatRows=1, hAlign='LEFT')
        t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(BLUE)),
                              ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                              ('LINEBELOW', (0, 1), (-1, -1), .3, colors.HexColor(GRAY)),
                              ('TOPPADDING', (0, 0), (-1, -1), 6),
                              ('BOTTOMPADDING', (0, 0), (-1, -1), 6)]))
        story.append(t)
    doc.build(story)
    return output.getvalue()


def main():
    import streamlit as st
    st.set_page_config(page_title='Bábolna Sped | Sofőrlétszám', layout='wide')
    st.title('Sofőrlétszám')
    st.caption('Bábolna Sped Kft. • Agroorg havi létszámkimutatás')
    uploaded = st.file_uploader('Agroorg-export feltöltése', type=['xls', 'xlsx', 'csv'])
    if uploaded is None:
        st.info('Töltsd fel a havi Agroorg-exportot. Csak a 8417 FEOR-kódú dolgozók jelennek meg. A 0 keresetűek láthatók maradnak, de kimaradnak a számításból.')
        return
    data = uploaded.getvalue()
    fingerprint = sha256(b'v2-feor-earnings' + data + uploaded.name.encode()).hexdigest()
    try:
        if st.session_state.get('source_id') != fingerprint:
            st.session_state['base'] = read_source(data, uploaded.name)
            st.session_state['source_id'] = fingerprint
            st.session_state['revision'] = 0
            st.session_state['period'] = suggested_month(uploaded.name)
        if st.session_state['base'].empty:
            st.info('A fájlban nincs 8417 FEOR-kódú dolgozó.')
            return
        period = st.date_input('Elszámolási hónap – a kiválasztott dátum hónapja számít', key='period')
        month = period.strftime('%Y-%m')
        with st.expander('Korábbi besorolások visszatöltése'):
            st.write('Csak a besorolásokat vesszük át törzsszám alapján. Az új személyek Nemzetközisként indulnak. Az aktuális havi 0 kereset mindig kötelező kihagyás. A korábbi automatikus nullás kizárást nem visszük át egy nem nullás hónapra.')
            previous = st.file_uploader('Korábbi Excel-export', type=['xlsx'], key='previous')
            if st.button('Besorolások visszatöltése', disabled=previous is None):
                restored, count = restore_categories(previous.getvalue(), st.session_state['base'])
                st.session_state['base'] = restored
                st.session_state['revision'] += 1
                st.success(f'{count} személy besorolása visszatöltve.')
        st.subheader('Személyenkénti besorolás')
        st.info('Csak 8417 FEOR-kódú dolgozók. A nem nulla keresetűek alapbesorolása Nemzetközis. A 0 keresetűek alább, külön listában láthatók; besorolásuk kötelező Kihagyás.')
        base = enforce_rules(st.session_state['base'])
        zero_rows = base.loc[base['Kereset'].eq(0)].copy()
        active_rows = base.loc[base['Kereset'].ne(0)].copy()
        edited_active = st.data_editor(active_rows, hide_index=True, width='stretch', height=520,
                               key=f'editor_{fingerprint}_{st.session_state["revision"]}',
                               disabled=['Törzsszám', 'Név', 'FEOR', 'Alaplétszám', 'Kereset', 'Kihagyás oka'],
                               column_config={'Besorolás': st.column_config.SelectboxColumn('Besorolás', options=CATEGORIES, required=True),
                                              'Alaplétszám': st.column_config.NumberColumn(format='%.4f'),
                                              'Kereset': st.column_config.NumberColumn(format='%.2f')})
        st.subheader(f'0 kereset miatt kihagyva – {len(zero_rows)} személy')
        if zero_rows.empty:
            st.caption('Nincs 0 keresetű, 8417 FEOR-kódú dolgozó.')
        else:
            st.dataframe(zero_rows, hide_index=True, width='stretch')
        edited = enforce_rules(pd.concat([edited_active, zero_rows]).sort_values('Név').reset_index(drop=True))
        calculated = calculate(edited)
        st.subheader('Eredmény')
        n, b, total = st.columns(3)
        n.metric('Nemzetközi súlyozott létszám', f'{calculated["Nemzetközi"].sum():.4f}'.replace('.', ',') + ' fő')
        b.metric('Belföldi súlyozott létszám', f'{calculated["Belföldi"].sum():.4f}'.replace('.', ',') + ' fő')
        total.metric('Összes súlyozott létszám', f'{calculated[["Nemzetközi", "Belföldi"]].sum().sum():.4f}'.replace('.', ',') + ' fő')
        included = edited['Besorolás'].ne('Kihagyás')
        st.caption(f'8417 FEOR: {len(edited)} személy • Besorolt: {included.sum()} • Pozitív alaplétszámú besorolt: '
                   f'{(included & edited["Alaplétszám"].gt(0)).sum()} • Kihagyott: {(~included).sum()}')
        st.caption('Hibrid: 80% nemzetközi, 20% belföldi. Váltó / Ónódi: 62,5% nemzetközi. A köztes értékeket nem kerekítjük.')
        with st.expander('Számítás részletei'):
            st.dataframe(calculated, hide_index=True, width='stretch')
        st.subheader('Export')
        # Mindkét export az aktuális felületi állapotból készül.
        left, right = st.columns(2)
        left.download_button('Excel letöltése', excel_export(edited, month, uploaded.name),
                             file_name=f'letszam_{month}.xlsx',
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        right.download_button('PDF letöltése', pdf_export(edited, month, uploaded.name),
                              file_name=f'letszam_{month}.pdf', mime='application/pdf')
        st.caption('A besorolásokat az Excel-export őrzi meg. A böngésző bezárása után ebből tölthetők vissza.')
    except Exception as error:
        st.error(f'A feldolgozás nem fejezhető be: {error}')


if __name__ == '__main__':
    main()
