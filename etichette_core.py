"""Lettura dati e creazione dei PDF per le etichette locali."""
import csv
import datetime
import html
import json
import math
import os
import re
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
import reportlab

PAPER_PRESETS = {
    'UT101700 · 70 × 36 mm · 24 per A4': {
        'pw': 210, 'ph': 297, 'w': 70, 'h': 36, 'left': 0, 'top': 6,
        'dx': 70, 'dy': 36, 'cols': 3, 'rows': 8,
    },
    '99,1 × 34 mm · 16 per A4': {
        'pw': 210, 'ph': 297, 'w': 99.1, 'h': 34, 'left': 4.65, 'top': 12.5,
        'dx': 101.6, 'dy': 34, 'cols': 2, 'rows': 8,
    },
}
CUSTOM_PAPER = 'Personalizzato'
POSTAL = 'Nome · Comune · Indirizzo'
PERSONAL = 'Anagrafica completa'
SELECTED = 'Campi selezionati'


def cell_text(value):
    if value is None:
        return ''
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime('%d/%m/%Y')
    if isinstance(value, datetime.time):
        return value.strftime('%H:%M:%S')
    return str(value).strip()


def detect_header(rows):
    best_index, best_score = 0, -1
    for i, row in enumerate(rows[:15]):
        values = [cell_text(value) for value in row if cell_text(value)]
        known = sum(bool(re.search(r'cognome|nominativo|indirizzo|data nascita|codice fiscale|comune nascita|nome|residenza', value, re.I)) for value in values)
        score = known * 5 + len(values)
        if score > best_score:
            best_index, best_score = i, score
    return best_index


def apply_header(rows, header_index):
    if not 0 <= header_index < len(rows):
        raise ValueError('Riga delle intestazioni non valida.')
    count = max((len(row) for row in rows[header_index:]), default=0)
    if not count or count > 200:
        raise ValueError('Tabella vuota o con più di 200 colonne.')
    header = [cell_text(rows[header_index][i]) if i < len(rows[header_index]) else '' for i in range(count)]
    header = [name or f'Campo {i + 1}' for i, name in enumerate(header)]
    data = []
    for row in rows[header_index + 1:]:
        values = [cell_text(row[i]) if i < len(row) else '' for i in range(count)]
        if any(values):
            data.append(values)
    return header, data


class _Tables(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.depth = 0
        self.tables = []
        self.row = None
        self.cell = None
        self.hidden = False

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == 'table':
            self.depth += 1
            if self.depth == 1:
                self.tables.append([])
        elif self.depth == 1 and tag == 'tr':
            self.row = []
        elif self.depth == 1 and tag in ('td', 'th') and self.row is not None:
            self.cell = []
            self.hidden = '_OLDKEY' in attrs.get('class', '').split() or bool(re.search(r'display\s*:\s*none', attrs.get('style', ''), re.I))
        elif self.cell is not None and tag == 'br':
            self.cell.append(' ')

    def handle_data(self, data):
        if self.depth == 1 and self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if self.depth == 1 and tag in ('td', 'th') and self.cell is not None:
            if not self.hidden:
                self.row.append(' '.join(''.join(self.cell).split()))
            self.cell = None
        elif self.depth == 1 and tag == 'tr' and self.row is not None:
            if self.row:
                self.tables[-1].append(self.row)
            self.row = None
        elif tag == 'table' and self.depth:
            self.depth -= 1


def read_tables(path):
    path = Path(path)
    if path.stat().st_size > 30_000_000:
        raise ValueError('Il file supera 30 MB: suddividilo in file più piccoli.')
    if path.suffix.lower() in ('.xlsx', '.xlsm'):
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            tables = []
            for sheet in workbook.worksheets:
                rows = [[cell_text(value) for value in row] for row in sheet.iter_rows(values_only=True)]
                tables.append((sheet.title, rows))
            return tables
        finally:
            workbook.close()
    content = path.read_bytes()
    if content.startswith(bytes.fromhex('D0CF11E0')):
        raise ValueError('Questo è un vecchio Excel binario .xls. Aprilo in Excel e salvalo come .xlsx.')
    try:
        text = content.decode('utf-8-sig')
    except UnicodeDecodeError:
        text = content.decode('cp1252')
    if re.search(r'<table[\s>]', text, re.I):
        parser = _Tables()
        parser.feed(text)
        return [(f'Tabella {i+1}', rows) for i, rows in enumerate(parser.tables) if rows]
    sample = text[:10000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=';\t,')
        delimiter = dialect.delimiter
    except csv.Error:
        first = sample.splitlines()[0] if sample else ''
        delimiter = max(';\t,', key=first.count)
    import io
    rows = [row for row in csv.reader(io.StringIO(text, newline=''), delimiter=delimiter) if any(cell_text(v) for v in row)]
    return [(path.name, rows)] if rows else []


def field_index(header, pattern):
    return next((i for i, value in enumerate(header) if re.search(pattern, value, re.I)), None)


def initial_fields(header):
    name_index = field_index(header, r'cognome.*nome|nome.*cognome|nominativo')
    address = field_index(header, r'indirizzo|\bvia\b')
    is_person = name_index is not None and address is not None
    fields = [{'i': i, 'name': name, 'enabled': not is_person or i in (name_index, address)}
              for i, name in enumerate(header)]
    return fields, is_person


def default_template(header, fields, preset):
    name = field_index(header, r'cognome.*nome|nome.*cognome|nominativo')
    birth = field_index(header, r'comune.*nasc|luogo.*nasc')
    date = field_index(header, r'data.*nasc')
    address = field_index(header, r'indirizzo|\bvia\b|residenza')
    token = lambda i: '{c' + str(i) + '}' if i is not None else ''
    if preset == SELECTED:
        lines = [token(field['i']) for field in fields if field['enabled']]
    elif preset == PERSONAL:
        lines = [token(name), '{nato} a ' + token(birth) if birth is not None else '',
                 'il ' + token(date) if date is not None else '', 'Residente a {residenza}', token(address)]
    else:
        lines = [token(name), '{residenza}', token(address)]
    return '\n'.join(line for line in lines if line)


def label_text(template, row, header, town):
    sex = field_index(header, r'^sesso$')
    def replace(match):
        key = match.group(1)
        if key == 'residenza':
            return town
        if key == 'nato':
            return 'Nata' if sex is not None and sex < len(row) and cell_text(row[sex]).upper() == 'F' else 'Nato'
        if re.fullmatch(r'c\d+', key):
            i = int(key[1:])
            return cell_text(row[i]) if i < len(row) else ''
        return match.group(0)
    return re.sub(r'\{([^}]+)\}', replace, template)


def validate_geometry(g, n_items):
    for key in ('pw', 'ph', 'w', 'h', 'left', 'top', 'dx', 'dy', 'cols', 'rows', 'font', 'padding', 'firstPad', 'skip', 'copies'):
        if key not in g or not isinstance(g[key], (int, float)) or not math.isfinite(g[key]) or g[key] < 0:
            raise ValueError('Misure non valide: controlla tutti i valori.')
    if min(g['pw'], g['ph'], g['w'], g['h']) <= 0 or any(int(g[k]) != g[k] or g[k] < 1 for k in ('cols', 'rows')):
        raise ValueError('Dimensioni, righe e colonne devono essere maggiori di zero.')
    if g['cols'] * g['rows'] > 500:
        raise ValueError('Sono ammesse al massimo 500 etichette per foglio.')
    if g['dx'] < g['w'] or g['dy'] < g['h']:
        raise ValueError('Il passo deve essere almeno pari alla dimensione dell’etichetta.')
    if g['left'] + (g['cols'] - 1) * g['dx'] + g['w'] > g['pw'] + .01 or g['top'] + (g['rows'] - 1) * g['dy'] + g['h'] > g['ph'] + .01:
        raise ValueError('Le etichette escono dal foglio: controlla margini, passi e quantità.')
    if g['padding'] > 15 or g['firstPad'] > 25 or max(g['padding'], g['firstPad']) + g['padding'] >= g['w'] or g['padding'] * 2 >= g['h']:
        raise ValueError('Lo spazio interno è troppo grande.')
    if not 5 <= g['font'] <= 30 or int(g['copies']) != g['copies'] or not 1 <= g['copies'] <= 100:
        raise ValueError('Controlla la dimensione del testo e il numero di copie.')
    if int(g['skip']) != g['skip'] or g['skip'] >= g['cols'] * g['rows']:
        raise ValueError('Le posizioni già usate devono essere meno delle etichette per foglio.')
    if n_items * g['copies'] > 10000:
        raise ValueError('Massimo 10.000 etichette per operazione.')


def register_fonts():
    if 'Etichette-Regular' not in pdfmetrics.getRegisteredFontNames():
        fontdir = Path(reportlab.__file__).parent / 'fonts'
        pdfmetrics.registerFont(TTFont('Etichette-Regular', str(fontdir / 'Vera.ttf')))
        pdfmetrics.registerFont(TTFont('Etichette-Bold', str(fontdir / 'VeraBd.ttf')))


def _split_line(line, font, size, available):
    if not line:
        return ['']
    result, current = [], ''
    for word in line.split():
        candidate = (current + ' ' + word).strip()
        if pdfmetrics.stringWidth(candidate, font, size) <= available:
            current = candidate
            continue
        if current:
            result.append(current)
            current = ''
        for character in word:
            if pdfmetrics.stringWidth(current + character, font, size) > available:
                if not current:
                    raise ValueError('Il carattere non entra nello spazio disponibile.')
                result.append(current)
                current = character
            else:
                current += character
    if current:
        result.append(current)
    return result or ['']


def label_lines(text, g, first_column):
    register_fonts()
    size = g['font']
    left_padding = max(g['padding'], g['firstPad']) if first_column else g['padding']
    available = (g['w'] - left_padding - g['padding']) * mm
    lines = []
    for index, line in enumerate(text.split('\n')):
        font = 'Etichette-Bold' if index == 0 else 'Etichette-Regular'
        lines.extend((part, font) for part in _split_line(line, font, size, available))
    ascent, descent = pdfmetrics.getAscentDescent('Etichette-Bold', size)
    height = (len(lines) - 1) * size * 1.18 + ascent - descent
    if height > (g['h'] - 2 * g['padding']) * mm + .01:
        raise ValueError('Il testo non entra: riduci il carattere, lo spazio interno oppure il testo.')
    return lines, left_padding, height


def page_count(g, count):
    if not count:
        return 0
    return math.ceil((g['skip'] + count * g['copies']) / (g['cols'] * g['rows']))


def page_layout(g, items, header, template, town, page):
    cap = g['cols'] * g['rows']
    total = len(items) * g['copies']
    first = max(0, page * cap - g['skip'])
    last = min(total, (page + 1) * cap - g['skip'])
    for n in range(first, last):
        slot = (n + g['skip']) % cap
        row_number = n // g['copies']
        data_row = items[row_number]
        text = label_text(template, data_row, header, town)
        try:
            lines, pad_left, content_height = label_lines(text, g, slot % g['cols'] == 0)
        except ValueError as error:
            raise ValueError(f'Etichetta {row_number + 1}: {error}') from error
        yield {
            'slot': slot, 'x': g['left'] + (slot % g['cols']) * g['dx'],
            'y': g['top'] + (slot // g['cols']) * g['dy'], 'lines': lines,
            'pad_left': pad_left, 'content_height': content_height,
        }


def create_labels_pdf(path, items, header, template, town, g):
    validate_geometry(g, len(items))
    if not items:
        raise ValueError('Nessuna persona selezionata.')
    if not template.strip():
        raise ValueError('Il testo delle etichette è vuoto.')
    # Verifica tutte le etichette prima di scrivere il PDF, senza tagliare testo.
    layouts = [list(page_layout(g, items, header, template, town, p)) for p in range(page_count(g, len(items)))]
    document = canvas.Canvas(str(path), pagesize=(g['pw'] * mm, g['ph'] * mm))
    ascent, _ = pdfmetrics.getAscentDescent('Etichette-Bold', g['font'])
    for page in layouts:
        for label in page:
            x = (label['x'] + label['pad_left']) * mm
            middle = (g['ph'] - label['y'] - g['h'] / 2) * mm
            baseline = middle + label['content_height'] / 2 - ascent
            for text, font in label['lines']:
                document.setFont(font, g['font'])
                document.drawString(x, baseline, text)
                baseline -= g['font'] * 1.18
        document.showPage()
    document.save()


def create_list_pdf(path, items, fields, town, include_town, title):
    register_fonts()
    active = [field for field in fields if field['enabled']]
    if include_town:
        active.append({'i': -1, 'name': 'Comune di residenza'})
    if not items or not active:
        raise ValueError('Seleziona almeno una persona e un campo per l’elenco.')
    if len(active) > 12:
        raise ValueError('Per un elenco leggibile seleziona al massimo 12 campi.')
    normal = ParagraphStyle('etichette', fontName='Etichette-Regular', fontSize=8, leading=10, alignment=TA_LEFT, splitLongWords=True)
    bold = ParagraphStyle('etichette-bold', parent=normal, fontName='Etichette-Bold')
    def paragraph(value, style):
        return Paragraph(html.escape(cell_text(value)).replace('\n', '<br/>') or ' ', style)
    rows = [[paragraph(field['name'], bold) for field in active]]
    for item in items:
        rows.append([paragraph(town if field['i'] == -1 else (item[field['i']] if field['i'] < len(item) else ''), normal) for field in active])
    document = SimpleDocTemplate(str(path), pagesize=(210*mm, 297*mm), leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm, title=title)
    table = Table(rows, colWidths=[186*mm/len(active)]*len(active), repeatRows=1, hAlign='LEFT')
    table.setStyle(TableStyle([('VALIGN', (0,0),(-1,-1),'TOP'),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#e8edf3')),('LINEBELOW',(0,0),(-1,0),.5,colors.grey),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3)]))
    title_style = ParagraphStyle('titolo-etichette', fontName='Etichette-Bold', fontSize=14, leading=18, spaceAfter=8)
    document.build([paragraph(title, title_style), paragraph(f'{len(items)} voci', normal), table])


def export_csv(path, items, fields, town, include_town):
    active = [field for field in fields if field['enabled']]
    if include_town:
        active.append({'i': -1, 'name': 'Comune di residenza'})
    if not items or not active:
        raise ValueError('Seleziona almeno una persona e un campo.')
    def safe(value):
        text = cell_text(value)
        return "'" + text if text.startswith(('=', '+', '-', '@', '\t', '\r')) else text
    with open(path, 'w', encoding='utf-8-sig', newline='') as output:
        writer = csv.writer(output, delimiter=';', quoting=csv.QUOTE_ALL)
        writer.writerow([safe(field['name']) for field in active])
        for row in items:
            writer.writerow([safe(town if field['i'] == -1 else (row[field['i']] if field['i'] < len(row) else '')) for field in active])
