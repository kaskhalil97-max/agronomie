import zipfile
import xml.etree.ElementTree as ET

NS = {
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
}

def qn(prefix, tag):
    return '{%s}%s' % (NS[prefix], tag)

def cell_value(cell):
    vtype = cell.get(qn('office', 'value-type'))
    if vtype == 'float' or vtype == 'currency' or vtype == 'percentage':
        v = cell.get(qn('office', 'value'))
        try:
            return float(v)
        except (TypeError, ValueError):
            pass
    if vtype == 'date':
        return cell.get(qn('office', 'date-value'))
    # fallback: text content
    texts = []
    for p in cell.findall(qn('text', 'p')):
        texts.append(''.join(p.itertext()))
    if texts:
        return '\n'.join(texts)
    return None

def read_ods(path, max_cols=200, with_formulas=False):
    z = zipfile.ZipFile(path)
    data = z.read('content.xml')
    root = ET.fromstring(data)
    body = root.find(qn('office', 'body'))
    spreadsheet = body.find(qn('office', 'spreadsheet'))
    sheets = {}
    formula_sheets = {}
    for table in spreadsheet.findall(qn('table', 'table')):
        name = table.get(qn('table', 'name'))
        rows = []
        frows = []
        for row in table.findall(qn('table', 'table-row')):
            row_repeat = int(row.get(qn('table', 'number-rows-repeated'), '1'))
            row_cells = []
            row_formulas = []
            for cell in row.findall(qn('table', 'table-cell')) + row.findall(qn('table', 'covered-table-cell')):
                col_repeat = int(cell.get(qn('table', 'number-columns-repeated'), '1'))
                val = cell_value(cell)
                formula = cell.get(qn('table', 'formula'))
                for _ in range(col_repeat):
                    if len(row_cells) >= max_cols:
                        break
                    row_cells.append(val)
                    row_formulas.append(formula)
            for _ in range(row_repeat):
                rows.append(list(row_cells))
                if with_formulas:
                    frows.append(list(row_formulas))
                if len(rows) > 5000:
                    break
        sheets[name] = rows
        if with_formulas:
            formula_sheets[name] = frows
    if with_formulas:
        return sheets, formula_sheets
    return sheets

if __name__ == '__main__':
    import sys
    path = sys.argv[1]
    sheets = read_ods(path)
    for name, rows in sheets.items():
        nonempty = sum(1 for r in rows if any(c is not None for c in r))
        print(f"Sheet: {name!r}  rows={len(rows)}  nonempty_rows={nonempty}")
