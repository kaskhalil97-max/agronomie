"""Minimal ODS cell-value patcher.

Edits specific cells of a .ods spreadsheet's content.xml in place,
splitting table:number-rows-repeated / table:number-columns-repeated
groups as needed so a single target cell can be isolated and modified,
while leaving every other cell (including all formulas) untouched.
"""
import copy
import shutil
import zipfile

import xml.etree.ElementTree as ET

NS = {
    'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def qn(prefix, tag):
    return '{%s}%s' % (NS[prefix], tag)


ROW_TAG = qn('table', 'table-row')
CELL_TAG = qn('table', 'table-cell')
COVERED_CELL_TAG = qn('table', 'covered-table-cell')
ROWS_REPEATED = qn('table', 'number-rows-repeated')
COLS_REPEATED = qn('table', 'number-columns-repeated')


class OdsDocument:
    def __init__(self, path):
        self.path = path
        self.zip_in = zipfile.ZipFile(path, 'r')
        self.content_bytes = self.zip_in.read('content.xml')
        self.root = ET.fromstring(self.content_bytes)
        body = self.root.find(qn('office', 'body'))
        self.spreadsheet = body.find(qn('office', 'spreadsheet'))

    def _get_table(self, sheet_name):
        for table in self.spreadsheet.findall(qn('table', 'table')):
            if table.get(qn('table', 'name')) == sheet_name:
                return table
        raise KeyError(f"Sheet {sheet_name!r} not found")

    def _isolate_row(self, table, row_idx):
        """Return the single <table:row> element that is row_idx (0-based),
        splitting any repeated-row group that currently covers it."""
        children = list(table)
        cum = 0
        for pos, row in enumerate(children):
            if row.tag != ROW_TAG:
                continue
            rep = int(row.get(ROWS_REPEATED, '1'))
            if cum <= row_idx < cum + rep:
                if rep > 1:
                    before = row_idx - cum
                    after = rep - before - 1
                    new_rows = []
                    if before > 0:
                        r1 = copy.deepcopy(row)
                        r1.set(ROWS_REPEATED, str(before))
                        new_rows.append(r1)
                    target = copy.deepcopy(row)
                    if ROWS_REPEATED in target.attrib:
                        del target.attrib[ROWS_REPEATED]
                    new_rows.append(target)
                    if after > 0:
                        r2 = copy.deepcopy(row)
                        r2.set(ROWS_REPEATED, str(after))
                        new_rows.append(r2)
                    # replace element at pos with new_rows
                    table.remove(row)
                    for offset, nr in enumerate(new_rows):
                        table.insert(pos + offset, nr)
                    target_index = pos + (1 if before > 0 else 0)
                    return list(table)[target_index]
                else:
                    return row
            cum += rep
        raise IndexError(f"Row {row_idx} not found in sheet (max scanned {cum})")

    def _isolate_cell(self, row, col_idx):
        """Return the single <table:table-cell> element at col_idx (0-based)
        within a row, splitting any repeated-column group covering it."""
        children = list(row)
        cum = 0
        for pos, cell in enumerate(children):
            if cell.tag not in (CELL_TAG, COVERED_CELL_TAG):
                continue
            rep = int(cell.get(COLS_REPEATED, '1'))
            if cum <= col_idx < cum + rep:
                if rep > 1:
                    before = col_idx - cum
                    after = rep - before - 1
                    new_cells = []
                    if before > 0:
                        c1 = copy.deepcopy(cell)
                        c1.set(COLS_REPEATED, str(before))
                        new_cells.append(c1)
                    target = copy.deepcopy(cell)
                    if COLS_REPEATED in target.attrib:
                        del target.attrib[COLS_REPEATED]
                    new_cells.append(target)
                    if after > 0:
                        c2 = copy.deepcopy(cell)
                        c2.set(COLS_REPEATED, str(after))
                        new_cells.append(c2)
                    row.remove(cell)
                    for offset, nc in enumerate(new_cells):
                        row.insert(pos + offset, nc)
                    target_index = pos + (1 if before > 0 else 0)
                    return list(row)[target_index]
                else:
                    return cell
            cum += rep
        raise IndexError(f"Column {col_idx} not found in row (max scanned {cum})")

    def set_cell_value(self, sheet_name, row_idx, col_idx, value, allow_overwrite_formula=False):
        table = self._get_table(sheet_name)
        row = self._isolate_row(table, row_idx)
        cell = self._isolate_cell(row, col_idx)

        formula = cell.get(qn('table', 'formula'))
        if formula and not allow_overwrite_formula:
            raise ValueError(
                f"Refusing to overwrite formula cell at {sheet_name}!R{row_idx}C{col_idx}: {formula}"
            )

        for p in cell.findall(qn('text', 'p')):
            cell.remove(p)

        cell.set(qn('office', 'value-type'), 'float')
        cell.set(qn('office', 'value'), repr(float(value)))
        p = ET.SubElement(cell, qn('text', 'p'))
        # French-style display isn't critical; LibreOffice/Excel will
        # reformat on open. Keep a plain numeric string.
        if float(value).is_integer():
            p.text = str(int(value))
        else:
            p.text = str(value)

    def save(self, out_path):
        new_content = ET.tostring(self.root, encoding='UTF-8', xml_declaration=True)
        shutil.copyfile(self.path, out_path)
        # Rewrite the zip, replacing only content.xml
        self.zip_in.close()
        src = zipfile.ZipFile(self.path, 'r')
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == 'content.xml':
                    data = new_content
                if item.filename == 'mimetype':
                    zout.writestr(item, data, compress_type=zipfile.ZIP_STORED)
                else:
                    zout.writestr(item, data)
        src.close()
