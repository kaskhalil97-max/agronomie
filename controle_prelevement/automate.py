"""Automation prototype: fills one month's raw readings into
CONTROLE_CONSIGNE_PRELEVEMENT.ods from the two monthly production
workbooks (Nord / Sud).

Design notes (reverse-engineered from avant/apres + the workbooks'
own formulas, not guesswork):

- Each commune sheet has a wide "historical" table (one row per month
  since 2022) plus one or two small "current year" mini-tables used as
  scratch input cells.
- For most communes (BEZIERS, LIGNAN, CORNEILHAN, ESPONDEILHAN) the
  historical-table cell for the target month IS the raw manual entry
  (no formula) -> we write there directly.
- For most other communes (CERS, MONTBLANC, SAUVIAN, SERIGNAN,
  SERVIAN x2, VALRAS, VALROS, VILLENEUVE) the historical-table cell is
  itself a formula like `of:=[.B65]`, pointing at the real input cell
  in a mini-table elsewhere on the same sheet. We detect this and
  redirect the write to the referenced cell automatically, instead of
  hardcoding each mini-table's position.
- The source side (which raw meters feed which control-sheet field) is
  taken from the *formulas* inside `PROD ... MENSUEL` sheets
  (e.g. `=IF(Data!W7<>"",Data!W7,"-")`), cross-checked against the real
  avant->apres diff for March 2026.

Run as a script to validate: it reproduces "apres.ods" from "avant.ods"
+ "nord.xlsx" + "sud.xlsx" for March 2026, then a companion diff
against the real "apres.ods" confirms every manually-entered cell
matches.
"""
import datetime
import re
from collections import Counter

import openpyxl

import recalc
from ods_reader import read_ods
from ods_writer import OdsDocument

SIMPLE_REF_RE = re.compile(r'^of:=\[\.([A-Z]+)(\d+)\]$')
SUM_TERM_RE = re.compile(r'\[\.([A-Z]+)(\d+)\]')


def col_letters_to_index(letters):
    """1-based column number from spreadsheet letters, e.g. 'B' -> 2."""
    idx = 0
    for c in letters:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx


# ---------------------------------------------------------------------
# Per-commune config.
#   offsets: list of (offset_from_date_col, source_columns) to sum and
#            write. offset_from_date_col=1 means the column right after
#            the date column.
#   data_sheet: which tab of the workbook holds the raw named values
#               ('Data' or 'Data_C1_R0').
# The actual destination cell is resolved at run time: direct if the
# historical cell has no formula, otherwise redirected through a
# simple single-cell same-sheet formula reference.
# ---------------------------------------------------------------------
CONFIG = {
    'LIGNAN': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (1, ['AI']),   # V cana secours
        (2, ['AG']),   # V Tabarka
    ]},
    'CORNEILHAN': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (1, ['W']),    # V Import Béziers
        (2, ['U']),    # V Import CCAM
    ]},
    'ESPONDEILHAN': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (1, ['K']),    # V Import Béziers
        (2, ['J']),    # V achat SIEVH
    ]},
    'MONTBLANC': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (1, ['N']),          # V Import Béziers
        (2, ['Q', 'P', 'O']),  # V forages locaux (Carals, Caramudes, Vacabelle)
    ]},
    'SERVIAN - VILLAGE': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (1, ['E']),       # V Import Béziers
        (2, ['C', 'D']),  # V Forages locaux
    ]},
    'LA BAUME - SERVIAN': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (2, ['F', 'G']),  # V Forages locaux (offset1 = Import Béziers, always 0, not sourced)
    ]},
    'VALROS': {'workbook': 'nord', 'data_sheet': 'Data', 'offsets': [
        (1, ['AY']),   # V F91
        (2, ['AZ']),   # V F2017
    ]},
    'SERIGNAN': {'workbook': 'sud', 'data_sheet': 'Data', 'offsets': [
        (1, ['W', 'AB']),  # V Import Béziers
        (2, ['C', 'M']),   # V Forages locaux
    ]},
    'SAUVIAN': {'workbook': 'sud', 'data_sheet': 'Data', 'offsets': [
        (1, ['AO', 'AY']),  # V Import Béziers
    ]},
    'VALRAS': {'workbook': 'sud', 'data_sheet': 'Data_C1_R0', 'offsets': [
        (1, ['AO']),              # forage principal (Compteur F2)
        (2, ['K', 'U', 'AE']),    # forages 2bis + 3 + 4
    ]},
    'VILLENEUVE': {'workbook': 'sud', 'data_sheet': 'Data_C1_R0', 'offsets': [
        (1, ['FF', 'CC']),  # V Import Béziers
        (2, ['BD', 'BN']),  # V Forages locaux
    ]},
    'CERS': {'workbook': 'sud', 'data_sheet': 'Data_C1_R0', 'offsets': [
        (1, ['FG', 'FH', 'FK']),  # V Import Béziers (réservoir neuf + vieux + bypass)
        (2, ['DA', 'DK']),        # V Forages locaux (forage Moulin + forage Port du Soleil)
    ]},
}

# BEZIERS is positional and always direct (no formula redirection needed).
BEZIERS_OFFSETS = {
    1: ['EI'], 2: ['ES'], 3: ['FC'], 4: ['FM'], 5: ['FW'], 6: ['GG'], 7: ['GQ'], 8: ['HA'],  # P1-P8 (Carlet)
    10: ['BL'], 11: ['BV'], 12: ['CF'],                                                      # P9-P11 (Rayssac)
    14: ['CU'], 15: ['DE'],                                                                  # G1-G2 (Tabarka)
}
BEZIERS_WORKBOOK = 'sud'
BEZIERS_DATA_SHEET = 'Data'


def find_date_col(rows, header_row_idx):
    counts = Counter()
    for r in rows[header_row_idx:header_row_idx + 150]:
        for j, v in enumerate(r):
            if isinstance(v, str) and len(v) == 10 and v[4] == '-' and v[7] == '-':
                counts[j] += 1
    return counts.most_common(1)[0][0]


def find_header_row(rows):
    for i, r in enumerate(rows):
        if any(isinstance(v, str) and v.strip().startswith('V ') and '(m³)' in v for v in r):
            return i
    raise ValueError("No 'V ...' header row found")


def find_target_row(rows, date_col, target_date_str):
    for i, r in enumerate(rows):
        if date_col < len(r) and r[date_col] == target_date_str:
            return i
    raise ValueError(f"No row found for date {target_date_str}")


def resolve_write_cells(row_idx, col_idx, frows, max_depth=4):
    """Chase same-sheet formula indirection to find the real input
    cell(s) to write to.

    Returns a list of (row, col) targets:
      - a single-element list for a direct cell or a chain of simple
        `=[.X1]` references,
      - a multi-element list if the chain ends on a pure sum of simple
        references (`=[.X1]+[.X2]+...`), one target per term.

    Raises if the formula is anything else (a real computed/derived
    value, e.g. involving multiplication) -- we must never overwrite
    those.
    """
    r, c = row_idx, col_idx
    for _ in range(max_depth):
        formula = frows[r][c] if c < len(frows[r]) else None
        if formula is None:
            return [(r, c)]
        m = SIMPLE_REF_RE.match(formula)
        if m:
            letters, num = m.groups()
            r, c = int(num) - 1, col_letters_to_index(letters) - 1
            continue
        terms = SUM_TERM_RE.findall(formula)
        rebuilt = '+'.join(f'[.{letters}{num}]' for letters, num in terms)
        if terms and f'of:={rebuilt}' == formula:
            return [(int(num) - 1, col_letters_to_index(letters) - 1) for letters, num in terms]
        raise ValueError(
            f"Cell R{r}C{c} has a non-trivial formula ({formula!r}); "
            f"refusing to guess a write target."
        )
    raise ValueError(f"Formula indirection too deep starting at R{row_idx}C{col_idx}")


def get_data_row_values(wb, sheet_name, target_date, col_letters_list):
    """Individual (not summed) values for each source column, so callers
    can distribute them 1:1 across multiple destination cells when
    needed. Missing/non-numeric cells come back as None."""
    ws = wb[sheet_name]
    base_year = ws.cell(row=1, column=1).value.year
    row = 5 + (target_date.year - base_year) * 12 + (target_date.month - 1)
    values = []
    for letters in col_letters_list:
        v = ws.cell(row=row, column=col_letters_to_index(letters)).value
        values.append(v if isinstance(v, (int, float)) else None)
    return values


def _fill_one_month(target_date, sheets, formulas, workbooks, record_write, report):
    """Write every commune's raw readings for target_date, taking the
    Nord/Sud workbooks as the sole source of truth: whatever value is
    already sitting in the control file for that month is overwritten
    if the source workbooks disagree (e.g. a meter reading that came in
    late and only appears in an updated Nord/Sud export after the
    month was first filled)."""
    target_date_str = target_date.isoformat()

    def write_field(commune, offset, source_cols, workbook_key, data_sheet):
        rows = sheets[commune]
        frows = formulas[commune]
        header_row_idx = find_header_row(rows)
        date_col = find_date_col(rows, header_row_idx)
        try:
            row_idx = find_target_row(rows, date_col, target_date_str)
        except ValueError as e:
            report.append(f"  [!] {commune} {target_date_str} offset{offset}: {e}")
            return
        col_idx = date_col + offset

        raw_values = get_data_row_values(workbooks[workbook_key], data_sheet, target_date, source_cols)
        if all(v is None for v in raw_values):
            report.append(f"  [-] {commune} {target_date_str} offset{offset}: no source data, skipped")
            return
        values = [v if v is not None else 0.0 for v in raw_values]

        try:
            targets = resolve_write_cells(row_idx, col_idx, frows)
        except ValueError as e:
            report.append(f"  [!] {commune} {target_date_str} offset{offset}: {e}")
            return

        if len(targets) == 1:
            wr, wc = targets[0]
            record_write(commune, wr, wc, sum(values))
            tag = "direct" if (wr, wc) == (row_idx, col_idx) else f"redirected->R{wr}C{wc}"
            report.append(f"  [OK] {commune} {target_date_str} offset{offset} ({tag}) = {sum(values)}")
        else:
            # Destination is a sum of several manual cells. Distribute
            # source values 1:1 when counts match; otherwise dump the
            # whole total into the first term and zero the rest, since
            # addition is commutative and the *displayed total* is what
            # matters (only the term-by-term split would be ambiguous).
            if len(values) != len(targets):
                distributed = [sum(values)] + [0.0] * (len(targets) - 1)
            else:
                distributed = values
            for (wr, wc), v in zip(targets, distributed):
                record_write(commune, wr, wc, v)
            report.append(
                f"  [OK] {commune} {target_date_str} offset{offset} (redirected->{len(targets)} cells)"
                f" = {distributed} (total {sum(values)})"
            )

    for commune, cfg in CONFIG.items():
        for offset, source_cols in cfg['offsets']:
            write_field(commune, offset, source_cols, cfg['workbook'], cfg['data_sheet'])

    for offset, source_cols in BEZIERS_OFFSETS.items():
        write_field('BEZIERS', offset, source_cols, BEZIERS_WORKBOOK, BEZIERS_DATA_SHEET)


def fill_month(avant_path, nord_path, sud_path, out_path, year, month, resync_previous=True):
    target_date = datetime.date(year, month, 1)

    sheets, formulas = read_ods(avant_path, with_formulas=True)
    doc = OdsDocument(avant_path)

    workbooks = {
        'nord': openpyxl.load_workbook(nord_path, data_only=True),
        'sud': openpyxl.load_workbook(sud_path, data_only=True),
    }

    report = []
    written = {}  # commune -> {(row, col): value} of every raw cell we touched

    def record_write(commune, row, col, value):
        doc.set_cell_value(commune, row, col, value)
        written.setdefault(commune, {})[(row, col)] = value

    if resync_previous:
        prev_year, prev_month = (year, month - 1) if month > 1 else (year - 1, 12)
        prev_date = datetime.date(prev_year, prev_month, 1)
        report.append(f"-- Resynchronisation de {prev_date.isoformat()} avec les fichiers Nord/Sud a jour --")
        _fill_one_month(prev_date, sheets, formulas, workbooks, record_write, report)

    report.append(f"-- Remplissage de {target_date.isoformat()} --")
    _fill_one_month(target_date, sheets, formulas, workbooks, record_write, report)

    # Bake correct cached results into every formula cell that depends
    # (directly or transitively) on something we just wrote, so the
    # file already displays right even if the spreadsheet app opening
    # it never recalculates on its own.
    for commune, overlay in written.items():
        values = [list(row) for row in sheets[commune]]
        for (r, c), v in overlay.items():
            while len(values[r]) <= c:
                values[r].append(None)
            values[r][c] = v
        updates = recalc.recalculate(values, formulas[commune])
        for (r, c), v in updates.items():
            doc.set_formula_cache(commune, r, c, v)
        if updates:
            report.append(f"  [recalc] {commune}: {len(updates)} formule(s) mise(s) a jour")

    doc.save(out_path)
    return report


def detect_next_month(avant_path):
    """Find the next month to fill: the first date row, after the last
    one that already has BEZIERS raw data (our most reliably "direct"
    commune), whose date is still empty."""
    sheets = read_ods(avant_path)
    rows = sheets['BEZIERS']
    header_row_idx = find_header_row(rows)
    date_col = find_date_col(rows, header_row_idx)
    p1_col = date_col + 1  # BEZIERS offset1 (P1, Carlet) is always a direct raw cell

    last_filled_date = None
    for r in rows[header_row_idx:]:
        date_val = r[date_col] if date_col < len(r) else None
        if not (isinstance(date_val, str) and len(date_val) == 10):
            continue
        p1_val = r[p1_col] if p1_col < len(r) else None
        if isinstance(p1_val, (int, float)) and p1_val:
            last_filled_date = date_val
        else:
            if last_filled_date is not None:
                break

    if last_filled_date is None:
        raise ValueError("Impossible de déterminer le dernier mois rempli dans ce fichier.")

    y, m, _ = (int(x) for x in last_filled_date.split('-'))
    m += 1
    if m > 12:
        m = 1
        y += 1
    return y, m


if __name__ == '__main__':
    import argparse
    import os
    import sys

    parser = argparse.ArgumentParser(
        description="Remplit automatiquement le mois manquant dans CONTROLE_CONSIGNE_PRELEVEMENT.ods "
                    "à partir des deux fichiers de production mensuelle Nord et Sud.")
    parser.add_argument('controle_ods', help="Fichier CONTROLE_CONSIGNE_PRELEVEMENT.ods du mois precedent")
    parser.add_argument('nord_xlsx', help="Fichier Volume Production Mensuelle Nord (.xlsx)")
    parser.add_argument('sud_xlsx', help="Fichier Volume Production Mensuelle Sud (.xlsx)")
    parser.add_argument('-o', '--sortie', help="Fichier de sortie (par defaut: <controle>_rempli.ods)")
    parser.add_argument('--annee', type=int, help="Annee du mois a remplir (par defaut: detection automatique)")
    parser.add_argument('--mois', type=int, help="Mois a remplir, 1-12 (par defaut: detection automatique)")
    parser.add_argument('--no-resync-precedent', action='store_true',
                         help="Ne pas re-verifier le mois precedent contre les fichiers Nord/Sud a jour "
                              "(par defaut, il est resynchronise si une valeur y a change)")
    args = parser.parse_args()

    if args.annee and args.mois:
        year, month = args.annee, args.mois
    else:
        year, month = detect_next_month(args.controle_ods)

    if args.sortie:
        out_path = args.sortie
    else:
        base, ext = os.path.splitext(args.controle_ods)
        out_path = f"{base}_rempli{ext}"

    print(f"Remplissage du mois {month:02d}/{year}...")
    try:
        report = fill_month(args.controle_ods, args.nord_xlsx, args.sud_xlsx, out_path, year, month,
                             resync_previous=not args.no_resync_precedent)
    except Exception as e:
        print(f"ERREUR: {e}", file=sys.stderr)
        sys.exit(1)

    print('\n'.join(report))
    print()
    print(f"Termine. Fichier cree : {out_path}")
    print("Ouvrez-le dans LibreOffice Calc (ou Excel) : les totaux et ratios se recalculent tout seuls.")
