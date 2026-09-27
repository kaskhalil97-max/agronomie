"""A tiny formula evaluator for the closed set of OpenFormula patterns
actually used in CONTROLE_CONSIGNE_PRELEVEMENT.ods (cell refs, +, -, *,
/, parentheses, numeric literals, and SUM() of a single-row or
single-column range). It never touches Excel/LibreOffice -- it's here
purely so the .ods file we write already carries *correct* cached
results for every formula, instead of relying on the spreadsheet
application to recalculate them when the file is opened.

This is not a general spreadsheet engine: unsupported formulas (cross-
sheet references, IF, #REF! errors, ...) are simply left untouched --
their cached value stays whatever it already was in the source file.
"""
import re

CELL_REF_RE = re.compile(r'\[\.\$?([A-Z]+)\$?(\d+)\]')
RANGE_REF_RE = re.compile(r'\[\.\$?([A-Z]+)\$?(\d+):\.\$?([A-Z]+)\$?(\d+)\]')


class RecalcError(Exception):
    pass


def col_letters_to_index(letters):
    idx = 0
    for c in letters:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx


def col_index_to_letters(idx):
    letters = ''
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _num(v):
    return v if isinstance(v, (int, float)) else 0.0


class SheetRecalculator:
    """Recomputes formula cells for one sheet, given its current
    (possibly just-edited) raw values and formulas, both as row-major
    lists of lists in the same shape as ods_reader.read_ods() returns."""

    def __init__(self, values, formulas):
        self.values = values
        self.formulas = formulas
        self.cache = {}
        self.visiting = set()
        self.updates = {}  # (row, col) -> newly computed value

    def _raw(self, row, col):
        if row < len(self.values) and col < len(self.values[row]):
            return self.values[row][col]
        return None

    def _formula(self, row, col):
        if row < len(self.formulas) and col < len(self.formulas[row]):
            return self.formulas[row][col]
        return None

    def get(self, row, col):
        key = (row, col)
        if key in self.cache:
            return self.cache[key]
        if key in self.visiting:
            raise RecalcError(f"Circular reference at R{row}C{col}")

        formula = self._formula(row, col)
        if not formula:
            value = _num(self._raw(row, col))
            self.cache[key] = value
            return value

        self.visiting.add(key)
        try:
            value = self._eval_formula(formula)
        finally:
            self.visiting.discard(key)

        self.cache[key] = value
        old = self._raw(row, col)
        if not isinstance(old, (int, float)) or abs(old - value) > 1e-6:
            self.updates[key] = value
        return value

    def _eval_formula(self, formula):
        if not formula.startswith('of:='):
            raise RecalcError(f"Unsupported formula dialect: {formula!r}")
        body = formula[len('of:='):]
        if '#REF' in body or '!' in body:
            # dangling reference / anything with a sheet-qualifier we
            # don't support -- leave untouched.
            raise RecalcError(f"Unsupported reference in: {formula!r}")

        def repl_range(m):
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            return f'__RANGE__({c1!r},{r1},{c2!r},{r2})'

        py_expr = RANGE_REF_RE.sub(repl_range, body)

        def repl_cell(m):
            letters, num = m.group(1), int(m.group(2))
            return f'__CELL__({letters!r},{num})'

        py_expr = CELL_REF_RE.sub(repl_cell, py_expr)

        if re.search(r'[A-Za-z_][A-Za-z0-9_]*\(', py_expr) and not re.match(
            r'^[\s()0-9+\-*/.,]*(__CELL__|__RANGE__|SUM)[\s\S]*$', py_expr
        ):
            raise RecalcError(f"Unsupported function in: {formula!r}")

        py_expr = py_expr.replace('SUM(', '__SUM__(')

        def cell(letters, row_1based):
            return self.get(row_1based - 1, col_letters_to_index(letters) - 1)

        def rng(c1, r1, c2, r2):
            col1, col2 = col_letters_to_index(c1), col_letters_to_index(c2)
            vals = []
            for r in range(min(r1, r2), max(r1, r2) + 1):
                for c in range(min(col1, col2), max(col1, col2) + 1):
                    vals.append(self.get(r - 1, c - 1))
            return vals

        def sum_(*args):
            flat = []
            for a in args:
                if isinstance(a, list):
                    flat.extend(a)
                else:
                    flat.append(a)
            return sum(flat)

        try:
            return float(eval(py_expr, {'__builtins__': {}}, {
                '__CELL__': cell, '__RANGE__': rng, '__SUM__': sum_,
            }))
        except RecalcError:
            raise
        except Exception as e:
            raise RecalcError(f"Could not evaluate {formula!r} ({py_expr!r}): {e}")


def recalculate(values, formulas, max_row=400):
    """Evaluate every formula cell in the first `max_row` rows and
    return {(row, col): new_value} for those whose cached value needs
    updating. Cells whose formula we don't support are silently
    skipped (their existing cached value is left as-is)."""
    calc = SheetRecalculator(values, formulas)
    limit = min(max_row, len(formulas))
    for r in range(limit):
        frow = formulas[r]
        for c, f in enumerate(frow):
            if not f:
                continue
            try:
                calc.get(r, c)
            except RecalcError:
                continue
    return calc.updates
