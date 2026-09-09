"""Læg to Excel-filer ved START.cmd, og kør importen."""
from pathlib import Path
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
import json
import os
import re
import shutil
import tempfile
from zipfile import BadZipFile

from excel import (ImportErrorDetail, account_number, build_account_mapping,
                   openpyxl, parse_amount, render_workbook, select_column)


def read_export(content):
    book = openpyxl.load_workbook(BytesIO(content), data_only=False)
    try:
        tables = [(sheet, row[0].row) for sheet in book
                  for row in sheet.iter_rows(max_row=min(sheet.max_row, 30))
                  if [str(c.value or "").strip() for c in row[:4]] ==
                  ["Nr.", "Navn", "Perioden", "År til dato"]]
        if len(tables) != 1:
            raise ImportErrorDetail("Saldobalancen skal indeholde én tabel med Nr., Navn, Perioden og År til dato.")
        sheet, header = tables[0]
        periods = []
        pattern = r"Saldobalance for perioden\s+(\d{2}\.\d{2}\.(?:\d{4}|\d{2}))\s*-\s*(\d{2}\.\d{2}\.(?:\d{4}|\d{2}))"
        for row in sheet.iter_rows(max_row=header - 1):
            for cell in row:
                match = re.fullmatch(pattern, str(cell.value or "").strip())
                if match:
                    dates = []
                    for text in match.groups():
                        day, month, year = (int(v) for v in text.split("."))
                        if year < 100:
                            year += date.today().year // 100 * 100
                        dates.append(date(year, month, day))
                    periods.append(dates)
        if len(periods) != 1:
            raise ImportErrorDetail("Perioden kunne ikke aflæses entydigt i saldobalancens overskrift.")
        start, end = periods[0]
        if start > end or start.year != end.year:
            raise ImportErrorDetail("Saldobalancens periode skal ligge inden for ét kalenderår.")
        amounts = {}
        for row in sheet.iter_rows(min_row=header + 1):
            if all(c.value in (None, "") for c in row):
                continue
            number_cell, name_cell, amount_cell = row[:3]
            try:
                number = account_number(number_cell.value)
                if amount_cell.data_type == "f":
                    raise ValueError("Beløbet er en formel")
                amount = parse_amount(amount_cell.value)
            except ValueError:
                raise ImportErrorDetail(f"Ugyldigt kontonummer eller beløb i saldobalancens række {number_cell.row}.") from None
            if number in amounts:
                raise ImportErrorDetail(f"Konto {number} optræder flere gange i saldobalancen.")
            amounts[number] = amount
        if not amounts:
            raise ImportErrorDetail("Saldobalancen indeholder ingen kontobeløb.")
        return start, end, amounts
    finally:
        book.close()


def build_plan(sheet, year, amounts, previous=None):
    column = select_column(sheet, year)
    mapping, issues = build_account_mapping(sheet, column)
    if any(issues[key] for key in ("duplicate_excel_accounts", "malformed_excel_accounts", "ambiguous_excel_rows")):
        raise ImportErrorDetail("Regnskabet har ugyldige eller tvetydige kontorækker. Ret dem før import.")
    matched = set(mapping) & set(amounts)
    if not matched:
        raise ImportErrorDetail("Ingen konti fra saldobalancen matcher første ark. Ingen fil ændret.")
    updates, balances, zeros = {}, {}, []
    for number, target in mapping.items():
        value = -amounts.get(number, Decimal(0))
        if value == 0:
            value = Decimal(0)
        old = sheet[target.cell].value
        if old is None or Decimal(str(old)) != value:
            updates[target.cell] = value
        balances[str(number)] = str(value)
        before = previous.get(str(number)) if previous is not None else old
        if value == 0 and before is not None and Decimal(str(before)) != 0:
            zeros.append({"account": number, "name": target.name, "previous": str(before), "cell": target.cell})
    known_subtotals = {"D117": (114, 115, 116), "D238": (237,), "D246": (245,), "H6": (5,)}
    detail_cells = {target.cell for target in mapping.values()}
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, column)
        if (sheet.cell(row, 2).value is not None or not isinstance(cell.value, (int, float))
                or not any(sheet.cell(row, other).data_type == "f" for other in (4, 8, 9) if other != column)):
            continue
        constituents = known_subtotals.get(cell.coordinate)
        if not constituents:
            raise ImportErrorDetail(f"Ukendt fast delsum i {cell.coordinate}; ingen fil ændret.")
        verified = False
        for other in (4, 8, 9):
            if other == column:
                continue
            letter = openpyxl.utils.get_column_letter(other)
            expected = {f"=SUM({letter}{constituents[0]}:{letter}{constituents[-1]})"}
            if len(constituents) == 1:
                expected.add(f"=SUM({letter}{constituents[0]})")
            verified = verified or sheet.cell(row, other).value in expected
        cells = [sheet.cell(r, column).coordinate for r in constituents]
        if not verified or not set(cells) <= detail_cells:
            raise ImportErrorDetail(f"Delsummen i {cell.coordinate} kunne ikke kontrolleres.")
        value = sum((Decimal(str(updates.get(c, sheet[c].value))) for c in cells), Decimal(0))
        if value != Decimal(str(cell.value)):
            updates[cell.coordinate] = value
    letter = openpyxl.utils.get_column_letter(column)
    result = sheet[f"{letter}254"]
    expected = {f"={letter}58-{letter}249", f"=-SUM({letter}249-{letter}58)", f"={letter}58+{letter}249"}
    if result.value not in expected:
        raise ImportErrorDetail(f"Resultatformlen i {result.coordinate} kunne ikke kontrolleres.")
    new_formula = f"={letter}58+{letter}249"
    formulas = {} if result.value == new_formula else {result.coordinate: {"old": result.value, "new": new_formula}}
    updates[f"{letter}1"] = f"Regnskab{year} pr. {date.today():%d.%m}"
    all_accounts = set(mapping)
    for coord in issues["protected_formula_cells"]:
        all_accounts.add(account_number(sheet.cell(sheet[coord].row, 2).value))
    report = {"year": year, "matched": len(matched), "missing": len(set(mapping) - set(amounts)),
              "unmatched": sorted(set(amounts) - all_accounts), "zeroed": zeros,
              "comparison": "sidste træk" if previous is not None else "inputregnskabet (første kørsel)"}
    return updates, formulas, balances, report


def atomic_write(path, content):
    with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".import_", suffix=".tmp", delete=False) as handle:
        temporary = Path(handle.name)
        try:
            handle.write(content)
        except BaseException:
            handle.close()
            temporary.unlink(missing_ok=True)
            raise
    try:
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def load_history(path, year):
    try:
        history = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
        if not isinstance(history, dict):
            raise ValueError()
        previous = history.get(str(year))
        if previous is not None:
            if not isinstance(previous, dict):
                raise ValueError()
            for number, value in previous.items():
                account_number(number)
                parse_amount(Decimal(value))
        return history, previous
    except (ValueError, TypeError, InvalidOperation):
        raise ImportErrorDetail("Historikken i program/historik.json kunne ikke læses. Ingen fil ændret.") from None


def run(folder):
    folder = Path(folder).resolve()
    source, balance = folder / "regnskab.xlsx", folder / "saldobalance.xlsx"
    for path in (source, balance):
        if not path.is_file():
            raise ImportErrorDetail(f"Filen {path.name} mangler. Læg den i samme mappe som START.cmd.")
    original, export = source.read_bytes(), balance.read_bytes()
    start, end, amounts = read_export(export)
    program = folder / "program"
    program.mkdir(exist_ok=True)
    history_path = program / "historik.json"
    history, previous = load_history(history_path, start.year)
    book = openpyxl.load_workbook(BytesIO(original), data_only=False)
    try:
        sheet = book.worksheets[0]
        updates, formulas, balances, report = build_plan(sheet, start.year, amounts, previous)
        completed = render_workbook(original, sheet.title, updates, formulas)
    finally:
        book.close()
    if source.read_bytes() != original or balance.read_bytes() != export:
        raise ImportErrorDetail("En inputfil blev ændret under importen. Prøv igen.")
    output = folder / "opdateret_regnskab.xlsx"
    if output.is_symlink() or (output.exists() and any(os.path.samefile(output, p) for p in (source, balance))):
        raise ImportErrorDetail("Outputfilen peger på en anden fil; originalerne må ikke overskrives.")
    if output.exists():
        backups = program / "sikkerhedskopier"
        backups.mkdir(exist_ok=True)
        backup = backups / f"opdateret_regnskab_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
        with output.open("rb") as src, backup.open("xb") as dst:
            shutil.copyfileobj(src, dst)
    atomic_write(output, completed)
    lines = [f"Færdig: {output.name}", f"Periode: {start:%d.%m.%Y} - {end:%d.%m.%Y}",
             f"{report['matched']} konti importeret. {report['missing']} manglende konti sat til nul."]
    if report["unmatched"]:
        lines.append("Konti i udtrækket uden række i regnskabet: " + ", ".join(map(str, report["unmatched"])))
    lines.append(f"Notits: {len(report['zeroed'])} konti er blevet nul siden {report['comparison']}.")
    for entry in report["zeroed"]:
        value = f"{Decimal(entry['previous']):,.2f}".translate(str.maketrans(",.", ".,"))
        lines.append(f"  {entry['account']} - {entry['name']}: {value} -> 0 kr.")
    lines.append("Åbn opdateret_regnskab.xlsx i Excel for at genberegne formlerne.")
    report["period"] = [start.isoformat(), end.isoformat()]
    try:
        atomic_write(program / "seneste_rapport.txt", ("\n".join(lines) + "\n").encode("utf-8"))
        atomic_write(program / "seneste_rapport.json", json.dumps(report, ensure_ascii=False, indent=2).encode("utf-8"))
        history[str(start.year)] = balances
        atomic_write(history_path, json.dumps(history, ensure_ascii=False, indent=2).encode("utf-8"))
    except OSError:
        lines.append("Bemærk: Regnskabet er gemt, men rapport/historik kunne ikke gemmes. Nulnotitsen kan gentages næste gang.")
    print("\n".join(lines))
    return output, report


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    try:
        run(Path(__file__).resolve().parent.parent)
        return 0
    except PermissionError:
        print("Kunne ikke skrive filen. Luk opdateret_regnskab.xlsx i Excel og prøv igen.")
    except (ImportErrorDetail, OSError, ValueError, BadZipFile, InvalidOperation) as error:
        print(f"Fejl: {error}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
