import unicodedata
from dataclasses import dataclass

import openpyxl

MONTH_COLUMNS = {
    1: "C", 2: "D", 3: "E", 4: "F", 5: "G", 6: "H",
    7: "I", 8: "J", 9: "K", 10: "L", 11: "M", 12: "N",
}
FIRST_DATA_ROW = 3
LAST_DATA_ROW = 173
RECETTE_LAST_ROW = 16  # rows 3-16 are receipts, rows after are expenses
UNMATCHED_SHEET = "À vérifier"


def normalize(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return " ".join(text.upper().split())


@dataclass
class TargetRow:
    row: int
    bloc: str
    label: str
    is_recette: bool


def load_workbook(path: str):
    return openpyxl.load_workbook(path)


def get_treasury_sheet(wb, year: int):
    sheet_name = f"Trésorerie {year}"
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Onglet '{sheet_name}' introuvable dans le classeur.")
    return wb[sheet_name]


def parse_target_rows(ws) -> list[TargetRow]:
    """Scans column A (bloc, forward-filled) and column B (fournisseur/label)
    to build the list of rows that can receive an amount. Rows where column B
    is empty are totals/separators and are skipped."""
    targets = []
    current_bloc = None
    for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
        bloc_cell = ws.cell(row=row, column=1).value
        label_cell = ws.cell(row=row, column=2).value
        if bloc_cell:
            current_bloc = bloc_cell
        if label_cell:
            targets.append(
                TargetRow(
                    row=row,
                    bloc=current_bloc or "",
                    label=label_cell,
                    is_recette=row <= RECETTE_LAST_ROW,
                )
            )
    return targets


def build_label_index(targets: list[TargetRow]) -> dict[str, list[TargetRow]]:
    index: dict[str, list[TargetRow]] = {}
    for t in targets:
        index.setdefault(normalize(t.label), []).append(t)
    return index


def write_month_amounts(ws, month: int, amounts_by_row: dict[int, float]) -> None:
    col_letter = MONTH_COLUMNS[month]
    for row, amount in amounts_by_row.items():
        ws[f"{col_letter}{row}"] = round(amount, 2)


def write_unmatched(wb, entries: list[dict]) -> None:
    if UNMATCHED_SHEET in wb.sheetnames:
        ws = wb[UNMATCHED_SHEET]
    else:
        ws = wb.create_sheet(UNMATCHED_SHEET)
        ws.append(["Date", "Libellé", "Montant", "Tiers résolu", "Catégorie Pennylane", "Raison"])
    for entry in entries:
        ws.append(
            [
                entry["date"],
                entry["label"],
                entry["amount"],
                entry.get("tiers", ""),
                entry.get("category", ""),
                entry.get("reason", ""),
            ]
        )
