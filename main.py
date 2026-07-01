import argparse
import json
import sys
from calendar import monthrange
from datetime import date
from pathlib import Path

from dotenv import load_dotenv
import os

from treasury_sync.pennylane_client import PennylaneClient
from treasury_sync.excel_treasury import (
    load_workbook,
    get_treasury_sheet,
    parse_target_rows,
    build_label_index,
    write_month_amounts,
    write_unmatched,
)
from treasury_sync.matcher import resolve

DEFAULT_MAPPING_PATH = Path(__file__).resolve().parent / "mapping.json"


def parse_args():
    parser = argparse.ArgumentParser(description="Met à jour l'onglet Trésorerie depuis Pennylane.")
    parser.add_argument("--month", required=True, help="Mois à synchroniser, format YYYY-MM")
    parser.add_argument("--excel", default=os.environ.get("EXCEL_PATH"), help="Chemin du fichier Excel")
    parser.add_argument("--token", default=os.environ.get("PENNYLANE_API_TOKEN"), help="Token API Pennylane")
    parser.add_argument("--mapping", default=str(DEFAULT_MAPPING_PATH), help="Chemin du fichier mapping.json")
    parser.add_argument("--output", default=None, help="Fichier de sortie (par défaut: copie _updated à côté de l'original)")
    parser.add_argument("--in-place", action="store_true", help="Écrase directement le fichier Excel d'origine")
    return parser.parse_args()


def resolve_tiers_name(client: PennylaneClient, tx: dict) -> str | None:
    if tx.get("supplier"):
        return client.get_supplier_name(tx["supplier"]["id"])
    if tx.get("customer"):
        return client.get_customer_name(tx["customer"]["id"])
    return None


def main():
    load_dotenv()
    args = parse_args()

    if not args.excel:
        sys.exit("Chemin du fichier Excel manquant (--excel ou EXCEL_PATH dans .env)")
    if not args.token:
        sys.exit("Token API Pennylane manquant (--token ou PENNYLANE_API_TOKEN dans .env)")

    year, month = (int(x) for x in args.month.split("-"))
    since = date(year, month, 1)
    until = date(year, month, monthrange(year, month)[1])

    mapping = json.loads(Path(args.mapping).read_text())

    client = PennylaneClient(args.token)
    wb = load_workbook(args.excel)
    ws = get_treasury_sheet(wb, year)
    targets = parse_target_rows(ws)
    label_index = build_label_index(targets)

    amounts_by_row: dict[int, float] = {}
    unmatched_entries = []
    matched_count = 0

    for tx in client.iter_transactions_since(since):
        tx_date = date.fromisoformat(tx["date"])
        if tx_date > until:
            continue

        tiers_name = resolve_tiers_name(client, tx)
        category_labels = [c["label"] for c in tx.get("categories", [])]
        amount = float(tx["amount"])

        target, reason, candidates = resolve(
            tiers_name, tx["label"], amount, category_labels, label_index, mapping
        )

        if target is None:
            unmatched_entries.append(
                {
                    "date": tx["date"],
                    "label": tx["label"],
                    "amount": amount,
                    "tiers": tiers_name or "",
                    "category": ", ".join(category_labels),
                    "reason": reason
                    + (
                        f" (candidats: {', '.join(c.label for c in candidates)})"
                        if candidates
                        else ""
                    ),
                }
            )
            continue

        signed_amount = amount if target.is_recette else abs(amount)
        amounts_by_row[target.row] = amounts_by_row.get(target.row, 0.0) + signed_amount
        matched_count += 1

    write_month_amounts(ws, month, amounts_by_row)
    write_unmatched(wb, unmatched_entries)

    if args.in_place:
        output_path = args.excel
    else:
        excel_path = Path(args.excel)
        output_path = args.output or str(excel_path.with_name(f"{excel_path.stem}_updated{excel_path.suffix}"))

    wb.save(output_path)

    print(f"{matched_count} transaction(s) affectée(s) sur {len(amounts_by_row)} ligne(s).")
    print(f"{len(unmatched_entries)} transaction(s) non affectée(s), listées dans l'onglet 'À vérifier'.")
    print(f"Fichier enregistré : {output_path}")


if __name__ == "__main__":
    main()
