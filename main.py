import argparse
import json
import os
import shutil
import sys
from datetime import date
from pathlib import Path

from dotenv import load_dotenv

from treasury_sync.pennylane_client import PennylaneClient
from treasury_sync.excel_treasury import (
    load_workbook,
    ensure_year_sheet,
    parse_target_rows,
    build_label_index,
    apply_month_amounts,
    write_unmatched,
)
from treasury_sync.matcher import resolve
from treasury_sync.state import load_state, save_state

DEFAULT_MAPPING_PATH = Path(__file__).resolve().parent / "mapping.json"


def parse_args():
    parser = argparse.ArgumentParser(
        description="Met à jour l'onglet Trésorerie depuis Pennylane, en reprenant "
        "automatiquement où la dernière synchronisation s'est arrêtée."
    )
    parser.add_argument(
        "--since",
        default=None,
        help="Optionnel : force un rattrapage à partir de cette date (YYYY-MM-DD), "
        "ignore la reprise automatique. À utiliser une seule fois pour rattraper "
        "des mois anciens.",
    )
    return parser.parse_args()


def resolve_tiers_name(client: PennylaneClient, tx: dict) -> str | None:
    if tx.get("supplier"):
        return client.get_supplier_name(tx["supplier"]["id"])
    if tx.get("customer"):
        return client.get_customer_name(tx["customer"]["id"])
    return None


def backup_file(path: Path) -> Path:
    backup_dir = path.parent / "sauvegardes_tresorerie"
    backup_dir.mkdir(exist_ok=True)
    stamp = date.today().isoformat()
    backup_path = backup_dir / f"{path.stem}_{stamp}{path.suffix}"
    counter = 1
    while backup_path.exists():
        backup_path = backup_dir / f"{path.stem}_{stamp}_{counter}{path.suffix}"
        counter += 1
    shutil.copy2(path, backup_path)
    return backup_path


def main():
    load_dotenv()
    args = parse_args()

    excel_path_str = os.environ.get("EXCEL_PATH")
    token = os.environ.get("PENNYLANE_API_TOKEN")
    if not excel_path_str:
        sys.exit("EXCEL_PATH manquant dans le fichier .env")
    if not token:
        sys.exit("PENNYLANE_API_TOKEN manquant dans le fichier .env")
    excel_path = Path(excel_path_str)
    if not excel_path.exists():
        sys.exit(f"Fichier Excel introuvable : {excel_path}")

    mapping = json.loads(DEFAULT_MAPPING_PATH.read_text())
    state = load_state()
    last_synced_id = state.get("last_synced_transaction_id")

    client = PennylaneClient(token)

    if args.since:
        since_date = date.fromisoformat(args.since)
        print(f"Rattrapage manuel depuis le {since_date.isoformat()}.")
        tx_iterator = client.iter_transactions_since(since_date)
    elif last_synced_id is None:
        since_date = date.today().replace(day=1)
        print(
            f"Première synchronisation : aucun historique connu, je pars du "
            f"{since_date.isoformat()} (début du mois en cours)."
        )
        print(
            "Pour rattraper des mois plus anciens, relance avec --since AAAA-MM-JJ."
        )
        tx_iterator = client.iter_transactions_since(since_date)
    else:
        tx_iterator = client.iter_new_transactions(last_synced_id)

    backup_path = backup_file(excel_path)
    print(f"Sauvegarde créée : {backup_path}")

    wb = load_workbook(str(excel_path))

    sheet_cache: dict[int, tuple] = {}
    amounts_by_key: dict[tuple[int, int, int], float] = {}  # (year, month, row) -> amount
    unmatched_by_year: dict[int, list[dict]] = {}
    max_id_seen = last_synced_id
    matched_count = 0
    processed_count = 0

    for tx in tx_iterator:
        processed_count += 1
        if max_id_seen is None or tx["id"] > max_id_seen:
            max_id_seen = tx["id"]

        tx_date = date.fromisoformat(tx["date"])
        year = tx_date.year
        if year not in sheet_cache:
            ws = ensure_year_sheet(wb, year)
            targets = parse_target_rows(ws)
            sheet_cache[year] = (ws, build_label_index(targets))
        ws, label_index = sheet_cache[year]

        tiers_name = resolve_tiers_name(client, tx)
        category_labels = [c["label"] for c in tx.get("categories", [])]
        amount = float(tx["amount"])

        target, reason, candidates = resolve(
            tiers_name, tx["label"], amount, category_labels, label_index, mapping
        )

        if target is None:
            unmatched_by_year.setdefault(year, []).append(
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
        key = (year, tx_date.month, target.row)
        amounts_by_key[key] = amounts_by_key.get(key, 0.0) + signed_amount
        matched_count += 1

    by_sheet_month: dict[tuple[int, int], dict[int, float]] = {}
    for (year, month, row), amount in amounts_by_key.items():
        by_sheet_month.setdefault((year, month), {})[row] = amount

    replaced_count = added_count = 0
    for (year, month), amounts in by_sheet_month.items():
        ws, _ = sheet_cache[year]
        replaced, added = apply_month_amounts(ws, month, amounts)
        replaced_count += replaced
        added_count += added

    for entries in unmatched_by_year.values():
        write_unmatched(wb, entries)

    wb.save(str(excel_path))

    if max_id_seen is not None:
        save_state({"last_synced_transaction_id": max_id_seen})

    total_unmatched = sum(len(v) for v in unmatched_by_year.values())
    print(f"\n{processed_count} nouvelle(s) transaction(s) trouvée(s).")
    print(f"{matched_count} affectée(s) automatiquement au tableau.")
    print(f"  dont {replaced_count} estimation(s) remplacée(s) par une vraie valeur")
    print(f"  et {added_count} cellule(s) déjà réelle(s) mise(s) à jour (montant ajouté)")
    print(f"{total_unmatched} à vérifier manuellement (onglet 'À vérifier').")
    print(f"Fichier mis à jour : {excel_path}")


if __name__ == "__main__":
    main()
