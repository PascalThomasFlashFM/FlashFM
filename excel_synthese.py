"""
Application de synthèse de fichiers Excel vers CSV.
Traite tous les fichiers Excel d'un répertoire et les consolide dans synthèse.csv.

Usage :
    python excel_synthese.py                  # lance l'interface graphique
    python excel_synthese.py /chemin/dossier  # mode ligne de commande
"""

import csv
import os
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl


OUTPUT_FILENAME = "synthèse.csv"


# ── Logique métier ────────────────────────────────────────────────────────────

def normalize(value):
    """Normalise une valeur pour la comparaison (strip + lowercase)."""
    return str(value).strip().lower() if value is not None else ""


def find_column_index(headers, *candidates):
    """Trouve l'index d'une colonne parmi plusieurs noms candidats (insensible à la casse)."""
    headers_lower = [h.strip().lower() if h else "" for h in headers]
    for candidate in candidates:
        if candidate.strip().lower() in headers_lower:
            return headers_lower.index(candidate.strip().lower())
    return None


def load_existing_csv(csv_path):
    """Charge le CSV existant. Retourne (fieldnames, existing_rows)."""
    if not Path(csv_path).exists():
        return [], []

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        fieldnames = list(reader.fieldnames or [])
        rows = list(reader)

    return fieldnames, rows


def check_row_status(existing_rows, new_row, nom_key, prenom_key):
    """
    Compare une nouvelle ligne aux lignes existantes.
    Retourne :
      'identical' – même nom/prénom ET mêmes données → ignorer
      'conflict'  – même nom/prénom mais données différentes → ignorer
      'new'       – couple nom/prénom inconnu → ajouter
    """
    nom_new = normalize(new_row.get(nom_key, ""))
    prenom_new = normalize(new_row.get(prenom_key, ""))

    for existing in existing_rows:
        nom_ex = normalize(existing.get(nom_key, "") or existing.get("nom", ""))
        prenom_ex = normalize(
            existing.get(prenom_key, "")
            or existing.get("prénom", "")
            or existing.get("prenom", "")
        )

        if nom_ex == nom_new and prenom_ex == prenom_new:
            all_same = all(
                normalize(existing.get(k, "")) == normalize(new_row.get(k, ""))
                for k in new_row
            )
            return "identical" if all_same else "conflict"

    return "new"


NEWSLETTER_COLUMN = "Souhaitez-vous recevoir notre newsletter ?"


def wants_newsletter(row_dict, headers):
    """Retourne True si la colonne newsletter vaut 'oui' (insensible à la casse et aux balises HTML)."""
    import re
    for h in headers:
        # Comparaison en supprimant les éventuelles balises HTML autour du texte de la colonne
        clean_h = re.sub(r"<[^>]+>", "", h).strip()
        if clean_h.lower() == NEWSLETTER_COLUMN.lower():
            return normalize(row_dict.get(h, "")) == "oui"
    return False


def process_excel_file(excel_path, existing_rows, fieldnames):
    """
    Lit toutes les feuilles d'un fichier Excel et retourne les nouvelles lignes à ajouter.
    Seules les lignes avec "Oui" dans la colonne newsletter sont traitées.
    Modifie existing_rows et fieldnames sur place.
    Retourne (new_rows, skipped_identical, skipped_conflict, skipped_no_newsletter).
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    new_rows = []
    skipped_identical = 0
    skipped_conflict = 0
    skipped_no_newsletter = 0

    for sheet in wb.worksheets:
        rows_iter = sheet.iter_rows(values_only=True)
        try:
            raw_headers = next(rows_iter)
        except StopIteration:
            continue

        headers = [str(h).strip() if h is not None else "" for h in raw_headers]

        # Compléter les fieldnames si nécessaire
        for h in headers:
            if h and h not in fieldnames:
                fieldnames.append(h)

        nom_idx = find_column_index(headers, "nom", "name", "last name", "lastname")
        prenom_idx = find_column_index(
            headers, "prénom", "prenom", "first name", "firstname", "first_name"
        )

        if nom_idx is None or prenom_idx is None:
            continue

        nom_key = headers[nom_idx]
        prenom_key = headers[prenom_idx]

        for row_values in rows_iter:
            if all(v is None or str(v).strip() == "" for v in row_values):
                continue

            row_dict = {
                h: (str(row_values[i]).strip() if i < len(row_values) and row_values[i] is not None else "")
                for i, h in enumerate(headers)
            }

            # Ignorer les personnes qui n'ont pas accepté la newsletter
            if not wants_newsletter(row_dict, headers):
                skipped_no_newsletter += 1
                continue

            status = check_row_status(existing_rows, row_dict, nom_key, prenom_key)
            if status == "identical":
                skipped_identical += 1
            elif status == "conflict":
                skipped_conflict += 1
            else:
                new_rows.append(row_dict)
                existing_rows.append(row_dict)

    wb.close()
    return new_rows, skipped_identical, skipped_conflict, skipped_no_newsletter


def run_synthesis(directory, log_callback=print):
    """
    Traite tous les fichiers Excel du répertoire et consolide dans synthèse.csv.
    log_callback reçoit les messages de progression.
    """
    dir_path = Path(directory)
    csv_path = dir_path / OUTPUT_FILENAME

    fieldnames, existing_rows = load_existing_csv(csv_path)

    excel_files = sorted(
        p for p in dir_path.iterdir()
        if p.suffix.lower() in (".xlsx", ".xls", ".xlsm")
        and p.name != OUTPUT_FILENAME
    )

    if not excel_files:
        log_callback("Aucun fichier Excel trouvé dans ce répertoire.")
        return

    all_new_rows = []
    total_identical = 0
    total_conflict = 0
    total_no_newsletter = 0

    for excel_path in excel_files:
        log_callback(f"Traitement : {excel_path.name}")
        try:
            new_rows, skipped_id, skipped_conf, skipped_nl = process_excel_file(
                excel_path, existing_rows, fieldnames
            )
            all_new_rows.extend(new_rows)
            total_identical += skipped_id
            total_conflict += skipped_conf
            total_no_newsletter += skipped_nl
            log_callback(
                f"  → {len(new_rows)} ligne(s) ajoutée(s), "
                f"{skipped_nl} refus newsletter ignoré(s), "
                f"{skipped_id} doublon(s) identique(s) ignoré(s), "
                f"{skipped_conf} conflit(s) ignoré(s)"
            )
        except Exception as e:
            log_callback(f"  ✗ Erreur lors du traitement : {e}")

    if all_new_rows:
        file_exists = csv_path.exists()
        with open(csv_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(
                f, fieldnames=fieldnames, delimiter=";", extrasaction="ignore"
            )
            if not file_exists:
                writer.writeheader()
            writer.writerows(all_new_rows)
        log_callback(f"\n✓ {len(all_new_rows)} nouvelle(s) ligne(s) écrite(s) dans {OUTPUT_FILENAME}")
    else:
        log_callback(f"\nAucune nouvelle ligne à ajouter dans {OUTPUT_FILENAME}")

    log_callback(
        f"Résumé final : {total_no_newsletter} refus newsletter ignoré(s), "
        f"{total_identical} doublon(s) identique(s) ignoré(s), "
        f"{total_conflict} conflit(s) de données ignoré(s)"
    )


# ── Interface graphique ───────────────────────────────────────────────────────

def launch_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Synthèse Excel → CSV")
            self.resizable(True, True)
            self.minsize(620, 420)
            self._build_ui()

        def _build_ui(self):
            pad = {"padx": 10, "pady": 6}

            frame_dir = ttk.LabelFrame(self, text="Répertoire source")
            frame_dir.pack(fill="x", **pad)

            self.dir_var = tk.StringVar()
            ttk.Entry(frame_dir, textvariable=self.dir_var, width=60).pack(
                side="left", padx=6, pady=6, fill="x", expand=True
            )
            ttk.Button(frame_dir, text="Parcourir…", command=self._browse).pack(
                side="left", padx=(0, 6), pady=6
            )

            ttk.Button(self, text="Lancer la synthèse", command=self._run).pack(**pad)

            frame_log = ttk.LabelFrame(self, text="Journal")
            frame_log.pack(fill="both", expand=True, **pad)

            self.log_widget = tk.Text(frame_log, state="disabled", wrap="word", height=15)
            scroll = ttk.Scrollbar(frame_log, command=self.log_widget.yview)
            self.log_widget.configure(yscrollcommand=scroll.set)
            scroll.pack(side="right", fill="y")
            self.log_widget.pack(fill="both", expand=True, padx=4, pady=4)

            self.progress = ttk.Progressbar(self, mode="indeterminate")
            self.progress.pack(fill="x", padx=10, pady=(0, 10))

        def _browse(self):
            directory = filedialog.askdirectory(title="Choisir un répertoire")
            if directory:
                self.dir_var.set(directory)

        def _log(self, message):
            self.log_widget.configure(state="normal")
            self.log_widget.insert("end", message + "\n")
            self.log_widget.see("end")
            self.log_widget.configure(state="disabled")
            self.update_idletasks()

        def _run(self):
            directory = self.dir_var.get().strip()
            if not directory or not os.path.isdir(directory):
                messagebox.showerror("Erreur", "Veuillez sélectionner un répertoire valide.")
                return

            self.log_widget.configure(state="normal")
            self.log_widget.delete("1.0", "end")
            self.log_widget.configure(state="disabled")
            self.progress.start(10)

            try:
                run_synthesis(directory, self._log)
            except Exception as e:
                self._log(f"Erreur inattendue : {e}")
            finally:
                self.progress.stop()

    App().mainloop()


# ── Point d'entrée ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Mode ligne de commande
        run_synthesis(sys.argv[1])
    else:
        try:
            launch_gui()
        except ImportError:
            print("tkinter non disponible. Usage : python excel_synthese.py <répertoire>")
            sys.exit(1)
