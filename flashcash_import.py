"""
FlashCash - Import CSV vers Google Sheets
Copie les inscriptions joueurs d'un fichier CSV vers la feuille Google Sheets Flash Cash.
"""

import csv
import json
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

# ── Configuration ────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1pAoPThNMkDNh7zUMTEAJiltKaLAMioNAMCT1-ELqQnU"
SHEET_GID = 821980008
CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Regex patterns
RE_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", re.I)
RE_PHONE = re.compile(
    r"(?:(?:\+33|0033)\s*)?(?:0?\s*[1-9])(?:[\s.\-]?\d{2}){4}"
)
RE_POSTAL_CITY = re.compile(r"\b(\d{5})\s+([A-ZÀ-Ÿa-zà-ÿ\s\-]+?)(?=[,;\n]|$)", re.I)
RE_POSTAL = re.compile(r"\b\d{5}\b")


# ── Parsing helpers ───────────────────────────────────────────────────────────

def normalise_phone(raw: str) -> str:
    """Formate un numéro de téléphone en 'xx xx xx xx xx'."""
    digits = re.sub(r"\D", "", raw)
    if digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]
    if len(digits) == 10:
        return " ".join(digits[i:i+2] for i in range(0, 10, 2))
    return digits  # fallback brut si format inconnu


def extract_fields(coordonnees: str) -> dict:
    """
    Extrait Prénom, Nom, téléphone, ville, email depuis la chaîne 'Coordonnees'.
    Gère les formats : virgule-séparé, espace-séparé, multi-lignes, ordre variable.
    """
    result = {"prenom": "", "nom": "", "telephone": "", "ville": "", "email": ""}
    if not coordonnees or not isinstance(coordonnees, str):
        return result

    # Normaliser les retours à la ligne en espaces
    text = re.sub(r"\s*[\r\n]+\s*", " ", coordonnees).strip()

    # ── email ────────────────────────────────────────────────────────────────
    email_match = RE_EMAIL.search(text)
    if email_match:
        result["email"] = email_match.group(0).strip()
        text = text[:email_match.start()] + " " + text[email_match.end():]

    # ── téléphone ────────────────────────────────────────────────────────────
    phone_match = RE_PHONE.search(text)
    if phone_match:
        result["telephone"] = normalise_phone(phone_match.group(0))
        text = text[:phone_match.start()] + " " + text[phone_match.end():]

    # Normaliser les espaces multiples
    text = re.sub(r"\s+", " ", text).strip()

    # ── ville via code postal (XXXXX NomVille) ───────────────────────────────
    city_match = RE_POSTAL_CITY.search(text)
    if city_match:
        result["ville"] = city_match.group(2).strip().title()
        # Conserver uniquement ce qui précède le code postal,
        # en retirant les mots d'adresse juste avant (ceux qui commencent par un chiffre)
        prefix = text[:city_match.start()].strip().rstrip(",;")
        suffix = text[city_match.end():].strip().lstrip(",;")
        # Retirer les tokens d'adresse (numéro de rue, etc.) en fin de préfixe
        prefix_words = prefix.split()
        clean_prefix_words = []
        hit_address = False
        for w in prefix_words:
            if re.match(r"^\d", w):
                hit_address = True
            if not hit_address:
                clean_prefix_words.append(w)
        text = " ".join(clean_prefix_words) + (" " + suffix if suffix else "")
        text = re.sub(r"\s+", " ", text).strip()
    else:
        # Code postal seul sans ville lisible → retirer
        postal_match = RE_POSTAL.search(text)
        if postal_match:
            text = (text[:postal_match.start()] + " " + text[postal_match.end():]).strip()

    # ── Nom / Prénom ─────────────────────────────────────────────────────────
    def assign_name_city(words):
        """words[0]=Nom, words[1]=Prénom, words[2:]=Ville (si ville pas encore connue)."""
        if words:
            result["nom"] = words[0].upper()
        if len(words) >= 2:
            result["prenom"] = words[1].title()
        if len(words) >= 3 and not result["ville"]:
            result["ville"] = " ".join(words[2:]).title()

    # Cas 1 : séparateurs virgule ou point-virgule → split par virgule
    if re.search(r"[,;]", text):
        parts = [p.strip() for p in re.split(r"[,;]+", text)
                 if p.strip() and len(p.strip()) > 2 and not re.match(r"^\d", p.strip())]
        if parts:
            first_words = parts[0].split()
            if len(first_words) >= 2 and len(parts) == 1:
                # Token unique multi-mots : Nom Prénom [Ville] tout en un
                assign_name_city(first_words)
            elif len(first_words) >= 3 and len(parts) >= 1:
                # Première partie a 3+ mots → words[0]=Nom, words[1]=Prénom, la suite à la ville
                result["nom"] = first_words[0].upper()
                result["prenom"] = " ".join(first_words[1:]).title()
                if not result["ville"] and len(parts) > 1:
                    result["ville"] = parts[1].title()
            else:
                result["nom"] = parts[0].upper()
                result["prenom"] = parts[1].title() if len(parts) > 1 else ""
                if not result["ville"] and len(parts) > 2:
                    result["ville"] = parts[2].title()
    else:
        # Cas 2 : tout espace-séparé → mot[0]=Nom, mot[1]=Prénom, reste=Ville
        words = [w for w in text.split() if len(w) > 1 and not re.match(r"^\d", w)]
        assign_name_city(words)

    return result


# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_worksheet():
    """Retourne la feuille Google Sheets cible."""
    creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    # Trouver l'onglet par son gid
    for ws in spreadsheet.worksheets():
        if ws.id == SHEET_GID:
            return ws
    raise ValueError(f"Onglet avec gid={SHEET_GID} introuvable.")


def find_first_empty_row(ws, empty_threshold: int = 20) -> int:
    """
    Retourne la première ligne disponible après la dernière zone de données
    en colonne B, définie comme la dernière ligne non vide suivie d'au moins
    `empty_threshold` lignes vides consécutives.
    """
    col_b = ws.col_values(2)  # colonne B (index 2)

    # Parcourir depuis la fin pour trouver le dernier groupe de données
    # suivi d'au moins `empty_threshold` lignes vides
    n = len(col_b)
    consecutive_empty = 0
    last_data_row = 0  # 1-indexé

    for i in range(n - 1, -1, -1):
        if col_b[i].strip():
            if consecutive_empty >= empty_threshold or last_data_row == 0:
                last_data_row = i + 1  # convertir en index 1-based
                break
            else:
                consecutive_empty = 0
        else:
            consecutive_empty += 1

    # Si on n'a pas encore trouvé (boucle terminée sans break)
    if last_data_row == 0:
        for i in range(n - 1, -1, -1):
            if col_b[i].strip():
                last_data_row = i + 1
                break

    return last_data_row + 1  # ligne suivant la dernière donnée


def write_to_sheet(records: list, log_callback=None) -> int:
    """
    Écrit les enregistrements dans le Google Sheet.
    Retourne le nombre de lignes ajoutées.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    log("Connexion à Google Sheets...")
    ws = get_worksheet()
    start_row = find_first_empty_row(ws)
    log(f"Première ligne vide : {start_row}")

    rows_to_append = []
    for rec in records:
        # On prépare une ligne avec les colonnes A→G
        # A=vide, B=Prénom, C=Nom, D=Téléphone, E=Ville, F=vide, G=Email
        row = ["", rec["prenom"], rec["nom"], rec["telephone"], rec["ville"], "", rec["email"]]
        rows_to_append.append(row)

    if not rows_to_append:
        log("Aucun enregistrement à écrire.")
        return 0

    # Écriture en batch
    end_row = start_row + len(rows_to_append) - 1
    cell_range = f"A{start_row}:G{end_row}"
    log(f"Écriture de {len(rows_to_append)} ligne(s) dans {cell_range}...")
    ws.update(cell_range, rows_to_append, value_input_option="USER_ENTERED")
    log(f"✓ {len(rows_to_append)} ligne(s) ajoutée(s) avec succès.")
    return len(rows_to_append)


# ── Lecture CSV ───────────────────────────────────────────────────────────────

def parse_csv(filepath: str) -> list:
    """
    Lit le CSV et retourne une liste de dicts avec les champs extraits.
    Gère les encodages UTF-8 et Latin-1, les délimiteurs ; et ,
    et les champs multi-lignes entre guillemets.
    """
    records = []
    encodings = ["utf-8-sig", "latin-1", "utf-8"]

    rows = None
    last_error = None

    for enc in encodings:
        # Essayer d'abord ';' (exports français), puis ',' puis détection auto
        for delimiter in [";", ",", "\t"]:
            try:
                with open(filepath, newline="", encoding=enc) as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    candidate = list(reader)
                # Valider : au moins 2 colonnes et au moins 1 ligne de données
                if candidate and len(candidate[0].keys()) >= 2:
                    rows = candidate
                    break
            except Exception as exc:
                last_error = exc
        if rows is not None:
            break

    if rows is None:
        raise ValueError(
            f"Impossible de lire le fichier CSV.\n"
            f"Dernière erreur : {last_error}"
        )

    # Trouver la colonne Coordonnees (insensible à la casse, ignore BOM/espaces)
    coord_col = None
    for col in rows[0].keys():
        col_clean = col.strip().lstrip("﻿").lower()
        if "coordonn" in col_clean:
            coord_col = col
            break

    if coord_col is None:
        cols = list(rows[0].keys())
        raise ValueError(
            f"Colonne 'Coordonnees' introuvable dans le CSV.\n"
            f"Colonnes détectées ({len(cols)}) : {cols}"
        )

    for row in rows:
        raw = row.get(coord_col, "") or ""
        # Nettoyer le contenu multi-lignes (retours à la ligne dans le champ)
        raw = " ".join(raw.splitlines()).strip()
        if raw:
            fields = extract_fields(raw)
            fields["_raw"] = raw
            records.append(fields)

    return records


# ── Interface graphique ───────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FlashCash – Import CSV → Google Sheets")
        self.resizable(True, True)
        self.minsize(720, 500)
        self._csv_path = tk.StringVar()
        self._records: list = []
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 800, 600
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # ── En-tête ──────────────────────────────────────────────────────────
        header = tk.Frame(self, bg="#1a1a2e", pady=14)
        header.pack(fill="x")
        tk.Label(
            header, text="FlashCash  –  Import Joueurs",
            font=("Helvetica", 17, "bold"), fg="#e8c547", bg="#1a1a2e"
        ).pack()
        tk.Label(
            header, text="Importation CSV vers Google Sheets",
            font=("Helvetica", 10), fg="#aaaacc", bg="#1a1a2e"
        ).pack()

        # ── Sélection fichier ─────────────────────────────────────────────────
        file_frame = tk.LabelFrame(self, text=" Fichier CSV ", padx=10, pady=8)
        file_frame.pack(fill="x", padx=16, pady=(14, 4))

        entry = tk.Entry(file_frame, textvariable=self._csv_path, state="readonly",
                         width=60, font=("Courier", 10))
        entry.pack(side="left", expand=True, fill="x", padx=(0, 8))

        tk.Button(
            file_frame, text="Parcourir…", command=self._browse,
            bg="#3a3a5c", fg="white", relief="flat", padx=10, pady=4,
            cursor="hand2"
        ).pack(side="left")

        # ── Prévisualisation ──────────────────────────────────────────────────
        preview_frame = tk.LabelFrame(self, text=" Aperçu des données parsées ", padx=10, pady=6)
        preview_frame.pack(fill="both", expand=True, padx=16, pady=4)

        cols = ("Prénom", "Nom", "Téléphone", "Ville", "Email")
        self._tree = ttk.Treeview(preview_frame, columns=cols, show="headings", height=8)
        widths = [110, 130, 120, 130, 210]
        for col, w in zip(cols, widths):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=w, anchor="w")

        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # ── Journal ───────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self, text=" Journal ", padx=10, pady=6)
        log_frame.pack(fill="x", padx=16, pady=4)

        self._log = scrolledtext.ScrolledText(
            log_frame, height=6, state="disabled",
            font=("Courier", 9), bg="#0d0d1a", fg="#88ff88"
        )
        self._log.pack(fill="x")

        # ── Boutons d'action ──────────────────────────────────────────────────
        btn_frame = tk.Frame(self, pady=10)
        btn_frame.pack()

        self._btn_import = tk.Button(
            btn_frame, text="▶  Écrire dans Google Sheets",
            command=self._import, state="disabled",
            font=("Helvetica", 12, "bold"),
            bg="#e8c547", fg="#1a1a2e", relief="flat",
            padx=20, pady=8, cursor="hand2"
        )
        self._btn_import.pack(side="left", padx=8)

        tk.Button(
            btn_frame, text="Réinitialiser", command=self._reset,
            bg="#555577", fg="white", relief="flat",
            padx=12, pady=8, cursor="hand2"
        ).pack(side="left", padx=8)

    # ── Actions ───────────────────────────────────────────────────────────────

    def _log_msg(self, msg: str):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")
        self.update_idletasks()

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if not path:
            return
        self._csv_path.set(path)
        self._load_preview(path)

    def _load_preview(self, path: str):
        self._tree.delete(*self._tree.get_children())
        self._records = []
        try:
            self._records = parse_csv(path)
            for rec in self._records:
                self._tree.insert("", "end", values=(
                    rec["prenom"], rec["nom"],
                    rec["telephone"], rec["ville"], rec["email"]
                ))
            self._log_msg(f"✓ Fichier chargé : {len(self._records)} inscrit(s) détecté(s).")
            self._btn_import.configure(state="normal")
        except Exception as exc:
            messagebox.showerror("Erreur de lecture CSV", str(exc))
            self._log_msg(f"✗ Erreur : {exc}")
            self._btn_import.configure(state="disabled")

    def _import(self):
        if not self._records:
            messagebox.showwarning("Aucune donnée", "Veuillez d'abord charger un fichier CSV.")
            return

        if not messagebox.askyesno(
            "Confirmation",
            f"Écrire {len(self._records)} ligne(s) dans Google Sheets ?\n\n"
            "Les données seront ajoutées à la suite de la liste existante."
        ):
            return

        self._btn_import.configure(state="disabled", text="⏳  En cours…")
        self.update_idletasks()

        try:
            count = write_to_sheet(self._records, log_callback=self._log_msg)
            messagebox.showinfo(
                "Import réussi",
                f"{count} joueur(s) ajouté(s) dans Google Sheets avec succès."
            )
            self._btn_import.configure(text="✓  Import terminé", bg="#44aa44", fg="white")
        except Exception as exc:
            messagebox.showerror("Erreur d'import", str(exc))
            self._log_msg(f"✗ Erreur : {exc}")
            self._btn_import.configure(
                state="normal", text="▶  Écrire dans Google Sheets",
                bg="#e8c547", fg="#1a1a2e"
            )

    def _reset(self):
        self._csv_path.set("")
        self._records = []
        self._tree.delete(*self._tree.get_children())
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        self._btn_import.configure(
            state="disabled", text="▶  Écrire dans Google Sheets",
            bg="#e8c547", fg="#1a1a2e"
        )


# ── Point d'entrée ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not CREDENTIALS_FILE.exists():
        print(f"ERREUR : fichier de credentials introuvable : {CREDENTIALS_FILE}", file=sys.stderr)
        sys.exit(1)
    app = App()
    app.mainloop()
