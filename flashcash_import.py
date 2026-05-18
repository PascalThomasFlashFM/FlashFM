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
    Les champs peuvent être dans n'importe quel ordre, séparés par virgules ou
    points-virgules.
    """
    result = {"prenom": "", "nom": "", "telephone": "", "ville": "", "email": ""}
    if not coordonnees or not isinstance(coordonnees, str):
        return result

    text = coordonnees.strip()

    # ── email ────────────────────────────────────────────────────────────────
    email_match = RE_EMAIL.search(text)
    if email_match:
        result["email"] = email_match.group(0).strip()
        text = text[:email_match.start()] + text[email_match.end():]

    # ── téléphone ────────────────────────────────────────────────────────────
    phone_match = RE_PHONE.search(text)
    if phone_match:
        result["telephone"] = normalise_phone(phone_match.group(0))
        text = text[:phone_match.start()] + text[phone_match.end():]

    # ── ville (code postal + ville ou ville seule connue) ────────────────────
    city_match = RE_POSTAL_CITY.search(text)
    if city_match:
        result["ville"] = city_match.group(2).strip().title()
        # Supprimer code postal + ville + éventuelle adresse de rue avant
        start = city_match.start()
        end = city_match.end()
        # On retire aussi l'adresse qui précède le code postal sur la même "partie"
        before = text[:start]
        after = text[end:]
        # Retirer la portion d'adresse qui se trouve entre la dernière virgule/point-virgule et le code postal
        before_clean = re.sub(r"[^,;]*$", "", before)
        text = before_clean + after
    else:
        # Tenter de trouver un code postal seul et retirer
        postal_match = RE_POSTAL.search(text)
        if postal_match:
            text = text[:postal_match.start()] + text[postal_match.end():]

    # ── Nom / Prénom depuis les fragments restants ───────────────────────────
    parts = [p.strip() for p in re.split(r"[,;]+", text) if p.strip()]
    name_parts = []
    for p in parts:
        if re.match(r"^\d", p):  # adresse avec numéro de rue
            continue
        if len(p) <= 2:
            continue
        name_parts.append(p)

    if name_parts:
        first_token = name_parts[0]
        words = first_token.split()
        if len(words) >= 2 and len(name_parts) == 1:
            # "DUPONT Jean" ou "Marie MARTIN" : détecter lequel est le nom (MAJUSCULES)
            upper_words = [w for w in words if w.isupper()]
            title_words = [w for w in words if w.istitle() and not w.isupper()]
            if upper_words and title_words:
                result["nom"] = " ".join(upper_words)
                result["prenom"] = " ".join(title_words).title()
            else:
                # Heuristique : dernier mot = nom, reste = prénom
                result["nom"] = words[-1].upper()
                result["prenom"] = " ".join(words[:-1]).title()
        else:
            result["nom"] = name_parts[0].strip().upper() if len(name_parts) >= 1 else ""
            result["prenom"] = name_parts[1].strip().title() if len(name_parts) >= 2 else ""

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


def find_first_empty_row(ws) -> int:
    """Retourne le numéro de la première ligne vide en colonne B."""
    col_b = ws.col_values(2)  # colonne B (index 2)
    # Chercher la dernière ligne non vide
    last = len(col_b)
    while last > 0 and not col_b[last - 1].strip():
        last -= 1
    return last + 1  # +1 pour la ligne suivante (1-indexé)


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
    Gère les encodages UTF-8 et Latin-1.
    """
    records = []
    encodings = ["utf-8-sig", "latin-1", "utf-8"]
    rows = None

    for enc in encodings:
        try:
            with open(filepath, newline="", encoding=enc) as f:
                sample = f.read(4096)
            # Détecter le délimiteur
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            with open(filepath, newline="", encoding=enc) as f:
                reader = csv.DictReader(f, dialect=dialect)
                rows = list(reader)
            break
        except Exception:
            continue

    if rows is None:
        raise ValueError("Impossible de lire le fichier CSV.")

    # Trouver la colonne Coordonnees (insensible à la casse)
    coord_col = None
    if rows:
        for col in rows[0].keys():
            if "coordonnee" in col.lower() or "coordonn" in col.lower():
                coord_col = col
                break

    if coord_col is None:
        raise ValueError(
            "Colonne 'Coordonnees' introuvable dans le CSV.\n"
            f"Colonnes disponibles : {list(rows[0].keys()) if rows else '(fichier vide)'}"
        )

    for row in rows:
        raw = row.get(coord_col, "")
        if raw.strip():
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
