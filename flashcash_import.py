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
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

# ── Configuration ────────────────────────────────────────────
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
RE_POSTAL_CITY = re.compile(r"\b(\d{5})\s+([A-Z\xc0-Ÿa-z\xe0-\xff\s\-]+?)(?=[,;\n]|$)", re.I)
RE_POSTAL = re.compile(r"\b\d{5}\b")
RE_DATE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")


# ── Parsing helpers ───────────────────────────────────────────────

def normalise_phone(raw: str) -> str:
    digits = re.sub(r"\D", "", raw)
    if digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]
    if len(digits) == 10:
        return " ".join(digits[i:i+2] for i in range(0, 10, 2))
    return digits


def extract_fields(coordonnees: str) -> dict:
    result = {"prenom": "", "nom": "", "telephone": "", "ville": "", "email": ""}
    if not coordonnees or not isinstance(coordonnees, str):
        return result

    text = re.sub(r"\s*[\r\n]+\s*", " ", coordonnees).strip()

    email_match = RE_EMAIL.search(text)
    if email_match:
        result["email"] = email_match.group(0).strip()
        text = text[:email_match.start()] + " " + text[email_match.end():]

    phone_match = RE_PHONE.search(text)
    if phone_match:
        result["telephone"] = normalise_phone(phone_match.group(0))
        text = text[:phone_match.start()] + " " + text[phone_match.end():]

    text = re.sub(r"\s+", " ", text).strip()

    city_match = RE_POSTAL_CITY.search(text)
    if city_match:
        result["ville"] = city_match.group(2).strip().title()
        prefix = text[:city_match.start()].strip().rstrip(",;")
        suffix = text[city_match.end():].strip().lstrip(",;")
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
        postal_match = RE_POSTAL.search(text)
        if postal_match:
            text = (text[:postal_match.start()] + " " + text[postal_match.end():]).strip()

    def assign_name_city(words):
        if words:
            result["nom"] = words[0].upper()
        if len(words) >= 2:
            result["prenom"] = words[1].title()
        if len(words) >= 3 and not result["ville"]:
            result["ville"] = " ".join(words[2:]).title()

    if re.search(r"[,;]", text):
        parts = [p.strip() for p in re.split(r"[,;]+", text)
                 if p.strip() and len(p.strip()) > 2 and not re.match(r"^\d", p.strip())]
        if parts:
            first_words = parts[0].split()
            if len(first_words) >= 2 and len(parts) == 1:
                assign_name_city(first_words)
            elif len(first_words) >= 3 and len(parts) >= 1:
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
        words = [w for w in text.split() if len(w) > 1 and not re.match(r"^\d", w)]
        assign_name_city(words)

    return result


def _phone_key(phone: str) -> str:
    digits = re.sub(r"\D", "", phone)
    if digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]
    return digits


def get_existing_keys(all_values: list) -> set:
    keys = set()
    for row in all_values:
        if len(row) < 4:
            continue
        prenom = row[1].strip() if len(row) > 1 else ""
        nom    = row[2].strip() if len(row) > 2 else ""
        tel    = row[3].strip() if len(row) > 3 else ""
        email  = row[6].strip().lower() if len(row) > 6 else ""
        phone_digits = _phone_key(tel)
        if phone_digits and len(phone_digits) >= 9:
            keys.add("tel:" + phone_digits)
        if email and "@" in email:
            keys.add("email:" + email)
        if nom or prenom:
            keys.add("name:" + (nom + prenom).lower().replace(" ", ""))
    return keys


def get_history_index(all_values: list) -> dict:
    index = {}
    for row in all_values:
        col_a  = row[0].strip() if len(row) > 0 else ""
        prenom = row[1].strip() if len(row) > 1 else ""
        nom    = row[2].strip() if len(row) > 2 else ""
        if not RE_DATE.match(col_a):
            continue
        if not nom and not prenom:
            continue
        key = (nom.lower(), prenom.lower())
        index.setdefault(key, []).append(col_a)
    return index


def _parse_date(date_str: str):
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            pass
    return None


def record_key(rec: dict) -> list:
    keys = []
    phone_digits = _phone_key(rec.get("telephone", ""))
    if phone_digits and len(phone_digits) >= 9:
        keys.append("tel:" + phone_digits)
    email = rec.get("email", "").strip().lower()
    if email and "@" in email:
        keys.append("email:" + email)
    nom    = rec.get("nom", "").strip()
    prenom = rec.get("prenom", "").strip()
    if nom or prenom:
        keys.append("name:" + (nom + prenom).lower().replace(" ", ""))
    return keys


def get_worksheet():
    creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    for ws in spreadsheet.worksheets():
        if ws.id == SHEET_GID:
            return ws
    raise ValueError(f"Onglet avec gid={SHEET_GID} introuvable.")


def find_first_empty_row(ws, empty_threshold: int = 20) -> int:
    col_b = ws.col_values(2)
    n = len(col_b)
    consecutive_empty = 0
    last_data_row = 0

    for i in range(n - 1, -1, -1):
        if col_b[i].strip():
            if consecutive_empty >= empty_threshold or last_data_row == 0:
                last_data_row = i + 1
                break
            else:
                consecutive_empty = 0
        else:
            consecutive_empty += 1

    if last_data_row == 0:
        for i in range(n - 1, -1, -1):
            if col_b[i].strip():
                last_data_row = i + 1
                break

    return last_data_row + 1


def write_to_sheet(records: list, log_callback=None) -> tuple:
    def log(msg):
        if log_callback:
            log_callback(msg)

    log("Connexion \xe0 Google Sheets...")
    ws = get_worksheet()

    log("Lecture des donn\xe9es existantes dans la feuille...")
    all_values = ws.get_all_values()
    existing_keys = get_existing_keys(all_values)
    history_index = get_history_index(all_values)
    log(f"  → {len(existing_keys)} cl\xe9s d'identification trouv\xe9es.")

    start_row = find_first_empty_row(ws)
    log(f"Premi\xe8re ligne disponible : {start_row}")

    rows_to_append = []
    skipped = 0
    for rec in records:
        keys = record_key(rec)
        if any(k in existing_keys for k in keys):
            skipped += 1
            continue

        nom_lower    = rec["nom"].strip().lower()
        prenom_lower = rec["prenom"].strip().lower()
        hist_dates   = history_index.get((nom_lower, prenom_lower), [])

        col_h = ""
        col_i = ""
        if hist_dates:
            col_h = str(len(hist_dates))
            parsed = [_parse_date(d) for d in hist_dates]
            parsed = [d for d in parsed if d is not None]
            if parsed:
                last_dt = max(parsed)
                col_i = last_dt.strftime("%d/%m/%Y")

        row = ["", rec["prenom"], rec["nom"], rec["telephone"], rec["ville"], "", rec["email"], col_h, col_i]
        rows_to_append.append(row)
        for k in keys:
            existing_keys.add(k)

    if skipped:
        log(f"  → {skipped} inscrit(s) d\xe9j\xe0 pr\xe9sent(s) ignor\xe9(s).")

    if not rows_to_append:
        log("Aucun nouvel enregistrement \xe0 \xe9crire.")
        return 0, skipped

    end_row = start_row + len(rows_to_append) - 1
    cell_range = f"A{start_row}:I{end_row}"
    log(f"\xc9criture de {len(rows_to_append)} nouvelle(s) ligne(s) dans {cell_range}...")
    ws.update(cell_range, rows_to_append, value_input_option="USER_ENTERED")
    log(f"✓ {len(rows_to_append)} joueur(s) ajout\xe9(s) avec succ\xe8s.")
    return len(rows_to_append), skipped


# ── Lecture CSV ─────────────────────────────────────────────────────────────────

def parse_csv(filepath: str) -> list:
    records = []
    encodings = ["utf-8-sig", "latin-1", "utf-8"]

    rows = None
    last_error = None

    for enc in encodings:
        for delimiter in [";", ",", "\t"]:
            try:
                with open(filepath, newline="", encoding=enc) as f:
                    reader = csv.DictReader(f, delimiter=delimiter)
                    candidate = list(reader)
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
            f"Derni\xe8re erreur : {last_error}"
        )

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
            f"Colonnes d\xe9tect\xe9es ({len(cols)}) : {cols}"
        )

    for row in rows:
        raw = row.get(coord_col, "") or ""
        raw = " ".join(raw.splitlines()).strip()
        if raw:
            fields = extract_fields(raw)
            fields["_raw"] = raw
            records.append(fields)

    return records


# ── Interface graphique ───────────────────────────────────────────────────────────────

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
        BG       = "#1e1e2e"
        BG2      = "#2a2a3e"
        FG       = "#cdd6f4"
        FG_DIM   = "#aaaacc"
        ACCENT   = "#e8c547"
        BTN_BROWSE  = "#4a4a7a"
        BTN_RESET   = "#5a3a6a"
        BTN_IMPORT  = "#2d6a4f"
        BTN_DONE    = "#1a5c38"

        self.configure(bg=BG)

        header = tk.Frame(self, bg="#12122a", pady=14)
        header.pack(fill="x")
        tk.Label(
            header, text="FlashCash  –  Import Joueurs",
            font=("Helvetica", 17, "bold"), fg=ACCENT, bg="#12122a"
        ).pack()
        tk.Label(
            header, text="Importation CSV vers Google Sheets",
            font=("Helvetica", 10), fg=FG_DIM, bg="#12122a"
        ).pack()

        file_frame = tk.LabelFrame(
            self, text=" Fichier CSV ",
            bg=BG, fg=FG, padx=10, pady=8
        )
        file_frame.pack(fill="x", padx=16, pady=(14, 4))

        entry = tk.Entry(
            file_frame, textvariable=self._csv_path, state="readonly",
            width=60, font=("Courier", 10),
            bg=BG2, fg=FG, insertbackground=FG,
            readonlybackground=BG2, relief="flat", bd=4
        )
        entry.pack(side="left", expand=True, fill="x", padx=(0, 8))

        tk.Button(
            file_frame, text="Parcourir…", command=self._browse,
            bg=BTN_BROWSE, fg="white", activebackground="#6060a0",
            activeforeground="white", relief="raised", padx=12, pady=5,
            font=("Helvetica", 10, "bold"), cursor="hand2", bd=2,
            highlightthickness=0
        ).pack(side="left")

        preview_frame = tk.LabelFrame(
            self, text=" Aper\xe7u des donn\xe9es pars\xe9es ",
            bg=BG, fg=FG, padx=10, pady=6
        )
        preview_frame.pack(fill="both", expand=True, padx=16, pady=4)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Dark.Treeview",
            background=BG2, foreground=FG,
            fieldbackground=BG2, rowheight=22,
            font=("Helvetica", 9)
        )
        style.configure("Dark.Treeview.Heading",
            background="#3a3a5e", foreground=ACCENT,
            font=("Helvetica", 9, "bold"), relief="flat"
        )
        style.map("Dark.Treeview", background=[("selected", "#3d5a80")])

        cols = ("Pr\xe9nom", "Nom", "T\xe9l\xe9phone", "Ville", "Email")
        self._tree = ttk.Treeview(
            preview_frame, columns=cols, show="headings",
            height=8, style="Dark.Treeview"
        )
        widths = [110, 130, 120, 130, 210]
        for col, w in zip(cols, widths):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=w, anchor="w")

        vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        log_frame = tk.LabelFrame(
            self, text=" Journal ",
            bg=BG, fg=FG, padx=10, pady=6
        )
        log_frame.pack(fill="x", padx=16, pady=4)

        self._log = scrolledtext.ScrolledText(
            log_frame, height=5, state="disabled",
            font=("Courier", 9), bg="#0d0d1a", fg="#88ff88",
            insertbackground="#88ff88", relief="flat", bd=2
        )
        self._log.pack(fill="x")

        btn_frame = tk.Frame(self, bg=BG, pady=12)
        btn_frame.pack()

        self._btn_import = tk.Button(
            btn_frame, text="▶  \xc9crire dans Google Sheets",
            command=self._import, state="disabled",
            font=("Helvetica", 12, "bold"),
            bg=BTN_IMPORT, fg="white",
            activebackground="#3d8b62", activeforeground="white",
            disabledforeground="#888888", disabledbackground="#1a3d2b",
            relief="raised", padx=22, pady=10, cursor="hand2", bd=2,
            highlightthickness=0
        )
        self._btn_import.pack(side="left", padx=8)
        self._BTN_IMPORT_COLOR = BTN_IMPORT
        self._BTN_DONE_COLOR   = BTN_DONE

        tk.Button(
            btn_frame, text="R\xe9initialiser", command=self._reset,
            bg=BTN_RESET, fg="white",
            activebackground="#7a4a8a", activeforeground="white",
            relief="raised", padx=14, pady=10,
            font=("Helvetica", 10, "bold"), cursor="hand2", bd=2,
            highlightthickness=0
        ).pack(side="left", padx=8)

    # ── Actions ───────────────────────────────────────────────────────────────────────

    def _log_msg(self, msg: str):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")
        self.update_idletasks()

    def _browse(self):
        path = filedialog.askopenfilename(
            title="S\xe9lectionner le fichier CSV",
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
            self._log_msg(f"✓ Fichier charg\xe9 : {len(self._records)} inscrit(s) d\xe9tect\xe9(s).")
            self._btn_import.configure(state="normal")
        except Exception as exc:
            messagebox.showerror("Erreur de lecture CSV", str(exc))
            self._log_msg(f"✗ Erreur : {exc}")
            self._btn_import.configure(state="disabled")

    def _import(self):
        if not self._records:
            messagebox.showwarning("Aucune donn\xe9e", "Veuillez d'abord charger un fichier CSV.")
            return

        if not messagebox.askyesno(
            "Confirmation",
            f"{len(self._records)} inscrit(s) dans le fichier CSV.\n\n"
            "L'application ignorera automatiquement ceux d\xe9j\xe0 pr\xe9sents\n"
            "dans Google Sheets et n'ajoutera que les nouveaux.\n\n"
            "Continuer ?"
        ):
            return

        self._btn_import.configure(state="disabled", text="⏳  En cours…")
        self.update_idletasks()

        try:
            added, skipped = write_to_sheet(self._records, log_callback=self._log_msg)
            msg = f"{added} nouveau(x) joueur(s) ajout\xe9(s) dans Google Sheets."
            if skipped:
                msg += f"\n{skipped} inscrit(s) d\xe9j\xe0 pr\xe9sent(s) ignor\xe9(s)."
            messagebox.showinfo("Import termin\xe9", msg)
            self._btn_import.configure(text="✓  Import termin\xe9", bg=self._BTN_DONE_COLOR, fg="white")
        except Exception as exc:
            messagebox.showerror("Erreur d'import", str(exc))
            self._log_msg(f"✗ Erreur : {exc}")
            self._btn_import.configure(
                state="normal", text="▶  \xc9crire dans Google Sheets",
                bg=self._BTN_IMPORT_COLOR, fg="white"
            )

    def _reset(self):
        self._csv_path.set("")
        self._records = []
        self._tree.delete(*self._tree.get_children())
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")
        self._btn_import.configure(
            state="disabled", text="▶  \xc9crire dans Google Sheets",
            bg=self._BTN_IMPORT_COLOR, fg="white"
        )


# ── Point d'entrée ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not CREDENTIALS_FILE.exists():
        print(f"ERREUR : fichier de credentials introuvable : {CREDENTIALS_FILE}", file=sys.stderr)
        sys.exit(1)
    app = App()
    app.mainloop()
