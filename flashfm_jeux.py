#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Flash FM – Gestionnaire de Jeux Concours
Tirage au sort avec vérification d'éligibilité, export Google Sheets,
envoi d'emails personnalisés.
Nécessite : pip install gspread google-auth
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json, os, csv, io, re, random, smtplib, threading, unicodedata, shutil, time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import quote

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

# ─── Chemins ──────────────────────────────────────────────
BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_PATH = os.path.join(BASE_DIR, "flashfm_templates.json")
CLUBS_PATH     = os.path.join(BASE_DIR, "flashfm_clubs.json")

# ─── SMTP OVH ─────────────────────────────────────────────
SMTP_SERVER    = "ssl0.ovh.net"
SMTP_PORT      = 465
FROM_ADDR      = "pascal@flashfm.fr"
FROM_NAME      = "Pascal Thomas Flash FM"
SMTP_LOGIN     = "pascal@flashfm.fr"
SMTP_PASSWORD  = "anton07"
TEST_EMAIL     = "pascal@flashfm.fr"
SPREADSHEET_ID = "1pAoPThNMkDNh7zUMTEAJiltKaLAMioNAMCT1-ELqQnU"

# ─── Couleurs ─────────────────────────────────────────────
C_RED   = "#C0392B"
C_DARK  = "#2C3E50"
C_WHITE = "#FFFFFF"
C_LIGHT = "#F0F2F5"
C_GREEN = "#27AE60"
C_BLUE  = "#2980B9"
C_GRAY  = "#95A5A6"
C_LGRAY = "#BDC3C7"

# ─── Regex CSV ────────────────────────────────────────────
EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+\.[a-z]{2,}', re.IGNORECASE)
PHONE_RE = re.compile(
    r'(?:(?:\+33|0033)?[\s.-]?[67]\d[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}[\s.-]?\d{2}'
    r'|0[1-9](?:[\s./]?\d{2}){4})'
)


# ══════════════════════════════════════════════════════════
#  Utilitaires
# ══════════════════════════════════════════════════════════

def normalize(s):
    """Minuscules, sans accents, espaces normalisés."""
    s = s.lower().strip()
    return ' '.join(
        ''.join(c for c in unicodedata.normalize('NFD', w)
                if unicodedata.category(c) != 'Mn')
        for w in s.split()
    )


def format_phone(phone):
    """Formate un numéro 10 chiffres en XX XX XX XX XX."""
    digits = re.sub(r'\D', '', phone)
    if len(digits) == 10:
        return ' '.join(digits[i:i+2] for i in range(0, 10, 2))
    return phone


def parse_date(s):
    """Tente de parser une date en plusieurs formats ; retourne None si échec."""
    s = s.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%d %m %Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


# ══════════════════════════════════════════════════════════
#  Parsing CSV
# ══════════════════════════════════════════════════════════

def parse_csv(filepath):
    """Parse un export CSV MGS ; retourne la liste de participants uniques."""
    with open(filepath, encoding="utf-8") as f:
        raw = f.read()
    reader = csv.reader(io.StringIO(raw), delimiter=';', quotechar='"')
    rows   = list(reader)

    participants, seen = [], {}
    for row in rows[1:]:
        if len(row) < 8:
            continue
        coordonnees = row[7].strip()
        phone_col   = row[6].strip() if len(row) > 6 else ""
        if not coordonnees:
            continue

        email_match = EMAIL_RE.search(coordonnees)
        email = email_match.group(0) if email_match else ""

        phone_match = PHONE_RE.search(coordonnees)
        if phone_match:
            phone = re.sub(r'[\s./]', '', phone_match.group(0))
            phone = re.sub(r'^(\+33|0033)', '0', phone)
        elif phone_col:
            phone = re.sub(r'[\s./]', '', phone_col)
            phone = re.sub(r'^(\+33|0033)', '0', phone)
        else:
            phone = ""

        text = coordonnees
        if email_match:
            text = text.replace(email_match.group(0), '')
        if phone_match:
            text = text.replace(phone_match.group(0), '')
        text = re.sub(r'\b\d{5}\b', '', text)
        text = re.sub(r'\bet\b', '', text)
        text = re.sub(r'[,;]+', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip(' .-')

        parts  = text.split()
        nom    = parts[0].upper()      if len(parts) >= 1 else "?"
        prenom = parts[1].capitalize() if len(parts) >= 2 else ""
        ville  = ' '.join(w.capitalize() for w in parts[2:] if len(w) > 1)

        key = (phone, email.lower())
        if key not in seen:
            seen[key] = True
            participants.append({
                'nom': nom, 'prenom': prenom,
                'ville': ville,
                'email': email, 'phone': phone,
            })
    return participants


# ══════════════════════════════════════════════════════════
#  Vérification d'éligibilité (Google Sheets)
# ══════════════════════════════════════════════════════════

def _sheets_client(creds_path):
    creds = Credentials.from_service_account_file(
        creds_path,
        scopes=["https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"]
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID)


def _find_col(header_row, keywords):
    for i, h in enumerate(header_row):
        if any(kw in normalize(h) for kw in keywords):
            return i
    return None


def validate_and_register_winner(candidate, nom_jeu, sh, banned_rows,
                                  col_inserted):
    """
    Vérifie l'éligibilité ET enregistre le gagnant dans 'liste des gagnants'.
    Retourne (valide: bool, raison: str).

    col_inserted : liste à un élément [bool] partagée entre les appels.
    La colonne F n'est insérée qu'une seule fois (premier gagnant validé) ;
    les gagnants suivants réutilisent cette même colonne.
    """
    nom    = normalize(candidate['nom'])
    prenom = normalize(candidate['prenom'])
    today  = datetime.now()

    # ── Vérification 1 : liste des bannis ─────────────────
    for row in banned_rows:
        row_text = normalize(' '.join(row))
        if nom in row_text and prenom in row_text:
            return False, "Banni des jeux"

    # ── Vérification 2 : lecture fraîche de l'historique ──
    ws       = sh.worksheet("liste des gagnants")
    all_vals = ws.get_all_values()

    header_row = all_vals[2] if len(all_vals) > 2 else []
    nom_col    = _find_col(header_row, ['nom'])
    prenom_col = _find_col(header_row, ['prenom', 'prénom'])
    if nom_col    is None: nom_col    = 0
    if prenom_col is None: prenom_col = 1

    date_cols = [i for i in range(5, len(header_row)) if header_row[i].strip()]

    winner_row_1idx = None
    for i, row in enumerate(all_vals):
        if i < 3 or not any(c.strip() for c in row):
            continue
        n_val = row[nom_col].strip()    if nom_col    < len(row) else ""
        p_val = row[prenom_col].strip() if prenom_col < len(row) else ""
        if normalize(n_val) == nom and normalize(p_val) == prenom:
            winner_row_1idx = i + 1
            break

    six_months_ago = today - timedelta(days=183)

    if winner_row_1idx is None:
        # Nouveau gagnant : on ajoute une ligne
        first_empty = len(all_vals) + 1
        for i in range(3, len(all_vals)):
            if not any(c.strip() for c in all_vals[i]):
                first_empty = i + 1
                break
        ws.update(
            values=[[candidate['nom'], candidate['prenom'],
                     candidate.get('ville', ''), candidate['email'],
                     format_phone(candidate['phone'])]],
            range_name=f"A{first_empty}:E{first_empty}"
        )
        winner_row_1idx = first_empty
        reason = "OK"
    else:
        # Gagnant connu : vérifier dates récentes
        row = all_vals[winner_row_1idx - 1]
        recent = []
        for i in date_cols:
            if i < len(row) and row[i].strip():
                d = parse_date(row[i])
                if d and d > six_months_ago:
                    recent.append(d)
        if recent:
            last = max(recent).strftime('%d/%m/%Y')
            return False, f"A gagné le {last} (moins de 6 mois)"
        last_all = []
        for i in date_cols:
            if i < len(row) and row[i].strip():
                d = parse_date(row[i])
                if d:
                    last_all.append(d)
        reason = (f"OK (dernière victoire le {max(last_all).strftime('%d/%m/%Y')}, "
                  f"plus de 6 mois)") if last_all else "OK (déjà gagnant, sans date récente)"

    # ── Colonne F : insérée une seule fois pour le tirage ──
    if not col_inserted[0]:
        sh.batch_update({"requests": [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": 5, "endIndex": 6},
            "inheritFromBefore": False
        }}]})
        MOIS_FR = ["janvier","février","mars","avril","mai","juin",
                   "juillet","août","septembre","octobre","novembre","décembre"]
        mois_an = f"{MOIS_FR[today.month - 1]} {today.year}"
        ws.update(values=[[f"{nom_jeu} {mois_an}"]], range_name="F3")
        col_inserted[0] = True
    ws.update(values=[[today.strftime("%d/%m/%Y")]], range_name=f"F{winner_row_1idx}")

    return True, reason


def draw_with_checks(participants, n, creds_path, nom_jeu,
                     log_cb, done_cb, error_cb):
    """
    Tirage au sort aléatoire avec vérification d'éligibilité et enregistrement.
    Exécuté dans un thread secondaire.
    """
    sh = _sheets_client(creds_path)

    banned_ws  = sh.worksheet("Joueurs à bannir des jeux")
    all_banned = banned_ws.get_all_values()
    banned_rows = [row for row in all_banned[1:] if any(c.strip() for c in row)]
    log_cb(f"Données chargées : {len(banned_rows)} joueur(s) banni(s).")

    pool = [p for p in participants if p['email']]
    random.shuffle(pool)

    winners, excluded = [], 0
    col_inserted = [False]

    for candidate in pool:
        if len(winners) >= n:
            break
        try:
            valid, reason = validate_and_register_winner(
                candidate, nom_jeu, sh, banned_rows, col_inserted)
        except Exception as e:
            log_cb(f"⚠  Erreur {candidate['prenom']} {candidate['nom']} : {e}")
            continue
        if valid:
            winners.append(candidate)
            log_cb(f"✓  {candidate['prenom']} {candidate['nom']} — {reason}")
        else:
            excluded += 1
            log_cb(f"⚠  {candidate['prenom']} {candidate['nom']} — {reason} → remplacé")

    if len(winners) < n:
        error_cb(f"Attention : seulement {len(winners)}/{n} gagnants éligibles trouvés "
                 f"({excluded} exclus).")

    done_cb(winners)


# ══════════════════════════════════════════════════════════
#  Export Google Sheets
# ══════════════════════════════════════════════════════════

def export_to_sheets(creds_path, tab_name, winners, game_info):
    sh = _sheets_client(creds_path)

    try:
        sh.del_worksheet(sh.worksheet(tab_name))
    except gspread.exceptions.WorksheetNotFound:
        pass

    ws = sh.add_worksheet(title=tab_name, rows=len(winners) + 5, cols=6)

    title  = (f"{game_info['nom_jeu']}  –  "
              f"{game_info['date']}  –  {game_info['lieu']}")
    header = ["Nom", "Prénom", "Ville", "Email", "Téléphone"]
    data   = [
        [w['nom'], w['prenom'], w.get('ville', ''),
         w['email'], format_phone(w['phone'])]
        for w in winners
    ]

    ws.update("A1", [
        [title, "", "", "", ""],
        ["",    "", "", "", ""],
        header
    ] + data)

    sid = ws.id
    sh.batch_update({"requests": [
        # Fusion titre
        {"mergeCells": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 5},
            "mergeType": "MERGE_ALL"}},
        # Style titre
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.75, "green": 0.05, "blue": 0.05},
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontSize": 13,
                               "foregroundColor": {"red": 1, "green": 1, "blue": 1}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)"}},
        # Style en-têtes
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 3,
                      "startColumnIndex": 0, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.13, "green": 0.24, "blue": 0.49},
                "textFormat": {"bold": True,
                               "foregroundColor": {"red": 1, "green": 1, "blue": 1}}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
        # Hauteur ligne titre
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 42}, "fields": "pixelSize"}},
        # Largeur colonnes
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 5},
            "properties": {"pixelSize": 160}, "fields": "pixelSize"}},
    ]})

    return (f"https://docs.google.com/spreadsheets/d/"
            f"{SPREADSHEET_ID}/edit#gid={sid}")


# ══════════════════════════════════════════════════════════
#  Templates email
# ══════════════════════════════════════════════════════════

DEFAULT_TEMPLATE = {
    "name": "Concert / Spectacle (Zénith)",
    "subject": "Flash FM : {nom_jeu}",
    "plain": (
        "Bonjour {prenom},\n\n"
        "Suite à votre participation à notre jeu Flash FM « {nom_jeu} », "
        "j'ai le plaisir de vous annoncer que vous avez été tiré au sort "
        "pour remporter 2 invitations !\n\n"
        "Date : {date}\nLieu : {lieu}\n\n"
        "Vous êtes inscrit sur une liste d'invités Flash FM. Vos places sont à retirer "
        "le soir du concert au guichet Invité sur la gauche du Zénith – "
        "Vous devrez vous munir d'une pièce d'identité.\n\n"
        "Il n'est pas nécessaire de prendre la file d'attente du guichet principal, "
        "le guichet « invités » est sur la gauche à l'extérieur du Zénith.\n\n"
        "Merci de me confirmer en retour la bonne réception de ce mail.\n\n"
        "Vous retrouvez chaque semaine nos nouveaux jeux en écoutant Flash FM, "
        "sur notre site flashfm.fr, et sur notre application Flash FM France.\n\n"
        "Merci de votre fidélité et bon spectacle !\n\n"
        "PS : Si un jour vous êtes sollicité par téléphone par l'institut de Sondage "
        "Médiamétrie concernant les audiences radio, n'oubliez pas de répondre que "
        "vous écoutez Flash FM 😉\n\n"
        "Laissez-nous un avis Google : https://g.page/r/CbZFpSFwfcq7EBM/review\n\n"
        "Pascal Thomas\nFLASH FM – Gérant SARL PROXIMEDIA\n"
        "30-32 Cours Gay Lussac – 87000 LIMOGES\n05 55 31 00 00\n"
        "pascal@flashfm.fr | http://www.flashfm.fr\n\n"
        "FLASH FM 1ère radio musicale à LIMOGES et en HAUTE-VIENNE\n"
        "Limoges : 89.9 – Saint-Junien : 98.4 – Brive : 99.9 – Uzerche : 99.9 "
        "– Guéret : 97.7 – Montmorillon : 95\n"
        "DAB + : Haute-Vienne – Corrèze – Vienne"
    ),
    "html": (
        "<html><body style='font-family:Calibri,sans-serif;font-size:12pt;'>"
        "<p>Bonjour <strong>{prenom}</strong>,</p>"
        "<p>Suite à votre participation à notre jeu Flash FM «&nbsp;"
        "<strong>{nom_jeu}</strong>&nbsp;», j'ai le plaisir de vous annoncer que vous "
        "avez été tiré au sort pour remporter 2 invitations&nbsp;!</p>"
        "<p><strong>Date&nbsp;: {date}</strong><br>"
        "<strong>Lieu&nbsp;: {lieu}</strong></p>"
        "<p><strong>Vous êtes inscrit sur une liste d'invités Flash FM.</strong> "
        "Vos places sont à retirer <strong>le soir du concert au guichet Invité sur "
        "la gauche du Zénith</strong> – Vous devrez vous munir d'une pièce "
        "d'identité.</p>"
        "<p>Il n'est pas nécessaire de prendre la file d'attente du guichet principal, "
        "le guichet «&nbsp;invités&nbsp;» est sur la gauche à l'extérieur du "
        "Zénith.</p>"
        "<br><p style='text-align:center;font-size:14pt;color:red;'>"
        "<strong>Merci de me confirmer en retour la bonne réception de ce "
        "mail.</strong></p><br>"
        "<p>Vous retrouvez chaque semaine nos nouveaux jeux en écoutant Flash FM, "
        "sur notre site <a href='http://flashfm.fr'>flashfm.fr</a>, et sur notre "
        "application Flash FM France (gratuite sur tous les stores).</p>"
        "<p>Merci de votre fidélité et bon spectacle&nbsp;!</p><br>"
        "<p style='color:#7030A0;'><em><u>PS</u>&nbsp;: Si un jour vous êtes sollicité "
        "par téléphone par l'institut de Sondage Médiamétrie concernant les audiences "
        "radio, n'oubliez pas de répondre que vous écoutez Flash FM&nbsp;"
        "&#128521;</em></p>"
        "<p style='color:#C00000;'><strong><em>Vous souhaitez nous remercier pour ces "
        "places ou vous appréciez tout simplement Flash FM&nbsp;?</em></strong> "
        "<em><a href='https://g.page/r/CbZFpSFwfcq7EBM/review' "
        "style='color:#C00000;'>Laissez-nous un avis sur notre fiche Google "
        "ici&nbsp;!</a></em></p><br>"
        "<p><strong><em><span style='font-size:14pt;color:#1F497D;'>Pascal "
        "Thomas</span></em></strong><br>"
        "<strong><em><span style='font-size:14pt;color:#1F497D;'>FLASH "
        "FM</span></em></strong><br>"
        "<em><span style='font-size:8pt;color:#7030A0;'>Gérant SARL "
        "PROXIMEDIA</span></em><br>"
        "<em><span style='color:#17365D;'>Directeur</span></em><br>"
        "<span style='color:#1F497D;'>30-32 Cours Gay Lussac</span><br>"
        "<span style='color:#1F497D;'>87000 LIMOGES</span><br>"
        "<span style='color:#4A442A;'>05 55 31 00 00</span></p>"
        "<p><a href='mailto:pascal@flashfm.fr' style='color:purple;'>"
        "pascal@flashfm.fr</a>&nbsp;|&nbsp;"
        "<a href='http://www.flashfm.fr' style='color:purple;'>"
        "http://www.flashfm.fr</a></p><br>"
        "<p><em><span style='color:#1F497D;'>FLASH FM 1<sup>ère</sup> radio musicale "
        "à LIMOGES et en HAUTE-VIENNE</span></em><br>"
        "Limoges&nbsp;: 89.9 – Saint-Junien&nbsp;: 98.4 – Brive&nbsp;: 99.9 – "
        "Uzerche&nbsp;: 99.9 – Guéret&nbsp;: 97.7 – Montmorillon&nbsp;: 95<br>"
        "DAB +&nbsp;: Haute-Vienne – Corrèze – Vienne – Et sur notre application "
        "Flash FM France</p>"
        "</body></html>"
    )
}


def load_templates():
    if os.path.exists(TEMPLATES_PATH):
        with open(TEMPLATES_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return [DEFAULT_TEMPLATE]


def save_templates(tpls):
    with open(TEMPLATES_PATH, 'w', encoding='utf-8') as f:
        json.dump(tpls, f, ensure_ascii=False, indent=2)


def apply_vars(text, variables):
    for k, v in variables.items():
        text = text.replace(f'{{{k}}}', v)
    return text


# ══════════════════════════════════════════════════════════
#  Envoi email
# ══════════════════════════════════════════════════════════

def send_smtp(to_addr, subject, html_body, plain_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"{FROM_NAME} <{FROM_ADDR}>"
    msg["To"]      = to_addr
    msg["Cc"]      = FROM_ADDR
    msg.attach(MIMEText(plain_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body,  "html",  "utf-8"))
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=20) as srv:
        srv.login(SMTP_LOGIN, SMTP_PASSWORD)
        srv.sendmail(FROM_ADDR, [to_addr, FROM_ADDR], msg.as_string())


# ══════════════════════════════════════════════════════════
#  Bouton coloré cross-platform
# ══════════════════════════════════════════════════════════

class ColorButton(tk.Frame):
    """Bouton avec couleur de fond garantie sur Windows, macOS et Linux."""
    def __init__(self, parent, text="", command=None,
                 bg=C_DARK, fg=C_WHITE, font_size=10, padx=14, pady=7):
        super().__init__(parent, bg=bg)
        self._bg  = bg
        self._fg  = fg
        self._cmd = command
        self._on  = True
        self._lbl = tk.Label(self, text=text, bg=bg, fg=fg,
                             font=("", font_size, "bold"),
                             padx=padx, pady=pady, cursor="hand2")
        self._lbl.pack()
        for w in (self, self._lbl):
            w.bind("<ButtonRelease-1>", self._fire)
            w.bind("<Enter>",  lambda e: self._tint(0.82) if self._on else None)
            w.bind("<Leave>",  lambda e: self._tint(1.0))

    def _fire(self, _):
        if self._on and self._cmd:
            self._cmd()

    def _tint(self, factor):
        r, g, b = (int(self._bg[i:i+2], 16) for i in (1, 3, 5))
        c = "#{:02x}{:02x}{:02x}".format(
            min(255, int(r * factor)),
            min(255, int(g * factor)),
            min(255, int(b * factor)))
        self._lbl.config(bg=c)
        tk.Frame.config(self, bg=c)

    def config(self, **kw):
        state = kw.pop("state", None)
        text  = kw.pop("text",  None)
        if state == tk.DISABLED:
            self._on = False
            self._lbl.config(fg="#999999", cursor="")
            tk.Frame.config(self, cursor="")
        elif state == tk.NORMAL:
            self._on = True
            self._lbl.config(fg=self._fg, cursor="hand2")
            tk.Frame.config(self, cursor="hand2")
        if text is not None:
            self._lbl.config(text=text)
        if kw:
            tk.Frame.config(self, **kw)


# ══════════════════════════════════════════════════════════
#  Dialogue Éditeur de modèle
# ══════════════════════════════════════════════════════════

class TemplateEditor(tk.Toplevel):
    def __init__(self, parent, template=None, on_save=None):
        super().__init__(parent)
        self.title("Nouveau modèle" if template is None else "Modifier le modèle")
        self.geometry("760x640")
        self.resizable(True, True)
        self.configure(bg=C_LIGHT)
        self.on_save = on_save
        self.grab_set()

        frm = tk.Frame(self, bg=C_LIGHT)
        frm.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        tk.Label(frm, text="Nom du modèle :", bg=C_LIGHT,
                 font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        self.v_name = tk.StringVar(value=template['name'] if template else "")
        ttk.Entry(frm, textvariable=self.v_name, width=65).grid(
            row=0, column=1, sticky="ew", padx=8)

        tk.Label(frm, text="Sujet :", bg=C_LIGHT,
                 font=("", 10, "bold")).grid(row=1, column=0, sticky="w", pady=4)
        self.v_subject = tk.StringVar(
            value=template['subject'] if template else "Flash FM : {nom_jeu}")
        ttk.Entry(frm, textvariable=self.v_subject, width=65).grid(
            row=1, column=1, sticky="ew", padx=8)

        tk.Label(frm,
                 text="Variables disponibles : {prenom}  {nom}  {nom_jeu}  {date}  {lieu}",
                 bg=C_LIGHT, fg=C_GRAY, font=("", 9)).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(2, 8))

        nb = ttk.Notebook(frm)
        nb.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=4)
        frm.rowconfigure(3, weight=1)
        frm.columnconfigure(1, weight=1)

        html_f = ttk.Frame(nb)
        nb.add(html_f, text="  Corps HTML  ")
        self.w_html = scrolledtext.ScrolledText(
            html_f, wrap=tk.WORD, font=("Courier", 10))
        self.w_html.pack(fill=tk.BOTH, expand=True)
        if template:
            self.w_html.insert("1.0", template.get('html', ''))

        plain_f = ttk.Frame(nb)
        nb.add(plain_f, text="  Corps Texte brut  ")
        self.w_plain = scrolledtext.ScrolledText(
            plain_f, wrap=tk.WORD, font=("Courier", 10))
        self.w_plain.pack(fill=tk.BOTH, expand=True)
        if template:
            self.w_plain.insert("1.0", template.get('plain', ''))

        bf = tk.Frame(frm, bg=C_LIGHT)
        bf.grid(row=4, column=0, columnspan=2, pady=14)
        ColorButton(bf, text="Annuler", command=self.destroy,
                    bg=C_LGRAY, fg=C_DARK, padx=16, pady=6).pack(
            side=tk.RIGHT, padx=6)
        ColorButton(bf, text="✓  Enregistrer", command=self._save,
                    bg=C_GREEN, fg=C_WHITE, font_size=10, padx=16, pady=6).pack(
            side=tk.RIGHT)

    def _save(self):
        name = self.v_name.get().strip()
        if not name:
            messagebox.showwarning("Champ manquant",
                "Saisissez un nom pour le modèle.", parent=self)
            return
        result = {
            "name":    name,
            "subject": self.v_subject.get().strip(),
            "html":    self.w_html.get("1.0",  tk.END).strip(),
            "plain":   self.w_plain.get("1.0", tk.END).strip(),
        }
        if self.on_save:
            self.on_save(result)
        self.destroy()


# ══════════════════════════════════════════════════════════
#  Application principale
# ══════════════════════════════════════════════════════════

class FlashFMApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Flash FM – Gestionnaire de Jeux Concours")
        self.geometry("800x1020")
        self.minsize(700, 800)
        self.configure(bg=C_LIGHT)

        self.templates    = load_templates()
        self.participants = []
        self.winners      = []
        self.creds_path   = tk.StringVar()

        self._build_header()
        self._build_scroll_area()

    # ── En-tête ───────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self, bg=C_DARK)
        h.pack(fill=tk.X)
        tk.Label(h, text="FLASH FM", bg=C_DARK, fg=C_RED,
                 font=("", 22, "bold"), pady=12).pack(side=tk.LEFT, padx=20)
        tk.Label(h, text="Gestionnaire de Jeux Concours",
                 bg=C_DARK, fg=C_WHITE, font=("", 12)).pack(side=tk.LEFT, padx=4)
        ColorButton(h, text="⚽  E-Billets Sports",
                    command=lambda: SportsApp(self),
                    bg="#1A4B8C", fg=C_WHITE, font_size=9).pack(
            side=tk.RIGHT, padx=12)

    # ── Zone défilante ────────────────────────────────────
    def _build_scroll_area(self):
        container = tk.Frame(self, bg=C_LIGHT)
        container.pack(fill=tk.BOTH, expand=True)

        self._canvas = tk.Canvas(container, bg=C_LIGHT, highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._inner  = tk.Frame(self._canvas, bg=C_LIGHT)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self._inner, anchor="nw")

        self._inner.bind("<Configure>", self._on_inner_cfg)
        self._canvas.bind("<Configure>", self._on_canvas_cfg)
        self._canvas.bind_all("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._build_steps()
        tk.Frame(self._inner, bg=C_LIGHT, height=24).pack()

    def _on_inner_cfg(self, _):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_cfg(self, event):
        self._canvas.itemconfig(self._win_id, width=event.width)

    # ── Helpers UI ────────────────────────────────────────
    def _section(self, title, number, color=C_RED):
        outer = tk.Frame(self._inner, bg=C_LIGHT)
        outer.pack(fill=tk.X, padx=18, pady=(12, 2))
        hdr = tk.Frame(outer, bg=color)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"   {number}   {title}", bg=color, fg=C_WHITE,
                 font=("", 11, "bold"), anchor="w", pady=7).pack(fill=tk.X)
        body = tk.Frame(outer, bg=C_WHITE, bd=1, relief=tk.GROOVE)
        body.pack(fill=tk.X)
        inner = tk.Frame(body, bg=C_WHITE, padx=16, pady=12)
        inner.pack(fill=tk.X)
        return inner

    @staticmethod
    def _btn(parent, text, cmd, color=C_DARK, fg=C_WHITE, font_size=10):
        return ColorButton(parent, text=text, command=cmd,
                           bg=color, fg=fg, font_size=font_size)

    # ── Construction des étapes ───────────────────────────
    def _build_steps(self):
        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_step4()
        self._build_step5()
        self._build_step6()
        self._build_log()

    # ① Paramètres du jeu
    def _build_step1(self):
        f = self._section("Paramètres du jeu", "①")
        self.gv = {}
        for i, (lbl, key, ph) in enumerate([
            ("Nom du jeu :",    "nom_jeu", "Les Années 80, La Tournée"),
            ("Date et heure :", "date",    "samedi 6 juin 2026 à 20h00"),
            ("Lieu :",          "lieu",    "Zénith Limoges Métropole"),
        ]):
            tk.Label(f, text=lbl, bg=C_WHITE, width=18, anchor="w",
                     font=("", 10)).grid(row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=ph)
            self.gv[key] = var
            ttk.Entry(f, textvariable=var, width=52).grid(
                row=i, column=1, sticky="ew", padx=8)

        tk.Label(f, text="Nombre de gagnants :", bg=C_WHITE, width=18,
                 anchor="w", font=("", 10)).grid(row=3, column=0, sticky="w", pady=4)
        self.nb_winners = tk.IntVar(value=5)
        ttk.Spinbox(f, from_=1, to=100, textvariable=self.nb_winners,
                    width=6, font=("", 11)).grid(row=3, column=1, sticky="w", padx=8)
        f.columnconfigure(1, weight=1)

    # ② Fichier CSV
    def _build_step2(self):
        f = self._section("Fichier CSV des participants", "②")
        row = tk.Frame(f, bg=C_WHITE)
        row.pack(fill=tk.X)
        self._btn(row, "📂  Parcourir…", self._pick_csv,
                  color=C_DARK).pack(side=tk.LEFT)
        self.lbl_csv = tk.Label(row, text="Aucun fichier sélectionné",
                                bg=C_WHITE, fg=C_GRAY, font=("", 10, "italic"))
        self.lbl_csv.pack(side=tk.LEFT, padx=12)
        self.lbl_csv_info = tk.Label(f, text="", bg=C_WHITE,
                                     fg=C_GREEN, font=("", 10))
        self.lbl_csv_info.pack(anchor="w", pady=(4, 0))

    def _pick_csv(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV",
            filetypes=[("CSV", "*.csv"), ("Tous", "*.*")])
        if not path:
            return
        try:
            self.participants = parse_csv(path)
            self.lbl_csv.config(text=os.path.basename(path),
                                fg=C_DARK, font=("", 10))
            self.lbl_csv_info.config(
                text=f"✓  {len(self.participants)} participants uniques chargés")
            self._log(f"CSV : {len(self.participants)} participants "
                      f"({os.path.basename(path)})")
        except Exception as e:
            messagebox.showerror("Erreur CSV", str(e))

    # ③ Tirage au sort avec vérification d'éligibilité
    def _build_step3(self):
        f = self._section("Tirage au sort", "③")

        # Credentials (partagé avec étape ④)
        creds_row = tk.Frame(f, bg=C_WHITE)
        creds_row.pack(fill=tk.X, pady=(0, 10))
        tk.Label(creds_row, text="Credentials Google :",
                 bg=C_WHITE, font=("", 10)).pack(side=tk.LEFT)
        self._btn(creds_row, "📄  JSON…", self._pick_creds,
                  color="#555", font_size=9).pack(side=tk.LEFT, padx=8)
        self.lbl_creds = tk.Label(creds_row, text="Aucun fichier",
                                   bg=C_WHITE, fg=C_GRAY,
                                   font=("", 10, "italic"))
        self.lbl_creds.pack(side=tk.LEFT)

        # Info vérification
        tk.Label(f,
                 text="ℹ  Le tirage vérifie automatiquement les joueurs bannis "
                      "et les gains récents (< 6 mois) via Google Sheets.",
                 bg=C_WHITE, fg="#666", font=("", 9),
                 wraplength=680, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))

        # Bouton tirage
        self.btn_draw = self._btn(f, "🎲   LANCER LE TIRAGE",
                                   self._do_draw, color=C_RED, font_size=13)
        self.btn_draw.pack(anchor="w", pady=(0, 12))

        # Tableau résultats
        cols = ("nom", "prenom", "ville", "email", "telephone")
        self.tree = ttk.Treeview(f, columns=cols, show="headings",
                                  height=7, selectmode="browse")
        hdrs = {"nom": "Nom", "prenom": "Prénom", "ville": "Ville",
                "email": "Email", "telephone": "Téléphone"}
        wids = {"nom": 120, "prenom": 110, "ville": 130,
                "email": 200, "telephone": 120}
        for c in cols:
            self.tree.heading(c, text=hdrs[c])
            self.tree.column(c, width=wids[c], minwidth=50, anchor="w")

        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sb.pack(side=tk.LEFT, fill=tk.Y)

        self.lbl_draw = tk.Label(f, text="", bg=C_WHITE,
                                  fg=C_GREEN, font=("", 10))
        self.lbl_draw.pack(anchor="w", pady=(8, 0))

    def _pick_creds(self):
        path = filedialog.askopenfilename(
            title="Fichier credentials Google (.json)",
            filetypes=[("JSON", "*.json"), ("Tous", "*.*")])
        if path:
            self.creds_path.set(path)
            self.lbl_creds.config(text=os.path.basename(path),
                                   fg=C_DARK, font=("", 10))
            # Mise à jour du label dans l'étape ④ aussi
            if hasattr(self, 'lbl_creds4'):
                self.lbl_creds4.config(text=os.path.basename(path), fg=C_DARK)

    def _do_draw(self):
        if not self.participants:
            messagebox.showwarning("CSV manquant",
                "Chargez d'abord un fichier CSV (étape ②).")
            return
        n = self.nb_winners.get()

        if not self.creds_path.get():
            if not messagebox.askyesno("Credentials manquants",
                "Aucun fichier credentials Google sélectionné.\n\n"
                "La vérification des éligibilités ne sera PAS effectuée.\n\n"
                "Continuer quand même sans vérification ?"):
                return
            # Tirage simple sans vérification
            pool = [p for p in self.participants if p['email']]
            if n > len(pool):
                messagebox.showwarning("Pas assez de participants",
                    f"Seulement {len(pool)} participants avec email.")
                return
            self.winners = random.sample(pool, n)
            self._update_winners_tree(self.winners)
            self._log(f"Tirage simple (sans vérification) : {n} gagnants.")
            return

        if not GSPREAD_OK:
            messagebox.showerror("Module manquant",
                "Installez gspread :\npip install gspread google-auth")
            return

        # Tirage avec vérification (threaded)
        self.btn_draw.config(state=tk.DISABLED, text="⏳  Tirage en cours…")
        self.lbl_draw.config(text="Vérification des éligibilités en cours…",
                             fg=C_GRAY)
        self.update()

        nom_jeu = self.gv['nom_jeu'].get()

        def run():
            try:
                draw_with_checks(
                    self.participants, n,
                    self.creds_path.get(), nom_jeu,
                    log_cb   = lambda m: self.after(0, lambda msg=m: self._log(msg)),
                    done_cb  = lambda w: self.after(0, lambda wl=w: self._draw_done(wl)),
                    error_cb = lambda m: self.after(0, lambda msg=m:
                                   messagebox.showwarning("Tirage incomplet", msg))
                )
            except Exception as e:
                self.after(0, lambda: self.lbl_draw.config(
                    text=f"✗  Erreur : {e}", fg=C_RED))
                self.after(0, lambda: self._log(f"Erreur tirage : {e}"))
                self.after(0, lambda: self.btn_draw.config(
                    state=tk.NORMAL, text="🎲   LANCER LE TIRAGE"))

        threading.Thread(target=run, daemon=True).start()

    def _draw_done(self, winners):
        self.winners = winners
        self._update_winners_tree(winners)
        self.btn_draw.config(state=tk.NORMAL, text="🎲   LANCER LE TIRAGE")
        self.lbl_draw.config(
            text=f"✓  {len(winners)} gagnant(s) validé(s)", fg=C_GREEN)
        self._log(f"Tirage terminé : {len(winners)} gagnant(s) validé(s).")

    def _update_winners_tree(self, winners):
        self.tree.delete(*self.tree.get_children())
        for w in winners:
            self.tree.insert("", tk.END, values=(
                w['nom'], w['prenom'], w.get('ville', ''),
                w['email'], format_phone(w['phone'])
            ))

    # ④ Export Google Sheets
    def _build_step4(self):
        f = self._section("Export Google Sheets", "④", color="#1A6B35")

        # Affichage du fichier credentials (sélectionné en étape ③)
        row = tk.Frame(f, bg=C_WHITE)
        row.pack(fill=tk.X)
        tk.Label(row, text="Credentials :", bg=C_WHITE,
                 font=("", 10)).pack(side=tk.LEFT)
        self.lbl_creds4 = tk.Label(row,
            text="(sélectionner en étape ③)",
            bg=C_WHITE, fg=C_GRAY, font=("", 10, "italic"))
        self.lbl_creds4.pack(side=tk.LEFT, padx=8)

        self._btn(f, "📊   EXPORTER VERS GOOGLE SHEETS",
                  self._do_sheets, color="#1A6B35",
                  font_size=11).pack(anchor="w", pady=(10, 4))

        self.lbl_sheets = tk.Label(f, text="", bg=C_WHITE,
                                    fg=C_GREEN, font=("", 10))
        self.lbl_sheets.pack(anchor="w")

    def _do_sheets(self):
        if not self.winners:
            messagebox.showwarning("Tirage manquant",
                "Effectuez d'abord le tirage (étape ③).")
            return
        if not self.creds_path.get():
            messagebox.showwarning("Credentials manquants",
                "Sélectionnez le fichier credentials Google (étape ③).")
            return
        if not GSPREAD_OK:
            messagebox.showerror("Module manquant",
                "pip install gspread google-auth")
            return

        nom_jeu   = self.gv['nom_jeu'].get()
        MOIS_FR   = ["janvier","février","mars","avril","mai","juin",
                     "juillet","août","septembre","octobre","novembre","décembre"]
        now       = datetime.now()
        mois_an   = f"{MOIS_FR[now.month - 1]} {now.year}"
        tab_name  = f"{nom_jeu[:45]} {mois_an}"
        game_info = {k: v.get() for k, v in self.gv.items()}

        self.lbl_sheets.config(text="⏳  Export en cours…", fg=C_GRAY)
        self.update()

        def run():
            try:
                url = export_to_sheets(self.creds_path.get(), tab_name,
                                       self.winners, game_info)
                self.after(0, lambda: self.lbl_sheets.config(
                    text=f"✓  Onglet « {tab_name} » créé", fg=C_GREEN))
                self.after(0, lambda: self._log(
                    f"Google Sheets : onglet « {tab_name} » → {url}"))
            except Exception as e:
                self.after(0, lambda: self.lbl_sheets.config(
                    text=f"✗  Erreur : {e}", fg=C_RED))
                self.after(0, lambda: self._log(f"Erreur Sheets : {e}"))

        threading.Thread(target=run, daemon=True).start()

    # ⑤ Modèle d'email
    def _build_step5(self):
        f = self._section("Modèle d'email", "⑤", color=C_BLUE)
        row = tk.Frame(f, bg=C_WHITE)
        row.pack(fill=tk.X)
        tk.Label(row, text="Modèle :", bg=C_WHITE,
                 font=("", 10)).pack(side=tk.LEFT)
        self.tpl_var   = tk.StringVar()
        self.tpl_combo = ttk.Combobox(row, textvariable=self.tpl_var,
                                       state="readonly", width=36, font=("", 10))
        self.tpl_combo.pack(side=tk.LEFT, padx=8)
        self._refresh_templates()
        self._btn(row, "＋ Nouveau",   self._tpl_new,
                  color="#444", font_size=9).pack(side=tk.LEFT, padx=2)
        self._btn(row, "✏  Modifier",  self._tpl_edit,
                  color="#444", font_size=9).pack(side=tk.LEFT, padx=2)
        self._btn(row, "🗑  Supprimer", self._tpl_del,
                  color=C_RED,  font_size=9).pack(side=tk.LEFT, padx=2)

    def _refresh_templates(self):
        names = [t['name'] for t in self.templates]
        self.tpl_combo['values'] = names
        if names and not self.tpl_var.get():
            self.tpl_var.set(names[0])

    def _get_tpl(self):
        name = self.tpl_var.get()
        return next((t for t in self.templates if t['name'] == name), None)

    def _tpl_new(self):
        def on_save(tpl):
            self.templates.append(tpl)
            save_templates(self.templates)
            self._refresh_templates()
            self.tpl_var.set(tpl['name'])
            self._log(f"Modèle « {tpl['name']} » créé.")
        TemplateEditor(self, on_save=on_save)

    def _tpl_edit(self):
        tpl = self._get_tpl()
        if not tpl:
            messagebox.showwarning("Sélection", "Sélectionnez un modèle.")
            return
        def on_save(updated):
            idx = self.templates.index(tpl)
            self.templates[idx] = updated
            save_templates(self.templates)
            self._refresh_templates()
            self.tpl_var.set(updated['name'])
            self._log(f"Modèle « {updated['name']} » mis à jour.")
        TemplateEditor(self, template=tpl, on_save=on_save)

    def _tpl_del(self):
        tpl = self._get_tpl()
        if not tpl:
            return
        if messagebox.askyesno("Supprimer",
                f"Supprimer le modèle « {tpl['name']} » ?"):
            self.templates.remove(tpl)
            save_templates(self.templates)
            self._refresh_templates()
            self._log(f"Modèle « {tpl['name']} » supprimé.")

    # ⑥ Envoi
    def _build_step6(self):
        f = self._section("Envoi des emails", "⑥", color="#6D3B8A")
        row = tk.Frame(f, bg=C_WHITE)
        row.pack(fill=tk.X)
        self._btn(row, f"📧   Test → {TEST_EMAIL}",
                  self._send_test, color=C_BLUE,
                  font_size=11).pack(side=tk.LEFT, padx=(0, 12))
        self._btn(row, "🚀   Envoyer aux gagnants",
                  self._send_all, color=C_GREEN,
                  font_size=11).pack(side=tk.LEFT)
        self.lbl_send = tk.Label(f, text="", bg=C_WHITE,
                                  fg=C_GREEN, font=("", 10))
        self.lbl_send.pack(anchor="w", pady=(8, 0))

    def _variables(self, prenom="Pascal", nom="Thomas"):
        d = {k: v.get() for k, v in self.gv.items()}
        d["prenom"] = prenom
        d["nom"]    = nom
        return d

    def _send_test(self):
        tpl = self._get_tpl()
        if not tpl:
            messagebox.showwarning("Modèle", "Sélectionnez un modèle.")
            return

        # Prépare la liste des mails à envoyer (un par gagnant, ou un seul test)
        if self.winners:
            items = [(self._variables(prenom=w['prenom'], nom=w['nom']), w)
                     for w in self.winners]
        else:
            items = [(self._variables(), None)]

        self.lbl_send.config(
            text=f"⏳  Envoi de {len(items)} mail(s) de test…", fg=C_GRAY)
        self.update()

        def run():
            ok_count, errs = 0, []
            for v, w in items:
                subject = apply_vars(tpl['subject'], v)
                html    = apply_vars(tpl['html'],    v)
                plain   = apply_vars(tpl['plain'],   v)
                # Sujet préfixé pour distinguer les mails de test
                subject = f"[TEST] {subject}"
                try:
                    send_smtp(TEST_EMAIL, subject, html, plain)
                    ok_count += 1
                    name = f"{w['prenom']} {w['nom']}" if w else "test"
                    self.after(0, lambda n=name: self._log(
                        f"✓  Test ({n}) → {TEST_EMAIL}"))
                except Exception as e:
                    errs.append(str(e))
                    self.after(0, lambda err=str(e): self._log(
                        f"✗  Erreur SMTP : {err}"))

            def finish():
                if not errs:
                    self.lbl_send.config(
                        text=f"✓  {ok_count} mail(s) de test envoyé(s) à {TEST_EMAIL}",
                        fg=C_GREEN)
                    messagebox.showinfo("Tests envoyés",
                        f"{ok_count} mail(s) envoyé(s) à {TEST_EMAIL}\n"
                        f"(sujet préfixé [TEST])")
                else:
                    self.lbl_send.config(
                        text=f"✗  Erreur SMTP : {errs[0]}", fg=C_RED)
                    messagebox.showerror("Erreur SMTP", errs[0])
            self.after(0, finish)

        threading.Thread(target=run, daemon=True).start()

    def _send_all(self):
        if not self.winners:
            messagebox.showwarning("Tirage", "Effectuez d'abord le tirage (③).")
            return
        tpl = self._get_tpl()
        if not tpl:
            messagebox.showwarning("Modèle", "Sélectionnez un modèle.")
            return
        no_email = [w for w in self.winners if not w['email']]
        if no_email:
            noms = ", ".join(f"{w['prenom']} {w['nom']}" for w in no_email)
            if not messagebox.askyesno("Emails manquants",
                    f"Sans email :\n{noms}\n\nContinuer ?"):
                return
        if not messagebox.askyesno("Confirmation",
                f"Envoyer {len(self.winners)} emails aux gagnants ?"):
            return

        self.lbl_send.config(text="⏳  Envoi en cours…", fg=C_GRAY)
        self.update()

        def run():
            ok, errs = 0, []
            for w in self.winners:
                if not w['email']:
                    continue
                v       = self._variables(prenom=w['prenom'], nom=w['nom'])
                subject = apply_vars(tpl['subject'], v)
                html    = apply_vars(tpl['html'],    v)
                plain   = apply_vars(tpl['plain'],   v)
                try:
                    send_smtp(w['email'], subject, html, plain)
                    ok += 1
                    ww = dict(w)
                    self.after(0, lambda x=ww: self._log(
                        f"✓  {x['prenom']} {x['nom']} → {x['email']}"))
                except Exception as e:
                    errs.append(w['email'])
                    ww, ee = dict(w), str(e)
                    self.after(0, lambda x=ww, err=ee: self._log(
                        f"✗  {x['email']} : {err}"))
            msg   = f"✓  {ok} email(s) envoyé(s)"
            if errs:
                msg += f"  —  {len(errs)} erreur(s)"
            color = C_GREEN if not errs else C_RED
            self.after(0, lambda: self.lbl_send.config(text=msg, fg=color))
            self.after(0, lambda: self._log(
                f"Envoi : {ok} OK, {len(errs)} erreur(s)"))

        threading.Thread(target=run, daemon=True).start()

    # Journal
    def _build_log(self):
        f = self._section("Journal", "📋", color="#555")
        self.log = scrolledtext.ScrolledText(
            f, height=10, wrap=tk.WORD,
            bg="#1C1C1E", fg="#D4D4D4",
            font=("Courier", 10), insertbackground=C_WHITE)
        self.log.pack(fill=tk.BOTH, expand=True)
        self.log.config(state=tk.DISABLED)

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"[{ts}]  {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)


# ══════════════════════════════════════════════════════════
#  Clubs data & Dropbox utilities (Sports variant)
# ══════════════════════════════════════════════════════════

DEFAULT_ROOT       = "/Users/pascal/Dropbox/Billets"
SPORTS_CONFIG_PATH = os.path.join(BASE_DIR, "flashfm_sports.json")


def load_clubs():
    if os.path.exists(CLUBS_PATH):
        with open(CLUBS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_clubs(clubs):
    with open(CLUBS_PATH, 'w', encoding='utf-8') as f:
        json.dump(clubs, f, ensure_ascii=False, indent=2)


def load_sports_config():
    if os.path.exists(SPORTS_CONFIG_PATH):
        with open(SPORTS_CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_sports_config(cfg):
    with open(SPORTS_CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── OAuth Dropbox (refresh token permanent) ──────────────────────────────────

def dropbox_get_auth_url(app_key):
    """Retourne l'URL d'autorisation OAuth2 Dropbox (offline = refresh token)."""
    return (f"https://www.dropbox.com/oauth2/authorize"
            f"?client_id={app_key}&response_type=code&token_access_type=offline")


def dropbox_exchange_code(app_key, app_secret, code):
    """Échange le code d'autorisation contre access_token + refresh_token."""
    import urllib.request, urllib.parse, urllib.error as _ue, base64, ssl
    _ssl_ctx = ssl.create_default_context()
    _ssl_ctx.check_hostname = False
    _ssl_ctx.verify_mode = ssl.CERT_NONE
    credentials = base64.b64encode(f"{app_key}:{app_secret}".encode()).decode()
    data = urllib.parse.urlencode({
        'code': code.strip(),
        'grant_type': 'authorization_code',
    }).encode()
    req = urllib.request.Request(
        'https://api.dropbox.com/oauth2/token',
        data=data,
        headers={'Authorization': f'Basic {credentials}',
                 'Content-Type': 'application/x-www-form-urlencoded'},
        method='POST')
    with urllib.request.urlopen(req, timeout=20, context=_ssl_ctx) as r:
        return json.loads(r.read())


def dropbox_refresh_access_token(app_key, app_secret, refresh_token):
    """Obtient un nouvel access token à partir du refresh token (ne expire pas)."""
    import urllib.request, urllib.parse, urllib.error as _ue, base64, ssl
    _ssl_ctx = ssl.create_default_context()
    _ssl_ctx.check_hostname = False
    _ssl_ctx.verify_mode = ssl.CERT_NONE
    credentials = base64.b64encode(f"{app_key}:{app_secret}".encode()).decode()
    data = urllib.parse.urlencode({
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token',
    }).encode()
    req = urllib.request.Request(
        'https://api.dropbox.com/oauth2/token',
        data=data,
        headers={'Authorization': f'Basic {credentials}',
                 'Content-Type': 'application/x-www-form-urlencoded'},
        method='POST')
    with urllib.request.urlopen(req, timeout=20, context=_ssl_ctx) as r:
        resp = json.loads(r.read())
    return resp['access_token']


def _get_dropbox_root():
    """Lit ~/.dropbox/info.json pour détecter la racine locale Dropbox."""
    for candidate in (
        os.path.expanduser("~/.dropbox/info.json"),
        os.path.expanduser("~/.dropbox/instance1/info.json"),
    ):
        try:
            with open(candidate, encoding='utf-8') as f:
                data = json.load(f)
            return (data.get('personal', {}).get('path') or
                    data.get('business', {}).get('path'))
        except Exception:
            pass
    # Fallback : remonte le chemin racine pour trouver le dossier "Dropbox"
    parts = os.path.normpath(DEFAULT_ROOT).split(os.sep)
    for i, part in enumerate(parts):
        if part.lower() == 'dropbox':
            joined = os.sep.join(parts[:i + 1])
            return joined if joined else os.sep
    return None


def local_to_dropbox_api_path(local_path):
    """Convertit un chemin local en chemin API Dropbox (/Billets/xxx/...)."""
    dbx_root = _get_dropbox_root()
    if not dbx_root:
        return None
    norm_local = os.path.normpath(local_path)
    norm_root  = os.path.normpath(dbx_root)
    # Comparaison insensible à la casse (macOS case-insensitive)
    if norm_local.lower().startswith(norm_root.lower() + os.sep.lower()):
        rel = norm_local[len(norm_root):]  # garde la casse réelle
        return rel.replace(os.sep, '/')    # commence déjà par '/'
    return None


def get_or_create_dropbox_link(token, api_path):
    """
    Récupère ou crée un lien de partage public Dropbox.
    Lève une exception avec un message clair en cas d'échec.
    """
    import urllib.request, urllib.error as _ue, ssl

    # Contourne les problèmes de certificats SSL sur macOS
    _ssl_ctx = ssl.create_default_context()
    _ssl_ctx.check_hostname = False
    _ssl_ctx.verify_mode = ssl.CERT_NONE

    hdrs = {'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'}

    def _call(endpoint, payload):
        req = urllib.request.Request(
            f'https://api.dropboxapi.com/2/{endpoint}',
            data=json.dumps(payload).encode(),
            headers=hdrs, method='POST')
        with urllib.request.urlopen(req, timeout=20, context=_ssl_ctx) as r:
            return json.loads(r.read())

    # Créer le lien directement ; si 409 = lien déjà existant, l'URL est dans le corps de l'erreur
    try:
        resp = _call('sharing/create_shared_link_with_settings',
                     {'path': api_path,
                      'settings': {'requested_visibility': 'public'}})
        return resp.get('url', '')
    except _ue.HTTPError as e:
        body_txt = e.read().decode(errors='replace')
        if e.code == 409:
            # Dropbox inclut l'URL existante dans le corps de la réponse 409
            try:
                err_body = json.loads(body_txt)
                url = (err_body.get('error', {})
                               .get('shared_link_already_exists', {})
                               .get('metadata', {})
                               .get('url', ''))
                if url:
                    return url
            except Exception:
                pass
            raise RuntimeError(
                f"Lien déjà existant mais URL introuvable dans la réponse : {body_txt}") from None
        raise RuntimeError(
            f"create_shared_link HTTP {e.code} : {body_txt}") from None


# ── Google Sheets (sports) ────────────────────────────────

def extract_sports_winners(all_vals):
    """
    Extrait les paires (nom, prenom) de tous les tableaux d'un onglet sport.
    Ignore les lignes de titres et d'en-têtes.
    """
    SKIP = {'nom', 'prenom', 'prénom', 'ville', 'email', 'telephone',
            'téléphone', 'lien', 'dropbox', 'lien dropbox', 'lien e-billet'}
    winners = []
    for row in all_vals:
        nom    = row[0].strip() if len(row) > 0 else ''
        prenom = row[1].strip() if len(row) > 1 else ''
        if not nom or not prenom:
            continue
        if normalize(nom) in SKIP or normalize(prenom) in SKIP:
            continue
        winners.append((nom, prenom))
    return winners


def find_insert_row(all_vals):
    """
    Retourne la ligne (1-indexée) où commencer le prochain tableau :
    2 lignes vides après la dernière cellule non-vide de la colonne E (index 4).
    """
    last = 0
    for i, row in enumerate(all_vals):
        if len(row) > 4 and row[4].strip():
            last = i + 1
    return last + 3 if last else 1


def sports_export_to_sheet(ws, winners, nom_jeu, date, lieu, winner_urls):
    """Ajoute un tableau de gagnants dans l'onglet sélectionné."""
    all_vals  = ws.get_all_values()
    start_row = find_insert_row(all_vals)

    titre   = f"{nom_jeu}  –  {date}  –  {lieu}"
    headers = ["Nom", "Prénom", "Ville", "Email", "Téléphone", "Lien e-billet"]
    rows = (
        [[titre, "", "", "", "", ""]]
        + [headers]
        + [
            [w['nom'], w['prenom'], w.get('ville', ''),
             w['email'], format_phone(w['phone']),
             winner_urls.get(w['email'], '')]
            for w in winners
        ]
    )
    ws.update(values=rows,
              range_name=f"A{start_row}:F{start_row + len(rows) - 1}")

    sid    = ws.id
    n_data = len(winners)
    _NAVY  = {"red": 0.098, "green": 0.235, "blue": 0.467}   # #192F77 – titre
    _BLUE  = {"red": 0.180, "green": 0.380, "blue": 0.671}   # #2E61AB – en-têtes
    _GRAY  = {"red": 0.750, "green": 0.750, "blue": 0.750}   # bordures données
    _WHITE = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
    _ALT   = {"red": 0.898, "green": 0.937, "blue": 0.988}   # bleu très pâle

    def _border(color, w=1):
        s = {"style": "SOLID", "width": w, "color": color}
        return {"top": s, "bottom": s, "left": s, "right": s}

    ws.spreadsheet.batch_update({"requests": [
        # ── Titre ──────────────────────────────────────────────────────────
        {"mergeCells": {
            "range": {"sheetId": sid,
                      "startRowIndex": start_row - 1, "endRowIndex": start_row,
                      "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"}},
        {"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": start_row - 1, "endRowIndex": start_row,
                      "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _NAVY,
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": _WHITE}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}},
        # ── En-têtes ───────────────────────────────────────────────────────
        {"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": start_row, "endRowIndex": start_row + 1,
                      "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {
                "backgroundColor": _BLUE,
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": _WHITE},
                "borders": _border(_NAVY, 2)}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,borders)"}},
        # ── Données : alignement + bordures (sans couleur de fond = laissé au banding) ─
        {"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": start_row + 1, "endRowIndex": start_row + 1 + n_data,
                      "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontSize": 10},
                "borders": _border(_GRAY)}},
            "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,borders)"}},
        # ── Lignes alternées blanc / bleu pâle ────────────────────────────
        {"addBanding": {
            "bandedRange": {
                "range": {"sheetId": sid,
                          "startRowIndex": start_row + 1, "endRowIndex": start_row + 1 + n_data,
                          "startColumnIndex": 0, "endColumnIndex": 6},
                "rowProperties": {
                    "firstBandColorStyle":  {"rgbColor": _WHITE},
                    "secondBandColorStyle": {"rgbColor": _ALT}}}}},
        # ── Redimensionner colonnes A–E automatiquement (F = lien, largeur libre) ─
        {"autoResizeDimensions": {
            "dimensions": {"sheetId": sid, "dimension": "COLUMNS",
                           "startIndex": 0, "endIndex": 5}}},
    ]})
    return start_row


def sports_draw_with_checks(participants, n, creds_path, ws_name,
                             log_cb, done_cb, error_cb):
    """
    Tirage au sort pour la variante sports.
    Vérifie que le candidat n'a pas déjà gagné cette saison en lisant l'onglet sélectionné.
    N'écrit rien dans Google Sheets (l'export est une étape séparée).
    """
    sh       = _sheets_client(creds_path)
    ws       = sh.worksheet(ws_name)
    all_vals = ws.get_all_values()
    existing = extract_sports_winners(all_vals)
    log_cb(f"Onglet « {ws_name} » : {len(existing)} gagnant(s) trouvé(s) cette saison.")

    pool = [p for p in participants if p['email']]
    random.shuffle(pool)
    winners, excluded = [], 0

    for candidate in pool:
        if len(winners) >= n:
            break
        nom    = normalize(candidate['nom'])
        prenom = normalize(candidate['prenom'])
        already = any(normalize(en) == nom and normalize(ep) == prenom
                      for en, ep in existing)
        if already:
            excluded += 1
            log_cb(f"⚠  {candidate['prenom']} {candidate['nom']}"
                   f" — déjà gagnant cette saison → remplacé")
        else:
            winners.append(candidate)
            log_cb(f"✓  {candidate['prenom']} {candidate['nom']}")

    if len(winners) < n:
        error_cb(f"Seulement {len(winners)}/{n} gagnants éligibles "
                 f"({excluded} déjà gagnant(s) cette saison).")
    done_cb(winners)


# ══════════════════════════════════════════════════════════
#  Application E-Billets Clubs Sportifs
# ══════════════════════════════════════════════════════════

class SportsApp(tk.Toplevel):
    C_SPORT = "#1A4B8C"

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Flash FM – E-Billets Clubs Sportifs")
        self.geometry("840x1020")
        self.minsize(700, 780)
        self.configure(bg=C_LIGHT)
        self.resizable(True, True)

        cfg              = load_sports_config()
        self.clubs       = load_clubs()
        self.templates   = load_templates()
        self.participants = []
        self.winners     = []
        self.winner_urls = {}
        self.creds_path  = tk.StringVar()
        self.v_tab_name  = tk.StringVar()
        self.v_dbx_key     = tk.StringVar(value=cfg.get('dropbox_app_key', ''))
        self.v_dbx_secret  = tk.StringVar(value=cfg.get('dropbox_app_secret', ''))
        self.v_dbx_refresh = tk.StringVar(value=cfg.get('dropbox_refresh_token', ''))

        self._build_header()
        self._build_scroll_area()

    # ── En-tête ───────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self, bg=self.C_SPORT)
        h.pack(fill=tk.X)
        tk.Label(h, text="FLASH FM", bg=self.C_SPORT, fg=C_RED,
                 font=("", 20, "bold"), pady=10).pack(side=tk.LEFT, padx=20)
        tk.Label(h, text="E-Billets Clubs Sportifs",
                 bg=self.C_SPORT, fg=C_WHITE, font=("", 12)).pack(side=tk.LEFT)

    # ── Zone défilante ────────────────────────────────────
    def _build_scroll_area(self):
        container = tk.Frame(self, bg=C_LIGHT)
        container.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container, bg=C_LIGHT, highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._inner = tk.Frame(canvas, bg=C_LIGHT)
        win_id = canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        self._build_steps()
        tk.Frame(self._inner, bg=C_LIGHT, height=24).pack()

    def _section(self, title, number):
        outer = tk.Frame(self._inner, bg=C_LIGHT)
        outer.pack(fill=tk.X, padx=18, pady=(10, 2))
        hdr = tk.Frame(outer, bg=self.C_SPORT)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"   {number}   {title}", bg=self.C_SPORT,
                 fg=C_WHITE, font=("", 11, "bold"), anchor="w",
                 pady=7).pack(fill=tk.X)
        body = tk.Frame(outer, bg=C_WHITE, bd=1, relief=tk.GROOVE)
        body.pack(fill=tk.X)
        inner = tk.Frame(body, bg=C_WHITE, padx=16, pady=12)
        inner.pack(fill=tk.X)
        return inner

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.insert(tk.END, f"[{ts}]  {msg}\n")
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def _build_steps(self):
        self._build_step1()
        self._build_step2()
        self._build_step3()
        self._build_step4()
        self._build_step5()
        self._build_step6()
        self._build_step7()
        self._build_step8()
        self._build_log()

    # ① Paramètres du match
    def _build_step1(self):
        f = self._section("Paramètres du match", "①")
        self.gv = {}
        for i, (lbl, key, ph) in enumerate([
            ("Nom du jeu :",    "nom_jeu", "Limoges CSP vs Adversaire"),
            ("Date et heure :", "date",    "samedi 6 juin 2026 à 20h00"),
            ("Lieu :",          "lieu",    "Palais des Sports de Beaublanc"),
        ]):
            tk.Label(f, text=lbl, bg=C_WHITE, width=18, anchor="w",
                     font=("", 10)).grid(row=i, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=ph)
            self.gv[key] = var
            ttk.Entry(f, textvariable=var, width=52).grid(
                row=i, column=1, sticky="ew", padx=8)
        tk.Label(f, text="Nombre de gagnants :", bg=C_WHITE, width=18,
                 anchor="w", font=("", 10)).grid(row=3, column=0, sticky="w", pady=4)
        self.nb_winners = tk.IntVar(value=5)
        ttk.Spinbox(f, from_=1, to=100, textvariable=self.nb_winners,
                    width=6, font=("", 11)).grid(row=3, column=1, sticky="w", padx=8)
        f.columnconfigure(1, weight=1)

    # ② CSV participants
    def _build_step2(self):
        f = self._section("Fichier CSV des participants", "②")
        row = tk.Frame(f, bg=C_WHITE)
        row.pack(fill=tk.X)
        ColorButton(row, text="📂  Parcourir…", command=self._pick_csv,
                    bg=C_DARK).pack(side=tk.LEFT)
        self.lbl_csv = tk.Label(row, text="Aucun fichier sélectionné",
                                bg=C_WHITE, fg=C_GRAY, font=("", 10, "italic"))
        self.lbl_csv.pack(side=tk.LEFT, padx=12)
        self.lbl_csv_info = tk.Label(f, text="", bg=C_WHITE,
                                     fg=C_GREEN, font=("", 10))
        self.lbl_csv_info.pack(anchor="w", pady=(4, 0))

    def _pick_csv(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier CSV",
            filetypes=[("CSV", "*.csv"), ("Tous", "*.*")])
        if not path:
            return
        try:
            self.participants = parse_csv(path)
            self.lbl_csv.config(text=os.path.basename(path),
                                fg=C_DARK, font=("", 10))
            self.lbl_csv_info.config(
                text=f"✓  {len(self.participants)} participants uniques chargés")
        except Exception as e:
            messagebox.showerror("Erreur CSV", str(e), parent=self)

    # ③ Configuration
    def _build_step3(self):
        f = self._section("Configuration Google Sheets & Dropbox", "③")

        # Credentials Google
        cr = tk.Frame(f, bg=C_WHITE)
        cr.pack(fill=tk.X, pady=(0, 8))
        tk.Label(cr, text="Credentials Google :", bg=C_WHITE,
                 font=("", 10)).pack(side=tk.LEFT)
        ColorButton(cr, text="📄  JSON…", command=self._pick_creds,
                    bg="#555", font_size=9).pack(side=tk.LEFT, padx=8)
        self.lbl_creds = tk.Label(cr, text="Aucun fichier",
                                   bg=C_WHITE, fg=C_GRAY, font=("", 10, "italic"))
        self.lbl_creds.pack(side=tk.LEFT)

        # Onglet Google Sheets
        tab_row = tk.Frame(f, bg=C_WHITE)
        tab_row.pack(fill=tk.X, pady=(0, 10))
        tk.Label(tab_row, text="Onglet Sheets :", bg=C_WHITE,
                 font=("", 10)).pack(side=tk.LEFT)
        self.cb_tab = ttk.Combobox(tab_row, textvariable=self.v_tab_name,
                                    state="readonly", width=28, font=("", 10))
        self.cb_tab.pack(side=tk.LEFT, padx=8)
        ColorButton(tab_row, text="⟳  Charger les onglets",
                    command=self._load_tabs,
                    bg=self.C_SPORT, font_size=9).pack(side=tk.LEFT)

        # Connexion Dropbox (OAuth2 permanent)
        sep = tk.Frame(f, bg=C_LGRAY, height=1)
        sep.pack(fill=tk.X, pady=(4, 8))
        tk.Label(f, text="Connexion Dropbox (token permanent) :",
                 bg=C_WHITE, font=("", 10, "bold")).pack(anchor="w")
        tk.Label(f,
                 text="Sur https://www.dropbox.com/developers/apps → votre app → "
                      "onglet Settings → copiez App key et App secret.",
                 bg=C_WHITE, fg=C_GRAY, font=("", 9),
                 wraplength=680, justify=tk.LEFT).pack(anchor="w", pady=(2, 6))

        row_key = tk.Frame(f, bg=C_WHITE)
        row_key.pack(fill=tk.X, pady=2)
        tk.Label(row_key, text="App key :", bg=C_WHITE, width=12,
                 anchor="w", font=("", 10)).pack(side=tk.LEFT)
        ttk.Entry(row_key, textvariable=self.v_dbx_key,
                  width=36, font=("", 9)).pack(side=tk.LEFT)

        row_sec = tk.Frame(f, bg=C_WHITE)
        row_sec.pack(fill=tk.X, pady=2)
        tk.Label(row_sec, text="App secret :", bg=C_WHITE, width=12,
                 anchor="w", font=("", 10)).pack(side=tk.LEFT)
        ttk.Entry(row_sec, textvariable=self.v_dbx_secret,
                  width=36, show="*", font=("", 9)).pack(side=tk.LEFT)

        row_auth = tk.Frame(f, bg=C_WHITE)
        row_auth.pack(fill=tk.X, pady=(6, 2))
        ColorButton(row_auth, text="🌐  Étape 1 : Ouvrir l'autorisation Dropbox",
                    command=self._dbx_open_auth,
                    bg=self.C_SPORT, font_size=9).pack(side=tk.LEFT)

        row_code = tk.Frame(f, bg=C_WHITE)
        row_code.pack(fill=tk.X, pady=2)
        tk.Label(row_code, text="Code obtenu :", bg=C_WHITE, width=14,
                 anchor="w", font=("", 10)).pack(side=tk.LEFT)
        self.v_dbx_code = tk.StringVar()
        ttk.Entry(row_code, textvariable=self.v_dbx_code,
                  width=34, font=("", 9)).pack(side=tk.LEFT)
        ColorButton(row_code, text="✅  Étape 2 : Valider",
                    command=self._dbx_validate_code,
                    bg="#2e7d32", font_size=9).pack(side=tk.LEFT, padx=8)

        self.lbl_dbx_status = tk.Label(f, text="", bg=C_WHITE,
                                        font=("", 9, "italic"))
        self.lbl_dbx_status.pack(anchor="w", pady=(2, 0))
        self._dbx_refresh_status()

        # Dossier racine (affiché, non modifiable — hardcodé)
        tk.Frame(f, bg=C_LGRAY, height=1).pack(fill=tk.X, pady=(8, 4))
        tk.Label(f, text=f"Dossier racine : {DEFAULT_ROOT}",
                 bg=C_WHITE, fg=C_GRAY, font=("", 9, "italic")).pack(anchor="w")

    def _pick_creds(self):
        path = filedialog.askopenfilename(
            title="Credentials Google (.json)",
            filetypes=[("JSON", "*.json"), ("Tous", "*.*")])
        if path:
            self.creds_path.set(path)
            self.lbl_creds.config(text=os.path.basename(path),
                                   fg=C_DARK, font=("", 10))

    def _dbx_refresh_status(self):
        """Met à jour le label de statut Dropbox."""
        rt = self.v_dbx_refresh.get().strip()
        if rt:
            self.lbl_dbx_status.config(
                text="✅  Connexion Dropbox configurée (token permanent actif).",
                fg=C_GREEN)
        else:
            self.lbl_dbx_status.config(
                text="⚠  Pas encore autorisé — suivez les étapes ci-dessus.",
                fg="#e65100")

    def _dbx_open_auth(self):
        app_key = self.v_dbx_key.get().strip()
        if not app_key:
            messagebox.showwarning("App key manquante",
                "Saisissez l'App key Dropbox avant de continuer.", parent=self)
            return
        import webbrowser
        url = dropbox_get_auth_url(app_key)
        webbrowser.open(url)
        self._log("Navigateur ouvert → autorisez l'application puis collez le code.")

    def _dbx_validate_code(self):
        app_key    = self.v_dbx_key.get().strip()
        app_secret = self.v_dbx_secret.get().strip()
        code       = self.v_dbx_code.get().strip()
        if not (app_key and app_secret and code):
            messagebox.showwarning("Champs manquants",
                "App key, App secret et code sont requis.", parent=self)
            return
        try:
            resp = dropbox_exchange_code(app_key, app_secret, code)
        except Exception as e:
            messagebox.showerror("Erreur OAuth", str(e), parent=self)
            return
        refresh_token = resp.get('refresh_token', '')
        if not refresh_token:
            messagebox.showerror("Erreur OAuth",
                f"Pas de refresh_token dans la réponse :\n{resp}", parent=self)
            return
        self.v_dbx_refresh.set(refresh_token)
        cfg = load_sports_config()
        cfg['dropbox_app_key']      = app_key
        cfg['dropbox_app_secret']   = app_secret
        cfg['dropbox_refresh_token'] = refresh_token
        cfg.pop('dropbox_token', None)
        save_sports_config(cfg)
        self.v_dbx_code.set('')
        self._dbx_refresh_status()
        self._log("✅  Token permanent Dropbox enregistré. Plus besoin de renouveler.")

    def _get_dropbox_access_token(self):
        """Retourne un access token frais à partir du refresh token."""
        key     = self.v_dbx_key.get().strip()
        secret  = self.v_dbx_secret.get().strip()
        refresh = self.v_dbx_refresh.get().strip()
        if not (key and secret and refresh):
            return None
        return dropbox_refresh_access_token(key, secret, refresh)

    def _load_tabs(self):
        if not self.creds_path.get():
            messagebox.showwarning("Credentials manquants",
                "Sélectionnez d'abord les credentials Google.", parent=self)
            return
        if not GSPREAD_OK:
            messagebox.showerror("Module manquant",
                "pip install gspread google-auth", parent=self)
            return
        try:
            sh   = _sheets_client(self.creds_path.get())
            tabs = [ws.title for ws in sh.worksheets()]
            self.cb_tab['values'] = tabs
            if tabs:
                if self.v_tab_name.get() not in tabs:
                    self.v_tab_name.set(tabs[0])
            self._log(f"{len(tabs)} onglet(s) chargé(s).")
        except Exception as e:
            messagebox.showerror("Erreur Sheets", str(e), parent=self)

    # ④ Clubs partenaires & adversaires
    def _build_step4(self):
        f = self._section("Clubs partenaires & adversaires", "④")

        tk.Label(f, text="Club partenaire :", bg=C_WHITE,
                 font=("", 10, "bold")).pack(anchor="w")
        pr = tk.Frame(f, bg=C_WHITE)
        pr.pack(fill=tk.X, pady=(4, 2))
        self.v_partner = tk.StringVar()
        self.cb_partner = ttk.Combobox(pr, textvariable=self.v_partner,
                                        state="readonly", width=28, font=("", 10))
        self.cb_partner.pack(side=tk.LEFT)
        self.cb_partner.bind("<<ComboboxSelected>>",
                             lambda e: self._on_partner_change())
        ColorButton(pr, text="＋ Nouveau",  command=self._add_partner,
                    bg="#444", font_size=9).pack(side=tk.LEFT, padx=(8, 2))
        ColorButton(pr, text="✏ Renommer", command=self._rename_partner,
                    bg="#444", font_size=9).pack(side=tk.LEFT, padx=2)
        ColorButton(pr, text="🗑 Supprimer", command=self._del_partner,
                    bg=C_RED, font_size=9).pack(side=tk.LEFT, padx=2)
        self.lbl_partner_path = tk.Label(f, text="", bg=C_WHITE,
                                          fg=C_GRAY, font=("", 9, "italic"))
        self.lbl_partner_path.pack(anchor="w", pady=(2, 8))

        tk.Label(f, text="Club adverse :", bg=C_WHITE,
                 font=("", 10, "bold")).pack(anchor="w")
        ar = tk.Frame(f, bg=C_WHITE)
        ar.pack(fill=tk.X, pady=(4, 2))
        self.v_opposing = tk.StringVar()
        self.cb_opposing = ttk.Combobox(ar, textvariable=self.v_opposing,
                                         state="readonly", width=28, font=("", 10))
        self.cb_opposing.pack(side=tk.LEFT)
        self.cb_opposing.bind("<<ComboboxSelected>>",
                              lambda e: self._update_tickets_count())
        ColorButton(ar, text="＋ Nouveau",  command=self._add_opposing,
                    bg="#444", font_size=9).pack(side=tk.LEFT, padx=(8, 2))
        ColorButton(ar, text="✏ Renommer", command=self._rename_opposing,
                    bg="#444", font_size=9).pack(side=tk.LEFT, padx=2)
        ColorButton(ar, text="🗑 Supprimer", command=self._del_opposing,
                    bg=C_RED, font_size=9).pack(side=tk.LEFT, padx=2)
        self.lbl_opposing_path = tk.Label(f, text="", bg=C_WHITE,
                                           fg=C_GRAY, font=("", 9, "italic"))
        self.lbl_opposing_path.pack(anchor="w", pady=(2, 4))
        self.lbl_tickets_found = tk.Label(f, text="", bg=C_WHITE,
                                           fg=C_GRAY, font=("", 10))
        self.lbl_tickets_found.pack(anchor="w")

        self._refresh_partners()

    # ── Gestion clubs ─────────────────────────────────────
    def _refresh_partners(self):
        names = sorted(self.clubs.keys())
        self.cb_partner['values'] = names
        if names:
            if self.v_partner.get() not in names:
                self.v_partner.set(names[0])
            self._on_partner_change()

    def _on_partner_change(self):
        partner = self.v_partner.get()
        if partner:
            path = os.path.join(DEFAULT_ROOT, partner)
            self.lbl_partner_path.config(
                text=f"📁  {path}",
                fg=C_DARK if os.path.isdir(path) else C_RED)
        else:
            self.lbl_partner_path.config(text="")
        self._refresh_opposing()

    def _refresh_opposing(self):
        partner   = self.v_partner.get()
        opponents = sorted(self.clubs.get(partner, []))
        self.cb_opposing['values'] = opponents
        if opponents:
            if self.v_opposing.get() not in opponents:
                self.v_opposing.set(opponents[0])
        else:
            self.v_opposing.set("")
        self._update_tickets_count()

    def _update_tickets_count(self):
        partner  = self.v_partner.get()
        opposing = self.v_opposing.get()
        if not (partner and opposing):
            self.lbl_opposing_path.config(text="")
            self.lbl_tickets_found.config(text="")
            return
        path = os.path.join(DEFAULT_ROOT, partner, opposing)
        self.lbl_opposing_path.config(
            text=f"📁  {path}",
            fg=C_DARK if os.path.isdir(path) else C_RED)
        if os.path.isdir(path):
            n    = len(self._list_tickets(path))
            need = self.nb_winners.get() * 2
            self.lbl_tickets_found.config(
                text=f"{'✓' if n >= need else '⚠'}  {n} e-billet(s) disponible(s)"
                     f"  (besoin : {need})",
                fg=C_GREEN if n >= need else C_RED)
        else:
            self.lbl_tickets_found.config(text="⚠  Dossier non trouvé", fg=C_RED)

    @staticmethod
    def _list_tickets(folder):
        try:
            return sorted(
                f for f in os.listdir(folder)
                if os.path.isfile(os.path.join(folder, f))
                and not f.startswith('.')
            )
        except OSError:
            return []

    def _ask_name(self, title, label, default=""):
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.geometry("340x130")
        dlg.resizable(False, False)
        dlg.configure(bg=C_LIGHT)
        dlg.grab_set()
        result = [None]
        tk.Label(dlg, text=label, bg=C_LIGHT, font=("", 10)).pack(pady=(16, 4))
        var = tk.StringVar(value=default)
        e   = ttk.Entry(dlg, textvariable=var, width=36, font=("", 10))
        e.pack()
        e.focus()
        e.select_range(0, tk.END)
        def ok():
            result[0] = var.get().strip()
            dlg.destroy()
        e.bind("<Return>", lambda _: ok())
        bf = tk.Frame(dlg, bg=C_LIGHT)
        bf.pack(pady=12)
        ColorButton(bf, text="OK", command=ok,
                    bg=self.C_SPORT).pack(side=tk.LEFT, padx=6)
        ColorButton(bf, text="Annuler", command=dlg.destroy,
                    bg=C_LGRAY, fg=C_DARK).pack(side=tk.LEFT)
        dlg.wait_window()
        return result[0]

    def _add_partner(self):
        name = self._ask_name("Nouveau club partenaire", "Nom du club :")
        if not name:
            return
        if name in self.clubs:
            messagebox.showwarning("Existant", f"« {name} » existe déjà.",
                                   parent=self)
            return
        self.clubs[name] = []
        save_clubs(self.clubs)
        os.makedirs(os.path.join(DEFAULT_ROOT, name), exist_ok=True)
        self.v_partner.set(name)
        self._refresh_partners()
        self._log(f"Club partenaire ajouté : {name}")

    def _rename_partner(self):
        old = self.v_partner.get()
        if not old:
            return
        new = self._ask_name("Renommer le club", "Nouveau nom :", default=old)
        if not new or new == old:
            return
        self.clubs[new] = self.clubs.pop(old)
        save_clubs(self.clubs)
        src = os.path.join(DEFAULT_ROOT, old)
        dst = os.path.join(DEFAULT_ROOT, new)
        if os.path.isdir(src):
            os.rename(src, dst)
        self.v_partner.set(new)
        self._refresh_partners()
        self._log(f"Club partenaire renommé : {old} → {new}")

    def _del_partner(self):
        name = self.v_partner.get()
        if not name:
            return
        if not messagebox.askyesno("Supprimer",
                f"Supprimer le club « {name} » ?\n"
                "(Le dossier local n'est PAS supprimé.)", parent=self):
            return
        self.clubs.pop(name, None)
        save_clubs(self.clubs)
        self._refresh_partners()
        self._log(f"Club partenaire supprimé : {name}")

    def _add_opposing(self):
        partner = self.v_partner.get()
        if not partner:
            messagebox.showwarning("Club manquant",
                "Sélectionnez d'abord un club partenaire.", parent=self)
            return
        name = self._ask_name("Nouveau club adverse", "Nom du club adverse :")
        if not name:
            return
        if name in self.clubs.get(partner, []):
            messagebox.showwarning("Existant", f"« {name} » existe déjà.",
                                   parent=self)
            return
        self.clubs.setdefault(partner, []).append(name)
        save_clubs(self.clubs)
        folder = os.path.join(DEFAULT_ROOT, partner, name)
        os.makedirs(folder, exist_ok=True)
        self._refresh_opposing()
        self.v_opposing.set(name)
        self._log(f"Club adverse ajouté : {partner} / {name} — dossier créé")

    def _rename_opposing(self):
        partner = self.v_partner.get()
        old     = self.v_opposing.get()
        if not (partner and old):
            return
        new = self._ask_name("Renommer le club adverse", "Nouveau nom :", default=old)
        if not new or new == old:
            return
        lst = self.clubs.get(partner, [])
        if old in lst:
            lst[lst.index(old)] = new
        save_clubs(self.clubs)
        src = os.path.join(DEFAULT_ROOT, partner, old)
        dst = os.path.join(DEFAULT_ROOT, partner, new)
        if os.path.isdir(src):
            os.rename(src, dst)
        self._refresh_opposing()
        self.v_opposing.set(new)
        self._log(f"Club adverse renommé : {old} → {new}")

    def _del_opposing(self):
        partner  = self.v_partner.get()
        opposing = self.v_opposing.get()
        if not (partner and opposing):
            return
        if not messagebox.askyesno("Supprimer",
                f"Supprimer « {opposing} » ?\n"
                "(Le dossier local n'est PAS supprimé.)", parent=self):
            return
        lst = self.clubs.get(partner, [])
        if opposing in lst:
            lst.remove(opposing)
        save_clubs(self.clubs)
        self._refresh_opposing()
        self._log(f"Club adverse supprimé : {opposing}")

    # ⑤ Tirage au sort
    def _build_step5(self):
        f = self._section("Tirage au sort — vérification saison", "⑤")
        tk.Label(f,
                 text="ℹ  Vérifie que chaque candidat n'a pas déjà gagné cette "
                      "saison en lisant l'onglet Google Sheets sélectionné en ③. "
                      "Les gagnants ne sont PAS écrits dans « liste des gagnants ».",
                 bg=C_WHITE, fg="#666", font=("", 9),
                 wraplength=680, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))

        self.btn_draw = ColorButton(f, text="🎲   LANCER LE TIRAGE",
                                     command=self._do_draw,
                                     bg=C_RED, font_size=13)
        self.btn_draw.pack(anchor="w", pady=(0, 12))

        cols = ("nom", "prenom", "ville", "email", "telephone")
        self.tree = ttk.Treeview(f, columns=cols, show="headings",
                                  height=6, selectmode="browse")
        for c, h, w in zip(cols,
                           ["Nom", "Prénom", "Ville", "Email", "Téléphone"],
                           [120, 110, 120, 200, 115]):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, minwidth=50, anchor="w")
        sb = ttk.Scrollbar(f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sb.pack(side=tk.LEFT, fill=tk.Y)
        self.lbl_draw = tk.Label(f, text="", bg=C_WHITE, fg=C_GREEN, font=("", 10))
        self.lbl_draw.pack(anchor="w", pady=(8, 0))

    def _do_draw(self):
        if not self.participants:
            messagebox.showwarning("CSV manquant",
                "Chargez d'abord un fichier CSV (étape ②).", parent=self)
            return
        n = self.nb_winners.get()

        if not self.creds_path.get() or not self.v_tab_name.get():
            if not messagebox.askyesno("Sans vérification",
                    "Credentials Google ou onglet non configurés.\n"
                    "Tirage simple sans vérification de la saison ?", parent=self):
                return
            pool = [p for p in self.participants if p['email']]
            self.winners     = random.sample(pool, min(n, len(pool)))
            self.winner_urls = {}
            self._update_tree()
            self.lbl_draw.config(
                text=f"✓  {len(self.winners)} gagnant(s) (sans vérification)",
                fg=C_GREEN)
            return

        if not GSPREAD_OK:
            messagebox.showerror("Module manquant",
                "pip install gspread google-auth", parent=self)
            return

        self.btn_draw.config(state=tk.DISABLED, text="⏳  Tirage en cours…")
        self.lbl_draw.config(text="Lecture de l'onglet…", fg=C_GRAY)
        self.update()

        def run():
            try:
                sports_draw_with_checks(
                    self.participants, n,
                    self.creds_path.get(), self.v_tab_name.get(),
                    log_cb   = lambda m: self.after(0, lambda msg=m: self._log(msg)),
                    done_cb  = lambda w: self.after(0, lambda wl=w: self._draw_done(wl)),
                    error_cb = lambda m: self.after(0, lambda msg=m:
                                   messagebox.showwarning("Tirage incomplet", msg,
                                                          parent=self))
                )
            except Exception as e:
                self.after(0, lambda: self.lbl_draw.config(
                    text=f"✗  Erreur : {e}", fg=C_RED))
                self.after(0, lambda: self.btn_draw.config(
                    state=tk.NORMAL, text="🎲   LANCER LE TIRAGE"))

        threading.Thread(target=run, daemon=True).start()

    def _draw_done(self, winners):
        self.winners     = winners
        self.winner_urls = {}
        self._update_tree()
        self.btn_draw.config(state=tk.NORMAL, text="🎲   LANCER LE TIRAGE")
        self.lbl_draw.config(
            text=f"✓  {len(winners)} gagnant(s) validé(s)", fg=C_GREEN)
        self._log(f"Tirage terminé : {len(winners)} gagnant(s).")

    def _update_tree(self):
        self.tree.delete(*self.tree.get_children())
        for w in self.winners:
            self.tree.insert("", tk.END, values=(
                w['nom'], w['prenom'], w.get('ville', ''),
                w['email'], format_phone(w['phone'])
            ))

    # ⑥ Distribution des e-billets
    def _build_step6(self):
        f = self._section("Distribution des e-billets", "⑥")
        tk.Label(f,
                 text="Crée un sous-dossier par gagnant, déplace 2 e-billets et "
                      "récupère automatiquement le lien Dropbox via l'API.",
                 bg=C_WHITE, fg="#666", font=("", 9),
                 wraplength=680, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))
        self.btn_distrib = ColorButton(
            f, text="📂   CRÉER LES DOSSIERS & DISTRIBUER LES BILLETS",
            command=self._do_distribute,
            bg=self.C_SPORT, font_size=11)
        self.btn_distrib.pack(anchor="w", pady=(0, 6))
        self.lbl_distrib = tk.Label(f, text="", bg=C_WHITE,
                                     fg=C_GREEN, font=("", 10))
        self.lbl_distrib.pack(anchor="w")

    def _do_distribute(self):
        if not self.winners:
            messagebox.showwarning("Tirage manquant",
                "Effectuez d'abord le tirage (étape ⑤).", parent=self)
            return
        partner  = self.v_partner.get()
        opposing = self.v_opposing.get()
        token    = self._get_dropbox_access_token()
        if not (partner and opposing):
            messagebox.showwarning("Clubs manquants",
                "Sélectionnez club partenaire et club adverse (étape ④).",
                parent=self)
            return

        tickets_dir = os.path.join(DEFAULT_ROOT, partner, opposing)
        if not os.path.isdir(tickets_dir):
            messagebox.showerror("Dossier introuvable",
                f"Dossier introuvable :\n{tickets_dir}", parent=self)
            return

        tickets = self._list_tickets(tickets_dir)
        needed  = len(self.winners) * 2
        if len(tickets) < needed:
            messagebox.showerror("E-billets insuffisants",
                f"{needed} billets requis pour {len(self.winners)} gagnant(s), "
                f"seulement {len(tickets)} trouvé(s) dans :\n{tickets_dir}",
                parent=self)
            return

        if not messagebox.askyesno("Confirmation",
                f"Créer {len(self.winners)} dossiers et déplacer "
                f"{needed} e-billets ?\n\nDossier : {tickets_dir}", parent=self):
            return

        self.btn_distrib.config(state=tk.DISABLED, text="⏳  Distribution…")
        self.lbl_distrib.config(text="Distribution en cours…", fg=C_GRAY)
        self.update()

        def run():
            errors = []
            for i, winner in enumerate(self.winners):
                folder_name = f"{winner['prenom']} {winner['nom']}"
                winner_dir  = os.path.join(tickets_dir, folder_name)
                try:
                    os.makedirs(winner_dir, exist_ok=True)
                    for j in range(2):
                        shutil.move(
                            os.path.join(tickets_dir, tickets[i * 2 + j]),
                            os.path.join(winner_dir,  tickets[i * 2 + j]))
                    self.after(0, lambda m=f"✓  {folder_name} — billets déplacés": self._log(m))
                except Exception as e:
                    errors.append(folder_name)
                    self.after(0, lambda m=f"✗  {folder_name} : {e}": self._log(m))
                    continue

                # Récupération du lien Dropbox (étape séparée — n'annule pas le déplacement)
                url = ''
                if token:
                    try:
                        api_path = local_to_dropbox_api_path(winner_dir)
                        if api_path:
                            # Retry jusqu'à 5 fois (Dropbox peut prendre du temps à synchroniser)
                            for attempt in range(1, 6):
                                try:
                                    url = get_or_create_dropbox_link(token, api_path)
                                    self.after(0, lambda m=f"   🔗 {folder_name} → {url}": self._log(m))
                                    break
                                except RuntimeError as e_retry:
                                    if 'not_found' in str(e_retry) and attempt < 5:
                                        self.after(0, lambda m=f"   ⏳ Dropbox sync en cours, tentative {attempt}/5…": self._log(m))
                                        time.sleep(4)
                                    else:
                                        raise
                        else:
                            dbx_root = _get_dropbox_root()
                            self.after(0, lambda m=f"   ⚠ {folder_name} : chemin Dropbox introuvable (root={dbx_root!r})": self._log(m))
                    except Exception as e_dbx:
                        self.after(0, lambda m=f"   ✗ Dropbox {folder_name} : {e_dbx}": self._log(m))
                self.winner_urls[winner['email']] = url

            def finish():
                self.btn_distrib.config(state=tk.NORMAL,
                    text="📂   CRÉER LES DOSSIERS & DISTRIBUER LES BILLETS")
                if errors:
                    self.lbl_distrib.config(
                        text=f"✗  {len(errors)} erreur(s) — voir journal", fg=C_RED)
                else:
                    self.lbl_distrib.config(
                        text=f"✓  {len(self.winners)} dossier(s) créé(s), "
                             f"billets distribués", fg=C_GREEN)
            self.after(0, finish)

        threading.Thread(target=run, daemon=True).start()

    # ⑦ Export Google Sheets
    def _build_step7(self):
        f = self._section("Export vers Google Sheets", "⑦")
        tk.Label(f,
                 text="Ajoute le tableau des gagnants dans l'onglet sélectionné en ③, "
                      "à la suite des tableaux existants.",
                 bg=C_WHITE, fg="#666", font=("", 9),
                 wraplength=680, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))
        ColorButton(f, text="📊   EXPORTER VERS GOOGLE SHEETS",
                    command=self._do_export_sheet,
                    bg="#1A6B35", font_size=11).pack(anchor="w", pady=(0, 6))
        self.lbl_sheets = tk.Label(f, text="", bg=C_WHITE,
                                    fg=C_GREEN, font=("", 10))
        self.lbl_sheets.pack(anchor="w")

    def _do_export_sheet(self):
        if not self.winners:
            messagebox.showwarning("Tirage manquant",
                "Effectuez d'abord le tirage (étape ⑤).", parent=self)
            return
        if not self.creds_path.get() or not self.v_tab_name.get():
            messagebox.showwarning("Configuration manquante",
                "Configurez les credentials et l'onglet (étape ③).", parent=self)
            return
        if not GSPREAD_OK:
            messagebox.showerror("Module manquant",
                "pip install gspread google-auth", parent=self)
            return

        self.lbl_sheets.config(text="⏳  Export en cours…", fg=C_GRAY)
        self.update()

        nom_jeu = self.gv['nom_jeu'].get()
        date    = self.gv['date'].get()
        lieu    = self.gv['lieu'].get()
        tab     = self.v_tab_name.get()

        def run():
            try:
                sh  = _sheets_client(self.creds_path.get())
                ws  = sh.worksheet(tab)
                row = sports_export_to_sheet(
                    ws, self.winners, nom_jeu, date, lieu, self.winner_urls)
                self.after(0, lambda: self.lbl_sheets.config(
                    text=f"✓  Tableau ajouté dans « {tab} » à partir de la ligne {row}",
                    fg=C_GREEN))
                self.after(0, lambda: self._log(
                    f"Sheets : tableau exporté dans « {tab} » (ligne {row})"))
            except Exception as e:
                err = str(e)
                self.after(0, lambda: self.lbl_sheets.config(
                    text=f"✗  Erreur : {err}", fg=C_RED))
                self.after(0, lambda: self._log(f"Erreur export Sheets : {err}"))

        threading.Thread(target=run, daemon=True).start()

    # ⑧ Modèle d'email & Envoi
    def _build_step8(self):
        f = self._section("Modèle d'email & Envoi", "⑧")
        row = tk.Frame(f, bg=C_WHITE)
        row.pack(fill=tk.X)
        tk.Label(row, text="Modèle :", bg=C_WHITE,
                 font=("", 10)).pack(side=tk.LEFT)
        self.tpl_var   = tk.StringVar()
        self.tpl_combo = ttk.Combobox(row, textvariable=self.tpl_var,
                                       state="readonly", width=34, font=("", 10))
        self.tpl_combo.pack(side=tk.LEFT, padx=8)
        self._refresh_templates()
        tk.Label(f,
                 text="Variable supplémentaire disponible : {lien_dropbox}",
                 bg=C_WHITE, fg=C_GRAY, font=("", 9)).pack(anchor="w", pady=(4, 10))
        row2 = tk.Frame(f, bg=C_WHITE)
        row2.pack(fill=tk.X)
        ColorButton(row2, text=f"📧  Test → {TEST_EMAIL}",
                    command=self._send_test,
                    bg=C_BLUE, font_size=10).pack(side=tk.LEFT, padx=(0, 10))
        ColorButton(row2, text="🚀  Envoyer aux gagnants",
                    command=self._send_all,
                    bg=C_GREEN, font_size=10).pack(side=tk.LEFT)
        self.lbl_send = tk.Label(f, text="", bg=C_WHITE,
                                  fg=C_GREEN, font=("", 10))
        self.lbl_send.pack(anchor="w", pady=(8, 0))

    def _refresh_templates(self):
        names = [t['name'] for t in self.templates]
        self.tpl_combo['values'] = names
        if names and not self.tpl_var.get():
            self.tpl_var.set(names[0])

    def _get_tpl(self):
        name = self.tpl_var.get()
        return next((t for t in self.templates if t['name'] == name), None)

    def _variables(self, winner=None):
        d = {k: v.get() for k, v in self.gv.items()}
        if winner:
            d["prenom"]       = winner['prenom']
            d["nom"]          = winner['nom']
            d["lien_dropbox"] = self.winner_urls.get(winner['email'], '')
        else:
            d["prenom"]       = "Pascal"
            d["nom"]          = "Thomas"
            d["lien_dropbox"] = "(lien dropbox)"
        return d

    def _send_test(self):
        tpl = self._get_tpl()
        if not tpl:
            messagebox.showwarning("Modèle", "Sélectionnez un modèle.", parent=self)
            return

        if self.winners:
            items = [(self._variables(winner=w), w) for w in self.winners]
        else:
            items = [(self._variables(), None)]

        self.lbl_send.config(
            text=f"⏳  Envoi de {len(items)} mail(s) de test…", fg=C_GRAY)
        self.update()

        def run():
            ok_count, errs = 0, []
            for v, w in items:
                subject = f"[TEST] {apply_vars(tpl['subject'], v)}"
                html    = apply_vars(tpl['html'],  v)
                plain   = apply_vars(tpl['plain'], v)
                try:
                    send_smtp(TEST_EMAIL, subject, html, plain)
                    ok_count += 1
                    name = f"{w['prenom']} {w['nom']}" if w else "test"
                    self.after(0, lambda n=name: self._log(
                        f"✓  Test ({n}) → {TEST_EMAIL}"))
                except Exception as e:
                    errs.append(str(e))
                    self.after(0, lambda err=str(e): self._log(
                        f"✗  Erreur SMTP : {err}"))

            def finish():
                if not errs:
                    self.lbl_send.config(
                        text=f"✓  {ok_count} mail(s) de test envoyé(s) à {TEST_EMAIL}",
                        fg=C_GREEN)
                    messagebox.showinfo("Tests envoyés",
                        f"{ok_count} mail(s) envoyé(s) à {TEST_EMAIL}\n"
                        f"(sujet préfixé [TEST])", parent=self)
                else:
                    self.lbl_send.config(
                        text=f"✗  Erreur SMTP : {errs[0]}", fg=C_RED)
                    messagebox.showerror("Erreur SMTP", errs[0], parent=self)
            self.after(0, finish)

        threading.Thread(target=run, daemon=True).start()

    def _send_all(self):
        if not self.winners:
            messagebox.showwarning("Tirage", "Effectuez d'abord le tirage (⑤).",
                                   parent=self)
            return
        tpl = self._get_tpl()
        if not tpl:
            messagebox.showwarning("Modèle", "Sélectionnez un modèle.", parent=self)
            return
        if not self.winner_urls and not messagebox.askyesno(
                "Liens manquants",
                "La distribution n'a pas encore été effectuée.\n"
                "{lien_dropbox} sera vide. Continuer ?", parent=self):
            return
        if not messagebox.askyesno("Confirmation",
                f"Envoyer {len(self.winners)} email(s) aux gagnants ?",
                parent=self):
            return
        self.lbl_send.config(text="⏳  Envoi en cours…", fg=C_GRAY)
        self.update()

        def run():
            ok_count, errs = 0, []
            for w in self.winners:
                if not w['email']:
                    continue
                v       = self._variables(winner=w)
                subject = apply_vars(tpl['subject'], v)
                html    = apply_vars(tpl['html'],    v)
                plain   = apply_vars(tpl['plain'],   v)
                try:
                    send_smtp(w['email'], subject, html, plain)
                    ok_count += 1
                    ww = dict(w)
                    self.after(0, lambda x=ww: self._log(
                        f"✓  {x['prenom']} {x['nom']} → {x['email']}"))
                except Exception as e:
                    errs.append(w['email'])
                    ww, ee = dict(w), str(e)
                    self.after(0, lambda x=ww, err=ee: self._log(
                        f"✗  {x['email']} : {err}"))
            msg   = f"✓  {ok_count} email(s) envoyé(s)"
            if errs:
                msg += f"  —  {len(errs)} erreur(s)"
            self.after(0, lambda: self.lbl_send.config(
                text=msg, fg=C_GREEN if not errs else C_RED))

        threading.Thread(target=run, daemon=True).start()

    # Journal
    def _build_log(self):
        f = self._section("Journal", "📋")
        self.log_widget = scrolledtext.ScrolledText(
            f, height=8, wrap=tk.WORD,
            bg="#1C1C1E", fg="#D4D4D4",
            font=("Courier", 10), insertbackground=C_WHITE)
        self.log_widget.pack(fill=tk.BOTH, expand=True)
        self.log_widget.config(state=tk.DISABLED)

if __name__ == "__main__":
    app = FlashFMApp()
    app.mainloop()
