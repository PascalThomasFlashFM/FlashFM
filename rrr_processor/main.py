#!/usr/bin/env python3
"""
Flash FM - Traitement CA RRR
Application de traitement des appels à facture de la Régie Radio Régions.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import os

from processor import process_monthly_file

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
DEFAULT_SYNTHESIS_PATH = "/Users/Pascal/kDrive/Pascal3/Flash FM/Administration/RRR/Appels à factures/CA RRR.xlsx"


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Flash FM - Traitement CA RRR")
        self.geometry("800x600")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        self.config_data = load_config()
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # Title
        title = tk.Label(
            self,
            text="Traitement CA Régie Radio Régions",
            font=("Helvetica", 16, "bold"),
            bg="#f0f0f0",
            fg="#1a3a5c",
        )
        title.pack(pady=(15, 5))

        subtitle = tk.Label(
            self,
            text="Flash FM",
            font=("Helvetica", 11),
            bg="#f0f0f0",
            fg="#666666",
        )
        subtitle.pack(pady=(0, 15))

        # Frame fichier mensuel
        frame_monthly = ttk.LabelFrame(self, text="Fichier mensuel (Appel à facture RRR)")
        frame_monthly.pack(fill="x", **pad)

        self.monthly_var = tk.StringVar()
        ttk.Entry(frame_monthly, textvariable=self.monthly_var, width=70).pack(
            side="left", padx=5, pady=8, fill="x", expand=True
        )
        ttk.Button(
            frame_monthly, text="Parcourir…", command=self._browse_monthly
        ).pack(side="right", padx=5, pady=8)

        # Frame fichier synthèse
        frame_synth = ttk.LabelFrame(self, text="Fichier de synthèse (CA RRR.xlsx)")
        frame_synth.pack(fill="x", **pad)

        synthesis_path = self.config_data.get("synthesis_path", DEFAULT_SYNTHESIS_PATH)
        self.synthesis_var = tk.StringVar(value=synthesis_path)
        ttk.Entry(frame_synth, textvariable=self.synthesis_var, width=70).pack(
            side="left", padx=5, pady=8, fill="x", expand=True
        )
        ttk.Button(
            frame_synth, text="Parcourir…", command=self._browse_synthesis
        ).pack(side="right", padx=5, pady=8)

        # Bouton traitement
        self.btn_process = ttk.Button(
            self,
            text="▶  Lancer le traitement",
            command=self._run_processing,
        )
        self.btn_process.pack(pady=12)

        # Zone de log
        frame_log = ttk.LabelFrame(self, text="Journal de traitement")
        frame_log.pack(fill="both", expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(
            frame_log,
            wrap="word",
            font=("Courier", 10),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
            state="disabled",
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Tags pour couleurs
        self.log_text.tag_config("ok", foreground="#6adb6a")
        self.log_text.tag_config("warn", foreground="#f0c040")
        self.log_text.tag_config("err", foreground="#f04040")
        self.log_text.tag_config("info", foreground="#80c8f0")
        self.log_text.tag_config("head", foreground="#c080f0", font=("Courier", 10, "bold"))

        # Status bar
        self.status_var = tk.StringVar(value="Prêt.")
        status_bar = tk.Label(
            self,
            textvariable=self.status_var,
            bd=1,
            relief="sunken",
            anchor="w",
            bg="#e0e0e0",
        )
        status_bar.pack(fill="x", side="bottom")

    def _browse_monthly(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier mensuel RRR",
            filetypes=[("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.monthly_var.set(path)

    def _browse_synthesis(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier de synthèse CA RRR",
            filetypes=[("Fichiers Excel", "*.xlsx *.xlsm"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.synthesis_var.set(path)
            self.config_data["synthesis_path"] = path
            save_config(self.config_data)

    def _log(self, message, tag=""):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _run_processing(self):
        monthly = self.monthly_var.get().strip()
        synthesis = self.synthesis_var.get().strip()

        if not monthly:
            messagebox.showerror("Erreur", "Veuillez sélectionner le fichier mensuel.")
            return
        if not os.path.exists(monthly):
            messagebox.showerror("Erreur", f"Fichier mensuel introuvable :\n{monthly}")
            return
        if not synthesis:
            messagebox.showerror("Erreur", "Veuillez sélectionner le fichier de synthèse.")
            return
        if not os.path.exists(synthesis):
            messagebox.showerror("Erreur", f"Fichier de synthèse introuvable :\n{synthesis}")
            return

        # Save synthesis path
        self.config_data["synthesis_path"] = synthesis
        save_config(self.config_data)

        # Clear log
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.btn_process.configure(state="disabled")
        self.status_var.set("Traitement en cours…")

        def task():
            try:
                process_monthly_file(monthly, synthesis, self._log)
                self.after(0, lambda: self.status_var.set("Traitement terminé."))
                self.after(0, lambda: self._log("\n✓ Traitement terminé avec succès.", "ok"))
            except Exception as exc:
                self.after(0, lambda: self.status_var.set("Erreur lors du traitement."))
                self.after(0, lambda: self._log(f"\n✗ Erreur : {exc}", "err"))
            finally:
                self.after(0, lambda: self.btn_process.configure(state="normal"))

        threading.Thread(target=task, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
