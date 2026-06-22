# -*- coding: utf-8 -*-
"""
app_dashboard_gui.py - Sakarya QC Dashboard (finestra grafica)

Flusso a 2 passi:
  1) Scegli il file Excel  -> [1. Genera e verifica]  -> semaforo + anteprima KPI
  2) [2. Pubblica su GitHub] -> git pull + commit + push -> dashboard online

- Se il file scelto non e' gia' nella cartella (o ha un nome diverso), viene
  copiato qui col nome corretto "Sakarya Inspection Overall Status as of MM.DD.YYYY.xlsx".
- Prima del push fa SEMPRE git pull (anti-conflitto OneDrive / piu' PC).
- Nessuna finestra DOS: tutto l'output e' nel riquadro log.

Avvio: doppio clic su AVVIA_GUI.bat  (oppure  start "" pythonw app_dashboard_gui.py)
"""
import os
import re
import sys
import glob
import shutil
import subprocess
import threading
import webbrowser
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl

# ----------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
PATTERN = "Sakarya Inspection Overall Status as of *.xlsx"
FOGLI_RICHIESTI = ["Dashboard", "Repair", "Rejection"]
DASHBOARD_URL = "https://fanatics-hue.github.io/sakarya-dashboard/"
CRUSCOTTO = os.path.join(os.path.dirname(HERE), "Cruscotto Workspace.hta")

# colori semaforo
C_IDLE = "#5a6472"
C_OK = "#1e8e3e"
C_WARN = "#e8a200"
C_ERR = "#c5221f"
C_BUSY = "#1a73e8"

# nascondi la finestra nera dei processi git su Windows
_NO_WIN = 0x08000000 if os.name == "nt" else 0


def run(cmd):
    """Esegue un comando, ritorna (returncode, output_unito)."""
    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    p = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True,
                       encoding="utf-8", errors="replace",
                       creationflags=_NO_WIN, env=env)
    out = (p.stdout or "") + (p.stderr or "")
    return p.returncode, out.strip()


class App:
    def __init__(self, root):
        self.root = root
        self.excel_path = None      # file scelto dall'utente
        self.mm = self.dd = self.yyyy = None
        self.busy = False

        root.title("Sakarya Dashboard - Aggiornamento")
        root.geometry("680x560")
        root.minsize(620, 500)

        # ---- banner semaforo ----
        self.banner = tk.Label(root, text="Pronto. Scegli il file Excel per iniziare.",
                               bg=C_IDLE, fg="white", font=("Segoe UI", 12, "bold"),
                               anchor="w", padx=14, pady=12)
        self.banner.pack(fill="x")

        body = ttk.Frame(root, padding=12)
        body.pack(fill="both", expand=True)

        # ---- riga file ----
        rf = ttk.Frame(body)
        rf.pack(fill="x", pady=(0, 6))
        ttk.Label(rf, text="File Excel:", width=11).pack(side="left")
        self.lbl_file = ttk.Label(rf, text="(nessun file scelto)", foreground="#888")
        self.lbl_file.pack(side="left", fill="x", expand=True)
        ttk.Button(rf, text="Scegli file...", command=self.scegli_file).pack(side="right")

        # ---- riga data ----
        rd = ttk.Frame(body)
        rd.pack(fill="x", pady=(0, 10))
        ttk.Label(rd, text="Data report:", width=11).pack(side="left")
        self.ent_data = ttk.Entry(rd, width=14)
        self.ent_data.pack(side="left")
        ttk.Label(rd, text="(formato MM.DD.YYYY - dal nome del file, modificabile)",
                  foreground="#888").pack(side="left", padx=8)

        # ---- pulsanti azione ----
        rb = ttk.Frame(body)
        rb.pack(fill="x", pady=(0, 10))
        self.btn_gen = ttk.Button(rb, text="1.  Genera e verifica",
                                  command=self.t_genera, state="disabled")
        self.btn_gen.pack(side="left")
        self.btn_push = ttk.Button(rb, text="2.  Pubblica su GitHub",
                                   command=self.t_pubblica, state="disabled")
        self.btn_push.pack(side="left", padx=8)

        # ---- log ----
        ttk.Label(body, text="Dettagli:").pack(anchor="w")
        wrap = ttk.Frame(body)
        wrap.pack(fill="both", expand=True)
        self.txt = tk.Text(wrap, height=12, wrap="word", font=("Consolas", 9),
                           bg="#1e1e1e", fg="#d4d4d4", insertbackground="white")
        sb = ttk.Scrollbar(wrap, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set, state="disabled")
        self.txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ---- pulsanti in basso ----
        rbot = ttk.Frame(body)
        rbot.pack(fill="x", pady=(10, 0))
        ttk.Button(rbot, text="Apri dashboard online",
                   command=lambda: webbrowser.open(DASHBOARD_URL)).pack(side="left")
        ttk.Button(rbot, text="Apri cartella",
                   command=lambda: os.startfile(HERE)).pack(side="left", padx=8)
        if os.path.exists(CRUSCOTTO):
            ttk.Button(rbot, text="Cruscotto IQS",
                       command=lambda: os.startfile(CRUSCOTTO)).pack(side="right")

        self.log("Pronto. Repository: fanatics-hue/sakarya-dashboard (branch main).")

    # ------------------------------------------------------------------ utils
    def log(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def set_banner(self, text, color):
        self.banner.configure(text=text, bg=color)

    def lock(self, on):
        self.busy = on
        st = "disabled" if on else "normal"
        self.btn_gen.configure(state=st if self.excel_path else "disabled")
        # btn_push lo gestiamo a parte (resta off finche' non si genera)

    # --------------------------------------------------------------- scegli file
    def scegli_file(self):
        if self.busy:
            return
        f = filedialog.askopenfilename(
            title="Scegli il file Excel di status Sakarya",
            initialdir=HERE,
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")])
        if not f:
            return
        self.excel_path = f
        self.lbl_file.configure(text=os.path.basename(f), foreground="#000")
        # prova a leggere la data dal nome
        m = re.search(r"as of (\d{1,2})\.(\d{1,2})\.(\d{4})", os.path.basename(f))
        if m:
            self.mm, self.dd, self.yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
            self.log("Data rilevata dal nome file: %02d.%02d.%d" % (self.mm, self.dd, self.yyyy))
        else:
            now = datetime.now()
            self.mm, self.dd, self.yyyy = now.month, now.day, now.year
            self.log("Nome file senza data: precompilo con OGGI. Correggila se serve.")
        self.ent_data.delete(0, "end")
        self.ent_data.insert(0, "%02d.%02d.%d" % (self.mm, self.dd, self.yyyy))
        self.btn_gen.configure(state="normal")
        self.btn_push.configure(state="disabled")
        self.set_banner("File scelto. Premi  '1. Genera e verifica'.", C_IDLE)

    def _leggi_data_campo(self):
        s = self.ent_data.get().strip()
        m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", s)
        if not m:
            return None
        return int(m.group(1)), int(m.group(2)), int(m.group(3))

    # ------------------------------------------------------------------ genera
    def t_genera(self):
        threading.Thread(target=self.genera, daemon=True).start()

    def genera(self):
        try:
            self.lock(True)
            self.btn_push.configure(state="disabled")
            self.set_banner("Verifica e generazione in corso...", C_BUSY)

            # 1. data valida
            d = self._leggi_data_campo()
            if not d:
                self.set_banner("STOP: data non valida (usa MM.DD.YYYY).", C_ERR)
                self.log("[ERRORE] Data nel formato sbagliato: '%s'" % self.ent_data.get())
                return
            self.mm, self.dd, self.yyyy = d

            # 2. validazione fogli Excel
            self.log("-" * 56)
            self.log("Apro l'Excel per la verifica...")
            try:
                wb = openpyxl.load_workbook(self.excel_path, read_only=True)
                mancanti = [s for s in FOGLI_RICHIESTI if s not in wb.sheetnames]
                wb.close()
            except Exception as e:
                self.set_banner("STOP: impossibile aprire l'Excel.", C_ERR)
                self.log("[ERRORE] %s" % e)
                return
            if mancanti:
                self.set_banner("STOP: fogli mancanti nell'Excel.", C_ERR)
                self.log("[ERRORE] Mancano i fogli: %s" % ", ".join(mancanti))
                return
            self.log("Fogli richiesti presenti: %s" % ", ".join(FOGLI_RICHIESTI))

            # 3. copia nella cartella col nome corretto (se serve)
            target_name = "Sakarya Inspection Overall Status as of %02d.%02d.%d.xlsx" % (
                self.mm, self.dd, self.yyyy)
            target_path = os.path.join(HERE, target_name)
            if os.path.abspath(self.excel_path) != os.path.abspath(target_path):
                shutil.copy2(self.excel_path, target_path)
                self.log("File copiato nella cartella come: %s" % target_name)
            else:
                # tocca il file per renderlo il piu' recente
                os.utime(target_path, None)
                self.log("Uso il file gia' presente: %s" % target_name)
            self.excel_path = target_path

            # 4. genera i CSV richiamando genera_dati.py (logica gia' collaudata)
            self.log("Genero i dati (genera_dati.py)...")
            rc, out = run([sys.executable, os.path.join(HERE, "genera_dati.py")])
            for line in out.splitlines():
                self.log("   " + line)
            if rc != 0:
                self.set_banner("STOP: generazione dati fallita.", C_ERR)
                return

            # 5. anteprima KPI dal summary.csv
            kpi = self._leggi_kpi()
            self._mostra_kpi(kpi)

            # semaforo: warn se numeri sospetti
            if kpi.get("Incoming Plates", 0) in (0, None) or kpi.get("Overall Pass Rate (%)", 0) in (0, None):
                self.set_banner("ATTENZIONE: controlla i numeri prima di pubblicare.", C_WARN)
            else:
                self.set_banner("OK - dati pronti. Controlla i KPI, poi  '2. Pubblica'.", C_OK)

            self.btn_push.configure(state="normal")
        finally:
            self.lock(False)

    def _leggi_kpi(self):
        import csv
        kpi = {}
        path = os.path.join(DATA_DIR, "summary.csv")
        itp = 0
        in_itp = False
        try:
            with open(path, encoding="utf-8") as f:
                for row in csv.reader(f):
                    if not row:
                        continue
                    if row[0] in ("Purchase Order (PO) Qty", "Incoming Plates",
                                  "Pipes Accepted", "Pipes Rejected",
                                  "Repair / Rework", "Overall Pass Rate (%)",
                                  "Report Date"):
                        kpi[row[0]] = row[1] if len(row) > 1 else ""
                    if row[0] == "Seq. N°":
                        in_itp = True
                        continue
                    if in_itp and row[0].strip().isdigit():
                        itp += 1
            kpi["ITP Steps"] = itp
            # converti numerici utili al semaforo
            for k in ("Incoming Plates", "Overall Pass Rate (%)"):
                try:
                    kpi[k + "_n"] = float(str(kpi.get(k, "0")).replace(",", "."))
                except Exception:
                    kpi[k + "_n"] = 0
            kpi["Incoming Plates"] = kpi.get("Incoming Plates_n", 0)
            kpi["Overall Pass Rate (%)"] = kpi.get("Overall Pass Rate (%)_n", 0)
        except Exception as e:
            self.log("[avviso] non riesco a leggere summary.csv: %s" % e)
        return kpi

    def _mostra_kpi(self, kpi):
        self.log("-" * 56)
        self.log("ANTEPRIMA KPI (verifica prima di pubblicare):")
        self.log("   PO Qty            : %s" % kpi.get("Purchase Order (PO) Qty", "?"))
        self.log("   Incoming Plates   : %s" % kpi.get("Incoming Plates", "?"))
        self.log("   Pipes Accepted    : %s" % kpi.get("Pipes Accepted", "?"))
        self.log("   Pipes Rejected    : %s" % kpi.get("Pipes Rejected", "?"))
        self.log("   Repair / Rework   : %s" % kpi.get("Repair / Rework", "?"))
        self.log("   Overall Pass Rate : %s %%" % kpi.get("Overall Pass Rate (%)", "?"))
        self.log("   ITP Steps         : %s righe" % kpi.get("ITP Steps", "?"))
        self.log("   Report Date       : %s" % kpi.get("Report Date", "?"))
        self.log("-" * 56)

    # ----------------------------------------------------------------- pubblica
    def t_pubblica(self):
        if not messagebox.askyesno(
                "Conferma pubblicazione",
                "Pubblicare i dati su GitHub?\n\n"
                "La dashboard online verra' aggiornata per tutti i colleghi."):
            return
        threading.Thread(target=self.pubblica, daemon=True).start()

    def pubblica(self):
        try:
            self.lock(True)
            self.btn_push.configure(state="disabled")
            self.set_banner("Pubblicazione in corso...", C_BUSY)

            # git presente?
            rc, _ = run(["git", "--version"])
            if rc != 0:
                self.set_banner("STOP: Git non installato.", C_ERR)
                self.log("[ERRORE] Git non trovato. Installa da https://git-scm.com")
                return

            rc, branch = run(["git", "rev-parse", "--abbrev-ref", "HEAD"])
            branch = branch.strip() or "main"

            # 1. PULL (anti-conflitto)
            self.log("-" * 56)
            self.log("git pull origin %s ..." % branch)
            rc, out = run(["git", "pull", "--no-edit", "origin", branch])
            self.log(out or "(ok)")
            if rc != 0:
                self.set_banner("STOP: pull fallito (conflitto o rete).", C_ERR)
                self.log("[ERRORE] Risolvi il conflitto/connessione e riprova.")
                return

            # 2. ADD
            run(["git", "add", "-A"])
            rc, _ = run(["git", "diff", "--cached", "--quiet"])
            if rc == 0:
                self.set_banner("Niente da pubblicare: gia' aggiornata.", C_WARN)
                self.log("Nessuna modifica rispetto a GitHub.")
                return

            # 3. COMMIT
            msg = "Aggiornamento dashboard: %02d.%02d.%d" % (self.mm, self.dd, self.yyyy)
            rc, out = run(["git", "commit", "-m", msg])
            self.log(out)

            # 4. PUSH
            self.log("git push origin %s ..." % branch)
            rc, out = run(["git", "push", "origin", branch])
            self.log(out or "(ok)")
            if rc != 0:
                self.set_banner("STOP: push fallito (credenziali/rete).", C_ERR)
                self.log("[ERRORE] Controlla login GitHub e connessione.")
                return

            self.set_banner("FATTO! Online tra ~1 min (ricarica con CTRL+F5).", C_OK)
            self.log("=" * 56)
            self.log("COMPLETATO. Dashboard: %s" % DASHBOARD_URL)
        finally:
            self.lock(False)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
