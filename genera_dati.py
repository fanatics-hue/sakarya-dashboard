# -*- coding: utf-8 -*-
"""
genera_dati.py — Sakarya QC Dashboard
Legge il file Excel "Sakarya Inspection Overall Status as of *.xlsx" piu' recente
e genera 3 file CSV in ./data/ che il dashboard (index.html) legge direttamente:
    data/summary.csv        -> KPI + ITP Steps  (scheda Dashboard dell'Excel)
    data/defectsWeld.csv    -> difetti saldatura (scheda Repair)
    data/defectsFinal.csv   -> difetti final     (scheda Rejection)

NON serve aprire o modificare nulla a mano: doppio-clic su AGGIORNA_DASHBOARD.bat.
"""
import csv
import glob
import os
import re
import sys
from datetime import datetime

import openpyxl

# ======================================================================
#  CONFIG — l'unico valore che NON e' nell'Excel di status.
#  Modificalo qui solo se cambia (raramente).
# ======================================================================
PO_QTY = 15462          # Purchase Order (PO) Qty — quantita' totale ordine
# NB: "Pipes Rejected" NON e' un config: viene calcolato come numero di
#     righe della tabella Rejection (foglio Rejection -> defectsFinal),
#     replicando la logica del vecchio dashboard live.
# ======================================================================

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
PATTERN = "Sakarya Inspection Overall Status as of *.xlsx"


def log(msg):
    print(msg, flush=True)


def trovan_file_excel():
    """Trova il file di status piu' recente nella cartella."""
    candidati = [f for f in glob.glob(os.path.join(HERE, PATTERN))
                 if not os.path.basename(f).startswith("~$")]
    if not candidati:
        log("[ERRORE] Nessun file '%s' trovato nella cartella." % PATTERN)
        sys.exit(1)

    def chiave_data(path):
        # piu' recente per data "as of MM.DD.YYYY" nel nome file: non ci si
        # affida al mtime, che puo' essere alterato da copie/sync e far
        # scegliere per sbaglio un file di una settimana vecchia.
        m = re.search(r"as of (\d{1,2})\.(\d{1,2})\.(\d{4})", os.path.basename(path))
        if m:
            mm, dd, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return (yyyy, mm, dd)
        return (0, 0, 0)

    f = max(candidati, key=lambda p: (chiave_data(p), os.path.getmtime(p)))
    return f


def data_da_nome_file(path):
    """Estrae 'as of MM.DD.YYYY' dal nome file -> stringa DD/MM/YYYY."""
    m = re.search(r"as of (\d{1,2})\.(\d{1,2})\.(\d{4})", os.path.basename(path))
    if m:
        mm, dd, yyyy = m.group(1), m.group(2), m.group(3)
        return "%02d/%02d/%s" % (int(dd), int(mm), yyyy)
    return datetime.now().strftime("%d/%m/%Y")


def fmt_data(v):
    """Converte un valore data (datetime o stringa) in DD/MM/YYYY."""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    s = str(v or "").strip()
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return "%s/%s/%s" % (m.group(3), m.group(2), m.group(1))
    return s.split(" ")[0]


def num(v, default=0):
    """Estrae un numero da una cella (gestisce None, stringhe, errori #DIV/0)."""
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    return float(m.group(0)) if m else default


def righe_foglio(ws, max_col=12):
    """Restituisce tutte le righe del foglio come liste di celle (valori)."""
    return [list(r) for r in ws.iter_rows(values_only=True)]


# ----------------------------------------------------------------------
#  ESTRAZIONE
# ----------------------------------------------------------------------
def estrai_summary(wb, report_date, rejected):
    """Dal foglio 'Dashboard': KPI + 24 ITP Steps.
    `rejected` = n. righe tabella Rejection (calcolato a parte)."""
    ws = wb["Dashboard"]
    rows = righe_foglio(ws)

    def cella_dopo_label(label_sub):
        """Trova la prima riga con una cella che contiene label_sub e
        ritorna il primo valore numerico successivo nella stessa riga."""
        label_sub = label_sub.lower()
        for r in rows:
            testo = [("" if c is None else str(c)) for c in r]
            for j, c in enumerate(testo):
                if label_sub in c.lower():
                    for c2 in r[j + 1:]:
                        if isinstance(c2, (int, float)):
                            return c2
        return None

    total = cella_dopo_label("total pipes")
    if total is None:
        total = cella_dopo_label("input")
    accepted = cella_dopo_label("accept")
    passrate = cella_dopo_label("overall status")

    incoming = int(num(total, 0))
    accepted = int(num(accepted, 0))
    passrate = num(passrate, 0)
    if passrate <= 1.5:          # memorizzato come frazione (0.8765)
        passrate = passrate * 100
    passrate = round(passrate, 2)

    # --- ITP Steps: trova la riga header con "Seq"
    itp = []
    hdr_idx = -1
    hdr_off = 0
    for i, r in enumerate(rows):
        for j, c in enumerate(r):
            if c and "seq" in str(c).lower():
                hdr_idx = i
                hdr_off = j      # colonna dove inizia "Seq. n°"
                break
        if hdr_idx >= 0:
            break

    repair_count = 0
    if hdr_idx >= 0:
        for r in rows[hdr_idx + 1:]:
            base = r[hdr_off:]
            seq = base[0] if len(base) > 0 else None
            if not isinstance(seq, (int, float)):
                continue
            name = (base[1] if len(base) > 1 else "") or ""
            group = (base[2] if len(base) > 2 else "") or ""
            count = int(num(base[3] if len(base) > 3 else 0))
            yld = num(base[4] if len(base) > 4 else 0)
            if yld <= 1.5:
                yld = yld * 100
            loss = int(num(base[5] if len(base) > 5 else 0))
            cumul = int(num(base[6] if len(base) > 6 else 0))
            remarks = ""
            if len(base) > 7 and base[7]:
                remarks = str(base[7]).strip()
            itp.append([int(seq), str(name).strip(), str(group).strip(),
                        count, round(yld, 2), loss, cumul, remarks])
            if "welding repair" in str(name).strip().lower():
                repair_count = count

    # ---- scrivi CSV nel formato che il dashboard si aspetta ----
    out = []
    out.append(["SAKARYA GAS FIELD", ""])
    out.append(["KEY PERFORMANCE INDICATORS", ""])
    out.append(["Field", "Value"])
    out.append(["Purchase Order (PO) Qty", PO_QTY])
    out.append(["Incoming Plates", incoming])
    out.append(["Pipes Accepted", accepted])
    out.append(["Pipes Rejected", rejected])
    out.append(["Repair / Rework", repair_count])
    out.append(["Overall Pass Rate (%)", passrate])
    out.append(["Report Date", report_date])
    out.append([])
    out.append(["Seq. N°", "ITP Step", "Section", "Count",
                "Yield vs Input (%)", "Loss vs Input", "Cumul. Loss", "Remarks"])
    out.extend(itp)

    return out, dict(incoming=incoming, accepted=accepted,
                     passrate=passrate, repair=repair_count, itp=len(itp))


def estrai_difetti(wb, sheet_name, con_misure):
    """Dal foglio Repair/Rejection -> tabella difetti standard a 7 colonne."""
    ws = wb[sheet_name]
    rows = righe_foglio(ws)

    # trova header (riga con cella 'DATE') e offset colonna
    hdr_idx, off = -1, 0
    for i, r in enumerate(rows):
        for j, c in enumerate(r):
            if c and str(c).strip().lower().startswith("date"):
                hdr_idx, off = i, j
                break
        if hdr_idx >= 0:
            break
    if hdr_idx < 0:
        return [["Date", "Shift", "Pipe N°", "Defect", "Position", "Extension", "Remarks"]]

    out = [["Date", "Shift", "Pipe N°", "Defect", "Position", "Extension", "Remarks"]]
    for r in rows[hdr_idx + 1:]:
        b = r[off:]
        date = b[0] if len(b) > 0 else None
        pipe = b[2] if len(b) > 2 else None
        if not date or not pipe:
            continue
        shift = str(b[1]).strip() if len(b) > 1 and b[1] is not None else ""
        defect = str(b[3]).strip() if len(b) > 3 and b[3] is not None else ""
        if con_misure:
            pos = int(num(b[4] if len(b) > 4 else 0))
            ext = int(num(b[5] if len(b) > 5 else 0))
            rem = str(b[6]).strip() if len(b) > 6 and b[6] is not None else ""
        else:
            # Rejection: niente colonne Position/Extension
            pos, ext = "", ""
            rem = str(b[4]).strip() if len(b) > 4 and b[4] is not None else ""
        out.append([fmt_data(date), shift, str(pipe).strip(), defect, pos, ext, rem])
    return out


def scrivi_csv(nome, righe):
    path = os.path.join(DATA_DIR, nome)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for r in righe:
            w.writerow(r)
    return path


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    excel = trovan_file_excel()
    report_date = data_da_nome_file(excel)
    log("=" * 60)
    log("  SAKARYA DASHBOARD — generazione dati")
    log("=" * 60)
    log("File Excel  : %s" % os.path.basename(excel))
    log("Report Date : %s" % report_date)

    import warnings
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)

    # prima i difetti: "Pipes Rejected" = n. righe della tabella Rejection
    weld = estrai_difetti(wb, "Repair", con_misure=True)
    final = estrai_difetti(wb, "Rejection", con_misure=False)
    rejected = len(final) - 1   # -1 per l'header

    summary, stats = estrai_summary(wb, report_date, rejected)

    scrivi_csv("summary.csv", summary)
    scrivi_csv("defectsWeld.csv", weld)
    scrivi_csv("defectsFinal.csv", final)

    log("-" * 60)
    log("KPI estratti:")
    log("  Incoming Plates  : %s" % stats["incoming"])
    log("  Pipes Accepted   : %s" % stats["accepted"])
    log("  Pass Rate        : %s %%" % stats["passrate"])
    log("  Repair / Rework  : %s" % stats["repair"])
    log("  ITP Steps        : %s righe" % stats["itp"])
    log("  PO Qty (config)  : %s" % PO_QTY)
    log("  Pipes Rejected   : %s  (= righe Rejection)" % rejected)
    log("Difetti Weld (Repair)    : %s righe" % (len(weld) - 1))
    log("Difetti Final (Rejection): %s righe" % (len(final) - 1))
    log("-" * 60)
    log("OK — file scritti in: %s" % DATA_DIR)


if __name__ == "__main__":
    main()
