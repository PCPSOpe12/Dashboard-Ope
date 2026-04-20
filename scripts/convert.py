#!/usr/bin/env python3
"""
Convertit l'export DPS (Excel) en data.json pour le dashboard.
Lancé automatiquement par GitHub Actions à chaque push du fichier Excel.
"""
import json, sys, os
from pathlib import Path
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    print("openpyxl manquant. Lancer: pip install openpyxl")
    sys.exit(1)

EXCEL_PATH = Path("data/export_dps.xlsx")
JSON_PATH  = Path("data.json")

def parse_date(val):
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, datetime) else datetime.combine(val, datetime.min.time())
    if isinstance(val, str):
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try: return datetime.strptime(val.strip(), fmt)
            except: pass
    return None

def safe_float(v, default=0.0):
    try:
        if v is None or v == "": return default
        return float(str(v).replace(",", "."))
    except: return default

def safe_str(v):
    return str(v).strip() if v is not None else ""

def main():
    if not EXCEL_PATH.exists():
        print(f"Fichier introuvable : {EXCEL_PATH}")
        sys.exit(1)

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    # En-têtes
    headers = [str(h).strip() if h else "" for h in next(rows_iter)]

    def col(row, *names):
        for name in names:
            for i, h in enumerate(headers):
                if h.lower().replace(" ", "").replace("(", "").replace(")", "").replace("'", "") == \
                   name.lower().replace(" ", "").replace("(", "").replace(")", "").replace("'", ""):
                    return row[i] if i < len(row) else None
        return None

    data = []
    max_date = None

    for row in rows_iter:
        if all(v is None for v in row):
            continue

        date_val = col(row, "Début", "Debut", "début")
        d = parse_date(date_val)
        if not d:
            continue

        annee = d.year
        mois  = d.month
        if annee < 2020:
            continue

        if max_date is None or d > max_date:
            max_date = d

        data.append({
            "annee":    annee,
            "mois":     mois,
            "date":     d.strftime("%d/%m/%Y"),
            "libelle":  safe_str(col(row, "Libellé", "Libelle")),
            "lieu":     safe_str(col(row, "Lieu")),
            "dps":      safe_str(col(row, "DPS")) or "-",
            "statut":   safe_str(col(row, "Ouvert.", "Ouvert")) or "cloturé",
            "inscrits": int(safe_float(col(row, "Inscrits"))),
            "heures":   round(safe_float(col(row, "Heures")), 1),
            "pec":      int(safe_float(col(row, "Priseesencharge", "Prise(s)encharge", "Priseencharge"))),
            "duree":    round(safe_float(col(row, "Durée", "Duree")), 2),
        })

    output = {
        "rows":         data,
        "lastModified": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "maxDate":      max_date.strftime("%d/%m/%Y") if max_date else "—",
        "totalRows":    len(data),
        "years":        sorted(list(set(r["annee"] for r in data))),
    }

    JSON_PATH.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK — {len(data)} lignes converties → {JSON_PATH}")
    print(f"Années : {output['years']}")
    print(f"Dernière date : {output['maxDate']}")

if __name__ == "__main__":
    main()
