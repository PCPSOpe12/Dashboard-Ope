#!/usr/bin/env python3
"""
Convertit les exports DPS et Garde (Excel) en JSON pour le dashboard.
Lancé automatiquement par GitHub Actions à chaque push.
"""
import json, sys, os
from pathlib import Path
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    print("openpyxl manquant. Lancer: pip install openpyxl")
    sys.exit(1)

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

def col(row, headers, *names):
    for name in names:
        for i, h in enumerate(headers):
            hn = h.lower().replace(" ", "").replace("(", "").replace(")", "").replace("'", "").replace("é","e").replace("è","e").replace("ê","e")
            nn = name.lower().replace(" ", "").replace("(", "").replace(")", "").replace("'", "").replace("é","e").replace("è","e").replace("ê","e")
            if hn == nn:
                return row[i] if i < len(row) else None
    return None

def convert_file(excel_path, json_path, file_type):
    if not excel_path.exists():
        print(f"  Fichier introuvable : {excel_path} — ignoré")
        return False

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h else "" for h in next(rows_iter)]

    data = []
    max_date = None

    for row in rows_iter:
        if all(v is None for v in row):
            continue
        date_val = col(row, headers, "Début", "Debut", "début")
        d = parse_date(date_val)
        if not d or d.year < 2020:
            continue
        if max_date is None or d > max_date:
            max_date = d

        if file_type == "dps":
            entry = {
                "annee":   d.year,
                "mois":    d.month,
                "date":    d.strftime("%d/%m/%Y"),
                "libelle": safe_str(col(row, headers, "Libellé", "Libelle")),
                "lieu":    safe_str(col(row, headers, "Lieu")),
                "dps":     safe_str(col(row, headers, "DPS")) or "-",
                "statut":  safe_str(col(row, headers, "Ouvert.", "Ouvert")) or "cloturé",
                "inscrits": int(safe_float(col(row, headers, "Inscrits"))),
                "heures":   round(safe_float(col(row, headers, "Heures")), 1),
                "pec":      int(safe_float(col(row, headers, "Prise(s) en charge", "Priseesencharge"))),
            }
        else:  # garde
            entry = {
                "annee":   d.year,
                "mois":    d.month,
                "date":    d.strftime("%d/%m/%Y"),
                "libelle": safe_str(col(row, headers, "Libellé", "Libelle")),
                "lieu":    safe_str(col(row, headers, "Lieu")),
                "inscrits": int(safe_float(col(row, headers, "Inscrits"))),
                "heures":   round(safe_float(col(row, headers, "Heures")), 1),
                "pec":      int(safe_float(col(row, headers, "Intervention(s)", "Interventions"))),
            }
        data.append(entry)

    output = {
        "rows":         data,
        "lastModified": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "maxDate":      max_date.strftime("%d/%m/%Y") if max_date else "—",
        "totalRows":    len(data),
        "years":        sorted(list(set(r["annee"] for r in data))),
        "type":         file_type,
    }
    json_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  OK — {len(data)} lignes → {json_path} (années: {output['years']}, dernier: {output['maxDate']})")
    return True

def main():
    convert_file(Path("data/export_dps.xlsx"),   Path("data_dps.json"),   "dps")
    convert_file(Path("data/export_garde.xlsx"), Path("data_garde.json"), "garde")
    print("Conversion terminée.")

if __name__ == "__main__":
    main()
