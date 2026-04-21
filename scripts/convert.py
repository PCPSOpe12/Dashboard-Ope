#!/usr/bin/env python3
"""
Convertit les exports DPS, Garde et SR (Excel) en JSON pour le dashboard.
"""
import json
from pathlib import Path
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    print("openpyxl manquant. Lancer: pip install openpyxl")
    import sys; sys.exit(1)

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
    def norm(s): return s.lower().replace(" ","").replace("(","").replace(")","").replace("'","").replace("é","e").replace("è","e").replace("ê","e")
    for name in names:
        for i, h in enumerate(headers):
            if norm(h) == norm(name):
                return row[i] if i < len(row) else None
    return None

def read_sheet(excel_path):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h else "" for h in next(rows_iter)]
    rows = [r for r in rows_iter if not all(v is None for v in r)]
    return headers, rows

# ── DPS / Garde ────────────────────────────────────────────────────────────
def convert_activity(excel_path, json_path, file_type):
    if not excel_path.exists():
        print(f"  Fichier introuvable : {excel_path} — ignoré")
        return False
    headers, rows = read_sheet(excel_path)
    data, max_date = [], None
    for row in rows:
        d = parse_date(col(row, headers, "Début", "Debut", "début"))
        if not d or d.year < 2020: continue
        if max_date is None or d > max_date: max_date = d
        if file_type == "dps":
            entry = {
                "annee": d.year, "mois": d.month, "date": d.strftime("%d/%m/%Y"),
                "libelle": safe_str(col(row, headers, "Libellé", "Libelle")),
                "lieu":    safe_str(col(row, headers, "Lieu")),
                "dps":     safe_str(col(row, headers, "DPS")) or "-",
                "statut":  safe_str(col(row, headers, "Ouvert.", "Ouvert")) or "cloturé",
                "inscrits": int(safe_float(col(row, headers, "Inscrits"))),
                "heures":   round(safe_float(col(row, headers, "Heures")), 1),
                "pec":      int(safe_float(col(row, headers, "Prise(s) en charge", "Priseesencharge"))),
            }
        else:
            entry = {
                "annee": d.year, "mois": d.month, "date": d.strftime("%d/%m/%Y"),
                "libelle": safe_str(col(row, headers, "Libellé", "Libelle")),
                "lieu":    safe_str(col(row, headers, "Lieu")),
                "inscrits": int(safe_float(col(row, headers, "Inscrits"))),
                "heures":   round(safe_float(col(row, headers, "Heures")), 1),
                "pec":      int(safe_float(col(row, headers, "Intervention(s)", "Interventions"))),
            }
        data.append(entry)
    now = datetime.now()
    output = {
        "rows": data,
        "lastModified": now.strftime("%d/%m/%Y"),
        "maxDate": max_date.strftime("%d/%m/%Y") if max_date else "—",
        "totalRows": len(data),
        "years": sorted(list(set(r["annee"] for r in data))),
        "type": file_type,
    }
    json_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  OK — {len(data)} lignes → {json_path} (dernier: {output['maxDate']})")
    return True

# ── SR ─────────────────────────────────────────────────────────────────────
COMP_CIBLES = {'PSE1','PSE2','CE','CP','CEPS','CDD'}
TODAY = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

def convert_sr(comp_path, part_path, json_path):
    if not comp_path.exists() and not part_path.exists():
        print(f"  Fichiers SR introuvables — ignoré"); return False

    # Compétences
    comp_pivot = {}
    if comp_path.exists():
        headers, rows = read_sheet(comp_path)
        for row in rows:
            typ = safe_str(col(row, headers, "TYPE")).strip()
            if typ not in COMP_CIBLES: continue
            nom    = safe_str(col(row, headers, "NOM")).upper()
            prenom = safe_str(col(row, headers, "prenom", "prénom", "Prénom"))
            key = nom + " " + prenom
            if key not in comp_pivot:
                comp_pivot[key] = {"nom": nom, "prenom": prenom}
            obt = safe_str(col(row, headers, "Obtention"))
            exp = safe_str(col(row, headers, "Expiration"))
            if obt in ("-","nan",""): obt = ""
            if exp in ("-","nan",""): exp = ""
            comp_pivot[key][typ] = {"obt": obt, "exp": exp}
        print(f"  Compétences SR : {len(comp_pivot)} personnes")

    # Participations
    stats = {}
    max_date = None
    min_date = None
    if part_path.exists():
        headers, rows = read_sheet(part_path)
        for row in rows:
            name = safe_str(col(row, headers, "Personnel"))
            if not name: continue
            parts = name.strip().split(" ", 1)
            key = parts[0].upper() + " " + parts[1] if len(parts) == 2 else name.upper()
            if key not in stats:
                stats[key] = {
                    "nb_dps":0,"nb_gar":0,"last_dps":None,"last_gar":None,
                    "heures":0,"inscrit_dps":False,"inscrit_gar":False,
                    "activites_dps":[], "activites_gar":[]
                }
            code = safe_str(col(row, headers, "Code"))
            d    = parse_date(col(row, headers, "Début", "Debut"))
            h    = safe_float(col(row, headers, "Présence", "Presence"))
            lib  = safe_str(col(row, headers, "Evenement", "Évènement", "Libellé", "Libelle"))
            if d:
                if max_date is None or d > max_date: max_date = d
                if min_date is None or d < min_date: min_date = d

            if code == "DPS":
                if d and d < TODAY:
                    stats[key]["nb_dps"] += 1
                    stats[key]["heures"] += h
                    if not stats[key]["last_dps"] or d > stats[key]["last_dps"]:
                        stats[key]["last_dps"] = d
                    stats[key]["activites_dps"].append({
                        "date": d.strftime("%d/%m/%Y"), "lib": lib, "heures": round(h,1),
                        "ts": d.timestamp()
                    })
                elif d and d >= TODAY:
                    stats[key]["inscrit_dps"] = True
            elif code == "GAR":
                if d and d < TODAY:
                    stats[key]["nb_gar"] += 1
                    stats[key]["heures"] += h
                    if not stats[key]["last_gar"] or d > stats[key]["last_gar"]:
                        stats[key]["last_gar"] = d
                    stats[key]["activites_gar"].append({
                        "date": d.strftime("%d/%m/%Y"), "lib": lib, "heures": round(h,1),
                        "ts": d.timestamp()
                    })
                elif d and d >= TODAY:
                    stats[key]["inscrit_gar"] = True
        print(f"  Participations SR : {len(stats)} personnes")

    # Fusion — uniquement SR avec participations ET heures > 0
    # Fenêtre 12 mois glissants pour les activités affichées
    cutoff_12m = datetime(max_date.year - 1, max_date.month, max_date.day) if max_date else None

    secouristes = []
    for key in sorted(stats.keys()):
        c = comp_pivot.get(key, {})
        s = stats[key]
        if round(s["heures"], 1) == 0: continue  # Exclure SR sans heures
        np = key.split(" ", 1)
        def comp_val(typ):
            v = c.get(typ)
            return {"obt": v.get("obt",""), "exp": v.get("exp","")} if v else None

        # Filtrer activités sur 12 mois glissants, trier par date décroissante
        def filter_acts(acts):
            filtered = []
            for a in acts:
                if cutoff_12m:
                    from datetime import datetime as dt2
                    try:
                        ad = dt2.strptime(a["date"], "%d/%m/%Y")
                        if ad < cutoff_12m: continue
                    except: pass
                filtered.append({k2: v2 for k2, v2 in a.items() if k2 != "ts"})
            return sorted(filtered, key=lambda x: x["date"], reverse=True)

        adps = filter_acts(s["activites_dps"])
        agar = filter_acts(s["activites_gar"])

        # Dernière activité toutes catégories pour le feu
        last_any = None
        if s["last_dps"] and s["last_gar"]:
            last_any = max(s["last_dps"], s["last_gar"])
        elif s["last_dps"]:
            last_any = s["last_dps"]
        elif s["last_gar"]:
            last_any = s["last_gar"]

        last_any_str = last_any.strftime("%d/%m/%Y") if last_any else ""

        secouristes.append({
            "nom":    c.get("nom")    or (np[0] if np else key),
            "prenom": c.get("prenom") or (np[1] if len(np)>1 else ""),
            "PSE1": comp_val("PSE1"), "PSE2": comp_val("PSE2"),
            "CE":   comp_val("CE"),   "CP":   comp_val("CP"),
            "CEPS": comp_val("CEPS"), "CDD":  comp_val("CDD"),
            "nb_dps": s["nb_dps"],   "nb_gar": s["nb_gar"],
            "last_dps": s["last_dps"].strftime("%d/%m/%Y") if s["last_dps"] else "",
            "last_gar": s["last_gar"].strftime("%d/%m/%Y") if s["last_gar"] else "",
            "last_any": last_any_str,
            "heures": round(s["heures"], 1),
            "inscrit_dps": s["inscrit_dps"],
            "inscrit_gar": s["inscrit_gar"],
            "activites_dps": adps,
            "activites_gar": agar,
        })

    now = datetime.now()
    output = {
        "secouristes":  secouristes,
        "lastModified": now.strftime("%d/%m/%Y"),
        "minDate":      min_date.strftime("%d/%m/%Y") if min_date else "—",
        "maxDate":      max_date.strftime("%d/%m/%Y") if max_date else "—",
        "totalRows":    len(secouristes),
    }
    json_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  OK — {len(secouristes)} secouristes → {json_path}")
    return True

def main():
    convert_activity(Path("data/export_dps.xlsx"),   Path("data_dps.json"),   "dps")
    convert_activity(Path("data/export_garde.xlsx"), Path("data_garde.json"), "garde")
    convert_sr(
        Path("data/export_competences.xlsx"),
        Path("data/export_participations.xlsx"),
        Path("data_sr.json")
    )
    print("Conversion terminée.")

if __name__ == "__main__":
    main()
