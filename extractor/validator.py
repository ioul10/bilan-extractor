# ============================================================
# validator.py
# Valide la cohérence comptable des données extraites
#
# Pipeline :
#   1. mapped_rows → data_dict (ligne → valeurs numériques)
#   2. compute()   → recalcule chaque total selon les règles
#   3. compare     → extrait vs calculé
#   4. rapport     → ✅ / ❌ par ligne + équilibre global
# ============================================================

import re
from extractor.formulas import (
    ACTIF_RULES, PASSIF_RULES, CPC_RULES,
    ACTIF_LABELS, PASSIF_LABELS, CPC_LABELS,
    compute, verify_equilibre, verify_cross_checks,
)


# ── CONVERSION VALEUR NUMÉRIQUE ──────────────────────────────

def parse_number(val: str) -> float | None:
    """
    Convertit une valeur extraite du PDF en float.
    Formats supportés :
      - "1.234.567,89"  (format marocain)
      - "1 234 567,89"  (espaces)
      - "-1.234.567,89" (négatif)
      - "(1.234.567,89)" (négatif entre parenthèses)
    """
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ("-", "—", ""):
        return None

    # Négatif entre parenthèses
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    if s.startswith("-"):
        neg, s = True, s[1:]

    # Supprimer espaces et points de milliers
    s = s.replace("\xa0", "").replace(" ", "")

    # Format marocain : 1.234.567,89
    # → détecter si la virgule est le séparateur décimal
    if "," in s and "." in s:
        # ex: 1.234.567,89 → supprimer points, remplacer virgule
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        # ex: 218.585,17 ou 218585,17
        s = s.replace(",", ".")
    elif "." in s and "," not in s:
        # ex: 218585.17 (déjà format anglais)
        pass
    # else: entier pur

    try:
        result = float(s)
        return -result if neg else result
    except ValueError:
        return None


# ── CONSTRUCTION DU DATA DICT ────────────────────────────────

def build_data_dict(mapped_rows: list[dict], section: str) -> dict:
    """
    Convertit les lignes mappées en dictionnaire de données.

    mapped_rows : sortie de map_section()
    section     : "Actif" | "Passif" | "CPC"

    Retourne :
    {
        ligne: {col: float, ...}
    }

    Colonnes selon section :
      Actif  → B=Brut | C=Amort | D=Net N | E=Net N-1
      Passif → B=Exercice N | C=Exercice N-1
      CPC    → C=Propres | D=Précédents | E=Total N | F=Total N-1
    """
    if section == "Actif":
        col_names = ["B", "C", "D", "E"]
    elif section == "Passif":
        col_names = ["B", "C"]
    elif section == "CPC":
        col_names = ["C", "D", "E", "F"]
    else:
        return {}

    data = {}
    seen_lines = {}  # Pour gérer les doublons (même ligne mappée 2x)

    for row in mapped_rows:
        ligne = row.get("ligne")
        if ligne is None:
            continue

        values = row.get("values", [])

        # Parser les valeurs numériques
        parsed = {}
        for i, col in enumerate(col_names):
            val = values[i] if i < len(values) else None
            num = parse_number(val) if val else None
            if num is not None:
                parsed[col] = num

        if not parsed:
            continue

        # Gérer les doublons : garder celui avec le meilleur score
        if ligne in seen_lines:
            prev_score = seen_lines[ligne]
            curr_score = row.get("score", 0)
            if curr_score <= prev_score:
                continue  # On garde l'ancien

        data[ligne] = parsed
        seen_lines[ligne] = row.get("score", 0)

    return data


# ── VALIDATION ───────────────────────────────────────────────

TOLERANCE = 2.0  # Tolérance d'arrondi en MAD


def validate_section(data: dict, rules: dict, labels: dict,
                     section: str) -> dict:
    """
    Valide une section en comparant valeurs extraites vs calculées.

    Retourne un rapport structuré.
    """
    results = compute(rules, data)

    ligne_reports = []
    nb_ok    = 0
    nb_error = 0
    nb_miss  = 0

    for ligne, res in results.items():
        label    = labels.get(ligne, f"Ligne {ligne}")
        extrait  = res["extrait"]
        calcule  = res["calculé"]
        ecart    = res["écart"]

        if extrait is None:
            status = "missing"
            nb_miss += 1
        elif ecart is not None and ecart <= TOLERANCE:
            status = "ok"
            nb_ok += 1
        else:
            status = "error"
            nb_error += 1

        ligne_reports.append({
            "ligne":    ligne,
            "label":    label,
            "extrait":  extrait,
            "calculé":  calcule,
            "écart":    ecart,
            "status":   status,
        })

    # Trier par numéro de ligne
    ligne_reports.sort(key=lambda x: x["ligne"])

    return {
        "section":   section,
        "nb_ok":     nb_ok,
        "nb_error":  nb_error,
        "nb_miss":   nb_miss,
        "total":     len(results),
        "score":     round(nb_ok / len(results) * 100, 1) if results else 0,
        "lignes":    ligne_reports,
    }


def validate_all(actif_mapped:  list[dict],
                 passif_mapped: list[dict],
                 cpc_mapped:    list[dict]) -> dict:
    """
    Validation complète : Actif + Passif + CPC + Équilibre.

    Retourne un rapport global.
    """
    # Construire les dicts de données
    actif_data  = build_data_dict(actif_mapped,  "Actif")
    passif_data = build_data_dict(passif_mapped, "Passif")
    cpc_data    = build_data_dict(cpc_mapped,    "CPC")

    # Valider chaque section
    actif_report  = validate_section(actif_data,  ACTIF_RULES,
                                     ACTIF_LABELS,  "Actif")
    passif_report = validate_section(passif_data, PASSIF_RULES,
                                     PASSIF_LABELS, "Passif")
    cpc_report    = validate_section(cpc_data,    CPC_RULES,
                                     CPC_LABELS,    "CPC")

    # Équilibre fondamental
    equilibre = verify_equilibre(actif_data, passif_data)

    # Vérifications croisées
    cross = verify_cross_checks(cpc_data, passif_data)

    # Score global
    total_ok    = (actif_report["nb_ok"] +
                   passif_report["nb_ok"] +
                   cpc_report["nb_ok"])
    total_rules = (actif_report["total"] +
                   passif_report["total"] +
                   cpc_report["total"])
    score_global = round(total_ok / total_rules * 100, 1) if total_rules else 0

    return {
        "score_global": score_global,
        "equilibre":    equilibre,
        "cross_checks": cross,
        "actif":        actif_report,
        "passif":       passif_report,
        "cpc":          cpc_report,
        "data": {
            "actif":  actif_data,
            "passif": passif_data,
            "cpc":    cpc_data,
        }
    }


# ── AFFICHAGE RAPPORT ────────────────────────────────────────

def print_report(rapport: dict):
    """Affiche le rapport de validation dans le terminal."""

    print("\n" + "=" * 60)
    print(f"  RAPPORT DE VALIDATION BILAN")
    print(f"  Score global : {rapport['score_global']}%")
    print("=" * 60)

    # Équilibre
    eq = rapport["equilibre"]
    print(f"\n{'─'*40}")
    print(f"  ÉQUILIBRE FONDAMENTAL")
    print(f"{'─'*40}")
    print(f"  {eq['message']}")
    if eq.get("total_actif"):
        print(f"  Actif  : {eq['total_actif']:>20,.2f}")
        print(f"  Passif : {eq['total_passif']:>20,.2f}")

    # Vérifications croisées
    if rapport["cross_checks"]:
        print(f"\n{'─'*40}")
        print(f"  VÉRIFICATIONS CROISÉES")
        print(f"{'─'*40}")
        for check in rapport["cross_checks"]:
            print(f"  {check['message']} — {check['label']}")

    # Sections
    for section_key in ["actif", "passif", "cpc"]:
        rep = rapport[section_key]
        print(f"\n{'─'*40}")
        print(f"  {rep['section'].upper()} — Score : {rep['score']}%")
        print(f"  ✅ {rep['nb_ok']} OK  |  "
              f"❌ {rep['nb_error']} Erreurs  |  "
              f"⚠️  {rep['nb_miss']} Manquants")
        print(f"{'─'*40}")

        for l in rep["lignes"]:
            icon = {"ok": "✅", "error": "❌", "missing": "⚠️ "}.get(
                l["status"], "  ")
            label = l["label"][:38]
            if l["status"] == "ok":
                print(f"  {icon} L{l['ligne']:2d} {label:<40} "
                      f"{l['calculé']:>18,.2f}")
            elif l["status"] == "error":
                print(f"  {icon} L{l['ligne']:2d} {label:<40} "
                      f"extrait={l['extrait']:>15,.2f}  "
                      f"calculé={l['calculé']:>15,.2f}  "
                      f"écart={l['écart']:>10,.2f}")
            else:
                print(f"  {icon} L{l['ligne']:2d} {label:<40} "
                      f"(non extrait — calculé={l['calculé']:>15,.2f})")

    print("\n" + "=" * 60)
