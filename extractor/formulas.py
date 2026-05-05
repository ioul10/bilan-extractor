# ============================================================
# formulas.py
# Règles comptables du bilan marocain (Modèle Normal loi 9-88)
# Source : MODELE.xlsx exporté via VBA
#
# Structure :
#   - ACTIF_RULES  : règles de calcul pour chaque cellule
#   - PASSIF_RULES : idem
#   - CPC_RULES    : idem
#   - EQUILIBRE    : règle d'équilibre fondamental
#
# Convention colonnes :
#   Actif  → B=Brut | C=Amort&Prov | D=Net N | E=Net N-1
#   Passif → B=Exercice N | C=Exercice N-1
#   CPC    → C=Propres | D=Précédents | E=Total N | F=Total N-1
# ============================================================


# ── MAPPING LIGNE → LIBELLÉ OFFICIEL ────────────────────────
# Chaque numéro de ligne correspond à une rubrique du modèle

ACTIF_LABELS = {
    # Actif immobilisé
    6:  "Immobilisations en non valeurs [A]",
    7:  "Frais préliminaires",
    8:  "Charges à répartir sur plusieurs exercices",
    9:  "Primes de remboursement des obligations",
    10: "Immobilisations incorporelles [B]",
    11: "Immobilisations en Recherche et Développement",
    12: "Brevets, marques, droits et valeurs similaires",
    13: "Fonds commercial",
    14: "Autres immobilisations incorporelles",
    15: "Immobilisations corporelles [C]",
    16: "Terrains",
    17: "Constructions",
    18: "Installations techniques, matériel et outillage",
    19: "Matériel de transport",
    20: "Mobilier, Mat. de bureau, Aménagement divers",
    21: "Autres immobilisations corporelles",
    22: "Immobilisations corporelles en cours",
    23: "Immobilisations financières [D]",
    24: "Prêts immobilisés",
    25: "Autres créances financières",
    26: "Titres de participation",
    27: "Autres titres immobilisés",
    28: "Écarts de conversion actif [E]",
    29: "Diminution des créances immobilisées",
    30: "Augmentations des dettes de financement",
    31: "TOTAL I (A+B+C+D+E)",
    # Actif circulant
    32: "Stocks [F]",
    33: "Marchandises",
    34: "Matières et fournitures consommables",
    35: "Produits en cours",
    36: "Produits intermédiaires et produits résiduels",
    37: "Produits finis",
    38: "Créances de l'actif circulant [G]",
    39: "Fournisseurs débiteurs, avances et acomptes",
    40: "Clients et comptes rattachés",
    41: "Personnel",
    42: "État",
    43: "Comptes d'associés",
    44: "Autres débiteurs",
    45: "Comptes de régularisation - Actif",
    46: "Titres et valeurs de placement [H]",
    47: "Écarts de conversion actif [I]",
    48: "TOTAL II (F+G+H+I)",
    # Trésorerie
    49: "Trésorerie - Actif",
    50: "Chèques et valeurs à encaisser",
    51: "Banques, T.G et C.C.P",
    52: "Caisse, Régie d'avances et accréditifs",
    53: "TOTAL III",
    54: "TOTAL GÉNÉRAL I+II+III",
}

PASSIF_LABELS = {
    # Financement permanent
    5:  "CAPITAUX PROPRES",
    6:  "Capital social ou personnel",
    7:  "Capital social ou personnel",
    8:  "Moins : actionnaires, capital souscrit non appelé",
    9:  "Capital appelé",
    10: "Dont versé",
    11: "Prime d'émission, de fusion, d'apport",
    12: "Écarts de réévaluation",
    13: "Réserve légale",
    14: "Autres réserves",
    15: "Report à nouveau",
    16: "Résultat en instance d'affectation",
    17: "Résultat net de l'exercice",
    18: "Total des capitaux propres (A)",
    19: "Capitaux propres assimilés (B)",
    20: "Subvention d'investissement",
    21: "Provisions réglementées",
    22: "Dettes de financement (C)",
    23: "Emprunts obligataires",
    24: "Autres dettes de financement",
    25: "Provisions durables pour risques et charges (D)",
    26: "Provisions pour risques",
    27: "Provisions pour charges",
    28: "Écarts de conversion passif (E)",
    29: "Augmentation des créances immobilisées",
    30: "Diminution des dettes de financement",
    31: "TOTAL I (A+B+C+D+E)",
    # Passif circulant
    32: "Dettes du passif circulant (F)",
    33: "Fournisseurs et comptes rattachés",
    34: "Clients créditeurs, avances et acomptes",
    35: "Personnel",
    36: "Organismes sociaux",
    37: "État",
    38: "Comptes d'associés",
    39: "Autres créanciers",
    40: "Comptes de régularisation passif",
    41: "Autres provisions pour risques et charges (G)",
    42: "Écarts de conversion passif (H)",
    43: "TOTAL II (F+G+H)",
    # Trésorerie
    44: "Trésorerie Passif",
    45: "Crédits d'escompte",
    46: "Crédits de trésorerie",
    47: "Banques (soldes créditeurs)",
    48: "TOTAL III",
    49: "TOTAL GÉNÉRAL I+II+III",
}

CPC_LABELS = {
    # Produits exploitation
    5:  "PRODUITS D'EXPLOITATION",
    6:  "Ventes de marchandises (en l'état)",
    7:  "Ventes de marchandises (en l'état)",
    8:  "Ventes de biens et services produits",
    9:  "Chiffre d'affaires",
    10: "Variation de stocks de produits",
    11: "Immobilisations produites par l'entreprise",
    12: "Subventions d'exploitation",
    13: "Autres produits d'exploitation",
    14: "Reprises d'exploitation ; transferts de charges",
    15: "TOTAL I — Produits d'exploitation",
    # Charges exploitation
    16: "CHARGES D'EXPLOITATION",
    17: "Achats revendus de marchandises",
    18: "Achats consommés de matières et fournitures",
    19: "Autres charges externes",
    20: "Impôts et taxes",
    21: "Charges de personnel",
    22: "Autres charges d'exploitation",
    23: "Dotations d'exploitation",
    24: "TOTAL II — Charges d'exploitation",
    25: "RÉSULTAT D'EXPLOITATION (I-II)",
    # Financier
    26: "PRODUITS FINANCIERS",
    27: "Produits des titres de participation",
    28: "Gains de change",
    29: "Intérêts et autres produits financiers",
    30: "Reprises financières ; transferts de charges",
    31: "TOTAL IV — Produits financiers",
    32: "CHARGES FINANCIÈRES",
    33: "Charges d'intérêts",
    34: "Pertes de change",
    35: "Autres charges financières",
    36: "Dotations financières",
    37: "TOTAL V — Charges financières",
    38: "RÉSULTAT FINANCIER (IV-V)",
    39: "RÉSULTAT COURANT (III+VI)",
    # Non courant
    40: "PRODUITS NON COURANTS",
    41: "Produits des cessions d'immobilisations",
    42: "Subventions d'équilibre",
    43: "Reprises sur subventions d'investissement",
    44: "Autres produits non courants",
    45: "Reprises non courantes ; transferts de charges",
    46: "TOTAL VIII — Produits non courants",
    47: "CHARGES NON COURANTES",
    48: "Valeurs nettes d'amortissement des immob. cédées",
    49: "Subventions accordées",
    50: "Autres charges non courantes",
    51: "Dotations non courantes",
    52: "TOTAL IX — Charges non courantes",
    53: "RÉSULTAT NON COURANT (VIII-IX)",
    54: "RÉSULTAT AVANT IMPÔTS (VII+X)",
    55: "IMPÔTS SUR LES RÉSULTATS",
    56: "RÉSULTAT NET (XI-XII)",
    57: "TOTAL DES PRODUITS (I+IV+VIII)",
    58: "TOTAL DES CHARGES (II+V+IX+XII)",
    59: "RÉSULTAT NET — Vérification (XIV-XV)",
}


# ── RÈGLES DE CALCUL ────────────────────────────────────────
# Format : { ligne: ("type", paramètres) }
#
# Types :
#   "sum_range"  → somme d'une plage de lignes
#   "sum_lines"  → somme de lignes spécifiques
#   "net"        → B - C (Brut - Amort)
#   "diff"       → ligne_a - ligne_b
#   "add"        → ligne_a + ligne_b
#   "equal"      → copie d'une autre ligne

ACTIF_RULES = {
    # ── NET = BRUT - AMORT (colonne D) ──────────────────────
    7:  ("net", {}),
    8:  ("net", {}),
    9:  ("net", {}),
    11: ("net", {}),
    12: ("net", {}),
    13: ("net", {}),
    14: ("net", {}),
    16: ("net", {}),
    17: ("net", {}),
    18: ("net", {}),
    19: ("net", {}),
    20: ("net", {}),
    21: ("net", {}),
    22: ("net", {}),
    24: ("net", {}),
    26: ("net", {}),
    27: ("net", {}),
    29: ("net", {}),
    30: ("net", {}),
    33: ("net", {}),
    34: ("net", {}),
    35: ("net", {}),
    36: ("net", {}),
    37: ("net", {}),
    39: ("net", {}),
    40: ("net", {}),
    41: ("net", {}),
    42: ("net", {}),
    43: ("net", {}),
    44: ("net", {}),
    45: ("net", {}),
    46: ("net", {}),
    47: ("net", {}),
    50: ("net", {}),
    51: ("net", {}),
    52: ("net", {}),

    # ── TOTAUX RUBRIQUES (toutes colonnes) ──────────────────
    6:  ("sum_range", {"from": 7,  "to": 9}),   # [A]
    10: ("sum_range", {"from": 11, "to": 14}),  # [B]
    15: ("sum_range", {"from": 16, "to": 22}),  # [C]
    23: ("sum_range", {"from": 24, "to": 27}),  # [D]
    28: ("sum_range", {"from": 29, "to": 30}),  # [E]
    32: ("sum_range", {"from": 33, "to": 37}),  # [F] Stocks
    38: ("sum_range", {"from": 39, "to": 45}),  # [G] Créances
    49: ("sum_range", {"from": 50, "to": 52}),  # Trésorerie

    # ── TOTAUX PRINCIPAUX ────────────────────────────────────
    31: ("sum_lines", {"lines": [6, 10, 15, 23, 28]}),   # TOTAL I
    48: ("sum_lines", {"lines": [32, 38, 46, 47]}),      # TOTAL II
    53: ("equal",     {"line": 49}),                      # TOTAL III
    54: ("sum_lines", {"lines": [31, 48, 53]}),           # TOTAL GÉNÉRAL
}

PASSIF_RULES = {
    # ── TOTAUX RUBRIQUES ────────────────────────────────────
    18: ("sum_lines", {"lines": [7, 12, 13, 14, 15, 16, 17]}),  # (A)
    19: ("sum_lines", {"lines": [20, 21]}),                      # (B)
    22: ("sum_lines", {"lines": [23, 24]}),                      # (C)
    25: ("sum_lines", {"lines": [26, 27]}),                      # (D)
    28: ("sum_lines", {"lines": [29, 30]}),                      # (E)
    32: ("sum_range", {"from": 33, "to": 40}),                   # (F)

    # ── TOTAUX PRINCIPAUX ────────────────────────────────────
    31: ("sum_lines", {"lines": [18, 19, 22, 25, 28]}),  # TOTAL I
    43: ("sum_lines", {"lines": [32, 41, 42]}),           # TOTAL II
    44: ("sum_range", {"from": 45, "to": 47}),            # Trésorerie
    48: ("equal",     {"line": 44}),                      # TOTAL III
    49: ("sum_lines", {"lines": [31, 43, 48]}),           # TOTAL GÉNÉRAL
}

CPC_RULES = {
    # ── TOTAUX ──────────────────────────────────────────────
    15: ("sum_range", {"from": 7,  "to": 14}),  # I  — Produits exploit.
    24: ("sum_range", {"from": 17, "to": 23}),  # II — Charges exploit.
    31: ("sum_range", {"from": 27, "to": 30}),  # IV — Produits financiers
    37: ("sum_range", {"from": 33, "to": 36}),  # V  — Charges financières
    46: ("sum_range", {"from": 41, "to": 45}),  # VIII — Produits NC
    52: ("sum_range", {"from": 48, "to": 51}),  # IX  — Charges NC

    # ── RÉSULTATS ───────────────────────────────────────────
    25: ("diff", {"a": 15, "b": 24}),            # III — Résultat exploit.
    38: ("diff", {"a": 31, "b": 37}),            # VI  — Résultat financier
    39: ("add",  {"a": 25, "b": 38}),            # VII — Résultat courant
    53: ("diff", {"a": 46, "b": 52}),            # X   — Résultat NC
    54: ("add",  {"a": 39, "b": 53}),            # XI  — Résultat avant IS
    56: ("diff", {"a": 54, "b": 55}),            # XIII — Résultat net

    # ── TOTAUX GÉNÉRAUX ─────────────────────────────────────
    57: ("sum_lines", {"lines": [15, 31, 46]}),          # XIV — Total produits
    58: ("sum_lines", {"lines": [24, 37, 52, 55]}),      # XV  — Total charges
    59: ("diff", {"a": 57, "b": 58}),                    # XVI — Vérification
}


# ── ÉQUILIBRE FONDAMENTAL ────────────────────────────────────
EQUILIBRE = {
    "actif_ligne":   54,   # TOTAL GÉNÉRAL Actif  (colonne D = Net N)
    "passif_ligne":  49,   # TOTAL GÉNÉRAL Passif (colonne B = Exercice N)
    "tolerance":     1.0,  # tolérance d'arrondi en MAD
}


# ── VÉRIFICATIONS CROISÉES ───────────────────────────────────
# Valeurs qui doivent être identiques entre sections
CROSS_CHECKS = [
    {
        "label":       "Résultat net CPC = Résultat net Passif",
        "cpc_ligne":   56,   # XIII Résultat net CPC    (col E = Total N)
        "passif_ligne": 17,  # Résultat net de l'exercice Passif (col B)
        "tolerance":   1.0,
    }
]


# ============================================================
# MOTEUR DE CALCUL
# ============================================================

def compute(rules: dict, data: dict) -> dict:
    """
    Applique les règles de calcul sur les données extraites.

    data   : { ligne: {"B": val, "C": val, "D": val, "E": val} }
             ou pour Passif : { ligne: {"B": val, "C": val} }
             ou pour CPC    : { ligne: {"C": val, "D": val, "E": val, "F": val} }

    Retourne : { ligne: {"calculé": val, "extrait": val, "ecart": val} }
    """
    results = {}

    for ligne, (rule_type, params) in rules.items():
        extrait  = _get_val(data, ligne)
        calcule  = _apply_rule(rule_type, params, data, ligne)

        if calcule is not None:
            ecart = abs(extrait - calcule) if extrait is not None else None
            results[ligne] = {
                "calculé": calcule,
                "extrait": extrait,
                "écart":   ecart,
                "ok":      ecart is None or ecart <= 1.0,
            }

    return results


def _get_val(data: dict, ligne: int, col: str = None) -> float | None:
    """Récupère une valeur dans data. Col par défaut = première colonne valeur."""
    if ligne not in data:
        return None
    row = data[ligne]
    if col:
        return row.get(col)
    # Priorité : D (Net N) pour Actif, B (Exercice N) pour Passif/CPC Total N
    for c in ["D", "E", "B", "C"]:
        if c in row and row[c] is not None:
            return row[c]
    return None


def _apply_rule(rule_type: str, params: dict,
                data: dict, ligne: int) -> float | None:
    """Applique une règle et retourne la valeur calculée."""
    try:
        if rule_type == "net":
            b = _get_val(data, ligne, "B")
            c = _get_val(data, ligne, "C")
            if b is not None and c is not None:
                return b - c
            return None

        elif rule_type == "sum_range":
            total = 0.0
            found = False
            for r in range(params["from"], params["to"] + 1):
                v = _get_val(data, r)
                if v is not None:
                    total += v
                    found = True
            return total if found else None

        elif rule_type == "sum_lines":
            total = 0.0
            found = False
            for r in params["lines"]:
                v = _get_val(data, r)
                if v is not None:
                    total += v
                    found = True
            return total if found else None

        elif rule_type == "diff":
            a = _get_val(data, params["a"])
            b = _get_val(data, params["b"])
            if a is not None and b is not None:
                return a - b
            return None

        elif rule_type == "add":
            a = _get_val(data, params["a"])
            b = _get_val(data, params["b"])
            if a is not None and b is not None:
                return a + b
            return None

        elif rule_type == "equal":
            return _get_val(data, params["line"])

    except Exception:
        return None

    return None


def verify_equilibre(actif_data: dict, passif_data: dict) -> dict:
    """
    Vérifie l'équilibre fondamental : Total Actif = Total Passif.
    Retourne un dict avec le résultat.
    """
    total_actif  = _get_val(actif_data,  EQUILIBRE["actif_ligne"],  "D")
    total_passif = _get_val(passif_data, EQUILIBRE["passif_ligne"], "B")

    if total_actif is None or total_passif is None:
        return {"ok": None, "message": "Totaux non disponibles"}

    ecart = abs(total_actif - total_passif)
    ok    = ecart <= EQUILIBRE["tolerance"]

    return {
        "ok":           ok,
        "total_actif":  total_actif,
        "total_passif": total_passif,
        "écart":        ecart,
        "message":      "✅ ÉQUILIBRE OK" if ok else f"❌ DÉSÉQUILIBRE — écart : {ecart:,.2f}",
    }


def verify_cross_checks(cpc_data: dict, passif_data: dict) -> list:
    """
    Vérifie les cohérences croisées entre sections.
    Ex: Résultat net CPC = Résultat net Passif
    """
    results = []
    for check in CROSS_CHECKS:
        val_cpc    = _get_val(cpc_data,    check["cpc_ligne"],    "E")
        val_passif = _get_val(passif_data, check["passif_ligne"], "B")

        if val_cpc is None or val_passif is None:
            results.append({
                "label": check["label"],
                "ok":    None,
                "message": "Valeurs non disponibles"
            })
            continue

        ecart = abs(val_cpc - val_passif)
        ok    = ecart <= check["tolerance"]
        results.append({
            "label":      check["label"],
            "ok":         ok,
            "val_cpc":    val_cpc,
            "val_passif": val_passif,
            "écart":      ecart,
            "message":    "✅ OK" if ok else f"❌ Écart : {ecart:,.2f}",
        })

    return results
