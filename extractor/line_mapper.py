# ============================================================
# line_mapper.py
# Fait le lien entre les libellés extraits du PDF
# et les numéros de lignes du modèle (ACTIF_LABELS, etc.)
#
# Stratégie en 3 niveaux :
#   1. Match exact après normalisation
#   2. Match par mots-clés
#   3. Match flou (score de similarité)
# ============================================================

import re
import unicodedata
from extractor.formulas import ACTIF_LABELS, PASSIF_LABELS, CPC_LABELS


# ── NETTOYAGE ────────────────────────────────────────────────

def normalize(text: str) -> str:
    """
    Normalise un libellé pour comparaison :
    - Minuscules
    - Supprime accents
    - Supprime ponctuation et caractères parasites
    - Supprime les lettres isolées en début (A, B, C, I, etc.)
    - Supprime les préfixes numériques (I, II, III, IV...)
    - Supprime les points, tirets, astérisques en début
    - Compresse les espaces
    """
    if not text:
        return ""

    t = text.strip()

    # Supprimer les lettres/chiffres romains isolés en début
    # ex: "I IMMOBILISATIONS" → "IMMOBILISATIONS"
    # ex: "A . Frais" → "Frais"
    # ex: "E -Ventes" → "Ventes"
    t = re.sub(r'^[A-Z]{1,4}\s*[\.\-\*]?\s*', '', t)

    # Supprimer préfixes romains restants (I, II, III, IV, V, VI...)
    t = re.sub(r'^(XIV|XIII|XII|XI|X|IX|VIII|VII|VI|V|IV|III|II|I)\s+', '', t)

    # Supprimer * et . en début
    t = re.sub(r'^[\*\.\-\(\)]+\s*', '', t)

    # Supprimer les notes entre parenthèses en fin : (1), (2), (+/-), etc.
    t = re.sub(r'\s*\([^)]*\)\s*$', '', t)
    t = re.sub(r'\s*\[\w\]\s*$', '', t)  # [A], [B]...

    # Supprimer les numéros en fin : ligne 6.1, 6.2
    t = re.sub(r'\s*\d+\.\d+\s*$', '', t)

    # Minuscules
    t = t.lower()

    # Supprimer accents
    t = unicodedata.normalize('NFD', t)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')

    # Supprimer ponctuation sauf espaces
    t = re.sub(r'[^\w\s]', ' ', t)

    # Compresser espaces
    t = re.sub(r'\s+', ' ', t).strip()

    return t


# ── DICTIONNAIRE D'ABRÉVIATIONS ──────────────────────────────
# Remplace les abréviations courantes avant normalisation

ABBREVIATIONS = {
    r'\bimmob\.?\b':          'immobilisations',
    r'\bmat\.?\b':            'materiel',
    r'\bamenag\.?\b':         'amenagement',
    r'\bamenagement\b':       'amenagement',
    r'\bregularis\.?\b':      'regularisation',
    r'\bregularisa\.?\b':     'regularisation',
    r'\bcptes?\b':            'comptes',
    r'\bfournis\.?\b':        'fournisseurs',
    r'\bprov\.?\b':           'provisions',
    r'\bamort\.?\b':          'amortissements',
    r'\bval\.?\b':            'valeurs',
    r'\bdts?\b':              'dettes',
    r'\bcap\.?\b':            'capital',
    r'\bexploit\.?\b':        'exploitation',
    r'\bfin\.?\b':            'financement',
    r'\bimmo\.?\b':           'immobilisations',
    r'\btresorerie\b':        'tresorerie',
    r'\btresor\.?\b':         'tresorerie',
    r'\bsoc\.?\b':            'sociaux',
    r'\bpers\.?\b':           'personnel',
    r'\breint\.?\b':          'reintegrations',
    r'\brepris\.?\b':         'reprises',
    r'\bcession\b':           'cessions',
    r'\bnet\b':               'net',
    r'\bcourant\b':           'courant',
    r'\bprecedent\b':         'precedent',
    r'\bexerc\.?\b':          'exercice',
}


def expand_abbreviations(text: str) -> str:
    """Remplace les abréviations par les termes complets."""
    t = text.lower()
    for pattern, replacement in ABBREVIATIONS.items():
        t = re.sub(pattern, replacement, t, flags=re.IGNORECASE)
    return t


def full_normalize(text: str) -> str:
    """Normalisation complète : expand + normalize."""
    return normalize(expand_abbreviations(text))


# ── INDEX DE RECHERCHE ───────────────────────────────────────
# Construit un index normalisé → ligne pour chaque section

def _build_index(labels: dict) -> dict:
    """Construit l'index normalisé → numéro de ligne."""
    return {full_normalize(label): line for line, label in labels.items()}


ACTIF_INDEX  = _build_index(ACTIF_LABELS)
PASSIF_INDEX = _build_index(PASSIF_LABELS)
CPC_INDEX    = _build_index(CPC_LABELS)


# ── MOTS-CLÉS PRIORITAIRES ───────────────────────────────────
# Pour les cas où le match exact échoue
# Format : { mot_clé: numéro_ligne }

ACTIF_KEYWORDS = {
    # Actif immobilisé
    "non valeurs":                  6,
    "frais preliminaires":          7,
    "charges repartir":             8,
    "primes remboursement":         9,
    "incorporelles":                10,
    "recherche developpement":      11,
    "brevets marques":              12,
    "fonds commercial":             13,
    "autres immobilisations incorporelles": 14,
    "corporelles":                  15,
    "terrains":                     16,
    "constructions":                17,
    "installations techniques":     18,
    "materiel transport":           19,
    "mobilier":                     20,
    "autres immobilisations corporelles": 21,
    "immobilisations en cours":     22,
    "financieres":                  23,
    "prets immobilises":            24,
    "autres creances financieres":  25,
    "titres participation":         26,
    "autres titres":                27,
    "ecarts conversion actif":      28,
    "diminution creances":          29,
    "augmentation dettes":          30,
    "total i":                      31,
    # Actif circulant
    "stocks":                       32,
    "marchandises":                 33,
    "matieres fournitures":         34,
    "produits en cours":            35,
    "produits intermediaires":      36,
    "produits finis":               37,
    "creances actif circulant":     38,
    "fournisseurs debiteurs":       39,
    "clients comptes rattaches":    40,
    "personnel":                    41,
    "etat":                         42,
    "comptes associes":             43,
    "autres debiteurs":             44,
    "comptes regularisation actif": 45,
    "titres valeurs placement":     46,
    "ecarts conversion actif elements": 47,
    "total ii":                     48,
    # Trésorerie
    "tresorerie actif":             49,
    "cheques valeurs":              50,
    "banques":                      51,
    "caisse regie":                 52,
    "total iii":                    53,
    "total general":                54,
}

PASSIF_KEYWORDS = {
    "capitaux propres":             18,
    "capital social":               7,
    "moins actionnaires":           8,
    "prime emission":               11,
    "ecarts reevaluation":          12,
    "reserve legale":               13,
    "autres reserves":              14,
    "report nouveau":               15,
    "resultats instance":           16,
    "resultat net exercice":        17,
    "total capitaux propres":       18,
    "capitaux propres assimiles":   19,
    "subventions investissement":   20,
    "provisions reglementees":      21,
    "dettes financement":           22,
    "emprunts obligataires":        23,
    "autres dettes financement":    24,
    "provisions durables":          25,
    "provisions risques":           26,
    "provisions charges":           27,
    "ecarts conversion passif":     28,
    "augmentation creances":        29,
    "diminution dettes":            30,
    "total i":                      31,
    "dettes passif circulant":      32,
    "fournisseurs comptes":         33,
    "clients crediteurs":           34,
    "personnel":                    35,
    "organismes sociaux":           36,
    "etat":                         37,
    "comptes associes":             38,
    "autres creanciers":            39,
    "comptes regularisation passif": 40,
    "autres provisions":            41,
    "ecarts conversion passif elements": 42,
    "total ii":                     43,
    "tresorerie passif":            44,
    "credits escompte":             45,
    "credits tresorerie":           46,
    "banques soldes":               47,
    "total iii":                    48,
    "total general":                49,
}

CPC_KEYWORDS = {
    "produits exploitation":        5,
    "ventes marchandises":          7,
    "ventes biens services":        8,
    "chiffre affaires":             9,
    "variation stocks produits":    10,
    "immobilisations produites":    11,
    "subventions exploitation":     12,
    "autres produits exploitation": 13,
    "reprises exploitation":        14,
    "total i":                      15,
    "charges exploitation":         16,
    "achats revendus marchandises": 17,
    "achats consommes matieres":    18,
    "autres charges externes":      19,
    "impots taxes":                 20,
    "charges personnel":            21,
    "autres charges exploitation":  22,
    "dotations exploitation":       23,
    "total ii":                     24,
    "resultat exploitation":        25,
    "produits financiers":          26,
    "produits titres participation": 27,
    "gains change":                 28,
    "interets autres produits":     29,
    "reprises financieres":         30,
    "total iv":                     31,
    "charges financieres":          32,
    "charges interets":             33,
    "pertes change":                34,
    "autres charges financieres":   35,
    "dotations financieres":        36,
    "total v":                      37,
    "resultat financier":           38,
    "resultat courant":             39,
    "produits non courants":        40,
    "cessions immobilisations":     41,
    "subventions equilibre":        42,
    "reprises subventions":         43,
    "autres produits non courants": 44,
    "reprises non courantes":       45,
    "total viii":                   46,
    "charges non courantes":        47,
    "valeurs nettes amort":         48,
    "subventions accordees":        49,
    "autres charges non courantes": 50,
    "dotations non courantes":      51,
    "total ix":                     52,
    "resultat non courant":         53,
    "resultat avant impots":        54,
    "impots resultats":             55,
    "resultat net":                 56,
    "total produits":               57,
    "total charges":                58,
    "resultat net verification":    59,
}


# ── LIGNES À IGNORER ─────────────────────────────────────────
# Libellés qui ne correspondent à aucune ligne du modèle

IGNORE_PATTERNS = [
    r'^\d+$',                    # Numéros seuls
    r'^page\s*\d',               # "Page 1/2"
    r'^rapport\s+financier',     # En-têtes
    r'^\d+\.\d+$',              # "6.1", "6.2"
    r'^bilan\s+au',              # "Bilan au 31..."
    r'^compte\s+de\s+produits',  # Titre CPC
    r'^nature$',                 # Header colonne
    r'^exercice',                # Header colonne
    r'^\*\s+',                   # Sous-détails non modélisés
    r'^dont\s+verse',
    r'^\(1\)',
    r'^\(2\)',
    r'^office\s+national',
    r'^www\.',
    r'^conforme',
]


def should_ignore(text: str) -> bool:
    """Retourne True si le libellé doit être ignoré."""
    t = text.strip().lower()
    if len(t) <= 2:
        return True
    for pattern in IGNORE_PATTERNS:
        if re.match(pattern, t, re.IGNORECASE):
            return True
    return False


# ── MATCHING ─────────────────────────────────────────────────

def _score(a: str, b: str) -> float:
    """
    Score de similarité entre deux chaînes normalisées.
    Basé sur les mots communs / mots totaux.
    """
    words_a = set(a.split())
    words_b = set(b.split())
    if not words_a or not words_b:
        return 0.0
    intersection = words_a & words_b
    union = words_a | words_b
    # Jaccard
    jaccard = len(intersection) / len(union)
    # Bonus si tous les mots de b sont dans a
    coverage = len(intersection) / len(words_b) if words_b else 0
    return (jaccard + coverage) / 2


def match_line(text: str, section: str) -> dict:
    """
    Trouve la ligne du modèle correspondant au libellé extrait.

    Retourne :
    {
        "ligne":      int | None,
        "label":      str,         # libellé officiel
        "score":      float,       # 1.0 = exact, 0.0 = aucun match
        "method":     str,         # "exact" | "keyword" | "fuzzy" | "none"
        "original":   str,
    }
    """
    if should_ignore(text):
        return {"ligne": None, "label": None, "score": 0,
                "method": "ignored", "original": text}

    # Choisir les ressources selon la section
    if section == "Actif":
        index    = ACTIF_INDEX
        keywords = ACTIF_KEYWORDS
        labels   = ACTIF_LABELS
    elif section == "Passif":
        index    = PASSIF_INDEX
        keywords = PASSIF_KEYWORDS
        labels   = PASSIF_LABELS
    elif section == "CPC":
        index    = CPC_INDEX
        keywords = CPC_KEYWORDS
        labels   = CPC_LABELS
    else:
        return {"ligne": None, "label": None, "score": 0,
                "method": "unknown_section", "original": text}

    normalized = full_normalize(text)

    # ── Niveau 1 : Match exact ────────────────────────────────
    if normalized in index:
        ligne = index[normalized]
        return {
            "ligne":    ligne,
            "label":    labels[ligne],
            "score":    1.0,
            "method":   "exact",
            "original": text,
        }

    # ── Niveau 2 : Match par mots-clés ───────────────────────
    best_kw_line  = None
    best_kw_score = 0.0
    for kw, ligne in keywords.items():
        kw_norm = full_normalize(kw)
        # Vérifier si tous les mots du keyword sont dans le texte normalisé
        kw_words = set(kw_norm.split())
        norm_words = set(normalized.split())
        if kw_words and kw_words.issubset(norm_words):
            score = len(kw_words) / max(len(norm_words), 1)
            if score > best_kw_score:
                best_kw_score = score
                best_kw_line  = ligne

    if best_kw_line and best_kw_score >= 0.3:
        return {
            "ligne":    best_kw_line,
            "label":    labels[best_kw_line],
            "score":    best_kw_score,
            "method":   "keyword",
            "original": text,
        }

    # ── Niveau 3 : Match flou ─────────────────────────────────
    best_line  = None
    best_score = 0.0
    for label_norm, ligne in index.items():
        s = _score(normalized, label_norm)
        if s > best_score:
            best_score = s
            best_line  = ligne

    if best_line and best_score >= 0.4:
        return {
            "ligne":    best_line,
            "label":    labels[best_line],
            "score":    round(best_score, 2),
            "method":   "fuzzy",
            "original": text,
        }

    # ── Aucun match ───────────────────────────────────────────
    return {
        "ligne":    None,
        "label":    None,
        "score":    0.0,
        "method":   "none",
        "original": text,
    }


def map_section(rows: list[list], section: str) -> list[dict]:
    """
    Mappe toutes les lignes extraites d'une section.

    rows    : liste de [libellé, val1, val2, ...]
    section : "Actif" | "Passif" | "CPC"

    Retourne une liste de dicts avec :
    {
        "ligne":    int | None,
        "label":    str,
        "score":    float,
        "method":   str,
        "original": str,
        "values":   [val1, val2, ...]   # valeurs numériques brutes
    }
    """
    results = []
    for row in rows:
        if not row:
            continue
        libelle = row[0] if row[0] else ""
        values  = row[1:] if len(row) > 1 else []

        match = match_line(libelle, section)
        match["values"] = values
        results.append(match)

    return results
