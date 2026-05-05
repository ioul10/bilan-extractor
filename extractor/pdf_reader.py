import pdfplumber


# ── Règles fixes de colonnes par section ─────────────────────────
# Structure standard du bilan marocain :
#   Actif  : Libellé | Brut | Amort&Prov | Net exercice | Net ex.précédent
#   Passif : Libellé | Exercice | Exercice précédent
#   CPC    : Libellé | Op.propres | Op.précédents | Total exercice | Ex.précédent

SECTION_HEADERS = {
    "Actif":  ["LIBELLE", "BRUT", "AMORT & PROV", "NET EXERCICE", "NET EX. PRECEDENT"],
    "Passif": ["LIBELLE", "EXERCICE", "EXERCICE PRECEDENT"],
    "CPC":    ["LIBELLE", "OP. PROPRES", "OP. PRECEDENTS", "TOTAL EXERCICE", "EX. PRECEDENT"],
}

# Nombre de colonnes de valeurs attendues (hors libellé)
SECTION_NB_VALUE_COLS = {
    "Actif":  4,
    "Passif": 2,
    "CPC":    4,
}

# ── Ratio de séparation des colonnes (proportion de la largeur page) ──
# Ces ratios définissent où se trouvent les séparateurs de colonnes
# Ils sont relatifs à la largeur de la page (0.0 → 1.0)
# Calculés sur la base des PDFs analysés — valables pour format A4 portrait
SECTION_COL_RATIOS = {
    "Actif":  [0.0, 0.39, 0.52, 0.64, 0.80],  # 5 colonnes
    "Passif": [0.0, 0.61, 0.81],               # 3 colonnes
    "CPC":    [0.0, 0.47, 0.61, 0.72, 0.84],   # 5 colonnes
}


def parse_page_indices(raw: str) -> list[int]:
    """Convertit '0', '2,3', '2, 3' en liste d'entiers (index 0-based)."""
    indices = []
    for part in raw.replace(" ", "").split(","):
        if part.isdigit():
            indices.append(int(part))
    return indices


def get_pdf_page_count(pdf_path: str) -> int:
    with pdfplumber.open(pdf_path) as pdf:
        return len(pdf.pages)


def extract_section(pdf_path: str, page_indices: list[int],
                    section_name: str = None,
                    crop: tuple = None) -> list[list]:
    """
    Extrait les données d'une section sur les pages indiquées.
    Utilise extract_words() + reconstruction par positions X.
    Force le nombre de colonnes selon section_name si fourni.

    crop : (x0, y0, x1, y1) en points — zone à extraire sur la page.
           Si None → page entière.
           Utile pour les pages où Actif et Passif sont côte à côte.
    """
    all_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)

        for i, idx in enumerate(page_indices):
            if idx < 0 or idx >= total_pages:
                continue

            page = pdf.pages[idx]

            # Appliquer le crop si demandé
            if crop:
                x0, y0, x1, y1 = crop
                # Limiter y1 à la hauteur réelle de la page
                y1 = min(y1, page.height)
                page = page.crop((x0, y0, x1, y1))

            rows = _reconstruct_from_words(page, section_name)
            all_rows.extend(rows)

            # Séparateur visuel entre pages
            if i < len(page_indices) - 1:
                all_rows.append(["--- PAGE SUIVANTE ---"])

    return all_rows


def get_page_width(pdf_path: str, page_idx: int) -> float:
    """Retourne la largeur d'une page — utile pour calculer le crop."""
    with pdfplumber.open(pdf_path) as pdf:
        return pdf.pages[page_idx].width


def _detect_col_boundaries(words: list, page_width: float,
                            section_name: str = None) -> list:
    """
    Détecte les frontières de colonnes à partir des positions X des mots.

    Stratégie :
    - Si section_name connu → utiliser les clusters de valeurs numériques
      pour trouver précisément les N colonnes de valeurs attendues
    - Sinon → gap analysis générique sur tous les mots
    """
    if not words:
        return [(0, page_width)]

    # ── Stratégie basée sur les valeurs numériques ────────────────
    if section_name and section_name in SECTION_NB_VALUE_COLS:
        expected_val_cols = SECTION_NB_VALUE_COLS[section_name]
        boundaries = _detect_by_numeric_clusters(
            words, page_width, expected_val_cols
        )
        if boundaries and len(boundaries) == expected_val_cols + 1:
            return boundaries
        # Fallback si détection numérique échoue

    # ── Stratégie générique : gap analysis ───────────────────────
    xs = sorted(set(round(w["x0"]) for w in words))
    gap_threshold = page_width * 0.04

    gaps = []
    for i in range(1, len(xs)):
        gap = xs[i] - xs[i - 1]
        if gap > gap_threshold:
            gaps.append((xs[i - 1] + xs[i]) / 2)

    boundaries = []
    prev = 0
    for mid in gaps:
        boundaries.append((prev, mid))
        prev = mid
    boundaries.append((prev, page_width))

    if section_name and section_name in SECTION_NB_VALUE_COLS:
        expected_total = SECTION_NB_VALUE_COLS[section_name] + 1
        boundaries = _force_column_count(boundaries, words, page_width, expected_total)

    return boundaries


def _detect_by_numeric_clusters(words: list, page_width: float,
                                  nb_val_cols: int) -> list:
    """
    Détecte les colonnes en cherchant les clusters de valeurs numériques.
    
    Principe : les colonnes de valeurs sont identifiées par la concentration
    de nombres à des positions X similaires. On cherche nb_val_cols clusters.
    La colonne libellé occupe le reste (côté gauche).
    """
    import re
    num_pattern = re.compile(r'^-?[\d]+[\d\s.]*[,.][\d]+$')

    # Collecter les x0 des valeurs numériques
    num_xs = []
    for w in words:
        clean = w["text"].replace(" ", "")
        if num_pattern.match(clean):
            num_xs.append(w["x0"])

    if len(num_xs) < nb_val_cols:
        return []  # Pas assez de valeurs pour détecter

    # Clustering simple : trier les x, chercher nb_val_cols groupes
    num_xs_sorted = sorted(num_xs)

    # Trouver les gaps entre valeurs numériques consécutives
    gaps = []
    for i in range(1, len(num_xs_sorted)):
        gap = num_xs_sorted[i] - num_xs_sorted[i - 1]
        gaps.append((gap, i))

    # Trier par gap décroissant → garder les (nb_val_cols - 1) plus grands
    gaps.sort(reverse=True)
    split_indices = sorted([idx for _, idx in gaps[:nb_val_cols - 1]])

    # Trouver le centre de chaque cluster
    clusters = []
    prev_i = 0
    for split_i in split_indices + [len(num_xs_sorted)]:
        cluster = num_xs_sorted[prev_i:split_i]
        if cluster:
            clusters.append(cluster)
        prev_i = split_i

    if len(clusters) != nb_val_cols:
        return []

    # Calculer les frontières :
    # - Colonne libellé : 0 → juste avant le premier cluster
    # - Colonnes valeurs : entre clusters
    boundaries = []

    # Frontière libellé / première colonne valeur
    first_val_start = min(clusters[0]) - 5
    # S'assurer que la colonne libellé est assez large (au moins 30% de la page)
    label_end = max(first_val_start, page_width * 0.25)
    boundaries.append((0, label_end))

    # Frontières entre colonnes de valeurs
    prev_end = label_end
    for i, cluster in enumerate(clusters):
        if i < len(clusters) - 1:
            # Milieu entre fin du cluster i et début du cluster i+1
            gap_start = max(cluster)
            gap_end   = min(clusters[i + 1])
            col_end   = (gap_start + gap_end) / 2
        else:
            col_end = page_width
        boundaries.append((prev_end, col_end))
        prev_end = col_end

    return boundaries


def _force_column_count(boundaries: list, words: list,
                         page_width: float, expected: int) -> list:
    """
    Si le nombre de colonnes détectées ne correspond pas à expected,
    on recalcule en cherchant les gaps les plus larges.
    """
    if len(boundaries) == expected:
        return boundaries

    # Chercher tous les gaps possibles
    xs = sorted(set(round(w["x0"]) for w in words))
    all_gaps = []
    for i in range(1, len(xs)):
        gap = xs[i] - xs[i - 1]
        all_gaps.append((gap, (xs[i - 1] + xs[i]) / 2))

    # Trier par taille de gap décroissant → garder les (expected-1) plus grands
    all_gaps.sort(reverse=True)
    top_gaps = sorted([mid for _, mid in all_gaps[:expected - 1]])

    # Reconstruire les boundaries
    boundaries = []
    prev = 0
    for mid in top_gaps:
        boundaries.append((prev, mid))
        prev = mid
    boundaries.append((prev, page_width))

    return boundaries


def _find_column(x: float, boundaries: list) -> int:
    """Retourne l'index de colonne pour une position X."""
    for i, (x_start, x_end) in enumerate(boundaries):
        if x_start <= x < x_end:
            return i
    return len(boundaries) - 1


def _reconstruct_from_words(page, section_name: str = None) -> list[list]:
    """
    Reconstruit le tableau ligne par ligne depuis extract_words().
    - Si section_name connu → utilise les ratios de colonnes fixes
    - Sinon → détection automatique par gap analysis
    """
    words = page.extract_words(
        x_tolerance=4,
        y_tolerance=4,
        keep_blank_chars=False,
        extra_attrs=["fontname", "size"]
    )

    if not words:
        return []

    # ── Définir les frontières de colonnes ───────────────────────
    if section_name and section_name in SECTION_COL_RATIOS:
        # Utiliser les ratios fixes — indépendants du contenu
        ratios = SECTION_COL_RATIOS[section_name]
        w = page.width
        # Convertir ratios en boundaries (x_start, x_end)
        xs = [r * w for r in ratios]
        col_boundaries = []
        for i in range(len(xs) - 1):
            col_boundaries.append((xs[i], xs[i+1]))
        col_boundaries.append((xs[-1], w))
    else:
        col_boundaries = _detect_col_boundaries(words, page.width, section_name)

    nb_cols = len(col_boundaries)

    # ── Grouper par ligne (Y arrondi) ────────────────────────────
    lines: dict = {}
    for word in words:
        y_key = round(word["top"] / 4) * 4
        if y_key not in lines:
            lines[y_key] = []
        lines[y_key].append(word)

    sorted_lines = sorted(lines.items())

    # ── Construire les lignes ────────────────────────────────────
    rows = []
    for _, line_words in sorted_lines:
        row = [""] * nb_cols
        for word in sorted(line_words, key=lambda w: w["x0"]):
            col_idx = _find_column(word["x0"], col_boundaries)
            text = word["text"].strip()
            if not text:
                continue
            if row[col_idx]:
                row[col_idx] += " " + text
            else:
                row[col_idx] = text

        # Ignorer lignes vraiment vides
        if any(cell.strip() for cell in row):
            rows.append(row)

    return rows


def preview_page(pdf_path: str, page_idx: int) -> list[list]:
    """Aperçu des 20 premières lignes d'une page."""
    rows = extract_section(pdf_path, [page_idx])
    return rows[:20]
