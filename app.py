import streamlit as st
import tempfile
import os
from extractor.pdf_reader import (
    extract_section,
    parse_page_indices,
    get_pdf_page_count,
    get_page_width,
    preview_page,
)
from extractor.excel_writer import build_excel

# ── Config page ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Bilan Extractor",
    page_icon="📊",
    layout="wide",
)

# ── CSS custom ────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1F3864;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1rem;
        color: #888;
        margin-bottom: 2rem;
    }
    .section-card {
        background: #f8fafc;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        border-left: 4px solid #2E75B6;
        margin-bottom: 0.5rem;
    }
    .step-badge {
        background: #1F3864;
        color: white;
        border-radius: 50%;
        width: 28px;
        height: 28px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        margin-right: 8px;
    }
    .info-box {
        background: #EBF5FB;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        font-size: 0.85rem;
        color: #1a5276;
    }
</style>
""", unsafe_allow_html=True)

# ── Titre ─────────────────────────────────────────────────────────
st.markdown('<div class="main-title">📊 Bilan Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Extraction de bilans financiers PDF → Excel</div>', unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────
if "pdf_path" not in st.session_state:
    st.session_state.pdf_path = None
if "page_count" not in st.session_state:
    st.session_state.page_count = 0
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "preview_data" not in st.session_state:
    st.session_state.preview_data = {}

# ═══════════════════════════════════════════════════════════════════
# ÉTAPE 1 — UPLOAD
# ═══════════════════════════════════════════════════════════════════
st.markdown("### 1️⃣ Upload du PDF")

uploaded_file = st.file_uploader(
    "Glissez votre rapport financier PDF ici",
    type=["pdf"],
    help="Rapport annuel, liasse fiscale, ou tout document contenant un bilan"
)

if uploaded_file:
    # Sauvegarder temporairement
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.write(uploaded_file.read())
    tmp.flush()
    st.session_state.pdf_path = tmp.name
    st.session_state.page_count = get_pdf_page_count(tmp.name)
    st.session_state.excel_bytes = None  # reset

    st.success(f"✅ **{uploaded_file.name}** chargé — **{st.session_state.page_count} pages** détectées")
    st.markdown(
        f'<div class="info-box">💡 Les index de pages vont de <b>0</b> à '
        f'<b>{st.session_state.page_count - 1}</b>. '
        f'Entrez les index tels que vous les voyez (ex : 0, 1, 2 ...)</div>',
        unsafe_allow_html=True
    )
    st.markdown("")

# ═══════════════════════════════════════════════════════════════════
# ÉTAPE 2 — CONFIGURATION DES PAGES
# ═══════════════════════════════════════════════════════════════════
if st.session_state.pdf_path:
    st.markdown("### 2️⃣ Configuration des pages")
    st.markdown("Indiquez les **index** des pages pour chaque section (plusieurs pages : séparées par une virgule)")

    # Option : Actif et Passif sur la même page (côte à côte)
    same_page = st.checkbox(
        "🔀 Actif et Passif sont sur la **même page** (côte à côte)",
        value=False,
        help="Cocher si le bilan Actif et Passif apparaissent côte à côte sur une seule page (ex: LabelVie)"
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown("**🔵 ACTIF**")
        actif_input = st.text_input(
            "Pages Actif (index)",
            placeholder="ex: 0",
            key="actif_pages",
            label_visibility="collapsed"
        )
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown("**🟢 PASSIF**")
        passif_input = st.text_input(
            "Pages Passif (index)",
            placeholder="ex: 1" if not same_page else "même page que Actif",
            key="passif_pages",
            label_visibility="collapsed",
            disabled=same_page
        )
        st.markdown('</div>', unsafe_allow_html=True)

    with col3:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown("**🟣 CPC**")
        cpc_input = st.text_input(
            "Pages CPC (index)",
            placeholder="ex: 2, 3",
            key="cpc_pages",
            label_visibility="collapsed"
        )
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Preview ──────────────────────────────────────────────────
    st.markdown("")
    with st.expander("🔍 Aperçu d'une page avant extraction", expanded=False):
        preview_col1, preview_col2 = st.columns([1, 3])
        with preview_col1:
            preview_idx = st.number_input(
                "Index de la page à prévisualiser",
                min_value=0,
                max_value=st.session_state.page_count - 1,
                value=0,
                step=1
            )
            if st.button("Voir aperçu"):
                with st.spinner("Chargement..."):
                    preview_rows = preview_page(st.session_state.pdf_path, preview_idx)
                    st.session_state.preview_data[preview_idx] = preview_rows

        with preview_col2:
            if preview_idx in st.session_state.preview_data:
                rows = st.session_state.preview_data[preview_idx]
                if rows:
                    st.markdown(f"**Page index {preview_idx} — {len(rows)} lignes détectées (20 premières)**")
                    # Afficher comme tableau
                    max_cols = max(len(r) for r in rows)
                    display_rows = []
                    for r in rows:
                        display_rows.append(r + [""] * (max_cols - len(r)))
                    st.dataframe(
                        display_rows,
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.warning("Aucun contenu détecté sur cette page.")

    # ═══════════════════════════════════════════════════════════════
    # ÉTAPE 3 — EXTRACTION
    # ═══════════════════════════════════════════════════════════════
    st.markdown("")
    st.markdown("### 3️⃣ Extraction")

    extract_btn = st.button("⚡ Lancer l'extraction", type="primary", use_container_width=True)

    if extract_btn:
        # Validation
        errors = []
        sections_config = {}

        for label, raw in [("Actif", actif_input), ("Passif", passif_input), ("CPC", cpc_input)]:
            if not raw.strip():
                errors.append(f"Pages **{label}** non renseignées.")
                continue
            indices = parse_page_indices(raw)
            if not indices:
                errors.append(f"Index invalides pour **{label}** : `{raw}`")
                continue
            # Vérifier que les index sont dans les limites
            out_of_range = [i for i in indices if i >= st.session_state.page_count or i < 0]
            if out_of_range:
                errors.append(f"Index hors limites pour **{label}** : {out_of_range} (max: {st.session_state.page_count - 1})")
                continue
            sections_config[label] = indices

        if errors:
            for err in errors:
                st.error(err)
        else:
            with st.spinner("Extraction en cours..."):
                extracted = {}
                progress = st.progress(0)
                section_list = list(sections_config.items())

                for i, (name, indices) in enumerate(section_list):
                    rows = extract_section(st.session_state.pdf_path, indices)
                    extracted[name] = rows
                    progress.progress((i + 1) / len(section_list))

                # Construire Excel
                excel_bytes = build_excel(extracted)
                st.session_state.excel_bytes = excel_bytes
                progress.progress(1.0)

            # Résumé
            st.success("✅ Extraction terminée !")
            cols = st.columns(3)
            names = ["Actif", "Passif", "CPC"]
            colors = ["🔵", "🟢", "🟣"]
            for i, (name, color) in enumerate(zip(names, colors)):
                with cols[i]:
                    count = len(extracted.get(name, []))
                    st.metric(f"{color} {name}", f"{count} lignes")

    # ═══════════════════════════════════════════════════════════════
    # ÉTAPE 4 — DOWNLOAD
    # ═══════════════════════════════════════════════════════════════
    if st.session_state.excel_bytes:
        st.markdown("")
        st.markdown("### 4️⃣ Téléchargement")

        filename = "bilan_extrait.xlsx"
        if uploaded_file:
            base = os.path.splitext(uploaded_file.name)[0]
            filename = f"{base}_bilan.xlsx"

        st.download_button(
            label="⬇️ Télécharger le fichier Excel",
            data=st.session_state.excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

        st.markdown(
            '<div class="info-box">📋 Le fichier Excel contient 4 onglets : '
            '<b>Identification</b> (vide), <b>Actif</b>, <b>Passif</b>, <b>CPC</b></div>',
            unsafe_allow_html=True
        )

# ── Footer ────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:#aaa; font-size:0.8rem;'>"
    "Bilan Extractor v0.1 — Étape 1 : Extraction brute"
    "</div>",
    unsafe_allow_html=True
)
