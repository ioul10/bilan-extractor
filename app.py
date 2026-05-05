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
from extractor.line_mapper import map_section
from extractor.validator import validate_all
from extractor.excel_writer_v2 import fill_model

st.set_page_config(page_title="Bilan Extractor", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .main-title  { font-size:2rem; font-weight:700; color:#1F3864; }
    .sub-title   { font-size:1rem; color:#888; margin-bottom:2rem; }
    .section-card{ background:#f8fafc; border-radius:10px; padding:1rem 1.2rem;
                   border-left:4px solid #2E75B6; margin-bottom:.5rem; }
    .info-box    { background:#EBF5FB; border-radius:8px; padding:.8rem 1rem;
                   font-size:.85rem; color:#1a5276; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Bilan Extractor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Extraction · Validation · Export vers MODELE.xlsx</div>', unsafe_allow_html=True)

for key, default in [
    ("pdf_path", None), ("page_count", 0), ("model_path", None),
    ("excel_bytes", None), ("rapport", None), ("preview_data", {}),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── ÉTAPE 1 — UPLOADS ────────────────────────────────────────
st.markdown("### 1️⃣ Upload des fichiers")
col_pdf, col_model = st.columns(2)

with col_pdf:
    uploaded_pdf = st.file_uploader("📄 Rapport financier PDF", type=["pdf"])
    if uploaded_pdf:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        tmp.write(uploaded_pdf.read())
        tmp.flush()
        st.session_state.pdf_path   = tmp.name
        st.session_state.page_count = get_pdf_page_count(tmp.name)
        st.session_state.excel_bytes = None
        st.session_state.rapport     = None
        st.success(f"✅ {uploaded_pdf.name} — **{st.session_state.page_count} pages**")

with col_model:
    uploaded_model = st.file_uploader("📋 MODELE.xlsx (template)", type=["xlsx"])
    if uploaded_model:
        tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp2.write(uploaded_model.read())
        tmp2.flush()
        st.session_state.model_path = tmp2.name
        st.success(f"✅ {uploaded_model.name} chargé")

# ── ÉTAPE 2 — IDENTIFICATION ──────────────────────────────────
if st.session_state.pdf_path:
    st.markdown("### 2️⃣ Identification")
    id1, id2, id3 = st.columns(3)
    with id1:
        raison_sociale     = st.text_input("Raison sociale",      placeholder="ex: ONCF")
    with id2:
        date_bilan         = st.text_input("Date de bilan",        placeholder="ex: 31/12/2024")
    with id3:
        identifiant_fiscal = st.text_input("Identifiant fiscal",   placeholder="ex: 3330419")

    # ── ÉTAPE 3 — PAGES ──────────────────────────────────────
    st.markdown("### 3️⃣ Configuration des pages")
    st.markdown(
        f'<div class="info-box">💡 Index de pages : <b>0</b> à <b>{st.session_state.page_count - 1}</b></div>',
        unsafe_allow_html=True
    )
    st.markdown("")

    same_page = st.checkbox("🔀 Actif et Passif sur la **même page** (côte à côte)", value=False)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="section-card"><b>🔵 ACTIF</b>', unsafe_allow_html=True)
        actif_input = st.text_input("Actif", placeholder="ex: 0",   key="actif_pages",  label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="section-card"><b>🟢 PASSIF</b>', unsafe_allow_html=True)
        passif_input = st.text_input("Passif", placeholder="ex: 1", key="passif_pages", label_visibility="collapsed", disabled=same_page)
        st.markdown('</div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="section-card"><b>🟣 CPC</b>', unsafe_allow_html=True)
        cpc_input = st.text_input("CPC", placeholder="ex: 2, 3",    key="cpc_pages",    label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("🔍 Aperçu d'une page", expanded=False):
        pa, pb = st.columns([1, 3])
        with pa:
            prev_idx = st.number_input("Index page", min_value=0,
                                       max_value=st.session_state.page_count - 1, value=0, step=1)
            if st.button("Voir aperçu"):
                st.session_state.preview_data[prev_idx] = preview_page(st.session_state.pdf_path, prev_idx)
        with pb:
            if prev_idx in st.session_state.preview_data:
                rows = st.session_state.preview_data[prev_idx]
                if rows:
                    max_c = max(len(r) for r in rows)
                    st.dataframe([r + [""]*(max_c-len(r)) for r in rows],
                                 use_container_width=True, hide_index=True)

    # ── ÉTAPE 4 — EXTRACTION + VALIDATION ────────────────────
    st.markdown("")
    st.markdown("### 4️⃣ Extraction & Validation")

    if st.button("⚡ Lancer", type="primary", use_container_width=True):
        errors = []
        sections_config = {}

        for label, raw in [("Actif", actif_input), ("Passif", passif_input), ("CPC", cpc_input)]:
            if same_page and label == "Passif":
                continue
            if not raw.strip():
                errors.append(f"Pages **{label}** non renseignées.")
                continue
            indices = parse_page_indices(raw)
            if not indices:
                errors.append(f"Index invalides pour **{label}** : `{raw}`")
                continue
            out = [i for i in indices if i >= st.session_state.page_count or i < 0]
            if out:
                errors.append(f"Index hors limites pour **{label}** : {out}")
                continue
            sections_config[label] = indices

        for err in errors:
            st.error(err)

        if not errors:
            with st.spinner("En cours..."):
                prog = st.progress(0)

                actif_crop = passif_crop = None
                if same_page and "Actif" in sections_config:
                    w   = get_page_width(st.session_state.pdf_path, sections_config["Actif"][0])
                    mid = w / 2
                    actif_crop  = (0,   0, mid, 9999)
                    passif_crop = (mid, 0, w,   9999)
                    sections_config["Passif"] = sections_config["Actif"]

                actif_rows  = extract_section(st.session_state.pdf_path, sections_config.get("Actif",  []), "Actif",  actif_crop)
                prog.progress(0.2)
                passif_rows = extract_section(st.session_state.pdf_path, sections_config.get("Passif", []), "Passif", passif_crop)
                prog.progress(0.4)
                cpc_rows    = extract_section(st.session_state.pdf_path, sections_config.get("CPC",    []), "CPC")
                prog.progress(0.6)

                actif_mapped  = map_section(actif_rows,  "Actif")
                passif_mapped = map_section(passif_rows, "Passif")
                cpc_mapped    = map_section(cpc_rows,    "CPC")
                prog.progress(0.8)

                rapport = validate_all(actif_mapped, passif_mapped, cpc_mapped)
                st.session_state.rapport = rapport
                prog.progress(1.0)

            # Résumé
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Score global", f"{rapport['score_global']}%")
            m2.metric("Actif",        f"{rapport['actif']['score']}%")
            m3.metric("Passif",       f"{rapport['passif']['score']}%")
            m4.metric("CPC",          f"{rapport['cpc']['score']}%")

            eq = rapport["equilibre"]
            if eq.get("ok"):
                st.success(f"✅ {eq['message']}")
            elif eq.get("ok") is False:
                st.error(f"❌ {eq['message']}")
            else:
                st.warning("⚠️ Équilibre non calculable")

            for check in rapport.get("cross_checks", []):
                if check.get("ok"):
                    st.success(f"✅ {check['label']}")
                else:
                    st.warning(f"⚠️ {check['label']} — {check.get('message','')}")

            with st.expander("📋 Détail validation", expanded=False):
                for sk in ["actif", "passif", "cpc"]:
                    rep = rapport[sk]
                    st.markdown(f"**{rep['section']} — {rep['score']}%** "
                                f"✅{rep['nb_ok']} ❌{rep['nb_error']} ⚠️{rep['nb_miss']}")
                    st.dataframe([{
                        "": {"ok":"✅","error":"❌","missing":"⚠️"}.get(l["status"],""),
                        "L": l["ligne"],
                        "Libellé": l["label"],
                        "Extrait": l["extrait"] or "",
                        "Calculé": l["calculé"] or "",
                        "Écart":   l["écart"]   or "",
                    } for l in rep["lignes"]], use_container_width=True, hide_index=True)

            # Générer Excel
            if st.session_state.model_path:
                excel_bytes = fill_model(
                    model_path=st.session_state.model_path,
                    actif_data=rapport["data"]["actif"],
                    passif_data=rapport["data"]["passif"],
                    cpc_data=rapport["data"]["cpc"],
                    rapport=rapport,
                    identification={
                        "raison_sociale":     raison_sociale,
                        "date_bilan":         date_bilan,
                        "identifiant_fiscal": identifiant_fiscal,
                        "format_pdf":         uploaded_pdf.name if uploaded_pdf else "",
                    },
                )
                st.session_state.excel_bytes = excel_bytes
            else:
                st.info("ℹ️ Uploadez le MODELE.xlsx pour générer le fichier final.")

    # ── ÉTAPE 5 — TÉLÉCHARGEMENT ──────────────────────────────
    if st.session_state.excel_bytes:
        st.markdown("---")
        st.markdown("### 5️⃣ Téléchargement")
        nom = raison_sociale.replace(" ", "_") if raison_sociale else "bilan"
        dt  = date_bilan.replace("/","") if date_bilan else ""
        st.download_button(
            label="⬇️ Télécharger le MODELE rempli",
            data=st.session_state.excel_bytes,
            file_name=f"{nom}_{dt}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.markdown(
            '<div class="info-box">📋 Onglets : '
            '<b>Identification</b> · <b>Bilan Actif</b> · <b>Bilan Passif</b> · '
            '<b>CPC</b> · <b>Rapport Validation</b></div>',
            unsafe_allow_html=True
        )

st.markdown("---")
st.markdown("<div style='text-align:center;color:#aaa;font-size:.8rem;'>Bilan Extractor v2.0</div>",
            unsafe_allow_html=True)
