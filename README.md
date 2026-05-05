# 📊 Bilan Extractor

Extraction de bilans financiers depuis des rapports PDF vers Excel.

## Fonctionnalités

- Upload d'un rapport financier PDF
- Configuration manuelle des pages (Actif / Passif / CPC)
- Extraction brute avec colonnes respectées
- Export Excel avec 4 onglets : Identification, Actif, Passif, CPC

## Installation locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Déploiement Streamlit Cloud

1. Fork ce repo sur GitHub
2. Aller sur [share.streamlit.io](https://share.streamlit.io)
3. New app → sélectionner le repo → `app.py`
4. Deploy

## Structure

```
bilan-extractor/
├── app.py                  # Interface Streamlit
├── extractor/
│   ├── pdf_reader.py       # Extraction pdfplumber
│   └── excel_writer.py     # Écriture openpyxl
├── requirements.txt
└── README.md
```

## Version

- **v0.1** : Extraction brute (Étape 1)
- **v0.2** (à venir) : Dictionnaire des clés + normalisation
- **v0.3** (à venir) : Modèle Excel unifié
