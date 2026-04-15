import fitz
import re
import pandas as pd
import streamlit as st

st.title("📄 → CSV ABBYY (Factures de vente)")

uploaded_files = st.file_uploader(
    "Importer vos factures PDF",
    type="pdf",
    accept_multiple_files=True
)

# -------- EXTRACTION TEXTE --------
def extract_text(file):
    text = ""
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    for page in pdf:
        text += page.get_text()
    return text


# -------- EXTRACTION --------
def extract_data(text):

    text_clean = text.replace("\n", " ")

    # CLIENT
    client_match = re.search(r"À l'attention de\s*(.+)", text_clean)
    client = client_match.group(1).strip() if client_match else "Client inconnu"

    # NUMERO
    num_match = re.search(r"Numéro de facture\s*([0-9]+)", text_clean)
    numero = "F" + num_match.group(1) if num_match else ""

    # DATE
    date_match = re.search(r"[0-9]{1,2}\s+[a-zéû\.]+\s+[0-9]{4}", text_clean, re.IGNORECASE)
    date = date_match.group(0) if date_match else ""

    # MONTANTS
    montants = re.findall(r"[0-9]+[.,][0-9]{2}", text_clean)

    if len(montants) >= 3:
        HT = float(montants[-3].replace(",", "."))
        TVA = float(montants[-2].replace(",", "."))
        TTC = float(montants[-1].replace(",", "."))
    else:
        HT = TVA = TTC = 0

    # TAUX TVA (fixe 20%)
    taux_tva = 20

    # TYPE DE VENTE (2 = vente standard)
    type_vente = 2

    return {
        "Client": client,
        "Référence (numéro)": numero,
        "Date de paiement": date,
        "Moyen de paiement": "Virement",
        "Montant HT": HT,
        "Taux de TVA": taux_tva,
        "Montant TTC": TTC,
        "Type de vente (1,2,3,4)": type_vente
    }


# -------- TRAITEMENT --------
if uploaded_files:
    if st.button("🚀 Générer CSV ABBYY"):

        results = []

        for file in uploaded_files:
            text = extract_text(file)
            results.append(extract_data(text))

        df = pd.DataFrame(results)

        # ordre EXACT des colonnes
        df = df[
            [
                "Client",
                "Référence (numéro)",
                "Date de paiement",
                "Moyen de paiement",
                "Montant HT",
                "Taux de TVA",
                "Montant TTC",
                "Type de vente (1,2,3,4)"
            ]
        ]

        st.dataframe(df)

        csv = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "📥 Télécharger CSV ABBYY",
            csv,
            "abbyy_import.csv",
            "text/csv"
        )
