import fitz
import re
import pandas as pd
import streamlit as st

st.title("📄 → CSV ABBYY (Factures de vente)")

uploaded_files = st.file_uploader(
    "Importer vos factures de vente PDF",
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


# -------- EXTRACTION DONNÉES --------
def extract_data(text):

    text_clean = text.replace("\n", " ")

    # DATE
    date_match = re.search(r"[0-9]{1,2}\s+[a-zéû\.]+\s+[0-9]{4}", text_clean, re.IGNORECASE)
    date = date_match.group(0) if date_match else ""

    # NUMERO
    num_match = re.search(r"Numéro de facture\s*([0-9]+)", text_clean)
    numero = num_match.group(1) if num_match else ""

    # MONTANTS (FIABLE car structure fixe)
    montants = re.findall(r"[0-9]+[.,][0-9]{2}", text_clean)

    if len(montants) >= 3:
        HT = float(montants[-3].replace(",", "."))
        TVA = float(montants[-2].replace(",", "."))
        TTC = float(montants[-1].replace(",", "."))
    else:
        HT = TVA = TTC = 0

    return {
        "Date": date,
        "InvoiceNumber": numero,
        "SupplierName": "LA BOUTIQUE HYDRO",
        "TotalHT": HT,
        "TaxAmount": TVA,
        "TotalAmount": TTC
    }


# -------- TRAITEMENT --------
if uploaded_files:
    if st.button("🚀 Générer CSV"):

        results = []

        for file in uploaded_files:
            text = extract_text(file)
            data = extract_data(text)
            results.append(data)

        df = pd.DataFrame(results)

        st.dataframe(df)

        # EXPORT CSV ABBYY
        csv = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "📥 Télécharger CSV ABBYY",
            csv,
            "factures_abbyy.csv",
            "text/csv"
        )
