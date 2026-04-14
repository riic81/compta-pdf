import pdfplumber
import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.title("📄 → 📊 Compta automatique")

uploaded_files = st.file_uploader(
    "Importer vos factures PDF",
    type="pdf",
    accept_multiple_files=True
)

# -------- DATE --------
def parse_french_date(date_str):
    mois_map = {
        "janv.": "01", "févr.": "02", "mars": "03", "avr.": "04",
        "mai": "05", "juin": "06", "juil.": "07", "août": "08",
        "sept.": "09", "oct.": "10", "nov.": "11", "déc.": "12"
    }
    try:
        parts = date_str.split()
        return datetime.strptime(
            f"{parts[0]}/{mois_map.get(parts[1],'01')}/{parts[2]}",
            "%d/%m/%Y"
        )
    except:
        return None

# -------- EXTRACTION --------
def extract_data(text):
    data = {}

    # 🔹 DATE (ex: 12 avr. 2026)
    date_match = re.search(r"(\d{1,2} [a-zéû]+\.? \d{4})", text, re.IGNORECASE)
    raw_date = date_match.group(1) if date_match else ""
    data["Date"] = raw_date
    data["Date_obj"] = parse_french_date(raw_date)

    # 🔹 NOM client
    name_match = re.search(r"À l'attention de\.?\s*([A-Z\- ]+)", text)
    data["Nom"] = name_match.group(1).strip() if name_match else ""

    # 🔹 NUMÉRO facture
    num_match = re.search(r"Numéro de facture\s*\n\s*(\d+)", text)
    data["Numéro"] = num_match.group(1) if num_match else ""

    # 🔹 HT (ligne TVA)
    ht_match = re.search(r"TVA de\s*([\d,]+)", text)
    data["HT"] = float(ht_match.group(1).replace(",", ".")) if ht_match else 0

    # 🔹 TVA
    tva_match = re.search(r"TVA de\s*[\d,]+\s*€\s*([\d,]+)", text)
    data["TVA"] = float(tva_match.group(1).replace(",", ".")) if tva_match else 0

    # 🔹 TTC
    total_match = re.search(r"Montant total\s*([\d,]+)", text)
    data["TTC"] = float(total_match.group(1).replace(",", ".")) if total_match else 0

    # 🔹 Fournisseur (plus fiable)
    if "LA BOUTIQUE HYDRO" in text:
        fournisseur = "LA BOUTIQUE HYDRO"
    else:
        fournisseur = "Autre"

    data["Fournisseur"] = fournisseur

    # 🔹 Type
    data["Type"] = "Vente" if fournisseur == "LA BOUTIQUE HYDRO" else "Achat"

    return data

# -------- TRAITEMENT --------
if uploaded_files:
    if st.button("🚀 Générer"):
        results = []

        for file in uploaded_files:
            with pdfplumber.open(file) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"

                results.append(extract_data(text))

        df = pd.DataFrame(results)

        df["Date_obj"] = pd.to_datetime(df["Date_obj"], errors="coerce")
        df = df[df["Date_obj"].notna()]
        df = df.sort_values("Date_obj")
        df["Mois"] = df["Date_obj"].dt.strftime("%Y-%m")

        st.dataframe(df)

        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Journal", index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
