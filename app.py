import pdfplumber
import re
import pandas as pd
import streamlit as st
from datetime import datetime

st.title("📄 → 📊 Compta automatique")

uploaded_files = st.file_uploader("Importer vos factures PDF", type="pdf", accept_multiple_files=True)

def parse_french_date(date_str):
    mois_map = {
        "janv.": "01", "févr.": "02", "mars": "03", "avr.": "04",
        "mai": "05", "juin": "06", "juil.": "07", "août": "08",
        "sept.": "09", "oct.": "10", "nov.": "11", "déc.": "12"
    }
    try:
        parts = date_str.split()
        return datetime.strptime(f"{parts[0]}/{mois_map.get(parts[1],'01')}/{parts[2]}", "%d/%m/%Y")
    except:
        return None

def extract_data(text):
    data = {}

    date_match = re.search(r"Date\s*\n\s*(.+)", text)
    raw_date = date_match.group(1).strip() if date_match else ""
    data["Date"] = raw_date
    data["Date_obj"] = parse_french_date(raw_date)

    name_match = re.search(r"À l'attention de\.\s*(.+)", text)
    data["Nom"] = name_match.group(1).strip() if name_match else ""

    num_match = re.search(r"Numéro de facture\s*\n\s*(\d+)", text)
    data["Numéro"] = num_match.group(1) if num_match else ""

    ht_match = re.search(r"TVA de\s*([\d,]+)", text)
    data["HT"] = float(ht_match.group(1).replace(",", ".")) if ht_match else 0

    tva_match = re.search(r"TVA de\s*[\d,]+\s*€\s*([\d,]+)", text)
    data["TVA"] = float(tva_match.group(1).replace(",", ".")) if tva_match else 0

    total_match = re.search(r"Montant total\s*([\d,]+)", text)
    data["TTC"] = float(total_match.group(1).replace(",", ".")) if total_match else 0

    fournisseur_match = re.search(r"\n([A-Z ]{5,})\nIBAN", text)
    fournisseur = fournisseur_match.group(1).strip() if fournisseur_match else ""
    data["Fournisseur"] = fournisseur

    data["Type"] = "Vente" if "LA BOUTIQUE HYDRO" in fournisseur else "Achat"

    return data

if uploaded_files:
    if st.button("🚀 Générer"):
        results = []

        for file in uploaded_files:
            with pdfplumber.open(file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
                results.append(extract_data(text))
df["Date_obj"] = pd.to_datetime(df["Date_obj"], errors="coerce")

        df = pd.DataFrame(results)
        df = df[df["Date_obj"].notna()]
        df = df.sort_values("Date_obj")
        df["Date_obj"] = pd.to_datetime(df["Date_obj"], errors="coerce")
df["Mois"] = df["Date_obj"].dt.strftime("%Y-%m")

        st.dataframe(df)

        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Journal", index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
