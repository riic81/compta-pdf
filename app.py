import fitz  # PyMuPDF
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

# -------- EXTRACTION TEXTE --------
def extract_text_from_pdf(file):
    text = ""
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    for page in pdf:
        text += page.get_text()
    return text

# -------- EXTRACTION DONNÉES --------
# -------- MONTANTS (KRAMP FIX) --------

# HT
ht_match = re.search(r"Montant H\.T\.\s*([\d,]+)", text)
data["HT"] = float(ht_match.group(1).replace(",", ".")) if ht_match else 0

# TVA (on prend le dernier nombre de la ligne TVA)
tva_line = re.search(r"TVA.*", text)
if tva_line:
    numbers = re.findall(r"[0-9,]
[0-9]+[.,][0-9]{2}", tva_line.group(0))
    if len(numbers) >= 2:
        data["TVA"] = float(numbers[-1].replace(",", "."))
    else:
        data["TVA"] = 0
else:
    data["TVA"] = 0

# TTC
ttc_match = re.search(r"Montant T\.T\.C\.\s*(?:EUR)?\s*([\d,]+)", text)
data["TTC"] = float(ttc_match.group(1).replace(",", ".")) if ttc_match else 0
# -------- TRAITEMENT --------
if uploaded_files:
    if st.button("🚀 Générer"):
        results = []

        for file in uploaded_files:
            text = extract_text_from_pdf(file)
            results.append(extract_data(text))

        df = pd.DataFrame(results)

        # Nettoyage dates
        df["Date_obj"] = pd.to_datetime(df["Date_obj"], errors="coerce")
        df = df[df["Date_obj"].notna()]
        df = df.sort_values("Date_obj")

        # Format date propre
        df["Date"] = df["Date_obj"].dt.strftime("%Y-%m-%d")

        # Supprimer colonne technique
        df = df.drop(columns=["Date_obj"], errors="ignore")

        st.dataframe(df)

        # Export Excel
        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Journal", index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
