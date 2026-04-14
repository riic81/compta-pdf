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

    # -------- DATE --------
    date_match = re.search(r"\d{1,2}\s+[a-zéû]+\.?\s+\d{4}", text, re.IGNORECASE)
    raw_date = date_match.group(0) if date_match else ""
    data["Date"] = raw_date
    data["Date_obj"] = parse_french_date(raw_date)

    # -------- NUMERO --------
    num_match = re.search(r"Numéro de facture\s*(\d+)", text)
    if not num_match:
        num_match = re.search(r"facture\s*(\d+)", text, re.IGNORECASE)
    data["Numéro"] = num_match.group(1) if num_match else ""

    # -------- MONTANTS --------
    # chercher ligne TVA
    tva_line = re.search(r"TVA.*", text)

    if tva_line:
        numbers = re.findall(r"\d+[.,]\d{2}", tva_line.group(0))
        if len(numbers) >= 2:
            data["HT"] = float(numbers[0].replace(",", "."))
            data["TVA"] = float(numbers[1].replace(",", "."))
        else:
            data["HT"] = 0
            data["TVA"] = 0
    else:
        data["HT"] = 0
        data["TVA"] = 0

    # TTC
    total_match = re.search(r"Montant total\s*(\d+[.,]\d{2})", text)
    data["TTC"] = float(total_match.group(1).replace(",", ".")) if total_match else 0

    # -------- NOM (zone haut du document) --------
    lines = text.split("\n")

    nom = ""
    for i in range(len(lines)):
        if "Facture" in lines[i]:
            # regarder les lignes suivantes
            for j in range(i+1, i+5):
                if j < len(lines):
                    line = lines[j].strip()
                    if len(line) > 3 and "France" not in line:
                        nom = line
                        break
            break

    data["Nom"] = nom

    # -------- FOURNISSEUR --------
    if "laboutiquehydro" in text.lower():
        fournisseur = "LA BOUTIQUE HYDRO"
    else:
        fournisseur = "Inconnu"

    data["Fournisseur"] = fournisseur

    # -------- TYPE --------
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
