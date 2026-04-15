import fitz
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

# -------- TEXTE PDF --------
def extract_text_from_pdf(file):
    text = ""
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    for page in pdf:
        text += page.get_text()
    return text

# -------- EXTRACTION --------
def extract_data(text):
    data = {}

    text_clean = text.replace("\n", " ")

    # -------- DATE --------
    date_match = re.search(r"[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4}", text_clean)
    if date_match:
        raw_date = date_match.group(0)
        data["Date_obj"] = pd.to_datetime(raw_date, dayfirst=True)
    else:
        data["Date_obj"] = None

    # -------- NUMERO --------
    num_match = re.search(r"No\. Facture\s*([0-9]+)", text_clean)
    data["Numéro"] = num_match.group(1) if num_match else ""

    # -------- HT --------
    ht_match = re.search(r"Montant H\.T\.\s*([0-9,]+)", text_clean)
    data["HT"] = float(ht_match.group(1).replace(",", ".")) if ht_match else 0

    # -------- TVA (FIX FINAL) --------
    tva_match = re.search(r"TVA\s*20%\s*de\s*[0-9,]+\s*([0-9,]+)", text_clean)
    if tva_match:
        data["TVA"] = float(tva_match.group(1).replace(",", "."))
    else:
        data["TVA"] = 0

    # -------- TTC (FIX FINAL) --------
    ttc_match = re.search(r"Montant T\.T\.C\.\s*(?:EUR)?\s*([0-9,]+)", text_clean)
    if ttc_match:
        data["TTC"] = float(ttc_match.group(1).replace(",", "."))
    else:
        data["TTC"] = 0

    # -------- NOM --------
    lines = text.split("\n")
    nom = ""
    for i in range(len(lines)):
        if "Client" in lines[i]:
            if i + 1 < len(lines):
                nom = lines[i + 1].strip()
                break

    data["Nom"] = nom

    # -------- FOURNISSEUR --------
    if "kramp" in text.lower():
        fournisseur = "KRAMP"
    elif "laboutiquehydro" in text.lower():
        fournisseur = "LA BOUTIQUE HYDRO"
    else:
        fournisseur = "Inconnu"

    data["Fournisseur"] = fournisseur
    data["Type"] = "Achat" if fournisseur != "LA BOUTIQUE HYDRO" else "Vente"

    return data
# -------- TRAITEMENT --------
if uploaded_files:
    if st.button("🚀 Générer"):
        results = []

        for file in uploaded_files:
            text = extract_text_from_pdf(file)
            results.append(extract_data(text))

        df = pd.DataFrame(results)

        df["Date_obj"] = pd.to_datetime(df["Date_obj"], errors="coerce")
        df = df[df["Date_obj"].notna()]
        df = df.sort_values("Date_obj")

        df["Date"] = df["Date_obj"].dt.strftime("%Y-%m-%d")
        df = df.drop(columns=["Date_obj"], errors="ignore")

        st.dataframe(df)

        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Journal", index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
