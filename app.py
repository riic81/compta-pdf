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
def extract_data(text):
    data = {}

    text_clean = text.replace("\n", " ")

    # -------- DATE (2 formats) --------
    date_match = re.search(r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}", text_clean)
    if date_match:
        raw_date = date_match.group(0)
        try:
            data["Date_obj"] = pd.to_datetime(raw_date, dayfirst=True)
        except:
            data["Date_obj"] = None
        data["Date"] = raw_date
    else:
        date_match = re.search(r"\d{1,2}\s+[a-zéû]+\.?\s+\d{4}", text_clean, re.IGNORECASE)
        raw_date = date_match.group(0) if date_match else ""
        data["Date"] = raw_date
        data["Date_obj"] = parse_french_date(raw_date)

    # -------- NUMERO --------
    num_match = re.search(r"No\. Facture\s*(\d+)", text_clean)
    if not num_match:
        num_match = re.search(r"Numéro de facture\s*(\d+)", text_clean)
    if not num_match:
        num_match = re.search(r"facture\s*(\d+)", text_clean, re.IGNORECASE)

    data["Numéro"] = num_match.group(1) if num_match else ""

    # -------- MONTANTS (KRAMP ou standard) --------

    # Cas KRAMP
    ht_match = re.search(r"Montant H\.T\.\s*([\d,]+)", text_clean)
    tva_match = re.search(r"TVA.*?([\d,]+)", text_clean)
    ttc_match = re.search(r"Montant T\.T\.C\.\s*(?:EUR)?\s*([\d,]+)", text_clean)

    if ht_match and ttc_match:
        data["HT"] = float(ht_match.group(1).replace(",", "."))
        data["TVA"] = float(tva_match.group(1).replace(",", ".")) if tva_match else 0
        data["TTC"] = float(ttc_match.group(1).replace(",", "."))
    else:
        # fallback ancien système
        tva_block = re.search(r"TVA.*?(\d+[.,]\d{2})\s*€.*?(\d+[.,]\d{2})\s*€", text_clean)

        if tva_block:
            data["HT"] = float(tva_block.group(1).replace(",", "."))
            data["TVA"] = float(tva_block.group(2).replace(",", "."))
        else:
            data["HT"] = 0
            data["TVA"] = 0

        total_match = re.search(r"Montant total\s*(\d+[.,]\d{2})", text_clean)
        data["TTC"] = float(total_match.group(1).replace(",", ".")) if total_match else 0

    # -------- NOM --------
    lines = text.split("\n")
    nom = ""

    for i in range(len(lines)):
        if "Client" in lines[i]:
            if i + 1 < len(lines):
                nom = lines[i + 1].strip()
                break

    # fallback
    if not nom:
        for i in range(len(lines)):
            if "Facture" in lines[i]:
                for j in range(i+1, i+5):
                    if j < len(lines):
                        line = lines[j].strip()
                        if len(line) > 3 and "France" not in line:
                            nom = line
                            break
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

    # -------- TYPE --------
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
