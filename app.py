import fitz
import re
import pandas as pd
import streamlit as st

st.title("📄 → 📊 Compta automatique")

uploaded_files = st.file_uploader(
    "Importer vos factures PDF",
    type="pdf",
    accept_multiple_files=True
)

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

    # -------- DATE --------
    date_match = re.search(r"[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4}", text_clean)
    data["Date"] = pd.to_datetime(date_match.group(0), dayfirst=True).strftime("%Y-%m-%d") if date_match else ""

    # -------- NUMERO --------
    num_match = re.search(r"facture\s*([0-9]+)", text_clean, re.IGNORECASE)
    data["Numéro"] = num_match.group(1) if num_match else ""

    # -------- FOURNISSEUR --------
    is_kramp = "kramp" in text.lower()

    # ==============================
    # 🔵 KRAMP
    # ==============================
    if is_kramp:

        ht_match = re.search(r"Montant H\.T\.\s*([0-9]+[.,][0-9]{2})", text_clean)
        ttc_match = re.search(r"Montant T\.T\.C\.\s*(?:EUR)?\s*([0-9]+[.,][0-9]{2})", text_clean)

        HT = float(ht_match.group(1).replace(",", ".")) if ht_match else 0
        TTC = float(ttc_match.group(1).replace(",", ".")) if ttc_match else 0
        TVA = round(TTC - HT, 2) if TTC else 0

    # ==============================
    # 🟢 HYDRO
    # ==============================
    else:

        # récupérer TOUS les montants
        montants = re.findall(r"[0-9]+[.,][0-9]{2}", text_clean)

        # prendre les 3 plus grands (souvent HT TVA TTC)
        montants_float = sorted([float(m.replace(",", ".")) for m in montants], reverse=True)

        if len(montants_float) >= 3:
            TTC = montants_float[0]
            TVA = montants_float[1]
            HT = montants_float[2]
        else:
            HT = TVA = TTC = 0

    data["HT"] = HT
    data["TVA"] = TVA
    data["TTC"] = TTC

    # -------- NOM --------
    lines = text.split("\n")
    nom = ""
    for i in range(len(lines)):
        if "Client" in lines[i]:
            if i + 1 < len(lines):
                nom = lines[i + 1].strip()
                break
    data["Nom"] = nom

    # -------- FOURNISSEUR FINAL --------
    if is_kramp:
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

        st.dataframe(df)

        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Journal", index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
