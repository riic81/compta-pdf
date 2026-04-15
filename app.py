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
        data["Date"] = pd.to_datetime(date_match.group(0), dayfirst=True).strftime("%Y-%m-%d")
    else:
        data["Date"] = ""

    # -------- NUMERO --------
    num_match = re.search(r"No\. Facture\s*([0-9]+)", text_clean)
    if not num_match:
        num_match = re.search(r"facture\s*([0-9]+)", text_clean, re.IGNORECASE)
    data["Numéro"] = num_match.group(1) if num_match else ""

    # -------- DETECTION TYPE --------
    is_kramp = "kramp" in text.lower()

    # ==============================
    # 🔵 CAS KRAMP
    # ==============================
    if is_kramp:

        # HT
        ht_match = re.search(r"Montant H\.T\.\s*([0-9]+[.,][0-9]{2})", text_clean)
        data["HT"] = float(ht_match.group(1).replace(",", ".")) if ht_match else 0

        # TTC
        ttc_match = re.search(r"Montant T\.T\.C\.\s*(?:EUR)?\s*([0-9]+[.,][0-9]{2})", text_clean)
        data["TTC"] = float(ttc_match.group(1).replace(",", ".")) if ttc_match else 0

        # TVA (calcul fiable)
        data["TVA"] = round(data["TTC"] - data["HT"], 2) if data["TTC"] else 0

    # ==============================
    # 🟢 AUTRES FACTURES (HYDRO)
    # ==============================
    else:

        # récupérer tous les montants
        montants = re.findall(r"[0-9]+[.,][0-9]{2}", text_clean)

        if len(montants) >= 3:
            data["HT"] = float(montants[-3].replace(",", "."))
            data["TVA"] = float(montants[-2].replace(",", "."))
            data["TTC"] = float(montants[-1].replace(",", "."))
        else:
            data["HT"] = 0
            data["TVA"] = 0
            data["TTC"] = 0

    # -------- NOM --------
    lines = text.split("\n")
    nom = ""

    for i in range(len(lines)):
        if "Client" in lines[i]:
            if i + 1 < len(lines):
                nom = lines[i + 1].strip()
                break

    if not nom:
        for line in lines:
            if len(line.strip()) > 5 and line.strip().isupper():
                nom = line.strip()
                break

    data["Nom"] = nom

    # -------- FOURNISSEUR --------
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
