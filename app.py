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


# -------- EXTRACTION DONNÉES --------
def extract_data(text):

    text_clean = text.replace("\n", " ")
    lines = text.split("\n")  # ⚠️ IMPORTANT (corrige ton bug)

    # -------- CLIENT --------
    client = ""

    for i in range(len(lines)):
        if "Facture" in lines[i]:
            for j in range(i+1, i+6):
                if j < len(lines):
                    line = lines[j].strip()
                    if len(line) > 3 and "France" not in line:
                        client = line
                        break
            break

    if not client:
        client = "Client inconnu"

    # -------- NUMERO --------
    num_match = re.search(r"Numéro de facture\s*([0-9]+)", text_clean)
    numero = "F" + num_match.group(1) if num_match else ""

    # -------- DATE --------
    date_match = re.search(r"Date\s*([0-9]{1,2}\s+[a-zéû\.]+\s+[0-9]{4})", text_clean, re.IGNORECASE)

    mois_map = {
        "janv.": "01", "févr.": "02", "mars": "03", "avr.": "04",
        "mai": "05", "juin": "06", "juil.": "07", "août": "08",
        "sept.": "09", "oct.": "10", "nov.": "11", "déc.": "12"
    }

    if date_match:
        parts = date_match.group(1).split()
        jour = parts[0]
        mois = mois_map.get(parts[1].lower(), "01")
        annee = parts[2]
        date = f"{annee}-{mois}-{jour.zfill(2)}"
    else:
        date = ""

    # -------- HT --------
    ht_match = re.search(r"20\s*%?\s*TVA\s*de\s*([0-9]+[.,][0-9]{2})", text_clean)
    HT = float(ht_match.group(1).replace(",", ".")) if ht_match else 0

    # -------- TVA --------
    tva_match = re.search(r"20\s*%?\s*TVA.*?([0-9]+[.,][0-9]{2})\s*€", text_clean)
    TVA = float(tva_match.group(1).replace(",", ".")) if tva_match else 0

    # -------- TTC --------
    ttc_match = re.search(r"Montant total\s*([0-9]+[.,][0-9]{2})", text_clean)
    TTC = float(ttc_match.group(1).replace(",", ".")) if ttc_match else 0

    # -------- MODE DE PAIEMENT (corrigé) --------
    paiement = ""

    for i in range(len(lines)):
        if "Mode de paiement" in lines[i]:
            if i + 1 < len(lines):
                paiement = lines[i + 1].strip().lower()

    # NORMALISATION
    if "carte" in paiement:
        paiement = "Carte bleue"
    elif "paypal" in paiement:
        paiement = "Paypal"
    elif "virement" in paiement:
        paiement = "Virement anticipé"
    else:
        paiement = "Virement anticipé"

    return {
        "Client": client,
        "Référence (numéro)": numero,
        "Date de paiement": date,
        "Moyen de paiement": paiement,
        "Montant HT": HT,
        "Taux de TVA": 20,
        "Montant TTC": TTC,
        "Type de vente (1,2,3,4)": 1
    }


# -------- TRAITEMENT --------
if uploaded_files:
    if st.button("🚀 Générer CSV ABBYY"):

        results = []

        for file in uploaded_files:
            text = extract_text(file)
            results.append(extract_data(text))

        df = pd.DataFrame(results)

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
