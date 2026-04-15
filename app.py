import fitz
import pandas as pd
import streamlit as st

st.title("📄 → 📊 Compta simple (fiable)")

uploaded_file = st.file_uploader("Importer une facture PDF", type="pdf")

# -------- EXTRACTION TEXTE --------
def extract_text(file):
    text = ""
    pdf = fitz.open(stream=file.read(), filetype="pdf")
    for page in pdf:
        text += page.get_text()
    return text

if uploaded_file:

    text = extract_text(uploaded_file)

    st.subheader("📄 Texte extrait (copie/colle ce dont tu as besoin)")
    st.text_area("Texte du PDF", text, height=300)

    st.subheader("✏️ Saisie des informations")

    col1, col2 = st.columns(2)

    with col1:
        date = st.text_input("Date (YYYY-MM-DD)")
        numero = st.text_input("Numéro facture")
        nom = st.text_input("Nom client / fournisseur")

    with col2:
        ht = st.number_input("Montant HT", step=0.01)
        tva = st.number_input("TVA", step=0.01)
        ttc = st.number_input("TTC", step=0.01)

    fournisseur = st.text_input("Fournisseur")
    type_op = st.selectbox("Type", ["Achat", "Vente"])

    if st.button("📥 Générer Excel"):

        data = [{
            "Date": date,
            "Numéro": numero,
            "HT": ht,
            "TVA": tva,
            "TTC": ttc,
            "Nom": nom,
            "Fournisseur": fournisseur,
            "Type": type_op
        }]

        df = pd.DataFrame(data)

        st.dataframe(df)

        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
