if uploaded_files:
    if st.button("🚀 Générer"):
        results = []

        for file in uploaded_files:
            with pdfplumber.open(file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() + "\n"
                results.append(extract_data(text))

        # ✅ créer df AVANT de l'utiliser
        df = pd.DataFrame(results)

        # ✅ convertir en date
        df["Date_obj"] = pd.to_datetime(df["Date_obj"], errors="coerce")

        # ✅ nettoyage + tri
        df = df[df["Date_obj"].notna()]
        df = df.sort_values("Date_obj")

        # ✅ mois
        df["Mois"] = df["Date_obj"].dt.strftime("%Y-%m")

        # affichage
        st.dataframe(df)

        # export Excel
        with pd.ExcelWriter("compta.xlsx", engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Journal", index=False)

        with open("compta.xlsx", "rb") as f:
            st.download_button("📥 Télécharger Excel", f, "compta.xlsx")
