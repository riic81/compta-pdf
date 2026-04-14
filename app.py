import pdfplumber
import re
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# -------- OUTILS --------
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

# -------- APP --------
files = []

def select_files():
    global files
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    label.config(text=f"{len(files)} fichier(s) sélectionné(s)")


def generate_excel():
    if not files:
        messagebox.showerror("Erreur", "Aucun fichier sélectionné")
        return

    results = []

    for file in files:
        with pdfplumber.open(file) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
            results.append(extract_data(text))

    df = pd.DataFrame(results)
    df = df[df["Date_obj"].notna()]
    df = df.sort_values("Date_obj")
    df["Mois"] = df["Date_obj"].dt.strftime("%Y-%m")

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Journal", index=False)

        for mois, group in df.groupby("Mois"):
            ventes = group[group["Type"] == "Vente"]
            achats = group[group["Type"] == "Achat"]

            bilan = pd.DataFrame({
                "Indicateur": ["TVA collectée", "TVA déductible", "TVA à payer"],
                "Montant": [ventes["TVA"].sum(), achats["TVA"].sum(), ventes["TVA"].sum() - achats["TVA"].sum()]
            })

            ventes.to_excel(writer, sheet_name=f"{mois}_ventes", index=False)
            achats.to_excel(writer, sheet_name=f"{mois}_achats", index=False)
            bilan.to_excel(writer, sheet_name=f"{mois}_TVA", index=False)

    messagebox.showinfo("Succès", "Fichier Excel généré !")


# -------- UI --------
root = tk.Tk()
root.title("Compta PDF simple")
root.geometry("400x200")

btn_select = tk.Button(root, text="📂 Choisir les factures PDF", command=select_files)
btn_select.pack(pady=10)

label = tk.Label(root, text="Aucun fichier sélectionné")
label.pack()

btn_generate = tk.Button(root, text="🚀 Générer le fichier Excel", command=generate_excel)
btn_generate.pack(pady=20)

root.mainloop()

# -------- INSTRUCTIONS --------
# pip install pdfplumber pandas openpyxl
# pyinstaller --onefile --noconsole app.py
