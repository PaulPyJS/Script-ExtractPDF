import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Listbox, MULTIPLE, END
import pandas as pd
import os

template_path = os.path.join(os.path.dirname(__file__), "template_attributaire.xlsx")

# === Script : FIXED MAPPING TO MERGE DATA FROM PDF EXTRACT ===
# = v2 : Based on template of mapping from Excel that can be filled by others
#
def appliquer_mapping_rapide():
    path_extraction = filedialog.askopenfilename(title="Fichier d'extraction", filetypes=[("Excel", "*.xlsx")])
    if not path_extraction:
        return

    xls = pd.ExcelFile(path_extraction, engine="openpyxl")
    feuilles = xls.sheet_names

    popup = Toplevel()
    popup.title("Sélection des feuilles")

    tk.Label(popup, text="Feuilles pour Essais in situ").pack()
    list_insitu = Listbox(popup, selectmode=MULTIPLE, exportselection=False)
    list_insitu.pack()
    for f in feuilles:
        list_insitu.insert(END, f)

    tk.Label(popup, text="Feuilles pour Essai laboratoire").pack()
    list_labo = Listbox(popup, selectmode=MULTIPLE, exportselection=False)
    list_labo.pack()
    for f in feuilles:
        list_labo.insert(END, f)

    def valider_feuilles():
        feuilles_insitu = [list_insitu.get(i) for i in list_insitu.curselection()]
        feuilles_labo = [list_labo.get(i) for i in list_labo.curselection()]
        popup.destroy()
        traiter_concat_filtrage(path_extraction, feuilles_insitu, feuilles_labo)

    tk.Button(popup, text="Valider", command=valider_feuilles).pack(pady=5)

def traiter_concat_filtrage(path_extraction, feuilles_insitu, feuilles_labo):
    if os.path.exists(template_path):
        path_attributaire = template_path
    else:
        path_attributaire = filedialog.askopenfilename(
            title="Fichier de table attributaire (avec headers)",
            filetypes=[("Excel", "*.xlsx")]
        )
    if not path_attributaire:
        return


    def charger_feuilles(path, feuilles):
        frames = []
        for feuille in feuilles:
            df = pd.read_excel(path, sheet_name=feuille, engine="openpyxl")
            df = df.iloc[1:]  # Ignorer la première ligne (en-têtes)
            df.columns = [f"col_{i}" for i in range(len(df.columns))]
            frames.append(df)
        return pd.concat(frames, ignore_index=True)

    df_insitu = charger_feuilles(path_extraction, feuilles_insitu)
    df_labo = charger_feuilles(path_extraction, feuilles_labo)

    mapping_df = pd.read_excel(path_attributaire, sheet_name=1, engine="openpyxl")
    valeurs_insitu = mapping_df.iloc[:, 2]
    valeurs_labo = mapping_df.iloc[:, 4]

    colonnes_insitu = []
    colonnes_labo = []
    colonnes_finales = []
    vide_count = 1

    for i in range(len(mapping_df)):
        val_in = valeurs_insitu.iloc[i]
        val_lab = valeurs_labo.iloc[i]

        if (pd.isna(val_in) or str(val_in).strip().lower() == "vide") and (
                pd.isna(val_lab) or str(val_lab).strip().lower() == "vide"):
            colonnes_finales.append(f"vide_{vide_count}")
            vide_count += 1
            continue

        if pd.notna(val_in) and str(val_in).strip().lower() != "vide":
            col = f"col_{int(val_in) - 1}"
            colonnes_insitu.append(col)
            colonnes_finales.append(col)

        if pd.notna(val_lab) and str(val_lab).strip().lower() != "vide":
            col = f"col_{int(val_lab) - 1}"
            col_labo = col + "_labo" if col in colonnes_insitu else col
            colonnes_labo.append(col_labo)
            colonnes_finales.append(col_labo)


    df_insitu_reduit = df_insitu[[col for col in colonnes_insitu if col in df_insitu.columns]].copy()
    df_labo_reduit = pd.DataFrame()
    for col in colonnes_labo:
        original = col.replace("_labo", "")
        if original in df_labo.columns:
            df_labo_reduit[col] = df_labo[original]


    df_insitu_final = pd.DataFrame(columns=colonnes_finales)
    df_labo_final = pd.DataFrame(columns=colonnes_finales)

    for col in colonnes_insitu:
        if col in df_insitu_reduit.columns:
            df_insitu_final[col] = df_insitu_reduit[col]
    for col in colonnes_labo:
        if col in df_labo_reduit.columns:
            df_labo_final[col] = df_labo_reduit[col]

    df_insitu_final.fillna("", inplace=True)
    df_labo_final.fillna("", inplace=True)

    # ======================================== DEBUG =======================================================
    # debug_path = os.path.join(os.path.dirname(path_extraction), "debug_raw_insitu_labo.xlsx")
    # with pd.ExcelWriter(debug_path, engine="openpyxl") as writer:
    #     df_insitu_final.to_excel(writer, sheet_name="InSitu_final", index=False)
    #     df_labo_final.to_excel(writer, sheet_name="Labo_final", index=False)
    # ======================================== DEBUG =======================================================

    df_fusion = df_insitu_final.copy().astype("object")
    df_labo_final = df_labo_final.astype("object")

    max_rows = max(len(df_fusion), len(df_labo_final))
    while len(df_fusion) < max_rows:
        df_fusion.loc[len(df_fusion)] = [""] * len(df_fusion.columns)
    while len(df_labo_final) < max_rows:
        df_labo_final.loc[len(df_labo_final)] = [""] * len(df_labo_final.columns)

    for i in range(max_rows):
        for j in range(len(colonnes_finales)):
            cell_insitu = df_fusion.iat[i, j]
            cell_labo = df_labo_final.iat[i, j]
            if (pd.isna(cell_insitu) or cell_insitu == "") and pd.notna(cell_labo) and cell_labo != "":
                df_fusion.iat[i, j] = str(cell_labo)


    noms_colonnes = mapping_df.iloc[:, 1].astype(str).tolist()
    while len(noms_colonnes) < len(df_fusion.columns):
        noms_colonnes.append(f"colonne_sans_nom_{len(noms_colonnes) + 1}")
    df_fusion.columns = noms_colonnes[:len(df_fusion.columns)]


    nom_base = os.path.splitext(os.path.basename(path_extraction))[0]
    output_path = os.path.join(os.path.dirname(path_extraction), f"{nom_base}_nettoyé.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_fusion.to_excel(writer, sheet_name="Fusion", index=False)

    messagebox.showinfo("Succès", f"Fichier final exporté :\n{output_path}")
    print("INSITU:", colonnes_insitu)
    print("LABO  :", colonnes_labo)
    print("FINALES:", colonnes_finales)





