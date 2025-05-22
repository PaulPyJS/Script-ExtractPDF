import os
import logging

import re
import camelot
import fitz  # PyMuPDF

import pandas as pd
from pandas import ExcelWriter



# === Script : EXTRACT TABLE FROM PDF ===
#
def extraire_pdf_vers_excel(pdf_path, output_dir):
    # Log message to supress
    logging.getLogger("pdfminer").setLevel(logging.ERROR)

    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_file = os.path.join(output_dir, f"{pdf_name}_extraction.xlsx")

    # TO BE MODIFIED TO LET THE USER INPUT WHAT HE WANT
    keywords = {
        "Essai laboratoire": ["essai laboratoire", "essais laboratoire", "essais laboratoires", "laboratoire", "laboratoires"],
        "Essais in situ": ["essai in situ", "essais in situ", "in situ", "insitu", "in-situ"],
        "Commentaire": ["commentaire", "commentaires"],
        "Problématique": ["problématique", "problématiques"],
    }

    def looks_like_data(row):
        non_empty = [cell for cell in row if cell and str(cell).strip()]
        return sum(bool(re.search(r'\d', str(cell))) for cell in non_empty) >= 3

    doc = fitz.open(pdf_path)
    target_pages = []

    for i in range(len(doc)):
        text = doc.load_page(i).get_text().lower()
        for canonical, variants in keywords.items():
            if any(variant in text for variant in variants):
                target_pages.append((str(i + 1), canonical))
                break

    print(f"🔍 Pages retenues : {target_pages}")

    # Ghost tables to be removed if size lower than :
    min_rows = 3
    min_cols = 4

    all_tables = []
    pages_sans_tableaux = []

    for page, keyword in target_pages:
        print(f"📄 Traitement de la page {page} - Type : {keyword}")

        # Using multiple option to detect tables
        # Try LATTICE
        tables = camelot.read_pdf(
            pdf_path,
            pages=page,
            flavor="lattice",
            line_scale=40,
            shift_text=["", ""],
            copy_text=["v"],
        )
        valid_tables = [t for t in tables if t.df.shape[0] >= min_rows and t.df.shape[1] >= min_cols]

        # Try STREAM -- only if LATTICE not valid
        if not valid_tables:
            print(f"⚠️ Re-tentative avec flavor=stream sur la page {page}")
            tables = camelot.read_pdf(
                pdf_path,
                pages=page,
                flavor="stream",
                strip_text="\n"
            )
            valid_tables = [t for t in tables if t.df.shape[0] >= min_rows and t.df.shape[1] >= min_cols]

        # If still no valid table
        if not valid_tables:
            print(f"⚠️ Aucun tableau valide détecté sur la page {page}")
            pages_sans_tableaux.append((int(page), keyword))
            continue

        tables = valid_tables


        for i, table in enumerate(tables):
            raw_table = table.df.values.tolist()

            # Normalize the table to make sure every row got the same columns number :
            #   To keep blanks
            expected_cols = max(len(row) for row in raw_table)
            normalized = [
                row + [""] * (expected_cols - len(row)) if len(row) < expected_cols else row
                for row in raw_table
            ]

            df = pd.DataFrame(normalized)
            if df.shape[0] < min_rows or df.shape[1] < min_cols:
                print(f"🚫 Tableau ignoré (trop petit) - Page {page} Table {i+1} ({df.shape[0]} lignes, {df.shape[1]} colonnes)")
                continue


            data_start_idx = None
            for idx, row in df.iterrows():
                if looks_like_data(row):
                    data_start_idx = idx
                    break

            if data_start_idx is not None:
                df_clean = df.iloc[data_start_idx:].copy()
                sheet_name = f"Page{page}_{keyword}"[:31]

                if keyword not in ["Essai laboratoire", "Essais in situ"]:
                    has_group_lines = any(
                        (row[0] and all(str(cell).strip() == "" for cell in row[1:]))
                        for row in df_clean.values
                    )

                    if has_group_lines:
                        df_clean = df_clean[~((df_clean[0].notna()) & (df_clean.iloc[:, 1:].isna().all(axis=1)))]
                        df_clean = df_clean.replace(r'\n', ' ', regex=True)


                num_header_rows = 3
                header_rows = df.iloc[:num_header_rows].values.tolist()

                fused_headers = []
                for col_idx in range(df.shape[1]):
                    parts = []
                    for row in header_rows:
                        if col_idx < len(row):
                            val = str(row[col_idx]).strip()
                            if val and val.lower() not in ["", "nan"]:
                                parts.append(val)
                    header = " ".join(parts).strip()
                    fused_headers.append(header if header else f"col_{col_idx}")


                if any(h.startswith("col_") is False for h in fused_headers):
                    df_clean.columns = fused_headers
                else:
                    df_clean.columns = [f"col_{j}" for j in range(df_clean.shape[1])]

                df_clean.reset_index(drop=True, inplace=True)

                all_tables.append((sheet_name, df_clean))
            else:
                print(f"🚫 Aucune ligne de données trouvée - Page {page} Table {i + 1}")
                pages_sans_tableaux.append((int(page), keyword))


    # Excel export
    if all_tables:
        with ExcelWriter(output_file, engine="openpyxl") as writer:
            for name, df in all_tables:
                df.to_excel(writer, sheet_name=name, index=False)
        print(f"\n✅ Export terminé dans : {output_file}")
    else:
        print("\n❌ Aucun tableau pertinent trouvé.")

    return target_pages, pages_sans_tableaux


