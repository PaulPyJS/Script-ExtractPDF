import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
import os
import threading

from image_table_extract import ocr_sur_images_decoupees, traiter_tableaux_image, detecter_tableaux_par_image
from pdf_table_extract import  extraire_pdf_vers_excel
from mapping_tool import appliquer_mapping_rapide



# === Script : UI to help users choose from extract or direct mapping ===
# == v2 : Ask users to check for table in images format
#
def lancer_extraction():
    pdf_paths = list(pdf_listbox.get(0, tk.END))
    if not pdf_paths:
        messagebox.showwarning("Aucun fichier", "Veuillez sélectionner des fichiers PDF.")
        return
    if not dossier_sortie.get():
        messagebox.showwarning("Aucun dossier", "Veuillez choisir un dossier de sortie.")
        return

    resultat = []

    def traitement():
        for pdf_path in pdf_paths:
            nom_fichier = os.path.basename(pdf_path)
            log.insert(tk.END, f"📄 Traitement : {nom_fichier}\n")
            try:
                pages_ret, pages_sans_tableaux = extraire_pdf_vers_excel(pdf_path, dossier_sortie.get())
                print("DEBUG - Pages sans tableau :", pages_sans_tableaux)
                resultat.append((pdf_path, pages_ret, pages_sans_tableaux))
            except Exception as e:
                log.insert(tk.END, f"❌ Erreur : {nom_fichier} → {e}\n\n")
            log.see(tk.END)

        root.after(100, suite_traitement)

    def suite_traitement():
        for pdf_path, pages_ret, pages_sans_tableaux in resultat:
            nom_fichier = os.path.basename(pdf_path)

            # Pages without tables
            if pages_sans_tableaux:
                log.insert(tk.END, f"⚠️ Certaines pages sans tableau détecté :\n")
                for p, k in pages_sans_tableaux:
                    log.insert(tk.END, f"   - Page {p} ({k})\n")

                if messagebox.askyesno("Tableaux image",
                                       f"{nom_fichier} contient des pages sans tableau détecté.\nSouhaitez-vous traiter ces pages comme images ?"):
                    reponse = simpledialog.askstring("Pages à traiter",
                                                     "Entrez les pages à traiter comme images (ex: 5, 7, 8) :")
                    if reponse:
                        page_nums = [int(p.strip()) for p in reponse.split(",") if p.strip().isdigit()]
                        traiter_tableaux_image(pdf_path, page_nums, dossier_sortie.get())

            # Pages with tables found
            if pages_ret:
                log.insert(tk.END, "📌 Pages retenues :\n")
                for p, k in pages_ret:
                    log.insert(tk.END, f"   - Page {p} : {k}\n")
            else:
                log.insert(tk.END, "📌 Aucune page pertinente détectée\n")

            log.insert(tk.END, f"✅ Terminé : {nom_fichier}\n\n")
            log.see(tk.END)

    threading.Thread(target=traitement).start()

def choisir_pdfs():
    fichiers = filedialog.askopenfilenames(filetypes=[("Fichiers PDF", "*.pdf")])
    for f in fichiers:
        if f not in pdf_listbox.get(0, tk.END):
            pdf_listbox.insert(tk.END, f)

def choisir_dossier():
    dossier = filedialog.askdirectory()
    dossier_sortie.set(dossier)

def vider_selection():
    pdf_listbox.delete(0, tk.END)

def tester_detection_opencv():
    pdf_path = filedialog.askopenfilename(title="Choisir un PDF", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return

    page_str = simpledialog.askstring("Page à traiter", "Numéro de page du PDF à extraire :")
    if not page_str or not page_str.isdigit():
        messagebox.showerror("Erreur", "Numéro de page invalide.")
        return

    page_num = int(page_str)
    try:
        dossier = dossier_sortie.get()
        if not dossier:
            messagebox.showwarning("Dossier", "Veuillez d’abord choisir un dossier de sortie.")
            return

        detecter_tableaux_par_image(pdf_path, [page_num], dossier)

        # OCR
        dossier_images = dossier
        ocr_output_file = os.path.join(dossier, f"{os.path.basename(pdf_path).split('.')[0]}_ocr_tables.xlsx")

        if not any(f.lower().endswith(".png") for f in os.listdir(dossier_images)):
            messagebox.showwarning("Pas d’images", "Aucune image détectée dans le dossier sélectionné.")
            return

        tableaux = ocr_sur_images_decoupees(dossier_images, ocr_output_file, zoom_factor=2)

        if tableaux:
            messagebox.showinfo("Succès", f"{len(tableaux)} tableaux extraits et enregistrés.")
        else:
            messagebox.showinfo("Aucun tableau", "Aucun tableau détecté.")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur pendant l'extraction :\n{e}")



# =============================================================
# ------------------------- Interface -------------------------
# =============================================================

root = tk.Tk()
root.title("Extracteur de Tableaux PDF → Excel")
root.geometry("500x700")

main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill=tk.BOTH, expand=True)
dossier_sortie = tk.StringVar(value="Aucun dossier de sortie sélectionné")


tk.Label(main_frame, text="Fichiers PDF sélectionnés :", font=("Segoe UI", 10, "bold")).pack(pady=10)
pdf_listbox = tk.Listbox(main_frame, height=5, width=100)
pdf_listbox.pack(pady=(0, 10))
tk.Button(main_frame, text="Ajouter PDF", command=choisir_pdfs).pack(pady=2)
tk.Button(main_frame, text="🗑️ Vider la sélection", command=vider_selection).pack(pady=2)


ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=5, pady=5)
tk.Label(main_frame, text="Traitement : PDF → Excel", font=("Segoe UI", 10, "bold")).pack(pady=5)
tk.Button(main_frame, text="📁 Choisir dossier de sortie", command=choisir_dossier).pack(pady=5)
tk.Label(main_frame, textvariable=dossier_sortie, fg="blue").pack()
tk.Button(main_frame, text="🚀 Lancer l'extraction", command=lancer_extraction, bg="green", fg="white").pack(pady=10)


ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=5, pady=5)
tk.Label(main_frame, text="Traitement : Image → Excel (via OCR)", font=("Segoe UI", 10, "bold")).pack(pady=5)
tk.Button(main_frame, text="🧩 Tester détection tableau image (OpenCV)", command=tester_detection_opencv).pack(pady=5)


ttk.Separator(main_frame, orient='horizontal').pack(fill='x', padx=5, pady=5)
tk.Label(main_frame, text="Traitement : Excel → Table attributaire", font=("Segoe UI", 10, "bold")).pack(pady=5)
tk.Button(main_frame, text="⚡ Mapping express vers table attributaire", command=appliquer_mapping_rapide, bg="green", fg="white").pack(pady=5)


# LOG
tk.Label(main_frame, text="📝 Log :", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(20, 0))
log = scrolledtext.ScrolledText(main_frame, height=15, width=100)
log.pack()

root.mainloop()