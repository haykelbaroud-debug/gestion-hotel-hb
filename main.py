import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox, simpledialog
from operator import itemgetter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import database as db
import calculs as calc

ctk.set_appearance_mode("light")

C_FOND = "#F5F7FA"
C_BLEU_NUIT = "#1F3A56"
C_BLEU_ACIER = "#2B4C6E"
C_CYAN = "#217387"
C_JAUNE_HDR = "#E6B432"
C_JAUNE_BTN = "#FFDC8C"
C_VERT = "#288C5F"
C_BRONZE = "#BE8C3C"
C_GRIS = "#7D8791"
C_ROUGE = "#B42828"

CODE_SECRET = "5227"

def val_idx(seq, i):
    return itemgetter(i)(seq)

class FrmSelectorPopup(ctk.CTkToplevel):
    def __init__(self, parent, titre, items, callback, avec_recherche=True):
        super().__init__(parent)
        self.title(titre)
        self.geometry("320x380")
        self.configure(fg_color="#E1E6EB")
        self.attributes("-topmost", True)
        self.callback = callback
        self.items_initiaux = list(items)

        if avec_recherche:
            self.entry_search = ctk.CTkEntry(self, placeholder_text="Rechercher...", fg_color="white", text_color="#1E2832")
            self.entry_search.pack(fill="x", padx=8, pady=(8, 4))
            self.entry_search.bind("<KeyRelease>", self.filtrer)

        self.frame_list = ctk.CTkScrollableFrame(self, fg_color="white")
        self.frame_list.pack(fill="both", expand=True, padx=8, pady=8)
        self.afficher_liste(self.items_initiaux)

    def afficher_liste(self, liste):
        for w in self.frame_list.winfo_children():
            w.destroy()
        for item in liste:
            btn = ctk.CTkButton(self.frame_list, text=item, anchor="w", fg_color="transparent",
                                text_color="#1E2832", hover_color="#DCE6F1", height=28,
                                command=lambda val=item: self.valider(val))
            btn.pack(fill="x", pady=1)

    def filtrer(self, event=None):
        saisie = self.entry_search.get().upper().strip()
        if not saisie:
            self.afficher_liste(self.items_initiaux)
        else:
            filtres = [x for x in self.items_initiaux if saisie in x.upper()]
            self.afficher_liste(filtres)

    def valider(self, val):
        self.callback(val)
        self.destroy()


class HotelERPApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        db.initialiser_base()
        self.title("GESTION DES ACHATS & CONSOMMATIONS F&B — PRO")
        self.geometry("1220x740")
        self.configure(fg_color=C_FOND)

        self.mode_deverrouille = False
        self.mois_conso_choisi = "Choisir Mois..."
        self.dept_conso_choisi = "Choisir Département..."
        self.dept_code_choisi = ""
        self.sf_conso_choisie = "Sous-Famille..."
        self.art_conso_choisi = "Article..."
        self.mode_bascule_conso = "QUANTITE"
        self.scope_bascule_conso = "ARTICLE"

        self.container = ctk.CTkFrame(self, fg_color=C_FOND)
        self.container.pack(fill="both", expand=True)

        self.afficher_accueil_achats()

    def nettoyer_ecran(self):
        for w in self.container.winfo_children():
            w.destroy()
    def creer_tableau_avec_scroll(self, parent, cols):
        scroll_y = ttk.Scrollbar(parent, orient="vertical")
        tree = ttk.Treeview(parent, columns=cols, show="headings", yscrollcommand=scroll_y.set)
        scroll_y.config(command=tree.yview)
        scroll_y.pack(side="right", fill="y", padx=(0, 5), pady=10)
        tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=10)
        return tree

    def verifier_code_secret(self, action_nom="modifier"):
        saisie = simpledialog.askstring("Zone Sécurisée", f"Veuillez saisir le code secret pour {action_nom} :", show="*", parent=self)
        if saisie == CODE_SECRET:
            return True
        elif saisie is not None:
            messagebox.showerror("Accès Refusé", "Code secret incorrect.")
        return False
    def afficher_accueil_achats(self):
        self.nettoyer_ecran()

        header = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=60, corner_radius=6)
        header.pack(fill="x", padx=60, pady=(25, 15))
        ctk.CTkLabel(header, text="TABLEAU DE BORD DE GESTION DES ACHATS",
                     font=("Segoe UI", 20, "bold"), text_color="white").pack(pady=15)

        grid_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        grid_frame.pack(fill="both", expand=True, padx=60, pady=10)
        grid_frame.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(grid_frame, text="SAISIES MENSUELLES", fg_color=C_JAUNE_HDR, text_color="white",
                     font=("Segoe UI", 12, "bold"), height=34, corner_radius=4).grid(row=0, column=0, padx=15, pady=(0, 10), sticky="ew")
        ctk.CTkLabel(grid_frame, text="ANALYSES STRATÉGIQUES", fg_color=C_BLEU_NUIT, text_color="white",
                     font=("Segoe UI", 12, "bold"), height=34, corner_radius=4).grid(row=0, column=1, padx=15, pady=(0, 10), sticky="ew")
        ctk.CTkLabel(grid_frame, text="BILANS & CONFIGURATION", fg_color="#5A6473", text_color="white",
                     font=("Segoe UI", 12, "bold"), height=34, corner_radius=4).grid(row=0, column=2, padx=15, pady=(0, 10), sticky="ew")

        for i, mois in enumerate(db.LISTE_MOIS):
            btn_s = ctk.CTkButton(grid_frame, text=mois, fg_color=C_JAUNE_BTN, text_color="#64460A",
                                  hover_color="#F5CE76", font=("Segoe UI", 11, "bold"), height=30,
                                  command=lambda m=mois: self.ouvrir_saisie_achats_mois(m))
            btn_s.grid(row=i+1, column=0, padx=25, pady=3, sticky="ew")

            btn_a = ctk.CTkButton(grid_frame, text=mois, fg_color=C_BLEU_ACIER, text_color="white",
                                  hover_color=C_BLEU_NUIT, font=("Segoe UI", 11, "bold"), height=30,
                                  command=lambda m=mois: self.ouvrir_strategie_mois(m))
            btn_a.grid(row=i+1, column=1, padx=25, pady=3, sticky="ew")

        actions_droite = [
            ("Bilan Annuel Global", C_VERT, self.ouvrir_bilan_annuel_achats),
            ("Suivi Mensuel Annuel", C_VERT, self.ouvrir_suivi_mensuel_achats),
            ("Importer Fichier Excel (.xlsm)", C_CYAN, self.action_importer_excel),
            ("Graphiques Achats", C_GRIS, self.ouvrir_graphiques_achats),
            ("Rapport Exécutif (Top 5)", C_BRONZE, self.ouvrir_rapport_executif),
            ("Mode Administrateur (5227)", "#555F6E", self.basculer_admin),
            ("ACCUEIL CONSOMMATION →", C_BLEU_NUIT, self.afficher_accueil_consommation)
        ]
        for idx, (txt, col, cmd) in enumerate(actions_droite):
            btn_d = ctk.CTkButton(grid_frame, text=txt, fg_color=col, text_color="white",
                                  font=("Segoe UI", 11, "bold"), height=36, command=cmd)
            btn_d.grid(row=idx+1, column=2, padx=25, pady=5, sticky="ew")

    def action_importer_excel(self):
        chemin = filedialog.askopenfilename(title="Sélectionnez votre fichier Excel (.xlsm / .xlsx)",
                                            filetypes=[("Fichiers Excel", "*.xlsm *.xlsx")])
        if chemin:
            try:
                nb_a, nb_c = db.importer_depuis_excel(chemin)
                messagebox.showinfo("Importation Réussie", f"Importation terminée :\n- {nb_a} lignes d'Achats\n- {nb_c} lignes de Consommation importées !")
            except Exception as e:
                messagebox.showerror("Erreur d'importation", f"Impossible de lire le fichier :\n{e}")

    def basculer_admin(self):
        if not self.mode_deverrouille:
            if self.verifier_code_secret("activer le mode Administrateur"):
                self.mode_deverrouille = True
                messagebox.showinfo("Accès Autorisé", "Mode Administrateur activé.")
        else:
            self.mode_deverrouille = False
            messagebox.showinfo("Sécurité Activée", "Application reverrouillée.")

    def ouvrir_saisie_achats_mois(self, mois):
        self.nettoyer_ecran()
        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_achats).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text=f"SAISIE DES ACHATS — {mois.upper()}", font=("Segoe UI", 16, "bold"), text_color="white").pack(side="left", padx=20)

        form = ctk.CTkFrame(self.container, fg_color="white")
        form.pack(fill="x", padx=20, pady=5)
        cb_sf = ctk.CTkComboBox(form, values=sorted(list(db.FAMILLES_VALIDES)), width=200)
        cb_sf.set("MEAT")
        cb_sf.pack(side="left", padx=10, pady=10)
        ent_fourn = ctk.CTkEntry(form, placeholder_text="Nom du Fournisseur", width=250)
        ent_fourn.pack(side="left", padx=10)
        ent_mnt = ctk.CTkEntry(form, placeholder_text="Montant HT (DT)", width=150)
        ent_mnt.pack(side="left", padx=10)

        def ajouter():
            try:
                mnt = float(ent_mnt.get().replace(",", "."))
                fourn = ent_fourn.get().strip()
                sf = cb_sf.get().strip()
                if not fourn:
                    return
                conn = db.get_connection()
                c = conn.cursor()
                c.execute("INSERT INTO achats (mois, code_article, sous_famille, fournisseur, montant_ht) VALUES (?, ?, ?, ?, ?)",
                          (mois, sf, sf, fourn, mnt))
                conn.commit()
                conn.close()
                ent_fourn.delete(0, "end")
                ent_mnt.delete(0, "end")
                charger_table()
            except ValueError:
                messagebox.showerror("Erreur", "Montant invalide.")

                ctk.CTkButton(form, text="+ AJOUTER L'ACHAT", fg_color=C_VERT, command=ajouter).pack(side="left", padx=10)

        def supprimer_achat():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Aucune sélection", "Cliquez d'abord sur une ligne du tableau à supprimer.")
                return
            if messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer cet achat ?"):
                id_achat = tree.item(sel[0])["values"][0]
                conn = db.get_connection()
                conn.cursor().execute("DELETE FROM achats WHERE id = ?", (id_achat,))
                conn.commit()
                conn.close()
                charger_table()

        ctk.CTkButton(form, text="SUPPRIMER", fg_color=C_ROUGE, command=supprimer_achat).pack(side="left", padx=5)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("ID", "Sous-Famille", "Fournisseur", "Montant HT (DT)")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, anchor="center" if col in ("ID", "Sous-Famille") else "w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def charger_table():
            for r in tree.get_children():
                tree.delete(r)
            conn = db.get_connection()
            c = conn.cursor()
            c.execute("SELECT id, sous_famille, fournisseur, montant_ht FROM achats WHERE mois = ? ORDER BY id DESC", (mois,))
            for rid, rsf, rfourn, rmnt in c.fetchall():
                tree.insert("", "end", values=(rid, rsf, rfourn, f"{rmnt:,.3f} DT"))
            conn.close()

        charger_table()

    def ouvrir_strategie_mois(self, mois):
        self.nettoyer_ecran()
        res = calc.calculer_strategie_mois(mois)

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_achats).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text=f"ANALYSE STRATÉGIQUE ACHATS — {mois.upper()}", font=("Segoe UI", 16, "bold"), text_color="white").pack(side="left", padx=20)
        
        def export_pdf():
            lignes = [[l["col1"], f"{l['col2']:,.3f} DT", f"{l['col3']*100:.1f}%", l["col4"]] for l in res["lignes"]]
            calc.exporter_tableau_pdf(f"Strategie_Achats_{mois}", ["Fournisseur / Famille", "Montant HT", "Part %", "Statut"], lignes)

        ctk.CTkButton(top, text="EXPORTER PDF", fg_color=C_GRIS, command=export_pdf).pack(side="right", padx=15)
     ctk.CTkButton(top, text="EXPORTER EXCEL", fg_color=C_VERT,
                   command=lambda: calc.exporter_tableau_excel(f"Consommations_{mois}", list(cols), lignes_pdf)).pack(side="right", padx=5)

        kpi = ctk.CTkFrame(self.container, fg_color="white")
        kpi.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(kpi, text=f"TOTAL ACHATS : {res['total_achats']:,.3f} DT", font=("Segoe UI", 13, "bold"), text_color=C_BLEU_NUIT).pack(side="left", padx=30, pady=12)
        ctk.CTkLabel(kpi, text=f"TOTAL DE CONTRÔLE : {res['total_controle']:,.3f} DT", font=("Segoe UI", 13, "bold"), text_color=C_BLEU_NUIT).pack(side="left", padx=30)
        col_ecart = C_VERT if abs(res["ecart"]) < 0.001 else C_ROUGE
        ctk.CTkLabel(kpi, text=f"ÉCART DE CONTRÔLE : {res['ecart']:,.3f} DT", font=("Segoe UI", 13, "bold"), text_color=col_ecart).pack(side="right", padx=30)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Fournisseur / Sous-Famille", "Montant HT", "Part (%)", "Alerte Dépendance")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        tree.tag_configure("HEADER", background="#2B3C50", foreground="white", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("ALERTE", foreground="#B42828", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("CONFORME", foreground="#288246")

        for l in res["lignes"]:
            tag = "HEADER" if l["type"] == "HEADER" else ("ALERTE" if "Alerte" in l["col4"] else "CONFORME")
            tree.insert("", "end", values=(l["col1"], f"{l['col2']:,.3f} DT", f"{l['col3']*100:.1f}%", l["col4"]), tags=(tag,))
        tree.pack(fill="both", expand=True, padx=10, pady=10)
    def ouvrir_bilan_annuel_achats(self):
        self.nettoyer_ecran()
        _, recap_sf, totaux_sf = calc.calculer_recap_annuel_achats()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_achats).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text="RÉCAPITULATIF DES ACHATS ANNUELS — VUE GLOBALE CUMULÉE", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Fournisseur / Famille", "Montant Cumulé HT", "Part Annuelle (%)")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        tree.tag_configure("HEADER", background="#2B3C50", foreground="white", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("ALERTE", foreground="#B42828", font=("Segoe UI", 10, "bold"))

        lignes_pdf = []
        for sf in sorted(recap_sf.keys()):
            tot = totaux_sf[sf]
            hdr = f"{sf}  |  Total Annuel : {tot:,.3f} DT"
            tree.insert("", "end", values=(hdr, "", ""), tags=("HEADER",))
            lignes_pdf.append([hdr, "", ""])
            for fourn, mnt in sorted(recap_sf[sf].items(), key=lambda x: val_idx(x, 1), reverse=True):
                part = (mnt / tot) if tot > 0 else 0.0
                tag = "ALERTE" if part >= 0.30 else ""
                tree.insert("", "end", values=(fourn, f"{mnt:,.3f} DT", f"{part*100:.1f}%"), tags=(tag,))
                lignes_pdf.append([fourn, f"{mnt:,.3f} DT", f"{part*100:.1f}%"])
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(top, text="EXPORTER PDF", fg_color=C_GRIS,
                      command=lambda: calc.exporter_tableau_pdf("Bilan_Annuel_Achats", list(cols), lignes_pdf)).pack(side="right", padx=15)

    def ouvrir_suivi_mensuel_achats(self):
        self.nettoyer_ecran()
        suivi, _, _ = calc.calculer_recap_annuel_achats()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_achats).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text="SUIVI MENSUEL ANNUEL DES ACHATS", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ["Famille", "Fournisseur"] + [m[:4] for m in db.LISTE_MOIS] + ["Total Général"]
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=75 if col not in ("Famille", "Fournisseur") else 140)

        lignes_pdf = []
        for (sf, fourn), vals in sorted(suivi.items()):
            tot = sum(vals)
            row_vals = [sf, fourn] + [f"{v:,.0f}" for v in vals] + [f"{tot:,.3f} DT"]
            tree.insert("", "end", values=row_vals)
            lignes_pdf.append(row_vals)
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        ctk.CTkButton(top, text="EXPORTER PDF", fg_color=C_GRIS,
                      command=lambda: calc.exporter_tableau_pdf("Suivi_Mensuel_Achats", cols, lignes_pdf, "PAYSAGE")).pack(side="right", padx=15)

    def ouvrir_graphiques_achats(self):
        self.nettoyer_ecran()
        _, recap_sf, _ = calc.calculer_recap_annuel_achats()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_achats).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text="GRAPHIQUES D'ANALYSE DES FOURNISSEURS PAR FAMILLE", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        familles = sorted(list(recap_sf.keys()))
        if not familles:
            ctk.CTkLabel(self.container, text="Aucune donnée d'achat disponible.", font=("Segoe UI", 14)).pack(pady=50)
            return

        chart_frame = ctk.CTkFrame(self.container, fg_color="white")
        chart_frame.pack(fill="both", expand=True, padx=20, pady=10)

        def tracer(sf_choisie):
            for w in chart_frame.winfo_children():
                w.destroy()
            data = recap_sf[sf_choisie]
            labels = list(data.keys())
            sizes = [max(0.0, v) for v in data.values()]
            fig, ax = plt.subplots(figsize=(7, 4.5), dpi=100)
            ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=140)
            ax.set_title(f"Répartition Fournisseurs — {sf_choisie}", fontweight="bold", color=C_BLEU_NUIT)
            canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)

        cb = ctk.CTkComboBox(top, values=familles, width=220, command=tracer)
        cb.set(val_idx(familles, 0))
        cb.pack(side="right", padx=20)
        tracer(val_idx(familles, 0))
    def ouvrir_rapport_executif(self):
        self.nettoyer_ecran()
        data = calc.generer_donnees_executives()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_achats).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text="RAPPORT EXÉCUTIF GLOBAL — TOP 5 STRATÉGIQUE", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        f1 = ctk.CTkFrame(self.container, fg_color="white")
        f1.pack(fill="x", padx=40, pady=10)
        ctk.CTkLabel(f1, text=f"TOP 5 — PLUS GRANDS FOURNISSEURS (ACHATS TOTAL HT : {data['total_achats']:,.3f} DT)",
                     fg_color="#184C78", text_color="white", font=("Segoe UI", 12, "bold"), height=32).pack(fill="x")
        for fourn, mnt in data["top5_fourn"]:
            part = (mnt / data["total_achats"] * 100) if data["total_achats"] > 0 else 0
            row = ctk.CTkFrame(f1, fg_color="transparent")
            row.pack(fill="x", padx=15, pady=3)
            ctk.CTkLabel(row, text=fourn, font=("Segoe UI", 11, "bold")).pack(side="left")
            ctk.CTkLabel(row, text=f"{mnt:,.3f} DT   ({part:.2f} %)", font=("Segoe UI", 11)).pack(side="right")

        f2 = ctk.CTkFrame(self.container, fg_color="white")
        f2.pack(fill="x", padx=40, pady=15)
        ctk.CTkLabel(f2, text=f"TOP 5 — SECTEURS (CONSO : {data['total_conso']:,.3f} DT | {data['total_nuitees']:,.0f} NUITÉES)",
                     fg_color=C_BRONZE, text_color="white", font=("Segoe UI", 12, "bold"), height=32).pack(fill="x")
        for dept, mnt in data["top5_dept"]:
            cout_n = (mnt / data["total_nuitees"]) if data["total_nuitees"] > 0 else 0
            row = ctk.CTkFrame(f2, fg_color="transparent")
            row.pack(fill="x", padx=15, pady=3)
            ctk.CTkLabel(row, text=dept, font=("Segoe UI", 11, "bold")).pack(side="left")
            ctk.CTkLabel(row, text=f"{mnt:,.3f} DT   |   Coût/Nuitée : {cout_n:,.3f} DT", font=("Segoe UI", 11)).pack(side="right")

    def afficher_accueil_consommation(self):
        self.nettoyer_ecran()

        header = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55, corner_radius=6)
        header.pack(fill="x", padx=40, pady=(20, 15))
        ctk.CTkLabel(header, text="SUIVI ET ANALYSE DE LA CONSOMMATION PAR DÉPARTEMENT",
                     font=("Segoe UI", 18, "bold"), text_color="white").pack(pady=12)

        grid = ctk.CTkFrame(self.container, fg_color="transparent")
        grid.pack(fill="both", expand=True, padx=40, pady=5)
        grid.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(grid, text="TRANSFERTS BRUTS", fg_color=C_JAUNE_HDR, text_color="white", font=("Segoe UI", 11, "bold"), height=30).grid(row=0, column=0, padx=8, pady=5, sticky="ew")
        ctk.CTkLabel(grid, text="INDICATEURS CONSO", fg_color=C_BLEU_NUIT, text_color="white", font=("Segoe UI", 11, "bold"), height=30).grid(row=0, column=1, padx=8, pady=5, sticky="ew")
        ctk.CTkLabel(grid, text="FILTRES DE NAVIGATION", fg_color=C_CYAN, text_color="white", font=("Segoe UI", 11, "bold"), height=30).grid(row=0, column=2, padx=8, pady=5, sticky="ew")
        ctk.CTkLabel(grid, text="SYNTHÈSES GLOBALES", fg_color="#5A6473", text_color="white", font=("Segoe UI", 11, "bold"), height=30).grid(row=0, column=3, padx=8, pady=5, sticky="ew")

        for i, mois in enumerate(db.LISTE_MOIS):
            ctk.CTkButton(grid, text=mois, fg_color="#FFE182", text_color="#3C3214", hover_color="#F5CE76",
                          font=("Segoe UI", 11, "bold"), height=28,
                          command=lambda m=mois: self.ouvrir_saisie_nuitees_transferts(m)).grid(row=i+1, column=0, padx=12, pady=2, sticky="ew")

            ctk.CTkButton(grid, text=mois, fg_color=C_BLEU_NUIT, text_color="white",
                          font=("Segoe UI", 11, "bold"), height=28,
                          command=lambda m=mois: self.ouvrir_rapport_conso_mois(m, None)).grid(row=i+1, column=1, padx=12, pady=2, sticky="ew")

        f_col3 = ctk.CTkFrame(grid, fg_color="white", corner_radius=8)
        f_col3.grid(row=1, column=2, rowspan=12, padx=8, pady=5, sticky="nsew")

        self.btn_sel_mois = ctk.CTkButton(f_col3, text=self.mois_conso_choisi, fg_color=C_CYAN, command=self.popup_mois)
        self.btn_sel_mois.pack(fill="x", padx=15, pady=(20, 8))

        self.btn_sel_dept = ctk.CTkButton(f_col3, text=self.dept_conso_choisi, fg_color="#2B3C50", command=self.popup_dept)
        self.btn_sel_dept.pack(fill="x", padx=15, pady=8)

        ctk.CTkButton(f_col3, text="VALIDER & VISUALISER", fg_color=C_VERT, font=("Segoe UI", 11, "bold"), height=38,
                      command=self.valider_dept_mois).pack(fill="x", padx=15, pady=(8, 25))

        self.btn_sel_sf = ctk.CTkButton(f_col3, text=self.sf_conso_choisie, fg_color="#285A82", command=self.popup_sf)
        self.btn_sel_sf.pack(fill="x", padx=15, pady=8)

        self.btn_sel_art = ctk.CTkButton(f_col3, text=self.art_conso_choisi, fg_color="#506E8C", command=self.popup_article)
        self.btn_sel_art.pack(fill="x", padx=15, pady=8)

        ctk.CTkButton(f_col3, text="CONSOMMATION PAR ARTICLE", fg_color=C_CYAN, font=("Segoe UI", 11, "bold"), height=38,
                      command=self.valider_analyse_article).pack(fill="x", padx=15, pady=8)

        btns_d = [
            ("CONSOMMATION ANNUELLE", C_VERT, self.ouvrir_conso_annuelle),
            ("BASE ARTICLES", C_GRIS, self.ouvrir_base_articles),
            ("BASE DÉPARTEMENTS", C_GRIS, self.ouvrir_base_depts),
            ("← ACCUEIL PRINCIPAL ACHATS", "#555F6E", self.afficher_accueil_achats)
        ]
        for idx, (txt, col, cmd) in enumerate(btns_d):
            ctk.CTkButton(grid, text=txt, fg_color=col, font=("Segoe UI", 11, "bold"), height=36,
                          command=cmd).grid(row=(idx*2)+1, column=3, rowspan=2, padx=12, pady=8, sticky="ew")
    def popup_mois(self):
        def set_m(val):
            self.mois_conso_choisi = val
            self.btn_sel_mois.configure(text=val)
        FrmSelectorPopup(self, "Choisir un Mois", db.LISTE_MOIS, set_m, avec_recherche=False)

    def popup_dept(self):
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT code_court, nom_long FROM departements ORDER BY nom_long")
        depts = {f"{nom} ({code})": code for code, nom in c.fetchall()}
        conn.close()
        def set_d(val):
            self.dept_conso_choisi = val
            self.dept_code_choisi = depts.get(val, "")
            self.btn_sel_dept.configure(text=val)
        FrmSelectorPopup(self, "Choisir un Département", list(depts.keys()), set_d, avec_recherche=True)

    def popup_sf(self):
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT DISTINCT sous_famille FROM articles WHERE sous_famille != '' ORDER BY sous_famille")
        sfs = [val_idx(r, 0) for r in c.fetchall()]
        conn.close()
        def set_sf(val):
            self.sf_conso_choisie = val
            self.art_conso_choisi = "Article..."
            self.btn_sel_sf.configure(text=val)
            self.btn_sel_art.configure(text="Article...")
        FrmSelectorPopup(self, "Choisir Sous-Famille", sfs, set_sf, avec_recherche=True)

    def popup_article(self):
        if self.sf_conso_choisie == "Sous-Famille...":
            messagebox.showwarning("Filtre Manquant", "Veuillez d'abord sélectionner une Sous-Famille.")
            return
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT DISTINCT designation FROM articles WHERE sous_famille = ? ORDER BY designation", (self.sf_conso_choisie,))
        arts = [val_idx(r, 0) for r in c.fetchall()]
        conn.close()
        def set_art(val):
            self.art_conso_choisi = val
            self.btn_sel_art.configure(text=val)
        FrmSelectorPopup(self, "Choisir Article", arts, set_art, avec_recherche=True)

    def valider_dept_mois(self):
        if self.mois_conso_choisi == "Choisir Mois..." or not self.dept_code_choisi:
            messagebox.showwarning("Sélection Incomplète", "Veuillez sélectionner un Mois ET un Département.")
            return
        self.ouvrir_rapport_conso_mois(self.mois_conso_choisi, self.dept_code_choisi)

    def ouvrir_saisie_nuitees_transferts(self, mois):
        self.nettoyer_ecran()
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT nb_nuitees FROM nuitees WHERE mois = ?", (mois,))
        res = c.fetchone()
        actuel = val_idx(res, 0) if res else 0.0
        conn.close()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_consommation).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text=f"TRANSFERTS & NUITÉES — {mois.upper()}", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        bar = ctk.CTkFrame(self.container, fg_color="white")
        bar.pack(fill="x", padx=20, pady=5)
        lbl_nuit = ctk.CTkLabel(bar, text=f"Nuitées du mois : {actuel:,.0f}", font=("Segoe UI", 13, "bold"), text_color=C_BLEU_NUIT)
        lbl_nuit.pack(side="left", padx=20, pady=10)

        def maj_nuitees():
            n_val = simpledialog.askfloat("Nuitées du mois", f"Nouveau nombre de nuitées pour {mois} :", parent=self)
            if n_val is not None:
                cn = db.get_connection()
                cn.cursor().execute("UPDATE nuitees SET nb_nuitees = ? WHERE mois = ?", (n_val, mois))
                cn.commit()
                cn.close()
                lbl_nuit.configure(text=f"Nuitées du mois : {n_val:,.0f}")

                ctk.CTkButton(bar, text="MODIFIER LES NUITÉES", fg_color=C_VERT, command=maj_nuitees).pack(side="left", padx=10)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Code", "Département", "Sous-Famille", "Article", "Qté Consommée", "Valeur Consommée")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT code_dept, nom_dept, sous_famille, designation, qte_conso, val_conso FROM transferts WHERE mois = ?", (mois,))
        for cd, d, sf, art, q, v in c.fetchall():
            tree.insert("", "end", values=(cd, d, sf, art, f"{q:,.3f}", f"{v:,.3f} DT"))
        conn.close()


    def ouvrir_rapport_conso_mois(self, mois, code_dept):
        self.nettoyer_ecran()
        conn = db.get_connection()
        c = conn.cursor()

        c.execute("SELECT nb_nuitees FROM nuitees WHERE mois = ?", (mois,))
        res_n = c.fetchone()
        nuitees = float(val_idx(res_n, 0)) if res_n and val_idx(res_n, 0) else 0.0

        if code_dept:
            nom_seul = self.dept_conso_choisi.split("(").pop(0).strip()
            c.execute("""
                SELECT COALESCE(nom_dept, code_dept), COALESCE(sous_famille, 'DIVERS'), designation,
                       SUM(qte_conso), SUM(val_conso)
                FROM transferts
                WHERE mois = ? AND (UPPER(code_dept) = UPPER(?) OR UPPER(nom_dept) LIKE UPPER(?))
                GROUP BY nom_dept, code_dept, sous_famille, designation
            """, (mois, code_dept, f"%{nom_seul}%"))
        else:
            c.execute("""
                SELECT COALESCE(nom_dept, code_dept), COALESCE(sous_famille, 'DIVERS'), designation,
                       SUM(qte_conso), SUM(val_conso)
                FROM transferts
                WHERE mois = ?
                GROUP BY nom_dept, code_dept, sous_famille, designation
            """, (mois,))

        rows = c.fetchall()
        conn.close()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_consommation).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text=f"RAPPORT COMPTABLE CONSOMMATIONS — {mois.upper()} | Nuitées : {nuitees:,.0f}",
                     font=("Segoe UI", 14, "bold"), text_color="white").pack(side="left", padx=20)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Désignation / Secteur", "Quantité Consommée", "Valeur Consommée", "Qté / Nuitée", "Valeur / Nuitée")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        tree.tag_configure("DEPT", background=C_CYAN, foreground="white", font=("Segoe UI", 10, "bold"))
        tree.tag_configure("SF", background="#2B3C50", foreground="white", font=("Segoe UI", 10, "bold"))
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        from collections import defaultdict
        arbre = defaultdict(lambda: defaultdict(list))
        for nom_d, sf, desig, qte, val in rows:
            arbre[str(nom_d or "SECTEUR")][str(sf or "DIVERS")].append((desig, float(qte or 0), float(val or 0)))

        lignes_pdf = []
        for dept, sfs in arbre.items():
            tot_d = sum(v for lst in sfs.values() for _, _, v in lst)
            rat_d = (tot_d / nuitees) if nuitees > 0 else 0.0
            tree.insert("", "end", values=(f"DÉPARTEMENT : {dept.upper()}", "", f"{tot_d:,.3f} DT", "", f"{rat_d:,.3f} DT"), tags=("DEPT",))
            lignes_pdf.append([f"DÉPARTEMENT : {dept.upper()}", "", f"{tot_d:,.3f} DT", "", f"{rat_d:,.3f} DT"])
            for sf, items in sfs.items():
                tree.insert("", "end", values=(f"Sous-Famille : {sf}", "", "", "", ""), tags=("SF",))
                lignes_pdf.append([f"Sous-Famille : {sf}", "", "", "", ""])
                for art, q, v in items:
                    qn = round(q / nuitees, 3) if nuitees > 0 else 0.0
                    vn = round(v / nuitees, 3) if nuitees > 0 else 0.0
                    r_vals = [str(art), f"{q:,.3f}", f"{v:,.3f} DT", f"{qn:,.3f}", f"{vn:,.3f} DT"]
                    tree.insert("", "end", values=r_vals)
                    lignes_pdf.append(r_vals)

        ctk.CTkButton(top, text="EXPORTER PDF", fg_color=C_GRIS,
                      command=lambda: calc.exporter_tableau_pdf(f"Consommations_{mois}", list(cols), lignes_pdf)).pack(side="right", padx=15)
    def valider_analyse_article(self):
        if self.art_conso_choisi == "Article...":
            messagebox.showwarning("Sélection Incomplète", "Veuillez sélectionner un Article.")
            return
        self.ouvrir_analyse_croisee_article()

    def ouvrir_analyse_croisee_article(self):
        self.nettoyer_ecran()
        cible = self.sf_conso_choisie if self.scope_bascule_conso == "SOUSFAM" else self.art_conso_choisi
        matrice = calc.calculer_analyse_article_ou_sf(cible, self.scope_bascule_conso, self.mode_bascule_conso)

        top = ctk.CTkFrame(self.container, fg_color=C_VERT if self.mode_bascule_conso == "VALEUR" else C_CYAN, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=90, command=self.afficher_accueil_consommation).pack(side="left", padx=10, pady=12)

        def basculer_val():
            self.mode_bascule_conso = "QUANTITE" if self.mode_bascule_conso == "VALEUR" else "VALEUR"
            self.ouvrir_analyse_croisee_article()

        def basculer_scope():
            self.scope_bascule_conso = "ARTICLE" if self.scope_bascule_conso == "SOUSFAM" else "SOUSFAM"
            self.ouvrir_analyse_croisee_article()

        ctk.CTkButton(top, text="VOIR QUANTITÉS" if self.mode_bascule_conso == "VALEUR" else "VOIR VALEURS (DT)",
                      fg_color=C_BLEU_NUIT, command=basculer_val).pack(side="left", padx=8)
        ctk.CTkButton(top, text="ARTICLE SEUL" if self.scope_bascule_conso == "SOUSFAM" else "TOUTE LA SOUS-FAMILLE",
                      fg_color="#DC7828", command=basculer_scope).pack(side="left", padx=8)

        ctk.CTkLabel(top, text=f"ANALYSE ({self.mode_bascule_conso}) : {cible.upper()}", font=("Segoe UI", 13, "bold"), text_color="white").pack(side="left", padx=15)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ["Département"] + [m[:4] for m in db.LISTE_MOIS]
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=75 if col != "Département" else 180)
        for dept, vals in sorted(matrice.items()):
            tree.insert("", "end", values=[dept] + [f"{v:,.3f}" for v in vals])
        tree.pack(fill="both", expand=True, padx=10, pady=10)

    def ouvrir_conso_annuelle(self):
        self.nettoyer_ecran()
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT SUM(nb_nuitees) FROM nuitees")
        res_n = c.fetchone()
        tot_nuit = val_idx(res_n, 0) if res_n and val_idx(res_n, 0) else 0.0
        c.execute("""
            SELECT COALESCE(d.nom_long, t.code_dept), COALESCE(a.sous_famille, 'DIVERS'), t.designation,
                   SUM(t.qte_conso), SUM(t.val_conso)
            FROM transferts t
            LEFT JOIN departements d ON t.code_dept = d.code_court
            LEFT JOIN articles a ON t.code_article = a.code
            GROUP BY d.nom_long, t.code_dept, a.sous_famille, t.designation
        """)
        rows = c.fetchall()
        conn.close()

        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_consommation).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text=f"CONSOMMATIONS ANNUELLES CUMULÉES | Total Nuitées : {tot_nuit:,.0f}", font=("Segoe UI", 14, "bold"), text_color="white").pack(side="left", padx=20)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Département", "Sous-Famille", "Article", "Qté Cumulée", "Valeur Cumulée", "Valeur / Nuitée")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        for dept, sf, art, qte, val in rows:
            rat = (val / tot_nuit) if tot_nuit > 0 else 0.0
            tree.insert("", "end", values=(dept, sf, art, f"{qte:,.3f}", f"{val:,.3f} DT", f"{rat:,.3f} DT"))
        tree.pack(fill="both", expand=True, padx=10, pady=10)

    def ouvrir_base_articles(self):
        self.nettoyer_ecran()
        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_consommation).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text="RÉFÉRENTIEL ARTICLES (DB_Données)", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Code Article", "Désignation", "Classe", "Sous-Famille")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT code, designation, classe, sous_famille FROM articles")
        for r in c.fetchall():
            tree.insert("", "end", values=r)
        conn.close()
        tree.pack(fill="both", expand=True, padx=10, pady=10)

    def ouvrir_base_depts(self):
        self.nettoyer_ecran()
        top = ctk.CTkFrame(self.container, fg_color=C_BLEU_NUIT, height=55)
        top.pack(fill="x", padx=20, pady=15)
        ctk.CTkButton(top, text="← ACCUEIL", fg_color=C_BLEU_ACIER, width=100, command=self.afficher_accueil_consommation).pack(side="left", padx=15, pady=12)
        ctk.CTkLabel(top, text="RÉFÉRENTIEL DÉPARTEMENTS (DB_Départements)", font=("Segoe UI", 15, "bold"), text_color="white").pack(side="left", padx=20)

        tree_frame = ctk.CTkFrame(self.container, fg_color="white")
        tree_frame.pack(fill="both", expand=True, padx=20, pady=10)
        cols = ("Code Court", "Nom Département")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT code_court, nom_long FROM departements")
        for r in c.fetchall():
            tree.insert("", "end", values=r)
        conn.close()
        tree.pack(fill="both", expand=True, padx=10, pady=10)


if __name__ == "__main__":
    app = HotelERPApp()
    app.mainloop()
