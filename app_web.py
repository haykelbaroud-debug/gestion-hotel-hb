import streamlit as st
import pandas as pd
import plotly.express as px
import tempfile
import os
from io import BytesIO
import database as db
import calculs as calc

st.set_page_config(
    page_title="Gestion Achats & Conso F&B",
    page_icon="🏨",
    layout="wide"
)

db.initialiser_base()

def generer_excel_memoire(titre, entetes, lignes):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rapport"
    ws.append([titre])
    ws["A1"].font = Font(size=14, bold=True, color="1F3A56")
    ws.append([])
    ws.append(entetes)
    fill_hdr = PatternFill(start_color="1F3A56", end_color="1F3A56", fill_type="solid")
    font_hdr = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(entetes) + 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.fill = fill_hdr
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center")
    for row in lignes:
        ws.append(row)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generer_pdf_memoire(titre, entetes, lignes, paysage=False):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    buf = BytesIO()
    page_format = landscape(A4) if paysage else A4
    doc = SimpleDocTemplate(buf, pagesize=page_format, rightMargin=20, leftMargin=20, topMargin=25, bottomMargin=20)
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{titre}</b>", styles["Title"]), Spacer(1, 10)]
    t = Table([entetes] + lignes, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A56")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DCE1")),
        ("FONTSIZE", (0, 1), (-1, -1), 7.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")])
    ]))
    elements.append(t)
    doc.build(elements)
    return buf.getvalue()
if "authentifie" not in st.session_state:
    st.session_state.authentifie = False

if not st.session_state.authentifie:
    st.title("🏨 Gestion des Achats & Consommations F&B")
    st.subheader("Accès Sécurisé")
    code = st.text_input("Entrez votre code d'accès :", type="password")
    if st.button("Se connecter", type="primary"):
        if code == "5227":
            st.session_state.authentifie = True
            st.rerun()
        else:
            st.error("Code secret incorrect.")
    st.stop()

st.sidebar.title("Navigation")
menu = st.sidebar.radio(
    "Choisir un module :",
    ["📊 Rapport Exécutif", "🛒 Achats & Stratégie", "🍽️ Consommations", "📥 Import Excel"]
)

if st.sidebar.button("🔒 Verrouiller"):
    st.session_state.authentifie = False
    st.rerun()

st.title(menu)

if menu == "📊 Rapport Exécutif":
    data = calc.generer_donnees_executives()
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Achats HT", f"{data['total_achats']:,.3f} DT")
    col2.metric("Total Consommation", f"{data['total_conso']:,.3f} DT")
    col3.metric("Total Nuitées", f"{data['total_nuitees']:,.0f}")

    st.divider()
    c_gauche, c_droite = st.columns(2)

    with c_gauche:
        st.subheader("🏆 Top 5 — Plus Grands Fournisseurs")
        if data["top5_fourn"]:
            lignes_f = []
            for fourn, mnt in data["top5_fourn"]:
                part = (mnt / data["total_achats"] * 100) if data["total_achats"] > 0 else 0
                lignes_f.append({"Fournisseur": fourn, "Montant HT (DT)": f"{mnt:,.3f}", "Part (%)": f"{part:.2f} %"})
            st.dataframe(pd.DataFrame(lignes_f), use_container_width=True, hide_index=True)

    with c_droite:
        st.subheader("🏢 Top 5 — Secteurs Consommateurs")
        if data["top5_dept"]:
            lignes_d = []
            for dept, mnt in data["top5_dept"]:
                cout_n = (mnt / data["total_nuitees"]) if data["total_nuitees"] > 0 else 0
                lignes_d.append({"Département": dept, "Conso (DT)": f"{mnt:,.3f}", "Coût / Nuitée (DT)": f"{cout_n:,.3f}"})
            st.dataframe(pd.DataFrame(lignes_d), use_container_width=True, hide_index=True)

elif menu == "🛒 Achats & Stratégie":
    tab1, tab2, tab3, tab4 = st.tabs([
        "🎯 Stratégie Mensuelle",
        "📅 Bilan Annuel",
        "📆 Suivi 12 Mois",
        "📈 Graphiques & Saisie"
    ])

    with tab1:
        mois_sel = st.selectbox("Sélectionner le mois :", db.LISTE_MOIS, key="strat_mois")
        res = calc.calculer_strategie_mois(mois_sel)

        k1, k2, k3 = st.columns(3)
        k1.metric("Total Achats", f"{res['total_achats']:,.3f} DT")
        k2.metric("Total Contrôle", f"{res['total_controle']:,.3f} DT")
        k3.metric("Écart", f"{res['ecart']:,.3f} DT")

        if res["lignes"]:
            df_strat = pd.DataFrame([
                {
                    "Fournisseur / Famille": l["col1"],
                    "Montant HT": f"{l['col2']:,.3f} DT",
                    "Part (%)": f"{l['col3']*100:.1f}%",
                    "Statut": l["col4"] if l["type"] == "ROW" else "TOTAL FAMILLE"
                }
                for l in res["lignes"]
            ])

            def colorer_lignes(row):
                if row["Statut"] == "TOTAL FAMILLE":
                    return ["background-color: #1F3A56; color: white; font-weight: bold"] * len(row)
                elif "Alerte" in row["Statut"]:
                    return ["color: #B42828; font-weight: bold"] * len(row)
                return [""] * len(row)

            st.dataframe(df_strat.style.apply(colorer_lignes, axis=1), use_container_width=True, hide_index=True, height=450)

            lignes_exp = [[l["col1"], f"{l['col2']:,.3f} DT", f"{l['col3']*100:.1f}%", l["col4"]] for l in res["lignes"]]
            entetes_exp = ["Fournisseur / Famille", "Montant HT", "Part %", "Statut"]
            c_b1, c_b2 = st.columns(2)
            c_b1.download_button("📄 Télécharger PDF", data=generer_pdf_memoire(f"Strategie_Achats_{mois_sel}", entetes_exp, lignes_exp),
                                 file_name=f"Strategie_Achats_{mois_sel}.pdf", mime="application/pdf")
            c_b2.download_button("📊 Télécharger Excel", data=generer_excel_memoire(f"Strategie_Achats_{mois_sel}", entetes_exp, lignes_exp),
                                 file_name=f"Strategie_Achats_{mois_sel}.xlsx")
        else:
            st.info(f"Aucun achat enregistré pour {mois_sel}.")
    with tab2:
        _, recap_sf, totaux_sf = calc.calculer_recap_annuel_achats()
        if recap_sf:
            lignes_bilan = []
            for sf in sorted(recap_sf.keys()):
                tot = totaux_sf[sf]
                lignes_bilan.append({"Fournisseur / Famille": f"▶ {sf} (Total : {tot:,.3f} DT)", "Montant Cumulé HT": "", "Part (%)": "", "Alerte": "FAMILLE"})
                for fourn, mnt in sorted(recap_sf[sf].items(), key=lambda x: x[1], reverse=True):
                    part = (mnt / tot) if tot > 0 else 0.0
                    lignes_bilan.append({
                        "Fournisseur / Famille": fourn,
                        "Montant Cumulé HT": f"{mnt:,.3f} DT",
                        "Part (%)": f"{part*100:.1f}%",
                        "Alerte": "⚠️ >30%" if part >= 0.30 else "OK"
                    })
            df_bilan = pd.DataFrame(lignes_bilan)
            st.dataframe(df_bilan, use_container_width=True, hide_index=True, height=500)
        else:
            st.info("Aucune donnée annuelle disponible.")

    with tab3:
        suivi, _, _ = calc.calculer_recap_annuel_achats()
        if suivi:
            cols_m = [m[:4] for m in db.LISTE_MOIS]
            lignes_suivi = []
            for (sf, fourn), vals in sorted(suivi.items()):
                tot = sum(vals)
                row_dict = {"Famille": sf, "Fournisseur": fourn}
                for idx, m_col in enumerate(cols_m):
                    row_dict[m_col] = round(vals[idx], 2)
                row_dict["Total (DT)"] = round(tot, 3)
                lignes_suivi.append(row_dict)
            df_suivi = pd.DataFrame(lignes_suivi)
            st.dataframe(df_suivi, use_container_width=True, hide_index=True, height=500)
        else:
            st.info("Aucune donnée mensuelle disponible.")

    with tab4:
        col_g, col_s = st.columns([1.2, 1])
        with col_g:
            st.subheader("Répartition par Famille")
            _, recap_sf, _ = calc.calculer_recap_annuel_achats()
            familles = sorted(list(recap_sf.keys()))
            if familles:
                sf_graph = st.selectbox("Choisir une famille :", familles)
                data_g = recap_sf[sf_graph]
                df_g = pd.DataFrame({"Fournisseur": list(data_g.keys()), "Montant": [max(0.0, v) for v in data_g.values()]})
                fig = px.pie(df_g, values="Montant", names="Fournisseur", title=f"Fournisseurs — {sf_graph}", hole=0.35)
                st.plotly_chart(fig, use_container_width=True)

        with col_s:
            st.subheader("Saisie Rapide d'un Achat")
            with st.form("form_ajout_achat", clear_on_submit=True):
                m_saisie = st.selectbox("Mois :", db.LISTE_MOIS)
                sf_saisie = st.selectbox("Sous-Famille :", sorted(list(db.FAMILLES_VALIDES)))
                fourn_saisie = st.text_input("Nom du Fournisseur :")
                mnt_saisie = st.number_input("Montant HT (DT) :", min_value=0.0, step=10.0, format="%.3f")
                if st.form_submit_button("➕ Enregistrer l'achat", type="primary"):
                    if fourn_saisie.strip() and mnt_saisie > 0:
                        conn = db.get_connection()
                        conn.cursor().execute(
                            "INSERT INTO achats (mois, code_article, sous_famille, fournisseur, montant_ht) VALUES (?, ?, ?, ?, ?)",
                            (m_saisie, sf_saisie, sf_saisie, fourn_saisie.strip().upper(), mnt_saisie)
                        )
                        conn.commit()
                        conn.close()
                        st.success("Achat ajouté !")
                        st.rerun()
                    else:
                        st.warning("Veuillez remplir le fournisseur et un montant supérieur à 0.")

elif menu == "📥 Import Excel":
    st.write("Chargez votre classeur Excel (`.xlsm` ou `.xlsx`) pour mettre à jour la base de données.")
    fichier = st.file_uploader("Sélectionnez votre fichier Excel :", type=["xlsm", "xlsx"])
    if fichier is not None:
        if st.button("🚀 Lancer l'importation", type="primary"):
            with st.spinner("Importation des données en cours..."):
                try:
                    suffixe = os.path.splitext(fichier.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffixe) as tmp:
                        tmp.write(fichier.getvalue())
                        chemin_tmp = tmp.name
                    nb_a, nb_c = db.importer_depuis_excel(chemin_tmp)
                    os.remove(chemin_tmp)
                    st.success(f"Importation terminée !\n\n- **{nb_a}** lignes d'Achats\n- **{nb_c}** lignes de Consommation")
                except Exception as e:
                    st.error(f"Erreur lors de la lecture du fichier : {e}")

elif menu == "🍽️ Consommations":
    ctab1, ctab2, ctab3, ctab4 = st.tabs([
        "📋 Rapport Mensuel & Nuitées",
        "🔍 Analyse Croisée (12 Mois)",
        "📅 Conso Annuelle",
        "📚 Référentiels"
    ])

    # --- ONGLET 1 : RAPPORT MENSUEL & NUITÉES ---
    with ctab1:
        col_f1, col_f2, col_f3 = st.columns([1, 1.5, 1])
        mois_c = col_f1.selectbox("Mois :", db.LISTE_MOIS, key="conso_mois")

        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT code_court, nom_long FROM departements ORDER BY nom_long")
        depts_rows = c.fetchall()
        c.execute("SELECT nb_nuitees FROM nuitees WHERE mois = ?", (mois_c,))
        res_n = c.fetchone()
        nuitees_actuelles = float(res_n[0]) if res_n and res_n[0] else 0.0
        conn.close()

        options_dept = ["TOUS LES DÉPARTEMENTS"] + [f"{nom} ({code})" for code, nom in depts_rows]
        dept_sel = col_f2.selectbox("Département :", options_dept)

        with col_f3:
            nouv_nuit = st.number_input("Nuitées du mois :", min_value=0.0, value=nuitees_actuelles, step=10.0, format="%.0f")
            if nouv_nuit != nuitees_actuelles:
                cn = db.get_connection()
                cn.cursor().execute("UPDATE nuitees SET nb_nuitees = ? WHERE mois = ?", (nouv_nuit, mois_c))
                cn.commit()
                cn.close()
                st.rerun()

        code_choisi = None if dept_sel == "TOUS LES DÉPARTEMENTS" else dept_sel.split("(")[-1].replace(")", "").strip()
        nuitees, arbre = calc.calculer_consommation_departement(mois_c, code_choisi)

        if arbre:
            lignes_c = []
            lignes_exp_c = []
            for dept, sfs in arbre.items():
                tot_d = sum(item["val"] for lst in sfs.values() for item in lst)
                rat_d = (tot_d / nuitees) if nuitees > 0 else 0.0
                lignes_c.append({"Désignation / Secteur": f"🏢 DÉPARTEMENT : {dept.upper()}", "Qté Conso": "", "Valeur (DT)": f"{tot_d:,.3f}", "Qté / Nuitée": "", "Val / Nuitée": f"{rat_d:,.3f}"})
                lignes_exp_c.append([f"DEPARTEMENT : {dept.upper()}", "", f"{tot_d:,.3f} DT", "", f"{rat_d:,.3f} DT"])
                for sf, items in sfs.items():
                    lignes_c.append({"Désignation / Secteur": f"📂 Sous-Famille : {sf}", "Qté Conso": "", "Valeur (DT)": "", "Qté / Nuitée": "", "Val / Nuitée": ""})
                    lignes_exp_c.append([f"Sous-Famille : {sf}", "", "", "", ""])
                    for it in items:
                        lignes_c.append({
                            "Désignation / Secteur": it["article"],
                            "Qté Conso": f"{it['qte']:,.3f}",
                            "Valeur (DT)": f"{it['val']:,.3f}",
                            "Qté / Nuitée": f"{it['qte_nuitee']:,.3f}",
                            "Val / Nuitée": f"{it['val_nuitee']:,.3f}"
                        })
                        lignes_exp_c.append([str(it["article"]), f"{it['qte']:,.3f}", f"{it['val']:,.3f} DT", f"{it['qte_nuitee']:,.3f}", f"{it['val_nuitee']:,.3f} DT"])

            st.dataframe(pd.DataFrame(lignes_c), use_container_width=True, hide_index=True, height=450)
            ent_c = ["Désignation / Secteur", "Quantité", "Valeur", "Qté / Nuitée", "Valeur / Nuitée"]
            b_c1, b_c2 = st.columns(2)
            b_c1.download_button("📄 Télécharger PDF Conso", data=generer_pdf_memoire(f"Consommations_{mois_c}", ent_c, lignes_exp_c),
                                 file_name=f"Consommations_{mois_c}.pdf", mime="application/pdf")
            b_c2.download_button("📊 Télécharger Excel Conso", data=generer_excel_memoire(f"Consommations_{mois_c}", ent_c, lignes_exp_c),
                                 file_name=f"Consommations_{mois_c}.xlsx")
        else:
            st.info(f"Aucune consommation enregistrée pour {mois_c}.")
    # --- ONGLET 2 : ANALYSE CROISÉE (12 MOIS) ---
    with ctab2:
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT DISTINCT sous_famille FROM articles WHERE sous_famille != '' ORDER BY sous_famille")
        sfs_list = [r[0] for r in c.fetchall()]
        
        c1, c2, c3, c4 = st.columns([1.2, 1.5, 1, 1])
        sf_cible = c1.selectbox("Sous-Famille :", sfs_list if sfs_list else ["Aucune"], key="croise_sf")
        
        c.execute("SELECT DISTINCT designation FROM articles WHERE sous_famille = ? ORDER BY designation", (sf_cible,))
        arts_list = [r[0] for r in c.fetchall()]
        conn.close()

        art_cible = c2.selectbox("Article :", arts_list if arts_list else ["Aucun"], key="croise_art")
        scope_sel = c3.radio("Périmètre :", ["ARTICLE", "SOUSFAM"], horizontal=True)
        val_sel = c4.radio("Unité :", ["QUANTITE", "VALEUR"], horizontal=True)

        cible_finale = sf_cible if scope_sel == "SOUSFAM" else art_cible
        if cible_finale and cible_finale != "Aucun":
            matrice = calc.calculer_analyse_article_ou_sf(cible_finale, scope_sel, val_sel)
            if matrice:
                cols_m = [m[:4] for m in db.LISTE_MOIS]
                lignes_mat = []
                for dept, vals in sorted(matrice.items()):
                    row_m = {"Département": dept}
                    for idx, m_col in enumerate(cols_m):
                        row_m[m_col] = round(vals[idx], 3)
                    row_m["Total"] = round(sum(vals), 3)
                    lignes_mat.append(row_m)
                st.dataframe(pd.DataFrame(lignes_mat), use_container_width=True, hide_index=True)
            else:
                st.info("Aucune consommation trouvée pour cette sélection.")

    # --- ONGLET 3 : CONSOMMATION ANNUELLE ---
    with ctab3:
        conn = db.get_connection()
        c = conn.cursor()
        c.execute("SELECT SUM(nb_nuitees) FROM nuitees")
        res_tn = c.fetchone()
        tot_nuit = float(res_tn[0]) if res_tn and res_tn[0] else 0.0
        c.execute("""
            SELECT COALESCE(d.nom_long, t.code_dept), COALESCE(a.sous_famille, t.sous_famille, 'DIVERS'), t.designation,
                   SUM(t.qte_conso), SUM(t.val_conso)
            FROM transferts t
            LEFT JOIN departements d ON t.code_dept = d.code_court
            LEFT JOIN articles a ON t.code_article = a.code
            GROUP BY d.nom_long, t.code_dept, a.sous_famille, t.sous_famille, t.designation
        """)
        rows_an = c.fetchall()
        conn.close()

        st.metric("Total Nuitées Annuelles", f"{tot_nuit:,.0f}")
        if rows_an:
            lignes_an = [
                {
                    "Département": r[0],
                    "Sous-Famille": r[1],
                    "Article": r[2],
                    "Qté Cumulée": round(r[3], 3),
                    "Valeur (DT)": round(r[4], 3),
                    "Coût / Nuitée (DT)": round(r[4] / tot_nuit, 3) if tot_nuit > 0 else 0.0
                }
                for r in rows_an
            ]
            st.dataframe(pd.DataFrame(lignes_an), use_container_width=True, hide_index=True, height=450)

    # --- ONGLET 4 : RÉFÉRENTIELS ---
    with ctab4:
        r1, r2 = st.columns(2)
        conn = db.get_connection()
        with r1:
            st.subheader("Base Départements")
            df_d = pd.read_sql_query("SELECT code_court AS Code, nom_long AS Département FROM departements", conn)
            st.dataframe(df_d, use_container_width=True, hide_index=True, height=350)
        with r2:
            st.subheader("Base Articles")
            df_a = pd.read_sql_query("SELECT code AS Code, designation AS Désignation, sous_famille AS Famille FROM articles", conn)
            st.dataframe(df_a, use_container_width=True, hide_index=True, height=350)
        conn.close()

