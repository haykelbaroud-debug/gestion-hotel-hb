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
db.initialiser_parametrage()
db.initialiser_referentiels_m2()
db.initialiser_achats_m3()

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
    ["📊 Rapport Exécutif", "🛒 Achats & Stratégie", "🍽️ Consommations", "⚙️ Paramétrage de base", "🗂️ Base de données", "🧾 Achats & Mouvements", "📥 Import Excel"]
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

elif menu == "⚙️ Paramétrage de base":
    from operator import itemgetter
    ptab1, ptab2, ptab3 = st.tabs(["📁 Classes", "📂 Familles", "🏷️ Sous-Familles"])

    with ptab1:
        col_form, col_liste = st.columns(2)
        classes = db.lister_classes()
        dict_classes = {itemgetter(0)(r): r for r in classes}

        with col_form:
            st.subheader("Fiche Classe")
            mode_cl = st.radio("Action :", ["➕ Ajouter", "✏️ Modifier / Statut / Supprimer"], key="mode_cl", horizontal=True)
            if mode_cl == "➕ Ajouter":
                c_code = st.text_input("Code Classe (ex: FB, GEN) :", key="add_cl_code")
                c_lib = st.text_input("Libellé Classe :", key="add_cl_lib")
                c_act = st.checkbox("Actif", value=True, key="add_cl_act")
                if st.button("Enregistrer la Classe", type="primary", key="btn_add_cl"):
                    if c_code.strip() and c_lib.strip():
                        ok, msg = db.sauvegarder_classe(c_code, c_lib, c_act, creation=True)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("Veuillez remplir le code et le libellé.")
            else:
                if dict_classes:
                    sel_cl = st.selectbox("Sélectionner une classe :", list(dict_classes.keys()),
                                          format_func=lambda x: f"{x} - {itemgetter(1)(dict_classes[x])}", key="sel_edit_cl")
                    _, lib_actuel, act_actuel = dict_classes[sel_cl]
                    new_lib = st.text_input("Libellé :", value=lib_actuel, key="edit_cl_lib")
                    new_act = st.checkbox("Actif", value=bool(act_actuel), key="edit_cl_act")
                    b1, b2 = st.columns(2)
                    if b1.button("💾 Mettre à jour", type="primary", key="btn_upd_cl"):
                        ok, msg = db.sauvegarder_classe(sel_cl, new_lib, new_act, creation=False)
                        st.success(msg)
                        st.rerun()
                    if b2.button("🗑️ Supprimer", key="btn_del_cl"):
                        ok, msg = db.supprimer_classe(sel_cl)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.info("Aucune classe enregistrée.")

        with col_liste:
            st.subheader("Liste des Classes")
            if classes:
                df_cl = pd.DataFrame([
                    {
                        "Code": itemgetter(0)(r),
                        "Libellé": itemgetter(1)(r),
                        "Statut": "✅ Actif" if itemgetter(2)(r) == 1 else "⛔ Désactivé"
                    }
                    for r in classes
                ])
                st.dataframe(df_cl, use_container_width=True, hide_index=True)
            else:
                st.info("Commencez par ajouter votre première classe à gauche.")

    with ptab2:
        col_form, col_liste = st.columns(2)
        classes_actives = db.lister_classes(seulement_actives=True)
        toutes_classes = db.lister_classes()
        familles = db.lister_familles()
        dict_fam = {itemgetter(0)(r): r for r in familles}

        with col_form:
            st.subheader("Fiche Famille")
            mode_fam = st.radio("Action :", ["➕ Ajouter", "✏️ Modifier / Statut / Supprimer"], key="mode_fam", horizontal=True)
            if mode_fam == "➕ Ajouter":
                if classes_actives:
                    map_cl_act = {itemgetter(0)(r): f"{itemgetter(0)(r)} - {itemgetter(1)(r)}" for r in classes_actives}
                    f_cl = st.selectbox("Classe parente :", list(map_cl_act.keys()), format_func=lambda x: map_cl_act[x], key="add_fam_cl")
                    f_code = st.text_input("Code Famille (ex: BOIS, VIAND) :", key="add_fam_code")
                    f_lib = st.text_input("Libellé Famille :", key="add_fam_lib")
                    f_act = st.checkbox("Actif", value=True, key="add_fam_act")
                    if st.button("Enregistrer la Famille", type="primary", key="btn_add_fam"):
                        if f_code.strip() and f_lib.strip():
                            ok, msg = db.sauvegarder_famille(f_code, f_lib, f_cl, f_act, creation=True)
                            if ok:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                        else:
                            st.warning("Veuillez remplir le code et le libellé.")
                else:
                    st.warning("Créez et activez d'abord au moins une Classe.")
            else:
                if dict_fam and toutes_classes:
                    sel_fam = st.selectbox("Sélectionner une famille :", list(dict_fam.keys()),
                                           format_func=lambda x: f"{x} - {itemgetter(1)(dict_fam[x])}", key="sel_edit_fam")
                    _, lib_f, cl_f, _, act_f = dict_fam[sel_fam]
                    map_cl_all = {itemgetter(0)(r): f"{itemgetter(0)(r)} - {itemgetter(1)(r)}" for r in toutes_classes}
                    idx_cl = list(map_cl_all.keys()).index(cl_f) if cl_f in map_cl_all else 0
                    new_cl_f = st.selectbox("Classe parente :", list(map_cl_all.keys()), index=idx_cl,
                                            format_func=lambda x: map_cl_all[x], key="edit_fam_cl")
                    new_lib_f = st.text_input("Libellé :", value=lib_f, key="edit_fam_lib")
                    new_act_f = st.checkbox("Actif", value=bool(act_f), key="edit_fam_act")
                    b1, b2 = st.columns(2)
                    if b1.button("💾 Mettre à jour", type="primary", key="btn_upd_fam"):
                        ok, msg = db.sauvegarder_famille(sel_fam, new_lib_f, new_cl_f, new_act_f, creation=False)
                        st.success(msg)
                        st.rerun()
                    if b2.button("🗑️ Supprimer", key="btn_del_fam"):
                        ok, msg = db.supprimer_famille(sel_fam)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.info("Aucune famille enregistrée.")

        with col_liste:
            st.subheader("Liste des Familles")
            if familles:
                df_fam = pd.DataFrame([
                    {
                        "Classe": f"{itemgetter(2)(r)} ({itemgetter(3)(r)})",
                        "Code Famille": itemgetter(0)(r),
                        "Libellé Famille": itemgetter(1)(r),
                        "Statut": "✅ Actif" if itemgetter(4)(r) == 1 else "⛔ Désactivé"
                    }
                    for r in familles
                ])
                st.dataframe(df_fam, use_container_width=True, hide_index=True)
            else:
                st.info("Aucune famille enregistrée.")
    with ptab3:
        col_form, col_liste = st.columns(2)
        fam_actives = db.lister_familles(seulement_actives=True)
        toutes_fam = db.lister_familles()
        sous_familles = db.lister_sous_familles()
        dict_sf = {itemgetter(0)(r): r for r in sous_familles}

        with col_form:
            st.subheader("Fiche Sous-Famille")
            mode_sf = st.radio("Action :", ["➕ Ajouter", "✏️ Modifier / Statut / Supprimer"], key="mode_sf", horizontal=True)
            if mode_sf == "➕ Ajouter":
                if fam_actives:
                    map_fam_act = {itemgetter(0)(r): f"[{itemgetter(2)(r)}] {itemgetter(0)(r)} - {itemgetter(1)(r)}" for r in fam_actives}
                    sf_fam = st.selectbox("Famille parente :", list(map_fam_act.keys()),
                                          format_func=lambda x: map_fam_act[x], key="add_sf_fam")
                    sf_code = st.text_input("Code Sous-Famille (ex: SOFT, MEAT) :", key="add_sf_code")
                    sf_lib = st.text_input("Libellé Sous-Famille :", key="add_sf_lib")
                    sf_act = st.checkbox("Actif", value=True, key="add_sf_act")
                    if st.button("Enregistrer la Sous-Famille", type="primary", key="btn_add_sf"):
                        if sf_code.strip() and sf_lib.strip():
                            ok, msg = db.sauvegarder_sous_famille(sf_code, sf_lib, sf_fam, sf_act, creation=True)
                            if ok:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                        else:
                            st.warning("Veuillez remplir le code et le libellé.")
                else:
                    st.warning("Créez et activez d'abord au moins une Famille.")
            else:
                if dict_sf and toutes_fam:
                    sel_sf = st.selectbox("Sélectionner une sous-famille :", list(dict_sf.keys()),
                                          format_func=lambda x: f"{x} - {itemgetter(1)(dict_sf[x])}", key="sel_edit_sf")
                    _, lib_sf, fam_sf, _, _, _, act_sf = dict_sf[sel_sf]
                    map_fam_all = {itemgetter(0)(r): f"[{itemgetter(2)(r)}] {itemgetter(0)(r)} - {itemgetter(1)(r)}" for r in toutes_fam}
                    idx_fam = list(map_fam_all.keys()).index(fam_sf) if fam_sf in map_fam_all else 0
                    new_fam_sf = st.selectbox("Famille parente :", list(map_fam_all.keys()), index=idx_fam,
                                              format_func=lambda x: map_fam_all[x], key="edit_sf_fam")
                    new_lib_sf = st.text_input("Libellé :", value=lib_sf, key="edit_sf_lib")
                    new_act_sf = st.checkbox("Actif", value=bool(act_sf), key="edit_sf_act")
                    b1, b2 = st.columns(2)
                    if b1.button("💾 Mettre à jour", type="primary", key="btn_upd_sf"):
                        ok, msg = db.sauvegarder_sous_famille(sel_sf, new_lib_sf, new_fam_sf, new_act_sf, creation=False)
                        st.success(msg)
                        st.rerun()
                    if b2.button("🗑️ Supprimer", key="btn_del_sf"):
                        ok, msg = db.supprimer_sous_famille(sel_sf)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.info("Aucune sous-famille enregistrée.")

        with col_liste:
            st.subheader("Liste des Sous-Familles")
            if sous_familles:
                df_sf = pd.DataFrame([
                    {
                        "Classe": f"{itemgetter(4)(r)} ({itemgetter(5)(r)})",
                        "Famille": f"{itemgetter(2)(r)} ({itemgetter(3)(r)})",
                        "Code Sous-Famille": itemgetter(0)(r),
                        "Libellé Sous-Famille": itemgetter(1)(r),
                        "Statut": "✅ Actif" if itemgetter(6)(r) == 1 else "⛔ Désactivé"
                    }
                    for r in sous_familles
                ])
                st.dataframe(df_sf, use_container_width=True, hide_index=True)
            else:
                st.info("Aucune sous-famille enregistrée.")
elif menu == "🗂️ Base de données":
    from operator import itemgetter
    btab1, btab2, btab3 = st.tabs(["📦 Base Articles", "🚚 Base Fournisseurs", "🏢 Base Départements"])

    # ---------------- ONGLET 1 : ARTICLES ----------------
    with btab1:
        col_form, col_liste = st.columns(2)
        articles = db.lister_articles_complet()
        dict_art = {itemgetter(0)(r): r for r in articles}
        sf_actives = db.lister_sous_familles(seulement_actives=True)
        toutes_sf = db.lister_sous_familles()

        with col_form:
            st.subheader("Fiche Article")
            mode_art = st.radio("Action :", ["➕ Ajouter", "✏️ Modifier / Statut / Supprimer"], key="mode_art", horizontal=True)
            unites_std = ["Kg", "Litre", "Pièce", "Carton", "Boîte", "Bouteille", "Sac", "Paquet", "Unité"]

            if mode_art == "➕ Ajouter":
                if sf_actives:
                    map_sf_act = {
                        itemgetter(0)(r): f"[{itemgetter(4)(r)} > {itemgetter(2)(r)}] {itemgetter(0)(r)} - {itemgetter(1)(r)}"
                        for r in sf_actives
                    }
                    a_sf = st.selectbox("Sous-famille (Classe > Famille > SF) :", list(map_sf_act.keys()),
                                        format_func=lambda x: map_sf_act[x], key="add_art_sf")
                    a_code = st.text_input("Code Article :", key="add_art_code")
                    a_des = st.text_input("Désignation :", key="add_art_des")
                    c_u, c_p, c_t = st.columns(3)
                    a_unite = c_u.selectbox("Unité :", unites_std, key="add_art_u")
                    a_prix = c_p.number_input("Prix Achat HT :", min_value=0.0, step=0.100, format="%.3f", key="add_art_p")
                    a_tva = c_t.selectbox("TVA (%) :", [19.0, 13.0, 7.0, 0.0], key="add_art_t")
                    a_act = st.checkbox("Actif", value=True, key="add_art_act")
                    if st.button("Enregistrer l'Article", type="primary", key="btn_add_art"):
                        if a_code.strip() and a_des.strip():
                            ok, msg = db.sauvegarder_article(a_code, a_des, a_sf, a_unite, a_prix, a_tva, a_act, creation=True)
                            if ok:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                        else:
                            st.warning("Veuillez remplir le code et la désignation.")
                else:
                    st.warning("Créez et activez d'abord au moins une Sous-famille dans 'Paramétrage de base'.")
            else:
                if dict_art:
                    sel_art = st.selectbox("Sélectionner un article :", list(dict_art.keys()),
                                           format_func=lambda x: f"{x} - {itemgetter(1)(dict_art[x])}", key="sel_edit_art")
                    _, des_a, sf_a, _, _, _, unite_a, prix_a, tva_a, act_a = dict_art[sel_art]
                    map_sf_all = {
                        itemgetter(0)(r): f"[{itemgetter(4)(r)} > {itemgetter(2)(r)}] {itemgetter(0)(r)} - {itemgetter(1)(r)}"
                        for r in toutes_sf
                    }
                    if sf_a and sf_a not in map_sf_all:
                        map_sf_all[sf_a] = f"{sf_a} (Import Excel)"
                    liste_keys_sf = list(map_sf_all.keys())
                    idx_sf = liste_keys_sf.index(sf_a) if sf_a in liste_keys_sf else 0
                    new_sf_a = st.selectbox("Sous-famille :", liste_keys_sf, index=idx_sf,
                                            format_func=lambda x: map_sf_all[x], key="edit_art_sf") if liste_keys_sf else sf_a
                    new_des_a = st.text_input("Désignation :", value=des_a, key="edit_art_des")
                    c_u, c_p, c_t = st.columns(3)
                    idx_u = unites_std.index(unite_a) if unite_a in unites_std else 0
                    new_u_a = c_u.selectbox("Unité :", unites_std, index=idx_u, key="edit_art_u")
                    new_p_a = c_p.number_input("Prix Achat HT :", value=float(prix_a), min_value=0.0, step=0.100, format="%.3f", key="edit_art_p")
                    tvas = [19.0, 13.0, 7.0, 0.0]
                    idx_t = tvas.index(float(tva_a)) if float(tva_a) in tvas else 0
                    new_t_a = c_t.selectbox("TVA (%) :", tvas, index=idx_t, key="edit_art_t")
                    new_act_a = st.checkbox("Actif", value=bool(act_a), key="edit_art_act")
                    b1, b2 = st.columns(2)
                    if b1.button("💾 Mettre à jour", type="primary", key="btn_upd_art"):
                        ok, msg = db.sauvegarder_article(sel_art, new_des_a, new_sf_a, new_u_a, new_p_a, new_t_a, new_act_a, creation=False)
                        st.success(msg)
                        st.rerun()
                    if b2.button("🗑️ Supprimer", key="btn_del_art"):
                        ok, msg = db.supprimer_article(sel_art)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.info("Aucun article enregistré.")

        with col_liste:
            st.subheader("Liste des Articles")
            filtre_art = st.text_input("🔍 Rechercher (code ou désignation) :", key="rech_art")
            if articles:
                lignes_art = []
                for r in articles:
                    if filtre_art.strip() == "" or filtre_art.lower() in itemgetter(0)(r).lower() or filtre_art.lower() in itemgetter(1)(r).lower():
                        lignes_art.append({
                            "Code": itemgetter(0)(r),
                            "Désignation": itemgetter(1)(r),
                            "Sous-Famille": itemgetter(2)(r),
                            "Unité": itemgetter(6)(r),
                            "Prix HT": f"{itemgetter(7)(r):,.3f}",
                            "TVA": f"{itemgetter(8)(r):.0f}%",
                            "Statut": "✅ Actif" if itemgetter(9)(r) == 1 else "⛔ Désactivé"
                        })
                st.dataframe(pd.DataFrame(lignes_art), use_container_width=True, hide_index=True, height=450)
            else:
                st.info("Aucun article enregistré.")
    # ---------------- ONGLET 2 : FOURNISSEURS ----------------
    with btab2:
        col_form, col_liste = st.columns(2)
        fournisseurs = db.lister_fournisseurs()
        dict_fourn = {itemgetter(0)(r): r for r in fournisseurs}

        with col_form:
            st.subheader("Fiche Fournisseur")
            mode_fourn = st.radio("Action :", ["➕ Ajouter", "✏️ Modifier / Statut / Supprimer"], key="mode_fourn", horizontal=True)
            if mode_fourn == "➕ Ajouter":
                fn_code = st.text_input("Code Fournisseur :", key="add_fn_code")
                fn_nom = st.text_input("Nom / Raison sociale :", key="add_fn_nom")
                fn_tel = st.text_input("Téléphone :", key="add_fn_tel")
                fn_mf = st.text_input("Matricule Fiscal :", key="add_fn_mf")
                fn_mail = st.text_input("Email :", key="add_fn_mail")
                fn_act = st.checkbox("Actif", value=True, key="add_fn_act")
                if st.button("Enregistrer le Fournisseur", type="primary", key="btn_add_fn"):
                    if fn_code.strip() and fn_nom.strip():
                        ok, msg = db.sauvegarder_fournisseur(fn_code, fn_nom, fn_tel, fn_mf, fn_mail, fn_act, creation=True)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("Veuillez remplir le code et le nom.")
            else:
                if dict_fourn:
                    sel_fn = st.selectbox("Sélectionner un fournisseur :", list(dict_fourn.keys()),
                                          format_func=lambda x: f"{x} - {itemgetter(1)(dict_fourn[x])}", key="sel_edit_fn")
                    _, nom_fn, tel_fn, mf_fn, mail_fn, act_fn = dict_fourn[sel_fn]
                    new_nom_fn = st.text_input("Nom / Raison sociale :", value=nom_fn, key="edit_fn_nom")
                    new_tel_fn = st.text_input("Téléphone :", value=tel_fn, key="edit_fn_tel")
                    new_mf_fn = st.text_input("Matricule Fiscal :", value=mf_fn, key="edit_fn_mf")
                    new_mail_fn = st.text_input("Email :", value=mail_fn, key="edit_fn_mail")
                    new_act_fn = st.checkbox("Actif", value=bool(act_fn), key="edit_fn_act")
                    b1, b2 = st.columns(2)
                    if b1.button("💾 Mettre à jour", type="primary", key="btn_upd_fn"):
                        ok, msg = db.sauvegarder_fournisseur(sel_fn, new_nom_fn, new_tel_fn, new_mf_fn, new_mail_fn, new_act_fn, creation=False)
                        st.success(msg)
                        st.rerun()
                    if b2.button("🗑️ Supprimer", key="btn_del_fn"):
                        ok, msg = db.supprimer_fournisseur(sel_fn)
                        st.success(msg)
                        st.rerun()
                else:
                    st.info("Aucun fournisseur enregistré.")

        with col_liste:
            st.subheader("Liste des Fournisseurs")
            if fournisseurs:
                df_fn = pd.DataFrame([
                    {
                        "Code": itemgetter(0)(r),
                        "Nom": itemgetter(1)(r),
                        "Téléphone": itemgetter(2)(r),
                        "Matricule Fiscal": itemgetter(3)(r),
                        "Statut": "✅ Actif" if itemgetter(5)(r) == 1 else "⛔ Désactivé"
                    }
                    for r in fournisseurs
                ])
                st.dataframe(df_fn, use_container_width=True, hide_index=True)
            else:
                st.info("Aucun fournisseur enregistré.")

    # ---------------- ONGLET 3 : DÉPARTEMENTS ----------------
    with btab3:
        col_form, col_liste = st.columns(2)
        depts = db.lister_departements_complet()
        dict_dept = {itemgetter(0)(r): r for r in depts}

        with col_form:
            st.subheader("Fiche Département")
            mode_dept = st.radio("Action :", ["➕ Ajouter", "✏️ Modifier / Statut / Supprimer"], key="mode_dept", horizontal=True)
            if mode_dept == "➕ Ajouter":
                dp_code = st.text_input("Code Court (ex: CUIS, BAR) :", key="add_dp_code")
                dp_nom = st.text_input("Nom Complet du Département :", key="add_dp_nom")
                dp_act = st.checkbox("Actif", value=True, key="add_dp_act")
                if st.button("Enregistrer le Département", type="primary", key="btn_add_dp"):
                    if dp_code.strip() and dp_nom.strip():
                        ok, msg = db.sauvegarder_departement(dp_code, dp_nom, dp_act, creation=True)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("Veuillez remplir le code court et le nom.")
            else:
                if dict_dept:
                    sel_dp = st.selectbox("Sélectionner un département :", list(dict_dept.keys()),
                                          format_func=lambda x: f"{x} - {itemgetter(1)(dict_dept[x])}", key="sel_edit_dp")
                    _, nom_dp, act_dp = dict_dept[sel_dp]
                    new_nom_dp = st.text_input("Nom Complet :", value=nom_dp, key="edit_dp_nom")
                    new_act_dp = st.checkbox("Actif", value=bool(act_dp), key="edit_dp_act")
                    b1, b2 = st.columns(2)
                    if b1.button("💾 Mettre à jour", type="primary", key="btn_upd_dp"):
                        ok, msg = db.sauvegarder_departement(sel_dp, new_nom_dp, new_act_dp, creation=False)
                        st.success(msg)
                        st.rerun()
                    if b2.button("🗑️ Supprimer", key="btn_del_dp"):
                        ok, msg = db.supprimer_departement(sel_dp)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.info("Aucun département enregistré.")

        with col_liste:
            st.subheader("Liste des Départements")
            if depts:
                df_dp = pd.DataFrame([
                    {
                        "Code": itemgetter(0)(r),
                        "Nom Département": itemgetter(1)(r),
                        "Statut": "✅ Actif" if itemgetter(2)(r) == 1 else "⛔ Désactivé"
                    }
                    for r in depts
                ])
                st.dataframe(df_dp, use_container_width=True, hide_index=True)
            else:
                st.info("Aucun département enregistré.")
elif menu == "🧾 Achats & Mouvements":
    from operator import itemgetter
    from datetime import date

    atab1, atab2, atab3, atab4, atab5, atab6 = st.tabs([
        "📝 Bon de Commande",
        "📦 Bon de Réception",
        "↩️ Retour Fournisseur",
        "➕ Entrée Directe",
        "➖ Sortie Directe",
        "🔐 Autorisations Dépôts"
    ])

    def afficher_ecran_document(type_doc, prefixe, titre_doc, besoin_fournisseur=True):
        articles_actifs = db.lister_articles_complet(seulement_actifs=True)
        fournisseurs_actifs = db.lister_fournisseurs(seulement_actifs=True)
        depts_autorises = db.lister_depts_achats()

        mode_doc = st.radio(
            "Opération :",
            ["➕ Nouveau Bon", "📋 Consulter / Modifier / Supprimer"],
            key=f"mode_{type_doc}",
            horizontal=True
        )

        cle_panier = f"panier_{type_doc}"
        if cle_panier not in st.session_state:
            st.session_state[cle_panier] = []

        if mode_doc == "➕ Nouveau Bon":
            if not articles_actifs:
                st.warning("Créez d'abord au moins un article dans 'Base de données > Base Articles'.")
                return
            if besoin_fournisseur and not fournisseurs_actifs:
                st.warning("Créez d'abord au moins un fournisseur dans 'Base de données > Base Fournisseurs'.")
                return

            num_auto = db.generer_numero_doc(prefixe, str(date.today().year))
            c1, c2, c3, c4 = st.columns(4)
            c1.text_input("N° Document (Auto) :", value=num_auto, disabled=True, key=f"num_{type_doc}")
            date_saisie = c2.date_input("Date :", value=date.today(), key=f"date_{type_doc}")

            map_depts = {itemgetter(0)(d): f"{itemgetter(0)(d)} - {itemgetter(1)(d)}" for d in depts_autorises}
            dept_sel = c3.selectbox("Dépôt / Département :", list(map_depts.keys()),
                                    format_func=lambda x: map_depts[x], key=f"dept_{type_doc}")

            if besoin_fournisseur:
                map_fourn = {itemgetter(0)(f): f"{itemgetter(0)(f)} - {itemgetter(1)(f)}" for f in fournisseurs_actifs}
                fourn_sel = c4.selectbox("Fournisseur :", list(map_fourn.keys()),
                                         format_func=lambda x: map_fourn[x], key=f"fourn_{type_doc}")
            else:
                fourn_sel = ""
                c4.text_input("Type Mouvement :", value=titre_doc, disabled=True, key=f"info_{type_doc}")

            c_ref, c_obs = st.columns(2)
            ref_ext = c_ref.text_input("N° BL / Facture / Référence :", key=f"ref_{type_doc}")
            obs_doc = c_obs.text_input("Observation / Motif :", key=f"obs_{type_doc}")

            st.markdown("#### Ajout d'articles au bon")
            map_art = {
                itemgetter(0)(a): f"{itemgetter(0)(a)} - {itemgetter(1)(a)} ({itemgetter(6)(a)})"
                for a in articles_actifs
            }
            dict_art_info = {itemgetter(0)(a): a for a in articles_actifs}

            ca1, ca2, ca3, ca4 = st.columns([3, 1, 1, 1])
            art_choisi = ca1.selectbox("Article :", list(map_art.keys()), format_func=lambda x: map_art[x], key=f"art_sel_{type_doc}")
            info_a = dict_art_info[art_choisi]
            qte_saisie = ca2.number_input("Quantité :", min_value=0.001, value=1.0, step=1.0, format="%.3f", key=f"qte_{type_doc}")
            pu_saisi = ca3.number_input("Prix Unit. HT :", min_value=0.0, value=float(itemgetter(7)(info_a)), step=0.100, format="%.3f", key=f"pu_{type_doc}")

            if ca4.button("➕ Ajouter ligne", key=f"add_line_{type_doc}"):
                st.session_state[cle_panier].append((
                    art_choisi,
                    itemgetter(1)(info_a),
                    itemgetter(6)(info_a),
                    qte_saisie,
                    pu_saisi,
                    float(itemgetter(8)(info_a))
                ))
                st.rerun()

            if st.session_state[cle_panier]:
                lignes_aff = []
                tot_ht, tot_ttc = 0.0, 0.0
                for idx, (ca, da, ua, qa, pa, ta) in enumerate(st.session_state[cle_panier]):
                    mht = round(qa * pa, 3)
                    mttc = round(mht * (1.0 + ta / 100.0), 3)
                    tot_ht += mht
                    tot_ttc += mttc
                    lignes_aff.append({
                        "Ligne": idx + 1,
                        "Code": ca,
                        "Désignation": da,
                        "Unité": ua,
                        "Qté": f"{qa:,.3f}",
                        "PU HT": f"{pa:,.3f}",
                        "TVA": f"{ta:.0f}%",
                        "Total HT": f"{mht:,.3f}",
                        "Total TTC": f"{mttc:,.3f}"
                    })
                st.dataframe(pd.DataFrame(lignes_aff), use_container_width=True, hide_index=True)

                k1, k2, k3 = st.columns(3)
                k1.metric("Total HT", f"{tot_ht:,.3f} DT")
                k2.metric("Total TTC", f"{tot_ttc:,.3f} DT")
                if k3.button("🧹 Vider les lignes", key=f"clear_{type_doc}"):
                    st.session_state[cle_panier] = []
                    st.rerun()

                if st.button(f"💾 Enregistrer {titre_doc} ({num_auto})", type="primary", key=f"save_{type_doc}"):
                    ok, msg = db.enregistrer_document_achat(
                        num_auto, type_doc, date_saisie, fourn_sel, dept_sel, ref_ext, obs_doc, st.session_state[cle_panier], creation=True
                    )
                    if ok:
                        st.session_state[cle_panier] = []
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
        else:
            docs = db.lister_documents_achat(type_doc)
            if not docs:
                st.info(f"Aucun document ({titre_doc}) enregistré pour le moment.")
                return

            df_docs = pd.DataFrame([
                {
                    "Numéro": itemgetter(0)(d),
                    "Date": itemgetter(1)(d),
                    "Fournisseur": itemgetter(2)(d) if besoin_fournisseur else "-",
                    "Dépôt": itemgetter(3)(d),
                    "Réf Externe": itemgetter(4)(d),
                    "Total HT": f"{itemgetter(5)(d):,.3f}",
                    "Total TTC": f"{itemgetter(6)(d):,.3f}",
                    "Observation": itemgetter(7)(d)
                }
                for d in docs
            ])
            st.dataframe(df_docs, use_container_width=True, hide_index=True)

            st.divider()
            dict_docs = {itemgetter(0)(d): d for d in docs}
            sel_num = st.selectbox("Sélectionner un document pour détail / modification / suppression :", list(dict_docs.keys()), key=f"sel_doc_{type_doc}")
            doc_info = dict_docs[sel_num]
            lignes_doc = db.lire_lignes_document(sel_num)

            st.markdown(f"**Détail du bon `{sel_num}` :**")
            st.dataframe(pd.DataFrame([
                {
                    "Code": itemgetter(0)(l),
                    "Désignation": itemgetter(1)(l),
                    "Unité": itemgetter(2)(l),
                    "Qté": f"{itemgetter(3)(l):,.3f}",
                    "PU HT": f"{itemgetter(4)(l):,.3f}",
                    "TVA": f"{itemgetter(5)(l):.0f}%",
                    "Montant HT": f"{itemgetter(6)(l):,.3f}",
                    "Montant TTC": f"{itemgetter(7)(l):,.3f}"
                }
                for l in lignes_doc
            ]), use_container_width=True, hide_index=True)

            with st.expander("✏️ Modifier l'entête ou les quantités de ce bon"):
                new_ref = st.text_input("Référence externe :", value=itemgetter(4)(doc_info), key=f"m_ref_{type_doc}")
                new_obs = st.text_input("Observation :", value=itemgetter(7)(doc_info), key=f"m_obs_{type_doc}")
                nouvelles_lignes = []
                for idx_l, l in enumerate(lignes_doc):
                    mc1, mc2, mc3 = st.columns([3, 1, 1])
                    mc1.text(f"{itemgetter(0)(l)} - {itemgetter(1)(l)}")
                    nq = mc2.number_input("Qté", value=float(itemgetter(3)(l)), min_value=0.0, step=1.0, format="%.3f", key=f"mq_{type_doc}_{idx_l}")
                    np = mc3.number_input("PU HT", value=float(itemgetter(4)(l)), min_value=0.0, step=0.100, format="%.3f", key=f"mp_{type_doc}_{idx_l}")
                    if nq > 0:
                        nouvelles_lignes.append((itemgetter(0)(l), itemgetter(1)(l), itemgetter(2)(l), nq, np, float(itemgetter(5)(l))))

                if st.button("💾 Enregistrer les modifications", type="primary", key=f"btn_mod_{type_doc}"):
                    ok, msg = db.enregistrer_document_achat(
                        sel_num, type_doc, itemgetter(1)(doc_info), itemgetter(8)(doc_info),
                        itemgetter(3)(doc_info), new_ref, new_obs, nouvelles_lignes, creation=False
                    )
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

            if st.button(f"🗑️ Supprimer définitivement le bon {sel_num}", key=f"del_doc_{type_doc}"):
                ok, msg = db.supprimer_document_achat(sel_num)
                st.success(msg)
                st.rerun()

    with atab1:
        afficher_ecran_document("BC", "BC", "Bon de Commande", besoin_fournisseur=True)
    with atab2:
        afficher_ecran_document("BR", "BR", "Bon de Réception", besoin_fournisseur=True)
    with atab3:
        afficher_ecran_document("RF", "RF", "Retour Fournisseur", besoin_fournisseur=True)
    with atab4:
        afficher_ecran_document("ED", "ED", "Entrée Directe", besoin_fournisseur=False)
    with atab5:
        afficher_ecran_document("SD", "SD", "Sortie Directe", besoin_fournisseur=False)
    with atab6:
        st.subheader("🔐 Départements autorisés à recevoir des Achats")
        st.info("Par défaut, seul l'Économat Central (ECO) reçoit les achats. Vous pouvez activer ou désactiver l'autorisation pour d'autres départements ci-dessous.")
        conn_tmp = db.get_connection()
        c_tmp = conn_tmp.cursor()
        c_tmp.execute("SELECT code_court, nom_long, COALESCE(autorise_achat, 0) FROM departements WHERE COALESCE(actif, 1) = 1 ORDER BY code_court")
        liste_dp_auth = c_tmp.fetchall()
        conn_tmp.close()

        for dp_code, dp_nom, dp_auth in liste_dp_auth:
            ca, cb, cc = st.columns([2, 3, 2])
            ca.write(f"**{dp_code}**")
            cb.write(dp_nom)
            est_auto = bool(dp_auth)
            if dp_code == "ECO":
                cc.success("✅ Dépôt principal (Actif)")
            else:
                if est_auto:
                    if cc.button("⛔ Désactiver", key=f"dis_auth_{dp_code}"):
                        db.basculer_autorisation_achat_dept(dp_code, 0)
                        st.rerun()
                else:
                    if cc.button("✅ Activer achats", key=f"ena_auth_{dp_code}"):
                        db.basculer_autorisation_achat_dept(dp_code, 1)
                        st.rerun()
