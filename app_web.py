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

@st.cache_resource
def _demarrer_base_une_seule_fois():
    db.initialiser_base()
    db.initialiser_achats_m3()
    db.initialiser_transferts_m4()
    db.initialiser_stock_inventaire_m5()
    db.initialiser_nuitees_m7()
    db.initialiser_autorisations_m8()
    db.creer_index_performance()
    return True

_demarrer_base_une_seule_fois()

@st.cache_data(ttl=300)
def _cache_articles_complets():
    return db.lister_articles_complet(seulement_actifs=True)

@st.cache_data(ttl=300)
def _cache_fournisseurs():
    return db.lister_fournisseurs(seulement_actifs=True)

@st.cache_data(ttl=300)
def _cache_depts_achats():
    return db.lister_depts_achats()

@st.cache_data(ttl=300)
def _cache_depts_actifs():
    return db.lister_departements_complet(seulement_actifs=True)
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
    ["📊 Rapport Exécutif", "🛒 Achats & Stratégie", "🍽️ Consommations", "⚙️ Paramétrage de base", "🗂️ Base de données", "🧾 Achats & Mouvements", "🔄 Transferts", "📦 Stock", "📋 Inventaire", "🛏️ Nuitées", "📈 Gestion", "💰 Centre des Coûts", "📥 Import Excel", "💾 Sauvegarde"]
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
            if st.button("🧹 Fusionner les classes en double (ex: 02 et 22)", key="btn_merge_cl"):
                    ok_m, msg_m = db.fusionner_doublons_classes()
                    if ok_m:
                        st.success(msg_m)
                        st.rerun()
                    else:
                        st.error(msg_m)
            if st.button("🔄 Reconstruire la Nomenclature Officielle Complète (14 Classes / 37 Familles / 95 Sous-Familles)", type="primary", key="btn_rebuild_nom"):
                    ok_r, msg_r = db.reconstruire_nomenclature_officielle()
                    if ok_r:
                        st.success(msg_r)
                        st.rerun()
                    else:
                        st.error(msg_r)
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
            st.divider()
    with st.expander("🧹 Outil de remise au propre de la Nomenclature (Classes 01..14 & Familles)", expanded=False):
        st.info("Fusionne les classes doublons (ex: 22 -> 02 ALIMENTATION, 33 -> 03, 44 -> 04) et remplace les libellés 'BIENS' des familles par leurs vrais noms sans toucher à vos 9 757 articles.")
        if st.button("🔧 Nettoyer et unifier la nomenclature maintenant", key="btn_clean_nom"):
            ok_n, msg_n = db.nettoyer_et_restaurer_nomenclature()
            if ok_n:
                st.cache_data.clear()
                st.success(msg_n)
                st.rerun()
            else:
                st.error(msg_n)

       st.divider()
    st.subheader("🔐 Autorisations par Classe, Famille, Sous-Famille & Article")
    st.info("Sélectionnez les Classes, Familles ou Sous-Familles AUTORISÉES pour ce Fournisseur ou Département (vide = tout autorisé), puis décochez les articles précis à exclure.")

    type_ent_sel = st.radio(
        "Type d'entité à paramétrer :",
        ["Fournisseur", "Département"],
        horizontal=True,
        key="rad_type_ent_m8"
    )

    if type_ent_sel == "Fournisseur":
        tous_fourn = _cache_fournisseurs()
        map_ent = {itemgetter(0)(f): f"{itemgetter(0)(f)} - {itemgetter(1)(f)}" for f in tous_fourn}
        type_ent_code = "FOURNISSEUR"
    else:
        tous_depts = _cache_depts_actifs()
        map_ent = {itemgetter(0)(d): f"{itemgetter(0)(d)} - {itemgetter(1)(d)}" for d in tous_depts}
        type_ent_code = "DEPARTEMENT"

    if map_ent:
        code_ent_sel = st.selectbox(
            f"Sélectionner le {type_ent_sel} :",
            list(map_ent.keys()),
            format_func=lambda x: map_ent[x],
            key="sel_ent_m8"
        )

        cl_ok_act, fam_ok_act, sf_ok_act, art_exc_act = db.lire_config_autorisations_entite(type_ent_code, code_ent_sel)
        tous_art_m8 = _cache_articles_complets()

        # Lecture des tables officielles (Classes, Familles, Sous-Familles)
        rows_cl = db.lister_classes(seulement_actives=True)
        rows_fam = db.lister_familles(seulement_actives=True)
        rows_sf = db.lister_sous_familles(seulement_actives=True)

        nom_cl = {str(itemgetter(0)(r)).strip().upper(): str(itemgetter(1)(r)).strip() for r in rows_cl}
        nom_fam = {str(itemgetter(0)(r)).strip().upper(): str(itemgetter(1)(r)).strip() for r in rows_fam}
        nom_sf = {str(itemgetter(0)(r)).strip().upper(): str(itemgetter(1)(r)).strip() for r in rows_sf}

        cl_to_fam = {}
        for r in rows_fam:
            f_c = str(itemgetter(0)(r)).strip().upper()
            c_c = str(itemgetter(2)(r)).strip().upper()
            if c_c and f_c:
                cl_to_fam.setdefault(c_c, set()).add(f_c)

        fam_to_sf = {}
        for r in rows_sf:
            s_c = str(itemgetter(0)(r)).strip().upper()
            f_c = str(itemgetter(2)(r)).strip().upper()
            if f_c and s_c:
                fam_to_sf.setdefault(f_c, set()).add(s_c)

        # Indexation des articles selon leur structure réelle :
        # (0:code, 1:designation, 2:sous_famille, 3:sf_lib, 4:famille, 5:classe)
        art_par_cl = {}
        art_par_fam = {}
        art_par_sf = {}
        for a in tous_art_m8:
            sf_u = str(itemgetter(2)(a) or "").strip().upper()
            fam_u = str(itemgetter(4)(a) or "").strip().upper()
            cl_u = str(itemgetter(5)(a) or "").strip().upper()
            if cl_u:
                art_par_cl.setdefault(cl_u, []).append(a)
            if fam_u:
                art_par_fam.setdefault(fam_u, []).append(a)
            if sf_u:
                art_par_sf.setdefault(sf_u, []).append(a)

        liste_cl_sorted = sorted(nom_cl.keys())
        map_cl = {c: f"{c} - {nom_cl[c]}" for c in liste_cl_sorted}

        # 1. Classes AUTORISÉES
        def_cl = [c for c in cl_ok_act if c in map_cl]
        sel_classes = st.multiselect(
            "1️⃣ Classes AUTORISÉES (vide = toutes autorisées) :",
            liste_cl_sorted,
            default=def_cl,
            format_func=lambda x: map_cl[x],
            key=f"ms_cl_{type_ent_code}_{code_ent_sel}"
        )

        # 2. Familles disponibles selon les classes choisies
        fam_dispo_set = set()
        for c in (sel_classes if sel_classes else liste_cl_sorted):
            fam_dispo_set.update(cl_to_fam.get(c, set()))
        fam_dispo = sorted(fam_dispo_set)
        map_fam = {f: f"{f} - {nom_fam.get(f, f)}" for f in fam_dispo}

        def_fam = [f for f in fam_ok_act if f in fam_dispo]
        sel_familles = st.multiselect(
            "2️⃣ Familles AUTORISÉES (optionnel : affine les classes choisies) :",
            fam_dispo,
            default=def_fam,
            format_func=lambda x: map_fam[x],
            key=f"ms_fam_{type_ent_code}_{code_ent_sel}"
        )

        # 3. Sous-Familles disponibles selon les familles ou classes choisies
        sf_dispo_set = set()
        for f in (sel_familles if sel_familles else fam_dispo):
            sf_dispo_set.update(fam_to_sf.get(f, set()))
        sf_dispo = sorted(sf_dispo_set)
        map_sf = {s: f"{s} - {nom_sf.get(s, s)}" if nom_sf.get(s, s).upper() != s else s for s in sf_dispo}

        def_sf = [s for s in sf_ok_act if s in sf_dispo]
        sel_sf = st.multiselect(
            "3️⃣ Sous-Familles AUTORISÉES (optionnel : ne sélectionnez que celles autorisées, ex. tout sauf OILF) :",
            sf_dispo,
            default=def_sf,
            format_func=lambda x: map_sf[x],
            key=f"ms_sf_{type_ent_code}_{code_ent_sel}"
        )

        # 4. Articles précis à DÉSACTIVER / EXCLURE (index 2 = sous_famille)
        articles_eligibles = []
        if sel_sf:
            for s in sel_sf:
                articles_eligibles.extend(art_par_sf.get(s, []))
        elif sel_familles:
            for f in sel_familles:
                articles_eligibles.extend(art_par_fam.get(f, []))
        elif sel_classes:
            for c in sel_classes:
                articles_eligibles.extend(art_par_cl.get(c, []))
        else:
            articles_eligibles = tous_art_m8

        map_art_m8 = {
            itemgetter(0)(a).upper(): f"{itemgetter(0)(a)} - {itemgetter(1)(a)} [{itemgetter(2)(a)}]"
            for a in articles_eligibles
        }
        def_exc = [c for c in art_exc_act if c in map_art_m8]
        sel_exclus = st.multiselect(
            f"4️⃣ Articles précis à DÉSACTIVER / EXCLURE ({len(articles_eligibles)} articles dans la sélection) :",
            list(map_art_m8.keys()),
            default=def_exc,
            format_func=lambda x: map_art_m8[x],
            key=f"ms_exc_{type_ent_code}_{code_ent_sel}"
        )

        col_b1, col_b2 = st.columns(2)
        if col_b1.button("💾 Enregistrer ce paramétrage", type="primary", key="btn_save_m8"):
            ok_m8, msg_m8 = db.enregistrer_config_autorisations_entite(
                type_ent_code, code_ent_sel, sel_classes, sel_familles, sel_sf, sel_exclus
            )
            if ok_m8:
                st.success(msg_m8)
            else:
                st.error(msg_m8)

        if col_b2.button("🔄 Tout réactiver (supprimer toutes les restrictions)", key="btn_reset_m8"):
            db.enregistrer_config_autorisations_entite(type_ent_code, code_ent_sel, [], [], [], [])
            st.success(f"Toutes les classes, familles et articles sont de nouveau actifs pour {map_ent[code_ent_sel]}.")
            st.rerun()
elif menu == "🗂️ Base de données":
    from operator import itemgetter
    btab1, btab2, btab3, btab4 = st.tabs(["📦 Base Articles", "🚚 Base Fournisseurs", "🏢 Base Départements", "📥 Import Base Articles"])

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
            st.divider()
            st.markdown("#### 📥 Importer la Base Fournisseurs (Excel .xlsx)")
            st.caption("Colonnes attendues : Code Fournisseur, Nom, Famille, Identifiant 1 (MF), Conventionné, Adresse.")
            fichier_fn = st.file_uploader("Choisir le fichier Excel des fournisseurs (.xlsx) :", type=["xlsx", "xlsm"], key="up_base_fourn")
            ecraser_fn = st.checkbox("Mettre à jour les fournisseurs existants (même code)", value=False, key="chk_maj_fn")
            if fichier_fn is not None:
                if st.button("🚀 Injecter les Fournisseurs", type="primary", key="btn_imp_fn"):
                    with st.spinner("Importation des fournisseurs en cours..."):
                        suf_fn = os.path.splitext(fichier_fn.name)[1]
                        with tempfile.NamedTemporaryFile(delete=False, suffix=suf_fn) as tmp_fn:
                            tmp_fn.write(fichier_fn.getvalue())
                            path_fn = tmp_fn.name
                        ok_fn, res_fn = db.importer_fournisseurs_excel(path_fn, ecraser_existants=ecraser_fn)
                        os.remove(path_fn)
                        if ok_fn:
                            aj, mj = res_fn
                            st.success(f"Importation terminée : **{aj}** fournisseur(s) ajouté(s), **{mj}** mis à jour !")
                            st.rerun()
                        else:
                            st.error(f"Erreur : {res_fn}")

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
            st.divider()
            if st.button("🔄 Charger & Activer les 30 Départements Officiels (01 à 30)", type="primary", key="btn_load_depts_off"):
                ok_dp, msg_dp = db.charger_departements_officiels()
                if ok_dp:
                    st.success(msg_dp)
                    st.rerun()
                else:
                    st.error(msg_dp)
                    # ---------------- ONGLET 4 : IMPORT BASE ARTICLES EXCEL ----------------
    with btab4:
        st.subheader("📥 Injection automatique de la Base Articles & Nomenclature")
        st.info("Chargez votre fichier Excel 'Base articles.xlsx' : l'application créera automatiquement les Classes, Familles, Sous-Familles et Articles sans toucher à vos mouvements existants.")
        fichier_ba = st.file_uploader("Choisir le fichier Excel (Base articles.xlsx) :", type=["xlsx", "xlsm"], key="up_base_art")
        mode_ecraser = st.checkbox("Mettre à jour les articles existants s'ils ont le même code (sinon ignorer les doublons)", value=False, key="chk_ecraser_art")

        if fichier_ba is not None:
            if st.button("🚀 Injecter la Base Articles maintenant", type="primary", key="btn_imp_ba"):
                with st.spinner("Injection de la nomenclature et des articles en cours..."):
                    suf = os.path.splitext(fichier_ba.name)[1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suf) as tmp_f:
                        tmp_f.write(fichier_ba.getvalue())
                        path_ba = tmp_f.name
                    ok_ba, res_ba = db.importer_base_articles_excel(path_ba, ecraser_existants=mode_ecraser)
                    os.remove(path_ba)
                    if ok_ba:
                        n_cl, n_fam, n_sf, n_art = res_ba
                        st.success(f"Injection réussie !\n- **{n_cl}** nouvelles Classes\n- **{n_fam}** nouvelles Familles\n- **{n_sf}** nouvelles Sous-Familles\n- **{n_art}** Articles injectés")
                    else:
                        st.error(f"Erreur lors de l'injection : {res_ba}")
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
        articles_actifs = _cache_articles_complets()
        fournisseurs_actifs = _cache_fournisseurs()
        depts_autorises = _cache_depts_achats()

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
            if besoin_fournisseur and fourn_sel:
                articles_filtres = db.filtrer_articles_autorises(articles_actifs, "FOURNISSEUR", fourn_sel)
            else:
                articles_filtres = db.filtrer_articles_autorises(articles_actifs, "DEPARTEMENT", dept_sel)
            map_art = {


                itemgetter(0)(a): f"{itemgetter(0)(a)} - {itemgetter(1)(a)} ({itemgetter(6)(a)})"
                for a in articles_filtres
            }
            dict_art_info = {itemgetter(0)(a): a for a in articles_filtres}

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

                
elif menu == "🔄 Transferts":
    from operator import itemgetter
    from datetime import date

    ttab1, ttab2, ttab3 = st.tabs([
        "📋 Demande de Transfert",
        "✅ Validation Transferts",
        "🚚 Bon de Transfert Direct"
    ])

    articles_actifs = _cache_articles_complets()
    depts_actifs = _cache_depts_actifs()
    map_depts_t = {itemgetter(0)(d): f"{itemgetter(0)(d)} - {itemgetter(1)(d)}" for d in depts_actifs}
    liste_codes_depts = list(map_depts_t.keys())
    idx_eco = liste_codes_depts.index("01") if "01" in liste_codes_depts else (liste_codes_depts.index("ECO") if "ECO" in liste_codes_depts else 0)

    with ttab1:
        mode_dt = st.radio("Opération :", ["➕ Nouvelle Demande", "📋 Consulter / Modifier / Supprimer"], key="mode_dt", horizontal=True)
        if "panier_DT" not in st.session_state:
            st.session_state["panier_DT"] = []

        if mode_dt == "➕ Nouvelle Demande":
            if not articles_actifs or len(liste_codes_depts) < 2:
                st.warning("Assurez-vous d'avoir au moins 1 article actif et 2 départements actifs.")
            else:
                num_dt = db.generer_numero_transfert("DT", str(date.today().year))
                c1, c2, c3, c4 = st.columns(4)
                c1.text_input("N° Demande :", value=num_dt, disabled=True, key="num_dt_new")
                date_dt = c2.date_input("Date Demande :", value=date.today(), key="date_dt_new")
                src_dt = c3.selectbox("Dépôt Source :", liste_codes_depts, index=idx_eco,
                                      format_func=lambda x: map_depts_t[x], key="src_dt_new")
                autres_depts = [d for d in liste_codes_depts if d != src_dt]
                dst_dt = c4.selectbox("Département Demandeur :", autres_depts,
                                      format_func=lambda x: map_depts_t[x], key="dst_dt_new")
                obs_dt = st.text_input("Observation / Besoin :", key="obs_dt_new")
                articles_dt_filtres = db.filtrer_articles_autorises(articles_actifs, "DEPARTEMENT", dst_dt)
                map_art_t = {itemgetter(0)(a): f"{itemgetter(0)(a)} - {itemgetter(1)(a)} ({itemgetter(6)(a)})" for a in articles_dt_filtres}
                dict_art_t = {itemgetter(0)(a): a for a in articles_dt_filtres}
                map_art_t = {itemgetter(0)(a): f"{itemgetter(0)(a)} - {itemgetter(1)(a)} ({itemgetter(6)(a)})" for a in articles_actifs}
                dict_art_t = {itemgetter(0)(a): a for a in articles_actifs}

                ca1, ca2, ca3, ca4 = st.columns([3, 1.5, 1.5, 1.5])
                art_dt = ca1.selectbox("Article demandé :", list(map_art_t.keys()), format_func=lambda x: map_art_t[x], key="art_sel_dt")
                info_a = dict_art_t[art_dt]
                stock_dispo_dt = db.calculer_stock_article_dept(art_dt, src_dt)
                ca2.metric(f"Stock {src_dt}", f"{stock_dispo_dt:,.3f} {itemgetter(6)(info_a)}")
                qte_dt = ca3.number_input("Qté demandée :", min_value=0.001, value=1.0, step=1.0, format="%.3f", key="qte_sel_dt")

                if ca4.button("➕ Ajouter", key="btn_add_line_dt"):
                    st.session_state["panier_DT"].append((
                        art_dt, itemgetter(1)(info_a), itemgetter(6)(info_a), qte_dt, qte_dt, float(itemgetter(7)(info_a))
                    ))
                    st.rerun()

                if st.session_state["panier_DT"]:
                    st.dataframe(pd.DataFrame([
                        {
                            "Code": itemgetter(0)(x),
                            "Désignation": itemgetter(1)(x),
                            "Unité": itemgetter(2)(x),
                            "Qté Demandée": f"{itemgetter(3)(x):,.3f}",
                            "Stock Actuel Source": f"{db.calculer_stock_article_dept(itemgetter(0)(x), src_dt):,.3f}"
                        }
                        for x in st.session_state["panier_DT"]
                    ]), use_container_width=True, hide_index=True)

                    b1, b2 = st.columns(2)
                    if b1.button("💾 Enregistrer la Demande", type="primary", key="save_dt"):
                        ok, msg = db.enregistrer_transfert(
                            num_dt, "DT", date_dt, src_dt, dst_dt, "EN ATTENTE", obs_dt, st.session_state["panier_DT"], creation=True
                        )
                        if ok:
                            st.session_state["panier_DT"] = []
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    if b2.button("🧹 Vider", key="clear_dt"):
                        st.session_state["panier_DT"] = []
                        st.rerun()
        else:
            demandes = db.lister_transferts("DT")
            if demandes:
                st.dataframe(pd.DataFrame([
                    {
                        "N° Demande": itemgetter(0)(d),
                        "Date": itemgetter(1)(d),
                        "Source": itemgetter(2)(d),
                        "Demandeur": itemgetter(3)(d),
                        "Statut": itemgetter(4)(d),
                        "Observation": itemgetter(5)(d)
                    }
                    for d in demandes
                ]), use_container_width=True, hide_index=True)

                dict_dt = {itemgetter(0)(d): d for d in demandes}
                sel_dt = st.selectbox("Sélectionner une demande :", list(dict_dt.keys()), key="sel_view_dt")
                info_dt = dict_dt[sel_dt]
                lignes_dt = db.lire_lignes_transfert(sel_dt)
                st.dataframe(pd.DataFrame([
                    {
                        "Code": itemgetter(0)(l),
                        "Désignation": itemgetter(1)(l),
                        "Unité": itemgetter(2)(l),
                        "Qté Demandée": f"{itemgetter(3)(l):,.3f}",
                        "Qté Accordée": f"{itemgetter(4)(l):,.3f}"
                    }
                    for l in lignes_dt
                ]), use_container_width=True, hide_index=True)

                if itemgetter(4)(info_dt) == "EN ATTENTE":
                    with st.expander("✏️ Modifier les quantités de cette demande en attente"):
                        new_obs_dt = st.text_input("Observation :", value=itemgetter(5)(info_dt), key="edit_obs_dt")
                        lignes_mod = []
                        for idx_l, l in enumerate(lignes_dt):
                            mc1, mc2 = st.columns(2)
                            mc1.text(f"{itemgetter(0)(l)} - {itemgetter(1)(l)}")
                            nq = mc2.number_input("Qté demandée", value=float(itemgetter(3)(l)), min_value=0.0, step=1.0, format="%.3f", key=f"ed_dt_q_{idx_l}")
                            if nq > 0:
                                lignes_mod.append((itemgetter(0)(l), itemgetter(1)(l), itemgetter(2)(l), nq, nq, float(itemgetter(5)(l))))
                        if st.button("💾 Enregistrer modification", type="primary", key="btn_upd_dt"):
                            ok, msg = db.enregistrer_transfert(
                                sel_dt, "DT", itemgetter(1)(info_dt), itemgetter(2)(info_dt), itemgetter(3)(info_dt),
                                "EN ATTENTE", new_obs_dt, lignes_mod, creation=False
                            )
                            st.success(msg)
                            st.rerun()

                    if st.button(f"🗑️ Supprimer la demande {sel_dt}", key="btn_del_dt"):
                        ok, msg = db.supprimer_transfert(sel_dt)
                        st.success(msg)
                        st.rerun()
                else:
                    st.info("Cette demande est déjà traitée (Générée ou Rejetée) et ne peut plus être modifiée.")
            else:
                st.info("Aucune demande de transfert enregistrée.")
    with ttab2:
        st.subheader("Demandes de transfert en attente de validation")
        demandes_attente = db.lister_transferts("DT", statut_filtre="EN ATTENTE")
        if not demandes_attente:
            st.success("✅ Aucune demande de transfert en attente.")
        else:
            map_att = {
                itemgetter(0)(d): f"{itemgetter(0)(d)} | {itemgetter(1)(d)} | {itemgetter(2)(d)} ➔ {itemgetter(3)(d)}"
                for d in demandes_attente
            }
            dict_att = {itemgetter(0)(d): d for d in demandes_attente}
            sel_val_dt = st.selectbox("Choisir la demande à traiter :", list(map_att.keys()), format_func=lambda x: map_att[x], key="sel_val_dt")
            d_info = dict_att[sel_val_dt]
            src_v = itemgetter(2)(d_info)
            lignes_v = db.lire_lignes_transfert(sel_val_dt)

            st.info(f"Dépôt source : **{src_v}** ➔ Département demandeur : **{itemgetter(3)(d_info)}** | Observation : {itemgetter(5)(d_info) or 'Aucune'}")
            lignes_a_valider = []
            tout_dispo = True

            for idx_v, l in enumerate(lignes_v):
                ca, da, ua, q_dem, _, pu, _ = l
                dispo = db.calculer_stock_article_dept(ca, src_v)
                vc1, vc2, vc3, vc4 = st.columns([3, 1.5, 1.5, 1.5])
                vc1.write(f"**{ca}** - {da} ({ua})")
                vc2.metric("Demandé", f"{q_dem:,.3f}")
                vc3.metric(f"Stock {src_v}", f"{dispo:,.3f}")
                q_acc = vc4.number_input("Qté Accordée :", min_value=0.0, value=float(q_dem), step=1.0, format="%.3f", key=f"val_q_{sel_val_dt}_{idx_v}")
                if q_acc > dispo:
                    tout_dispo = False
                    st.error(f"⚠️ Stock insuffisant pour {ca} : disponible = {dispo:,.3f} {ua}, accordé = {q_acc:,.3f} {ua}")
                lignes_a_valider.append((ca, da, ua, q_dem, q_acc, pu))

            obs_val = st.text_input("Note de validation / motif de rejet :", key="obs_val_dt")
            date_gen_bt = st.date_input("Date du Bon de Transfert généré :", value=date.today(), key="date_gen_bt")

            bv1, bv2, bv3 = st.columns(3)
            if bv1.button("✅ Générer & Valider le Transfert", type="primary", disabled=not tout_dispo, key="btn_gen_bt"):
                ok, msg = db.valider_demande_et_generer_bt(sel_val_dt, date_gen_bt, lignes_a_valider, obs_val)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

            if bv2.button("⛔ Rejeter la demande", key="btn_rej_dt"):
                ok, msg = db.changer_statut_demande(sel_val_dt, "REJETÉE")
                st.warning(msg)
                st.rerun()

            if bv3.button("🗑️ Supprimer la demande", key="btn_sup_val_dt"):
                ok, msg = db.supprimer_transfert(sel_val_dt)
                st.success(msg)
                st.rerun()
    with ttab3:
        mode_bt = st.radio("Opération :", ["➕ Nouveau Bon de Transfert Direct", "📋 Consulter / Modifier / Supprimer"], key="mode_bt", horizontal=True)
        if "panier_BT" not in st.session_state:
            st.session_state["panier_BT"] = []

        if mode_bt == "➕ Nouveau Bon de Transfert Direct":
            if not articles_actifs or len(liste_codes_depts) < 2:
                st.warning("Assurez-vous d'avoir au moins 1 article actif et 2 départements actifs.")
            else:
                num_bt = db.generer_numero_transfert("BT", str(date.today().year))
                c1, c2, c3, c4 = st.columns(4)
                c1.text_input("N° Bon Transfert :", value=num_bt, disabled=True, key="num_bt_new")
                date_bt = c2.date_input("Date Transfert :", value=date.today(), key="date_bt_new")
                src_bt = c3.selectbox("Département Source :", liste_codes_depts, index=idx_eco,
                                      format_func=lambda x: map_depts_t[x], key="src_bt_new")
                autres_bt = [d for d in liste_codes_depts if d != src_bt]
                dst_bt = c4.selectbox("Département Destinataire :", autres_bt,
                                      format_func=lambda x: map_depts_t[x], key="dst_bt_new")
                obs_bt = st.text_input("Observation :", key="obs_bt_new")
                articles_bt_filtres = db.filtrer_articles_autorises(articles_actifs, "DEPARTEMENT", dst_bt)
                map_art_bt = {itemgetter(0)(a): f"{itemgetter(0)(a)} - {itemgetter(1)(a)} ({itemgetter(6)(a)})" for a in articles_bt_filtres}
                dict_art_bt = {itemgetter(0)(a): a for a in articles_bt_filtres}

                ca1, ca2, ca3, ca4 = st.columns([3, 1.5, 1.5, 1.5])
                art_bt = ca1.selectbox("Article :", list(map_art_bt.keys()), format_func=lambda x: map_art_bt[x], key="art_sel_bt")
                info_a_bt = dict_art_bt[art_bt]
                stock_dispo_bt = db.calculer_stock_article_dept(art_bt, src_bt)
                ca2.metric(f"Stock {src_bt}", f"{stock_dispo_bt:,.3f} {itemgetter(6)(info_a_bt)}")
                qte_bt = ca3.number_input("Qté à transférer :", min_value=0.001, value=1.0, step=1.0, format="%.3f", key="qte_sel_bt")

                if ca4.button("➕ Ajouter", key="btn_add_line_bt"):
                    deja_panier = sum(itemgetter(4)(x) for x in st.session_state["panier_BT"] if itemgetter(0)(x) == art_bt)
                    if (qte_bt + deja_panier) > stock_dispo_bt:
                        st.error(f"Stock insuffisant dans {src_bt} ! Disponible : {stock_dispo_bt:,.3f} {itemgetter(6)(info_a_bt)}.")
                    else:
                        st.session_state["panier_BT"].append((
                            art_bt, itemgetter(1)(info_a_bt), itemgetter(6)(info_a_bt), qte_bt, qte_bt, float(itemgetter(7)(info_a_bt))
                        ))
                        st.rerun()

                if st.session_state["panier_BT"]:
                    st.dataframe(pd.DataFrame([
                        {
                            "Code": itemgetter(0)(x),
                            "Désignation": itemgetter(1)(x),
                            "Unité": itemgetter(2)(x),
                            "Qté Transférée": f"{itemgetter(4)(x):,.3f}",
                            "PU HT": f"{itemgetter(5)(x):,.3f}",
                            "Valeur HT": f"{itemgetter(4)(x) * itemgetter(5)(x):,.3f} DT"
                        }
                        for x in st.session_state["panier_BT"]
                    ]), use_container_width=True, hide_index=True)

                    b1, b2 = st.columns(2)
                    if b1.button(f"💾 Valider & Enregistrer ({num_bt})", type="primary", key="save_bt"):
                        ok, msg = db.enregistrer_transfert(
                            num_bt, "BT", date_bt, src_bt, dst_bt, "VALIDÉ", obs_bt, st.session_state["panier_BT"], creation=True
                        )
                        if ok:
                            st.session_state["panier_BT"] = []
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                    if b2.button("🧹 Vider", key="clear_bt"):
                        st.session_state["panier_BT"] = []
                        st.rerun()
        else:
            bons_bt = db.lister_transferts("BT")
            if bons_bt:
                st.dataframe(pd.DataFrame([
                    {
                        "N° Bon": itemgetter(0)(b),
                        "Date": itemgetter(1)(b),
                        "Source": itemgetter(2)(b),
                        "Destination": itemgetter(3)(b),
                        "Valeur HT": f"{itemgetter(7)(b):,.3f} DT",
                        "Origine": itemgetter(6)(b) or "Direct",
                        "Observation": itemgetter(5)(b)
                    }
                    for b in bons_bt
                ]), use_container_width=True, hide_index=True)

                dict_bt = {itemgetter(0)(b): b for b in bons_bt}
                sel_bt = st.selectbox("Sélectionner un bon de transfert :", list(dict_bt.keys()), key="sel_view_bt")
                info_bt = dict_bt[sel_bt]
                lignes_bt = db.lire_lignes_transfert(sel_bt)
                st.dataframe(pd.DataFrame([
                    {
                        "Code": itemgetter(0)(l),
                        "Désignation": itemgetter(1)(l),
                        "Unité": itemgetter(2)(l),
                        "Qté Transférée": f"{itemgetter(4)(l):,.3f}",
                        "PU HT": f"{itemgetter(5)(l):,.3f}",
                        "Montant HT": f"{itemgetter(6)(l):,.3f} DT"
                    }
                    for l in lignes_bt
                ]), use_container_width=True, hide_index=True)

                with st.expander("✏️ Modifier les quantités de ce bon de transfert"):
                    new_obs_bt = st.text_input("Observation :", value=itemgetter(5)(info_bt), key="edit_obs_bt")
                    lignes_mod_bt = []
                    for idx_l, l in enumerate(lignes_bt):
                        ca = itemgetter(0)(l)
                        dispo_mod = db.calculer_stock_article_dept(ca, itemgetter(2)(info_bt), ignorer_doc_transfert=sel_bt)
                        mc1, mc2, mc3 = st.columns([3, 1.5, 1.5])
                        mc1.text(f"{ca} - {itemgetter(1)(l)}")
                        mc2.metric("Stock Max", f"{dispo_mod:,.3f}")
                        nq = mc3.number_input("Nouvelle Qté", value=float(itemgetter(4)(l)), min_value=0.0, step=1.0, format="%.3f", key=f"ed_bt_q_{idx_l}")
                        if nq > 0:
                            lignes_mod_bt.append((ca, itemgetter(1)(l), itemgetter(2)(l), itemgetter(3)(l), nq, float(itemgetter(5)(l))))

                    if st.button("💾 Mettre à jour le Bon", type="primary", key="btn_upd_bt"):
                        ok, msg = db.enregistrer_transfert(
                            sel_bt, "BT", itemgetter(1)(info_bt), itemgetter(2)(info_bt), itemgetter(3)(info_bt),
                            "VALIDÉ", new_obs_bt, lignes_mod_bt, creation=False
                        )
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

                if st.button(f"🗑️ Supprimer le bon {sel_bt}", key="btn_del_bt"):
                    ok, msg = db.supprimer_transfert(sel_bt)
                    st.success(msg)
                    st.rerun()
            else:
                st.info("Aucun bon de transfert enregistré.")
elif menu == "📦 Stock":
    st.subheader("Visualisation & Valorisation du Stock en Temps Réel")
    donnees_stock = db.lister_etat_stock_global()

    if not donnees_stock:
        st.info("Aucun article ou département actif trouvé pour afficher le stock.")
    else:
        df_s = pd.DataFrame(donnees_stock)

        # Filtres rapides
        fc1, fc2, fc3, fc4 = st.columns(4)
        liste_dp = ["Tous"] + sorted(df_s["nom_dept"].unique().tolist())
        liste_cl = ["Toutes"] + sorted(df_s["classe"].unique().tolist())
        liste_fam = ["Toutes"] + sorted(df_s["famille"].unique().tolist())
        liste_sf = ["Toutes"] + sorted(df_s["sous_famille"].unique().tolist())

        sel_dp = fc1.selectbox("Département :", liste_dp, key="stk_f_dp")
        sel_cl = fc2.selectbox("Classe :", liste_cl, key="stk_f_cl")
        sel_fam = fc3.selectbox("Famille :", liste_fam, key="stk_f_fam")
        sel_sf = fc4.selectbox("Sous-famille :", liste_sf, key="stk_f_sf")

        masquer_zero = st.checkbox("Masquer les lignes avec stock à 0", value=False, key="stk_no_zero")

        df_filtre = df_s.copy()
        if sel_dp != "Tous":
            df_filtre = df_filtre[df_filtre["nom_dept"] == sel_dp]
        if sel_cl != "Toutes":
            df_filtre = df_filtre[df_filtre["classe"] == sel_cl]
        if sel_fam != "Toutes":
            df_filtre = df_filtre[df_filtre["famille"] == sel_fam]
        if sel_sf != "Toutes":
            df_filtre = df_filtre[df_filtre["sous_famille"] == sel_sf]
        if masquer_zero:
            df_filtre = df_filtre[df_filtre["qte"] != 0]

        # Vue d'analyse selon le niveau souhaité
        vue_niveau = st.radio(
            "Regrouper la visualisation :",
            ["📦 Par Article (Détail)", "🏢 Par Département", "📁 Par Classe", "📂 Par Famille", "🏷️ Par Sous-Famille"],
            horizontal=True,
            key="stk_vue_niv"
        )

        val_totale = df_filtre["valeur_ht"].sum()
        qte_totale = df_filtre["qte"].sum()
        k1, k2 = st.columns(2)
        k1.metric("Quantité Totale Cumulée", f"{qte_totale:,.3f}")
        k2.metric("Valeur Totale du Stock (HT)", f"{val_totale:,.3f} DT")

        if vue_niveau == "📦 Par Article (Détail)":
            df_aff = df_filtre[["nom_dept", "classe", "famille", "sous_famille", "code_article", "designation", "unite", "qte", "pu_ht", "valeur_ht"]].copy()
            df_aff.columns = ["Département", "Classe", "Famille", "Sous-Famille", "Code", "Désignation", "Unité", "Stock", "PU HT", "Valeur HT (DT)"]
        elif vue_niveau == "🏢 Par Département":
            df_aff = df_filtre.groupby(["nom_dept"], as_index=False).agg({"qte": "sum", "valeur_ht": "sum"})
            df_aff.columns = ["Département", "Quantité Totale", "Valeur Stock HT (DT)"]
        elif vue_niveau == "📁 Par Classe":
            df_aff = df_filtre.groupby(["classe"], as_index=False).agg({"qte": "sum", "valeur_ht": "sum"})
            df_aff.columns = ["Classe", "Quantité Totale", "Valeur Stock HT (DT)"]
        elif vue_niveau == "📂 Par Famille":
            df_aff = df_filtre.groupby(["classe", "famille"], as_index=False).agg({"qte": "sum", "valeur_ht": "sum"})
            df_aff.columns = ["Classe", "Famille", "Quantité Totale", "Valeur Stock HT (DT)"]
        else:
            df_aff = df_filtre.groupby(["classe", "famille", "sous_famille"], as_index=False).agg({"qte": "sum", "valeur_ht": "sum"})
            df_aff.columns = ["Classe", "Famille", "Sous-Famille", "Quantité Totale", "Valeur Stock HT (DT)"]

        st.dataframe(df_aff, use_container_width=True, hide_index=True, height=450)

        entetes_s = list(df_aff.columns)
        lignes_s = [[f"{v:,.3f}" if isinstance(v, float) else str(v) for v in row] for row in df_aff.values.tolist()]
        eb1, eb2 = st.columns(2)
        eb1.download_button("📄 Exporter Stock PDF", data=generer_pdf_memoire("Etat_du_Stock", entetes_s, lignes_s, paysage=True),
                            file_name="Etat_du_Stock.pdf", mime="application/pdf")
        eb2.download_button("📊 Exporter Stock Excel", data=generer_excel_memoire("Etat_du_Stock", entetes_s, lignes_s),
                            file_name="Etat_du_Stock.xlsx")
elif menu == "📋 Inventaire":
    from operator import itemgetter
    from datetime import date

    itab1, itab2 = st.tabs(["📝 Pointage de Stock", "🔒 Stock Physique & Clôture"])

    depts_actifs = db.lister_departements_complet(seulement_actifs=True)
    articles_actifs = db.lister_articles_complet(seulement_actifs=True)

    # ---------------- ONGLET 1 : POINTAGE DE STOCK ----------------
    with itab1:
        st.subheader("Saisie d'un nouveau pointage d'inventaire physique")
        if not depts_actifs or not articles_actifs:
            st.warning("Assurez-vous d'avoir au moins un département et un article actifs.")
        else:
            map_dp_inv = {itemgetter(0)(d): f"{itemgetter(0)(d)} - {itemgetter(1)(d)}" for d in depts_actifs}
            c1, c2, c3 = st.columns(3)
            num_inv = db.generer_numero_inventaire(str(date.today().year))
            c1.text_input("N° Inventaire :", value=num_inv, disabled=True, key="inv_num_new")
            date_inv = c2.date_input("Date du comptage :", value=date.today(), key="inv_date_new")
            dept_inv = c3.selectbox("Département inventorié :", list(map_dp_inv.keys()),
                                    format_func=lambda x: map_dp_inv[x], key="inv_dept_new")
            obs_inv = st.text_input("Observation (ex: Inventaire fin de mois) :", key="inv_obs_new")

            st.info(f"Saisissez les quantités physiques comptées dans **{dept_inv}** (le stock théorique est pré-rempli par défaut) :")

            # Préparation du tableau de pointage éditable
            lignes_prep = []
            for a in articles_actifs:
                ca = itemgetter(0)(a)
                da = itemgetter(1)(a)
                ua = itemgetter(6)(a)
                pu = float(itemgetter(7)(a))
                st_th = db.calculer_stock_article_dept(ca, dept_inv)
                lignes_prep.append({
                    "Code": ca,
                    "Désignation": da,
                    "Unité": ua,
                    "Stock Théorique": st_th,
                    "Stock Physique Compté": st_th,
                    "PU HT": pu
                })

            df_pointage = st.data_editor(
                pd.DataFrame(lignes_prep),
                disabled=["Code", "Désignation", "Unité", "Stock Théorique", "PU HT"],
                use_container_width=True,
                hide_index=True,
                key=f"editor_inv_{dept_inv}"
            )

            if st.button(f"💾 Enregistrer le pointage ({num_inv})", type="primary", key="btn_save_inv"):
                lignes_a_sauver = []
                for _, row in df_pointage.iterrows():
                    lignes_a_sauver.append((
                        row["Code"],
                        row["Désignation"],
                        row["Unité"],
                        float(row["Stock Théorique"]),
                        float(row["Stock Physique Compté"]),
                        float(row["PU HT"])
                    ))
                ok, msg = db.sauvegarder_inventaire(num_inv, date_inv, dept_inv, obs_inv, lignes_a_sauver, creation=True)
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
    # ---------------- ONGLET 2 : STOCK PHYSIQUE & CLÔTURE ----------------
    with itab2:
        st.subheader("Gestion des Inventaires Physiques (Consulter, Modifier, Clôturer, Déclôturer)")
        inventaires = db.lister_inventaires()
        if not inventaires:
            st.info("Aucun inventaire enregistré pour le moment. Commencez par l'onglet 'Pointage de Stock'.")
        else:
            st.dataframe(pd.DataFrame([
                {
                    "N° Inventaire": itemgetter(0)(i),
                    "Date": itemgetter(1)(i),
                    "Département": itemgetter(2)(i),
                    "Statut": "🔒 CLÔTURÉ" if itemgetter(3)(i) == "CLÔTURÉ" else "📝 BROUILLON",
                    "Écart Valorisé (HT)": f"{itemgetter(5)(i):,.3f} DT",
                    "Observation": itemgetter(4)(i)
                }
                for i in inventaires
            ]), use_container_width=True, hide_index=True)

            dict_inv = {itemgetter(0)(i): i for i in inventaires}
            sel_inv = st.selectbox("Sélectionner un inventaire :", list(dict_inv.keys()), key="sel_inv_manage")
            info_i = dict_inv[sel_inv]
            statut_i = itemgetter(3)(info_i)
            lignes_i = db.lire_lignes_inventaire(sel_inv)

            st.markdown(f"**Détail de `{sel_inv}` — Département `{itemgetter(2)(info_i)}` — Statut : `{statut_i}`**")
            df_det_inv = pd.DataFrame([
                {
                    "Code": itemgetter(0)(l),
                    "Désignation": itemgetter(1)(l),
                    "Unité": itemgetter(2)(l),
                    "Stock Théorique": float(itemgetter(3)(l)),
                    "Stock Physique": float(itemgetter(4)(l)),
                    "Écart Qté": float(itemgetter(5)(l)),
                    "PU HT": float(itemgetter(6)(l)),
                    "Valeur Écart (HT)": float(itemgetter(7)(l))
                }
                for l in lignes_i
            ])

            if statut_i == "BROUILLON":
                st.info("Cet inventaire est en mode BROUILLON : vous pouvez ajuster les quantités physiques ci-dessous avant de le clôturer.")
                new_obs_i = st.text_input("Observation :", value=itemgetter(4)(info_i), key="edit_obs_inv")
                df_mod_inv = st.data_editor(
                    df_det_inv,
                    disabled=["Code", "Désignation", "Unité", "Stock Théorique", "Écart Qté", "PU HT", "Valeur Écart (HT)"],
                    use_container_width=True,
                    hide_index=True,
                    key=f"mod_ed_{sel_inv}"
                )

                c_act1, c_act2, c_act3 = st.columns(3)
                if c_act1.button("💾 Enregistrer les modifications", type="primary", key="btn_upd_inv"):
                    lignes_upd = []
                    for _, r in df_mod_inv.iterrows():
                        lignes_upd.append((r["Code"], r["Désignation"], r["Unité"], float(r["Stock Théorique"]), float(r["Stock Physique"]), float(r["PU HT"])))
                    ok, msg = db.sauvegarder_inventaire(sel_inv, itemgetter(1)(info_i), itemgetter(2)(info_i), new_obs_i, lignes_upd, creation=False)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

                if c_act2.button("🔒 Clôturer cet inventaire", key="btn_clot_inv"):
                    ok, msg = db.cloturer_inventaire(sel_inv)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

                if c_act3.button("🗑️ Supprimer l'inventaire", key="btn_del_inv"):
                    ok, msg = db.supprimer_inventaire(sel_inv)
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
            else:
                st.dataframe(df_det_inv, use_container_width=True, hide_index=True)
                st.warning("🔒 Cet inventaire est CLÔTURÉ : toute modification ou suppression est verrouillée.")
                if st.button("🔓 Déclôturer pour modifier", key="btn_declot_inv"):
                    ok, msg = db.decloturer_inventaire(sel_inv)
                    if ok:
                        st.warning(msg)
                        st.rerun()
                    else:
                        st.error(msg)
elif menu == "📈 Gestion":
    from operator import itemgetter
    from datetime import date

    gtab1, gtab2 = st.tabs(["🛒 Gestion des Achats (ERP)", "🍽️ Gestion des Consommations (ERP)"])

    c_d1, c_d2 = st.columns(2)
    date_deb_g = c_d1.date_input("Du :", value=date(date.today().year, date.today().month, 1), key="gest_d1")
    date_fin_g = c_d2.date_input("Au :", value=date.today(), key="gest_d2")

    with gtab1:
        st.subheader("Suivi des Achats & Réceptions sur la période")
        achats_p = db.calculer_gestion_achats_periode(date_deb_g, date_fin_g)
        if not achats_p:
            st.info("Aucun mouvement d'achat (BR, ED, RF) enregistré sur cette période.")
        else:
            lignes_ga = []
            tot_ht_a, tot_ttc_a = 0.0, 0.0
            for tdoc, ddoc, num, fourn, dp, ca, da, sf, cl, qte, mht, mttc in achats_p:
                signe = -1.0 if tdoc == "RF" else 1.0
                tot_ht_a += signe * float(mht)
                tot_ttc_a += signe * float(mttc)
                lignes_ga.append({
                    "Type": tdoc,
                    "Date": ddoc,
                    "N° Doc": num,
                    "Fournisseur": fourn,
                    "Dépôt": dp,
                    "Classe": cl,
                    "Sous-Famille": sf,
                    "Article": f"{ca} - {da}",
                    "Qté": f"{signe * float(qte):,.3f}",
                    "Montant HT": f"{signe * float(mht):,.3f}",
                    "Montant TTC": f"{signe * float(mttc):,.3f}"
                })
            k1, k2 = st.columns(2)
            k1.metric("Total Achats Nets HT", f"{tot_ht_a:,.3f} DT")
            k2.metric("Total Achats Nets TTC", f"{tot_ttc_a:,.3f} DT")
            st.dataframe(pd.DataFrame(lignes_ga), use_container_width=True, hide_index=True, height=420)

    with gtab2:
        st.subheader("Suivi des Prélèvements & Transferts Consommateurs sur la période")
        conso_p = db.calculer_couts_estimatifs(date_deb_g, date_fin_g, code_dept_filtre="TOUS")
        if not conso_p:
            st.info("Aucune consommation (Transfert validé ou Sortie directe) enregistrée sur cette période.")
        else:
            df_gc = pd.DataFrame(conso_p)
            st.metric("Total Consommation Estimative (HT)", f"{df_gc['cout_ht'].sum():,.3f} DT")
            df_gc_aff = df_gc[["dept", "classe", "sous_famille", "code_article", "designation", "unite", "qte", "cout_ht"]].copy()
            df_gc_aff.columns = ["Département", "Classe", "Sous-Famille", "Code", "Désignation", "Unité", "Quantité", "Valeur HT (DT)"]
            st.dataframe(df_gc_aff, use_container_width=True, hide_index=True, height=420)
elif menu == "💰 Centre des Coûts":
    from operator import itemgetter
    from datetime import date

    st.subheader("Analyse des Coûts F&B — Estimatif vs Réel Inventaire")

    volet = st.radio(
        "Choisir le volet de calcul :",
        ["📊 Volet 1 : Estimatif (Prélèvements & Transferts)", "🔒 Volet 2 : Réel (Suite aux Inventaires Clôturés)"],
        horizontal=True,
        key="cc_volet"
    )

    depts_actifs = db.lister_departements_complet(seulement_actifs=True)
    map_dp_cc = {itemgetter(0)(d): f"{itemgetter(0)(d)} - {itemgetter(1)(d)}" for d in depts_actifs}

    donnees_cc = []
    titre_periode = ""

    if volet.startswith("📊"):
        cc1, cc2, cc3 = st.columns(3)
        d_deb_cc = cc1.date_input("Date début :", value=date(date.today().year, date.today().month, 1), key="cc_d1")
        d_fin_cc = cc2.date_input("Date fin :", value=date.today(), key="cc_d2")
        opts_dp = ["TOUS"] + list(map_dp_cc.keys())
        dp_sel_cc = cc3.selectbox("Département :", opts_dp,
                                  format_func=lambda x: "Tous les départements" if x == "TOUS" else map_dp_cc[x],
                                  key="cc_dp_est")
        donnees_cc = db.calculer_couts_estimatifs(d_deb_cc, d_fin_cc, code_dept_filtre=dp_sel_cc)
        titre_periode = f"Estimatif_{d_deb_cc}_au_{d_fin_cc}"
    else:
        if not map_dp_cc:
            st.warning("Aucun département actif trouvé.")
        else:
            dp_reel = st.selectbox("Département inventorié :", list(map_dp_cc.keys()),
                                   format_func=lambda x: map_dp_cc[x], key="cc_dp_reel")
            invs_clot = db.lister_inventaires_clotures_dept(dp_reel)
            if not invs_clot:
                st.warning(f"Aucun inventaire clôturé trouvé pour **{dp_reel}**. Clôturez d'abord un inventaire dans le module '📋 Inventaire'.")
            else:
                map_inv_c = {itemgetter(0)(i): f"{itemgetter(0)(i)} (du {itemgetter(1)(i)})" for i in invs_clot}
                liste_nums_inv = list(map_inv_c.keys())
                ic1, ic2 = st.columns(2)
                opts_deb = ["AUCUN"] + liste_nums_inv
                inv_deb = ic1.selectbox("Inventaire Initial (Stock Début) :", opts_deb,
                                        format_func=lambda x: "0 - Début d'activité (Stock initial = 0)" if x == "AUCUN" else map_inv_c[x],
                                        key="cc_inv_deb")
                inv_fin = ic2.selectbox("Inventaire Final (Stock Fin Clôturé) :", liste_nums_inv,
                                        index=len(liste_nums_inv) - 1,
                                        format_func=lambda x: map_inv_c[x],
                                        key="cc_inv_fin")
                if inv_deb != "AUCUN" and inv_deb == inv_fin:
                    st.error("L'inventaire initial et l'inventaire final doivent être différents.")
                else:
                    num_deb_reel = None if inv_deb == "AUCUN" else inv_deb
                    donnees_cc = db.calculer_couts_reels_inventaire(dp_reel, num_deb_reel, inv_fin)
                    titre_periode = f"Reel_{dp_reel}_{inv_fin}"
    st.divider()
    if not donnees_cc:
        st.info("Aucune donnée de coût disponible pour la sélection actuelle.")
    else:
        df_cc = pd.DataFrame(donnees_cc)
        cout_total_periode = df_cc["cout_ht"].sum()

        # Paramètre provisoire des nuitées (en attendant le module dédié Adultes/Enfants/Bébés/Visiteurs)
        kn1, kn2, kn3 = st.columns([1.5, 1.5, 2])
        synt_n_cc, _ = db.lister_nuitees_periode(d_deb_cc, d_fin_cc) if volet.startswith("📊") else db.lister_nuitees_periode("1900-01-01", "2099-12-31")
        mode_nuit_cc = kn1.radio("Base nuitées :", ["Pondérées F&B", "Brutes"], horizontal=True, key="cc_mode_nuit")
        nb_nuitees_auto = synt_n_cc["total_pondere"] if mode_nuit_cc == "Pondérées F&B" else synt_n_cc["total_brut"]
        nb_nuitees_cc = kn1.number_input("Nuitées retenues :", min_value=0.0, value=float(nb_nuitees_auto), step=10.0, format="%.2f", key="cc_nuit_prov")
        kn2.metric("Coût Total Consommé (HT)", f"{cout_total_periode:,.3f} DT")
        kn3.metric("Coût Moyen / Nuitée", f"{cout_nuitee_global:,.3f} DT")

        ctab1, ctab2, ctab3, ctab4 = st.tabs([
            "📦 Coût par Article",
            "🏷️ Coût par Sous-Famille",
            "📁 Coût par Classe",
            "🛏️ Coût par Nuitée (Synthèse)"
        ])

        with ctab1:
            df_art_cc = df_cc.groupby(["code_article", "designation", "unite", "classe", "sous_famille"], as_index=False).agg({"qte": "sum", "cout_ht": "sum"})
            df_art_cc["part_pct"] = (df_art_cc["cout_ht"] / cout_total_periode * 100.0) if cout_total_periode > 0 else 0.0
            df_art_cc["cout_nuitee"] = (df_art_cc["cout_ht"] / nb_nuitees_cc) if nb_nuitees_cc > 0 else 0.0
            df_art_cc = df_art_cc.sort_values(by="cout_ht", ascending=False)
            df_art_aff = df_art_cc[["code_article", "designation", "unite", "classe", "sous_famille", "qte", "cout_ht", "part_pct", "cout_nuitee"]].copy()
            df_art_aff.columns = ["Code", "Désignation", "Unité", "Classe", "Sous-Famille", "Qté Conso", "Coût Total HT (DT)", "Part (%)", "Coût / Nuitée (DT)"]
            st.dataframe(df_art_aff, use_container_width=True, hide_index=True, height=400)

        with ctab2:
            df_sf_cc = df_cc.groupby(["classe", "sous_famille"], as_index=False).agg({"qte": "sum", "cout_ht": "sum"})
            df_sf_cc["part_pct"] = (df_sf_cc["cout_ht"] / cout_total_periode * 100.0) if cout_total_periode > 0 else 0.0
            df_sf_cc["cout_nuitee"] = (df_sf_cc["cout_ht"] / nb_nuitees_cc) if nb_nuitees_cc > 0 else 0.0
            df_sf_cc = df_sf_cc.sort_values(by="cout_ht", ascending=False)
            df_sf_aff = df_sf_cc.copy()
            df_sf_aff.columns = ["Classe", "Sous-Famille", "Qté Cumulée", "Coût Total HT (DT)", "Part (%)", "Coût / Nuitée (DT)"]
            st.dataframe(df_sf_aff, use_container_width=True, hide_index=True, height=400)

        with ctab3:
            df_cl_cc = df_cc.groupby(["classe"], as_index=False).agg({"qte": "sum", "cout_ht": "sum"})
            df_cl_cc["part_pct"] = (df_cl_cc["cout_ht"] / cout_total_periode * 100.0) if cout_total_periode > 0 else 0.0
            df_cl_cc["cout_nuitee"] = (df_cl_cc["cout_ht"] / nb_nuitees_cc) if nb_nuitees_cc > 0 else 0.0
            df_cl_cc = df_cl_cc.sort_values(by="cout_ht", ascending=False)
            df_cl_aff = df_cl_cc.copy()
            df_cl_aff.columns = ["Classe", "Qté Cumulée", "Coût Total HT (DT)", "Part (%)", "Coût / Nuitée (DT)"]
            st.dataframe(df_cl_aff, use_container_width=True, hide_index=True, height=350)

        with ctab4:
            st.markdown("#### Ratio Coût / Nuitée par Département et par Classe")
            df_nuit_cc = df_cc.groupby(["dept", "classe"], as_index=False).agg({"cout_ht": "sum"})
            df_nuit_cc["cout_nuitee"] = (df_nuit_cc["cout_ht"] / nb_nuitees_cc) if nb_nuitees_cc > 0 else 0.0
            df_nuit_cc["part_pct"] = (df_nuit_cc["cout_ht"] / cout_total_periode * 100.0) if cout_total_periode > 0 else 0.0
            df_nuit_aff = df_nuit_cc.copy()
            df_nuit_aff.columns = ["Département", "Classe", "Coût Total HT (DT)", "Coût / Nuitée (DT)", "Part (%)"]
            st.dataframe(df_nuit_aff, use_container_width=True, hide_index=True)

            entetes_cc = list(df_nuit_aff.columns)
            lignes_cc = [[f"{v:,.3f}" if isinstance(v, float) else str(v) for v in row] for row in df_nuit_aff.values.tolist()]
            bc1, bc2 = st.columns(2)
            bc1.download_button("📄 Exporter Synthèse PDF", data=generer_pdf_memoire(f"Centre_Couts_{titre_periode}", entetes_cc, lignes_cc),
                                file_name=f"Centre_Couts_{titre_periode}.pdf", mime="application/pdf")
            bc2.download_button("📊 Exporter Synthèse Excel", data=generer_excel_memoire(f"Centre_Couts_{titre_periode}", entetes_cc, lignes_cc),
                                file_name=f"Centre_Couts_{titre_periode}.xlsx")
elif menu == "💾 Sauvegarde":
    from operator import itemgetter
    from datetime import datetime

    st.subheader("Sauvegarde & Restauration de la Base de Données")
    st.info("💡 Pensez à télécharger une sauvegarde avant chaque mise à jour de l'application sur GitHub ou en fin de journée.")

    col_sauv, col_rest = st.columns(2)

    with col_sauv:
        st.markdown("### 📤 1. Télécharger une sauvegarde")
        horodatage = datetime.now().strftime("%Y-%m-%d_%Hh%M")
        donnees_db = db.lire_fichier_base_octets()
        taille_ko = len(donnees_db) / 1024.0

        st.metric("Taille actuelle de la base", f"{taille_ko:,.1f} Ko")
        st.download_button(
            label="💾 Télécharger la Sauvegarde Complète (.db)",
            data=donnees_db,
            file_name=f"hotel_erp_sauvegarde_{horodatage}.db",
            mime="application/octet-stream",
            type="primary"
        )

    with col_rest:
        st.markdown("### 📥 2. Restaurer une sauvegarde")
        st.warning("⚠️ La restauration remplace les données actuelles par celles du fichier de sauvegarde choisi.")
        fichier_db = st.file_uploader("Choisir un fichier de sauvegarde (.db) :", type=["db", "sqlite", "sqlite3"], key="up_restore_db")
        if fichier_db is not None:
            if st.button("🔄 Lancer la restauration maintenant", type="primary", key="btn_restore_db"):
                ok, msg = db.restaurer_base_depuis_octets(fichier_db.getvalue())
                if ok:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

    st.divider()
    st.markdown("### 📊 Contenu actuel enregistré dans votre base")
    stats = db.statistique_tables_base()
    df_stats = pd.DataFrame([{"Module / Table": itemgetter(0)(s), "Nombre d'enregistrements": itemgetter(1)(s)} for s in stats])
    st.dataframe(df_stats, use_container_width=True, hide_index=True)
elif menu == "🛏️ Nuitées":
    from datetime import date

    ntab1, ntab2, ntab3 = st.tabs([
        "📝 Saisie & Consultation",
        "📊 Récapitulatif & Analyse",
        "⚙️ Coefficients F&B"
    ])

    coeffs_act = db.lire_coefficients_nuitees()

    # ---------------- ONGLET 1 : SAISIE JOURNALIÈRE ----------------
    with ntab1:
        st.subheader("Enregistrer ou modifier les nuitées d'une journée")
        c_d, c_obs = st.columns([1.5, 3])
        date_saisie_n = c_d.date_input("Date concernée :", value=date.today(), key="nuit_date_saisie")
        obs_n = c_obs.text_input("Observation (ex: Groupe, Séminaire, Tarif haute saison...) :", key="nuit_obs_saisie")

        # Pré-remplissage automatique si la date existe déjà
        _, existant = db.lister_nuitees_periode(date_saisie_n, date_saisie_n)
        def_pay = existant[0]["payants"] if existant else 0
        def_enf = existant[0]["enfants"] if existant else 0
        def_beb = existant[0]["bebes"] if existant else 0
        def_du = existant[0]["day_use"] if existant else 0
        def_tdu = existant[0]["tarif_du"] if existant else 0.0
        def_grat = existant[0]["gratuits"] if existant else 0

        if existant:
            st.info(f"ℹ️ Des données existent déjà pour le **{date_saisie_n}** : toute sauvegarde mettra à jour cette journée.")

        col1, col2, col3 = st.columns(3)
        n_pay = col1.number_input("Adultes Payants :", min_value=0, value=int(def_pay), step=1, key="n_pay")
        n_enf = col2.number_input("Enfants :", min_value=0, value=int(def_enf), step=1, key="n_enf")
        n_beb = col3.number_input("Bébés :", min_value=0, value=int(def_beb), step=1, key="n_beb")

        col4, col5, col6 = st.columns(3)
        n_du = col4.number_input("Day Use (Nombre) :", min_value=0, value=int(def_du), step=1, key="n_du")
        n_tdu = col5.number_input("Tarif Fixe Day Use (DT) :", min_value=0.0, value=float(def_tdu), step=5.0, format="%.3f", key="n_tdu")
        n_grat = col6.number_input("Visiteurs / Complimentary (Gratuits) :", min_value=0, value=int(def_grat), step=1, key="n_grat")

        tot_brut_prev = n_pay + n_enf + n_beb + n_du + n_grat
        tot_pond_prev = round(
            n_pay * coeffs_act["payants"] +
            n_enf * coeffs_act["enfants"] +
            n_beb * coeffs_act["bebes"] +
            n_du * coeffs_act["day_use"] +
            n_grat * coeffs_act["gratuits"], 2
        )
        rec_du_prev = round(n_du * n_tdu, 3)

        kp1, kp2, kp3 = st.columns(3)
        kp1.metric("Total Physique (Brut)", f"{tot_brut_prev} pers.")
        kp2.metric("Total Pondéré F&B (Éq. Adulte)", f"{tot_pond_prev:,.2f}")
        kp3.metric("Recette Day Use du jour", f"{rec_du_prev:,.3f} DT")

        bs1, bs2 = st.columns(2)
        if bs1.button("💾 Enregistrer la journée", type="primary", key="btn_save_nuit"):
            ok, msg = db.sauvegarder_nuitee_jour(date_saisie_n, n_pay, n_enf, n_beb, n_du, n_tdu, n_grat, obs_n)
            st.success(msg)
            st.rerun()
        if existant and bs2.button("🗑️ Supprimer cette journée", key="btn_del_nuit"):
            ok, msg = db.supprimer_nuitee_jour(date_saisie_n)
            st.warning(msg)
            st.rerun()

    # ---------------- ONGLET 2 : RÉCAPITULATIF ----------------
    with ntab2:
        st.subheader("Suivi & Récapitulatif des Nuitées")
        r1, r2 = st.columns(2)
        d_deb_n = r1.date_input("Du :", value=date(date.today().year, date.today().month, 1), key="rec_n_d1")
        d_fin_n = r2.date_input("Au :", value=date.today(), key="rec_n_d2")

        synthese_n, lignes_n = db.lister_nuitees_periode(d_deb_n, d_fin_n)
        if not lignes_n:
            st.info("Aucune nuitée enregistrée sur cette période.")
        else:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Payants", f"{synthese_n['payants']:,}")
            m2.metric("Enfants / Bébés", f"{synthese_n['enfants']:,} / {synthese_n['bebes']:,}")
            m3.metric("Day Use (Recette)", f"{synthese_n['day_use']:,} ({synthese_n['recette_day_use']:,.3f} DT)")
            m4.metric("Visiteurs / Complimentary", f"{synthese_n['gratuits']:,}")

            m5, m6 = st.columns(2)
            m5.metric("Total Nuitées Brutes", f"{synthese_n['total_brut']:,.0f}")
            m6.metric("Total Nuitées Pondérées F&B", f"{synthese_n['total_pondere']:,.2f}")

            df_n = pd.DataFrame(lignes_n)
            df_n_aff = df_n[["date", "payants", "enfants", "bebes", "day_use", "tarif_du", "recette_du", "gratuits", "total_brut", "total_pondere", "observation"]].copy()
            df_n_aff.columns = ["Date", "Payants", "Enfants", "Bébés", "Day Use", "Tarif DU (DT)", "Recette DU (DT)", "Gratuits/Compl.", "Total Brut", "Total Pondéré F&B", "Observation"]
            st.dataframe(df_n_aff, use_container_width=True, hide_index=True, height=400)

            entetes_n = list(df_n_aff.columns)
            lignes_exp_n = [[f"{v:,.3f}" if isinstance(v, float) else str(v) for v in row] for row in df_n_aff.values.tolist()]
            bn1, bn2 = st.columns(2)
            bn1.download_button("📄 Exporter PDF", data=generer_pdf_memoire(f"Nuitees_{d_deb_n}_au_{d_fin_n}", entetes_n, lignes_exp_n, paysage=True),
                                file_name=f"Nuitees_{d_deb_n}_au_{d_fin_n}.pdf", mime="application/pdf")
            bn2.download_button("📊 Exporter Excel", data=generer_excel_memoire(f"Nuitees_{d_deb_n}_au_{d_fin_n}", entetes_n, lignes_exp_n),
                                file_name=f"Nuitees_{d_deb_n}_au_{d_fin_n}.xlsx")

    # ---------------- ONGLET 3 : COEFFICIENTS F&B ----------------
    with ntab3:
        st.subheader("Paramétrage des Coefficients F&B (Équivalent Adulte)")
        st.info("Définissez le poids de consommation F&B de chaque catégorie pour le calcul du Coût / Nuitée.")
        cf1, cf2, cf3, cf4, cf5 = st.columns(5)
        c_pay = cf1.number_input("Coeff. Payants :", min_value=0.0, max_value=2.0, value=float(coeffs_act["payants"]), step=0.05, key="cf_pay")
        c_enf = cf2.number_input("Coeff. Enfants :", min_value=0.0, max_value=2.0, value=float(coeffs_act["enfants"]), step=0.05, key="cf_enf")
        c_beb = cf3.number_input("Coeff. Bébés :", min_value=0.0, max_value=2.0, value=float(coeffs_act["bebes"]), step=0.05, key="cf_beb")
        c_du = cf4.number_input("Coeff. Day Use :", min_value=0.0, max_value=2.0, value=float(coeffs_act["day_use"]), step=0.05, key="cf_du")
        c_gr = cf5.number_input("Coeff. Gratuits/Compl. :", min_value=0.0, max_value=2.0, value=float(coeffs_act["gratuits"]), step=0.05, key="cf_gr")

        if st.button("💾 Enregistrer les Coefficients", type="primary", key="btn_save_coeffs"):
            ok, msg = db.sauvegarder_coefficients_nuitees({
                "payants": c_pay, "enfants": c_enf, "bebes": c_beb, "day_use": c_du, "gratuits": c_gr
            })
            st.success(msg)
            st.rerun()
