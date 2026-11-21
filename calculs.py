from collections import defaultdict
from datetime import datetime
import os
from operator import itemgetter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from database import get_connection, LISTE_MOIS

def idx_val(seq, i):
    return itemgetter(i)(seq)

def calculer_strategie_mois(mois):
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT sous_famille, fournisseur, SUM(montant_ht) FROM achats WHERE mois = ? GROUP BY sous_famille, fournisseur", (mois,))
    rows = c.fetchall()
    conn.close()

    totaux_sf = defaultdict(float)
    achats_sf_fourn = defaultdict(dict)
    total_global = 0.0

    for sf, fourn, mnt in rows:
        totaux_sf[sf] += mnt
        achats_sf_fourn[sf][fourn] = mnt
        total_global += mnt

    resultats = []
    somme_verif = 0.0
    for sf in sorted(achats_sf_fourn.keys()):
        tot_sf = totaux_sf[sf]
        resultats.append({"type": "HEADER", "col1": f"{sf}  |  Total Général", "col2": tot_sf, "col3": 1.0, "col4": ""})
        for fourn, mnt in sorted(achats_sf_fourn[sf].items(), key=lambda x: idx_val(x, 1), reverse=True):
            part = (mnt / tot_sf) if tot_sf > 0 else 0.0
            statut = "Alerte (>30%)" if part > 0.30 else "Conforme"
            somme_verif += mnt
            resultats.append({"type": "ROW", "col1": fourn, "col2": mnt, "col3": part, "col4": statut})

    return {"total_achats": somme_verif, "total_controle": total_global, "ecart": somme_verif - total_global, "lignes": resultats}

def calculer_recap_annuel_achats():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT mois, sous_famille, fournisseur, SUM(montant_ht) FROM achats GROUP BY mois, sous_famille, fournisseur")
    rows = c.fetchall()
    conn.close()

    suivi_mensuel = defaultdict(lambda: [0.0]*12)
    recap_sf = defaultdict(lambda: defaultdict(float))
    totaux_sf = defaultdict(float)

    for mois, sf, fourn, mnt in rows:
        if mois in LISTE_MOIS:
            suivi_mensuel[(sf, fourn)][LISTE_MOIS.index(mois)] += mnt
        recap_sf[sf][fourn] += mnt
        totaux_sf[sf] += mnt

    return suivi_mensuel, recap_sf, totaux_sf

def calculer_consommation_departement(mois, code_dept=None):
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT nb_nuitees FROM nuitees WHERE mois = ?", (mois,))
    res_nuit = c.fetchone()
    nuitees = float(idx_val(res_nuit, 0)) if res_nuit and idx_val(res_nuit, 0) else 0.0

    query = """
        SELECT COALESCE(nom_dept, code_dept), COALESCE(sous_famille, 'DIVERS'), designation,
               SUM(qte_conso), SUM(val_conso)
        FROM transferts
        WHERE mois = ?
    """
    params = [mois]
    if code_dept:
        query += " AND (UPPER(code_dept) = UPPER(?) OR UPPER(nom_dept) = UPPER(?))"
        params.extend([code_dept, code_dept])
    query += " GROUP BY nom_dept, code_dept, sous_famille, designation"

    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    arbre = defaultdict(lambda: defaultdict(list))
    for nom_d, sf, desig, qte, val in rows:
        q_nuit = round(qte / nuitees, 3) if nuitees > 0 else 0.0
        v_nuit = round(val / nuitees, 3) if nuitees > 0 else 0.0
        arbre[nom_d][sf].append({"article": desig, "qte": qte, "val": val, "qte_nuitee": q_nuit, "val_nuitee": v_nuit})
    return nuitees, arbre

def calculer_analyse_article_ou_sf(cible, mode_scope="ARTICLE", mode_valeur="QUANTITE"):
    conn = get_connection()
    c = conn.cursor()
    col_sql = "SUM(val_conso)" if mode_valeur == "VALEUR" else "SUM(qte_conso)"
    if mode_scope == "SOUSFAM":
        c.execute(f"SELECT COALESCE(nom_dept, code_dept), mois, {col_sql} FROM transferts WHERE UPPER(sous_famille) = UPPER(?) GROUP BY nom_dept, code_dept, mois", (cible,))
    else:
        c.execute(f"SELECT COALESCE(nom_dept, code_dept), mois, {col_sql} FROM transferts WHERE UPPER(designation) = UPPER(?) GROUP BY nom_dept, code_dept, mois", (cible,))
    rows = c.fetchall()
    conn.close()

    matrice = defaultdict(lambda: [0.0]*12)
    for dept, mois, val in rows:
        if mois in LISTE_MOIS and val != 0:
            matrice[dept][LISTE_MOIS.index(mois)] += val
    return matrice

def generer_donnees_executives():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT fournisseur, SUM(montant_ht) FROM achats GROUP BY fournisseur ORDER BY SUM(montant_ht) DESC")
    fourn_rows = c.fetchall()
    total_achats = sum(idx_val(r, 1) for r in fourn_rows)
    top5_fourn = fourn_rows[:5]

    c.execute("SELECT COALESCE(nom_dept, code_dept), SUM(val_conso) FROM transferts GROUP BY nom_dept, code_dept ORDER BY SUM(val_conso) DESC")
    dept_rows = c.fetchall()
    total_conso = sum(idx_val(r, 1) for r in dept_rows)
    top5_dept = dept_rows[:5]

    c.execute("SELECT SUM(nb_nuitees) FROM nuitees")
    res_n = c.fetchone()
    total_nuitees = float(idx_val(res_n, 0)) if res_n and idx_val(res_n, 0) else 0.0
    conn.close()
    return {"total_achats": total_achats, "top5_fourn": top5_fourn, "total_conso": total_conso, "total_nuitees": total_nuitees, "top5_dept": top5_dept}

def exporter_tableau_pdf(titre_rapport, entetes, lignes_donnees, orientation="PORTRAIT"):
    dossier = os.path.join(os.getcwd(), "Rapports_PDF")
    os.makedirs(dossier, exist_ok=True)
    horodatage = datetime.now().strftime("%Y-%m-%d_%Hh%M")
    nom_propre = "".join(ch for ch in titre_rapport if ch.isalnum() or ch in (" ", "_", "-")).replace(" ", "_")
    chemin_pdf = os.path.join(dossier, f"{nom_propre}_{horodatage}.pdf")

    page_format = landscape(A4) if orientation == "PAYSAGE" else A4
    doc = SimpleDocTemplate(chemin_pdf, pagesize=page_format, rightMargin=25, leftMargin=25, topMargin=30, bottomMargin=25)
    styles = getSampleStyleSheet()
    elements = [Paragraph(f"<b>{titre_rapport}</b>", styles["Title"]), Spacer(1, 12)]

    t = Table([entetes] + lignes_donnees, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A56")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DCE1")),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")])
    ]))
    elements.append(t)
    doc.build(elements)
    os.startfile(chemin_pdf)
    return chemin_pdf
def exporter_tableau_excel(titre_rapport, entetes, lignes_donnees):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    dossier = os.path.join(os.getcwd(), "Exports_Excel")
    os.makedirs(dossier, exist_ok=True)
    horodatage = datetime.now().strftime("%Y-%m-%d_%Hh%M")
    nom_propre = "".join(ch for ch in titre_rapport if ch.isalnum() or ch in (" ", "_", "-")).replace(" ", "_")
    chemin_xlsx = os.path.join(dossier, f"{nom_propre}_{horodatage}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rapport"

    ws.append([titre_rapport])
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

    for row in lignes_donnees:
        ws.append(row)

    wb.save(chemin_xlsx)
    os.startfile(chemin_xlsx)
    return chemin_xlsx

