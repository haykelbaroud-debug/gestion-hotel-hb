import sqlite3
import openpyxl
from operator import itemgetter

DB_NAME = "hotel_erp_pro.db"

LISTE_MOIS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]

MOIS_VARIANTS = {
    "Janvier": ["Janvier", "Janv"],
    "Février": ["Février", "Fevrier", "Fev"],
    "Mars": ["Mars"],
    "Avril": ["Avril"],
    "Mai": ["Mai"],
    "Juin": ["Juin"],
    "Juillet": ["Juillet"],
    "Août": ["Août", "Aout"],
    "Septembre": ["Septembre", "Sept"],
    "Octobre": ["Octobre", "Oct"],
    "Novembre": ["Novembre", "Nov"],
    "Décembre": ["Décembre", "Decembre", "Dec"]
}

FAMILLES_VALIDES = {
    "SOFT", "WATR", "BEER", "WHIW", "ROSW", "REDW", "INTS", "NATS",
    "MEAT", "FISH", "POUL", "SAUS", "DAIR", "EGGS", "THEC", "MIXP",
    "BRKF", "PATI", "ICEC", "VFRE", "VFRO", "VCAN", "FFRE", "FFRO",
    "FCAN", "DRYP", "PAST", "BRED", "OILF", "SAUC", "IMPRIMES", "SPAW",
    "FRTS GUESTS", "ORAN", "FRTS BUREAUX", "PRODUITS BUANDERIE",
    "PRODUITS D'HYG & NET", "MATERIELS D'HYG & NE", "CARTOUCHES,TON&RUBAN",
    "MAINTENANCE", "EXTRA ARTICLES", "PRODUITS PISCINES"
}

def get_col(row, num_col):
    if row is None or len(row) <= num_col:
        return None
    return itemgetter(num_col)(row)

def nettoyer_nombre(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace("DT", "").replace("\xa0", "").replace(" ", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0

def get_connection():
    return sqlite3.connect(DB_NAME, timeout=30, check_same_thread=False)

def initialiser_base():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS articles (
            code TEXT PRIMARY KEY, designation TEXT, classe TEXT, sous_famille TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS departements (
            code_court TEXT PRIMARY KEY, nom_long TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS nuitees (
            mois TEXT PRIMARY KEY, nb_nuitees REAL DEFAULT 0
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS achats (
            id INTEGER PRIMARY KEY AUTOINCREMENT, mois TEXT, code_article TEXT,
            sous_famille TEXT, fournisseur TEXT, montant_ht REAL
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS transferts (
            id INTEGER PRIMARY KEY AUTOINCREMENT, mois TEXT, code_dept TEXT,
            nom_dept TEXT, sous_famille TEXT, code_article TEXT, designation TEXT,
            qte_conso REAL, val_conso REAL
        )""")
        for m in LISTE_MOIS:
            c.execute("INSERT OR IGNORE INTO nuitees (mois, nb_nuitees) VALUES (?, 0)", (m,))
        conn.commit()
        c.close()
    finally:
        conn.close()
def extraire_sous_famille_achat(code_article):
    chaine_verif = code_article.upper().strip()
    mots_longs = ["BUREAU", "HYG", "BUANDERIE", "GUESTS", "CARTOUCHES"]
    if any(mot in chaine_verif for mot in mots_longs):
        sf = code_article.strip()
    else:
        sf = code_article.split(" ").pop(0) if " " in code_article else code_article
    return sf.upper().replace("\xa0", " ").strip()

def trouver_feuille(wb, prefixe, mois):
    for var in MOIS_VARIANTS.get(mois, [mois]):
        cible = f"{prefixe}{var}".upper().replace(" ", "")
        for nom_reel in wb.sheetnames:
            if nom_reel.upper().replace(" ", "") == cible:
                return wb[nom_reel]
    return None

def importer_depuis_excel(chemin_excel):
    wb = openpyxl.load_workbook(chemin_excel, data_only=True, read_only=True)
    conn = get_connection()
    try:
        c = conn.cursor()

        if "DB_Données" in wb.sheetnames:
            ws = wb["DB_Données"]
            c.execute("DELETE FROM articles")
            for row in ws.iter_rows(min_row=5, values_only=True):
                c0 = get_col(row, 0)
                if c0 is not None and str(c0).strip() != "":
                    code = str(c0).strip()
                    if code.endswith(".0"):
                        code = code[:-2]
                    desig = str(get_col(row, 1) or "").strip()
                    classe = str(get_col(row, 2) or "").strip()
                    sf = str(get_col(row, 4) or "").strip()
                    c.execute("INSERT OR REPLACE INTO articles VALUES (?, ?, ?, ?)", (code, desig, classe, sf))

        nom_to_code = {}
        if "DB_Départements" in wb.sheetnames:
            ws = wb["DB_Départements"]
            c.execute("DELETE FROM departements")
            for row in ws.iter_rows(min_row=5, values_only=True):
                c0 = get_col(row, 0)
                c3 = get_col(row, 3)
                if c0 is not None and c3 is not None:
                    nom_long = str(c0).strip()
                    code_court = str(c3).upper().strip()
                    if nom_long != "" and code_court != "" and code_court != "CODE SAISIE (COURT)":
                        c.execute("INSERT OR REPLACE INTO departements VALUES (?, ?)", (code_court, nom_long))
                        nom_to_code[nom_long.upper()] = code_court

        c.execute("DELETE FROM achats")
        c.execute("DELETE FROM transferts")

        nb_achats, nb_conso = 0, 0
        for mois in LISTE_MOIS:
            ws_achat = trouver_feuille(wb, "DB_Achats_", mois)
            if ws_achat is not None:
                for row in ws_achat.iter_rows(min_row=2, values_only=True):
                    c0, c2, c12 = get_col(row, 0), get_col(row, 2), get_col(row, 12)
                    if c0 is not None and c2 is not None and c12 is not None:
                        code_art = str(c0).strip()
                        fourn = str(c2).strip()
                        mnt = nettoyer_nombre(c12)
                        sf = extraire_sous_famille_achat(code_art)
                        if mnt != 0 and sf in FAMILLES_VALIDES and fourn != "":
                            c.execute(
                                "INSERT INTO achats (mois, code_article, sous_famille, fournisseur, montant_ht) VALUES (?, ?, ?, ?, ?)",
                                (mois, code_art, sf, fourn, mnt)
                            )
                            nb_achats += 1

            ws_trans = trouver_feuille(wb, "DB_Transferts_", mois)
            if ws_trans is not None:
                nuit = 0.0
                for idx, row in enumerate(ws_trans.iter_rows(min_row=6, max_row=6, values_only=True)):
                    nuit = nettoyer_nombre(get_col(row, 1))
                c.execute("UPDATE nuitees SET nb_nuitees = ? WHERE mois = ?", (nuit, mois))

            ws_indie = trouver_feuille(wb, "DB_Indicateurs_", mois)
            if ws_indie is not None:
                dept_nom, dept_code, sf_actuelle = "", "", "DIVERS"
                for row in ws_indie.iter_rows(min_row=6, values_only=True):
                    val_a = str(get_col(row, 0) or "").replace("\xa0", " ").strip()
                    if val_a == "":
                        continue
                    val_up = val_a.upper()
                    if "DÉPARTEMENT" in val_up or "DEPARTEMENT" in val_up:
                        if ":" in val_a:
                            dept_nom = val_a.split(":", 1).pop(1).strip()
                        else:
                            dept_nom = val_a.replace("DÉPARTEMENT", "").replace("DEPARTEMENT", "").strip()
                        dept_code = nom_to_code.get(dept_nom.upper(), dept_nom.upper())
                    elif "SOUS-FAMILLE" in val_up or "SOUS FAMILLE" in val_up:
                        sf_actuelle = val_a.split(":", 1).pop(1).strip() if ":" in val_a else val_a
                    elif not val_up.startswith("TOTAL") and not val_up.startswith("RAPPORT") and dept_nom != "":
                        qte = nettoyer_nombre(get_col(row, 1))
                        val = nettoyer_nombre(get_col(row, 2))
                        if qte != 0 or val != 0:
                            c.execute("""INSERT INTO transferts 
                                (mois, code_dept, nom_dept, sous_famille, code_article, designation, qte_conso, val_conso)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                                (mois, dept_code, dept_nom, sf_actuelle, "", val_a, qte, val))
                            nb_conso += 1

        conn.commit()
        c.close()
        return nb_achats, nb_conso
    finally:
        wb.close()
        conn.close()
