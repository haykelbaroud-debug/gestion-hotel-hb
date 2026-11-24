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
# ==============================================================================
# MODULE 1 : PARAMÉTRAGE DE BASE (CLASSES, FAMILLES, SOUS-FAMILLES)
# ==============================================================================

def initialiser_parametrage():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS classes (
            code TEXT PRIMARY KEY,
            libelle TEXT NOT NULL,
            actif INTEGER DEFAULT 1
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS familles (
            code TEXT PRIMARY KEY,
            libelle TEXT NOT NULL,
            code_classe TEXT NOT NULL,
            actif INTEGER DEFAULT 1,
            FOREIGN KEY (code_classe) REFERENCES classes(code)
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS sous_familles (
            code TEXT PRIMARY KEY,
            libelle TEXT NOT NULL,
            code_famille TEXT NOT NULL,
            actif INTEGER DEFAULT 1,
            FOREIGN KEY (code_famille) REFERENCES familles(code)
        )""")
        conn.commit()
        c.close()
    finally:
        conn.close()

def lister_classes(seulement_actives=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        if seulement_actives:
            c.execute("SELECT code, libelle, actif FROM classes WHERE actif = 1 ORDER BY code")
        else:
            c.execute("SELECT code, libelle, actif FROM classes ORDER BY code")
        return c.fetchall()
    finally:
        conn.close()

def sauvegarder_classe(code, libelle, actif=1, creation=True):
    conn = get_connection()
    try:
        c = conn.cursor()
        code_p = code.strip().upper()
        lib_p = libelle.strip()
        if creation:
            c.execute("SELECT 1 FROM classes WHERE code = ?", (code_p,))
            if c.fetchone():
                return False, f"Le code classe '{code_p}' existe déjà."
            c.execute("INSERT INTO classes (code, libelle, actif) VALUES (?, ?, ?)", (code_p, lib_p, int(actif)))
        else:
            c.execute("UPDATE classes SET libelle = ?, actif = ? WHERE code = ?", (lib_p, int(actif), code_p))
        conn.commit()
        return True, "Classe enregistrée avec succès."
    finally:
        conn.close()

def supprimer_classe(code):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM familles WHERE code_classe = ?", (code,))
        if itemgetter(0)(c.fetchone()) > 0:
            return False, "Suppression impossible : des familles sont rattachées à cette classe."
        c.execute("DELETE FROM classes WHERE code = ?", (code,))
        conn.commit()
        return True, "Classe supprimée."
    finally:
        conn.close()

def lister_familles(seulement_actives=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = """
            SELECT f.code, f.libelle, f.code_classe, COALESCE(cl.libelle, ''), f.actif
            FROM familles f
            LEFT JOIN classes cl ON f.code_classe = cl.code
        """
        if seulement_actives:
            sql += " WHERE f.actif = 1"
        sql += " ORDER BY f.code_classe, f.code"
        c.execute(sql)
        return c.fetchall()
    finally:
        conn.close()

def sauvegarder_famille(code, libelle, code_classe, actif=1, creation=True):
    conn = get_connection()
    try:
        c = conn.cursor()
        code_p = code.strip().upper()
        lib_p = libelle.strip()
        if creation:
            c.execute("SELECT 1 FROM familles WHERE code = ?", (code_p,))
            if c.fetchone():
                return False, f"Le code famille '{code_p}' existe déjà."
            c.execute("INSERT INTO familles (code, libelle, code_classe, actif) VALUES (?, ?, ?, ?)",
                      (code_p, lib_p, code_classe, int(actif)))
        else:
            c.execute("UPDATE familles SET libelle = ?, code_classe = ?, actif = ? WHERE code = ?",
                      (lib_p, code_classe, int(actif), code_p))
        conn.commit()
        return True, "Famille enregistrée avec succès."
    finally:
        conn.close()

def supprimer_famille(code):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM sous_familles WHERE code_famille = ?", (code,))
        if itemgetter(0)(c.fetchone()) > 0:
            return False, "Suppression impossible : des sous-familles sont rattachées à cette famille."
        c.execute("DELETE FROM familles WHERE code = ?", (code,))
        conn.commit()
        return True, "Famille supprimée."
    finally:
        conn.close()

def lister_sous_familles(seulement_actives=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = """
            SELECT sf.code, sf.libelle, sf.code_famille, COALESCE(f.libelle, ''),
                   COALESCE(f.code_classe, ''), COALESCE(cl.libelle, ''), sf.actif
            FROM sous_familles sf
            LEFT JOIN familles f ON sf.code_famille = f.code
            LEFT JOIN classes cl ON f.code_classe = cl.code
        """
        if seulement_actives:
            sql += " WHERE sf.actif = 1"
        sql += " ORDER BY f.code_classe, sf.code_famille, sf.code"
        c.execute(sql)
        return c.fetchall()
    finally:
        conn.close()

def sauvegarder_sous_famille(code, libelle, code_famille, actif=1, creation=True):
    conn = get_connection()
    try:
        c = conn.cursor()
        code_p = code.strip().upper()
        lib_p = libelle.strip()
        if creation:
            c.execute("SELECT 1 FROM sous_familles WHERE code = ?", (code_p,))
            if c.fetchone():
                return False, f"Le code sous-famille '{code_p}' existe déjà."
            c.execute("INSERT INTO sous_familles (code, libelle, code_famille, actif) VALUES (?, ?, ?, ?)",
                      (code_p, lib_p, code_famille, int(actif)))
        else:
            c.execute("UPDATE sous_familles SET libelle = ?, code_famille = ?, actif = ? WHERE code = ?",
                      (lib_p, code_famille, int(actif), code_p))
        conn.commit()
        return True, "Sous-famille enregistrée avec succès."
    finally:
        conn.close()

def supprimer_sous_famille(code):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM sous_familles WHERE code = ?", (code,))
        conn.commit()
        return True, "Sous-famille supprimée."
    finally:
        conn.close()
# ==============================================================================
# MODULE 2 : BASE DE DONNÉES (ARTICLES, FOURNISSEURS, DÉPARTEMENTS)
# ==============================================================================

def initialiser_referentiels_m2():
    conn = get_connection()
    try:
        c = conn.cursor()
        # Mise à niveau douce des tables existantes (articles & departements)
        for col_sql in [
            "ALTER TABLE articles ADD COLUMN unite TEXT DEFAULT 'Unité'",
            "ALTER TABLE articles ADD COLUMN prix_achat REAL DEFAULT 0.0",
            "ALTER TABLE articles ADD COLUMN tva REAL DEFAULT 19.0",
            "ALTER TABLE articles ADD COLUMN actif INTEGER DEFAULT 1",
            "ALTER TABLE departements ADD COLUMN actif INTEGER DEFAULT 1"
        ]:
            try:
                c.execute(col_sql)
            except Exception:
                pass

        c.execute("""CREATE TABLE IF NOT EXISTS fournisseurs (
            code TEXT PRIMARY KEY,
            nom TEXT NOT NULL,
            telephone TEXT DEFAULT '',
            matricule_fiscal TEXT DEFAULT '',
            email TEXT DEFAULT '',
            actif INTEGER DEFAULT 1
        )""")
        conn.commit()
        c.close()
    finally:
        conn.close()

# --- ARTICLES ---
def lister_articles_complet(seulement_actifs=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = """
            SELECT a.code, a.designation, COALESCE(a.sous_famille, ''),
                   COALESCE(sf.libelle, ''), COALESCE(f.code, ''), COALESCE(cl.code, ''),
                   COALESCE(a.unite, 'Unité'), COALESCE(a.prix_achat, 0.0),
                   COALESCE(a.tva, 19.0), COALESCE(a.actif, 1)
            FROM articles a
            LEFT JOIN sous_familles sf ON UPPER(a.sous_famille) = UPPER(sf.code)
            LEFT JOIN familles f ON sf.code_famille = f.code
            LEFT JOIN classes cl ON f.code_classe = cl.code
        """
        if seulement_actifs:
            sql += " WHERE COALESCE(a.actif, 1) = 1"
        sql += " ORDER BY a.code"
        c.execute(sql)
        return c.fetchall()
    finally:
        conn.close()

def sauvegarder_article(code, designation, sous_famille, unite, prix_achat, tva, actif=1, creation=True):
    conn = get_connection()
    try:
        c = conn.cursor()
        code_p = code.strip().upper()
        des_p = designation.strip()
        # Récupère automatiquement la classe si la sous-famille est paramétrée
        c.execute("""
            SELECT COALESCE(f.code_classe, '')
            FROM sous_familles sf
            LEFT JOIN familles f ON sf.code_famille = f.code
            WHERE UPPER(sf.code) = UPPER(?)
        """, (sous_famille,))
        res_cl = c.fetchone()
        classe_p = itemgetter(0)(res_cl) if res_cl else ""

        if creation:
            c.execute("SELECT 1 FROM articles WHERE UPPER(code) = ?", (code_p,))
            if c.fetchone():
                return False, f"Le code article '{code_p}' existe déjà."
            c.execute("""INSERT INTO articles (code, designation, classe, sous_famille, unite, prix_achat, tva, actif)
                         VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                      (code_p, des_p, classe_p, sous_famille, unite, float(prix_achat), float(tva), int(actif)))
        else:
            c.execute("""UPDATE articles
                         SET designation = ?, classe = ?, sous_famille = ?, unite = ?, prix_achat = ?, tva = ?, actif = ?
                         WHERE code = ?""",
                      (des_p, classe_p, sous_famille, unite, float(prix_achat), float(tva), int(actif), code))
        conn.commit()
        return True, "Article enregistré avec succès."
    finally:
        conn.close()

def supprimer_article(code):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM achats WHERE UPPER(code_article) = UPPER(?)", (code,))
        if itemgetter(0)(c.fetchone()) > 0:
            return False, "Suppression impossible : cet article possède déjà un historique d'achats (désactivez-le plutôt)."
        c.execute("DELETE FROM articles WHERE code = ?", (code,))
        conn.commit()
        return True, "Article supprimé."
    finally:
        conn.close()

# --- FOURNISSEURS ---
def lister_fournisseurs(seulement_actifs=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = "SELECT code, nom, COALESCE(telephone, ''), COALESCE(matricule_fiscal, ''), COALESCE(email, ''), actif FROM fournisseurs"
        if seulement_actifs:
            sql += " WHERE actif = 1"
        sql += " ORDER BY nom"
        c.execute(sql)
        return c.fetchall()
    finally:
        conn.close()

def sauvegarder_fournisseur(code, nom, telephone, matricule_fiscal, email, actif=1, creation=True):
    conn = get_connection()
    try:
        c = conn.cursor()
        code_p = code.strip().upper()
        nom_p = nom.strip()
        if creation:
            c.execute("SELECT 1 FROM fournisseurs WHERE code = ?", (code_p,))
            if c.fetchone():
                return False, f"Le code fournisseur '{code_p}' existe déjà."
            c.execute("INSERT INTO fournisseurs (code, nom, telephone, matricule_fiscal, email, actif) VALUES (?, ?, ?, ?, ?, ?)",
                      (code_p, nom_p, telephone.strip(), matricule_fiscal.strip(), email.strip(), int(actif)))
        else:
            c.execute("""UPDATE fournisseurs
                         SET nom = ?, telephone = ?, matricule_fiscal = ?, email = ?, actif = ?
                         WHERE code = ?""",
                      (nom_p, telephone.strip(), matricule_fiscal.strip(), email.strip(), int(actif), code))
        conn.commit()
        return True, "Fournisseur enregistré avec succès."
    finally:
        conn.close()

def supprimer_fournisseur(code):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM fournisseurs WHERE code = ?", (code,))
        conn.commit()
        return True, "Fournisseur supprimé."
    finally:
        conn.close()

# --- DÉPARTEMENTS ---
def lister_departements_complet(seulement_actifs=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = "SELECT code_court, nom_long, COALESCE(actif, 1) FROM departements"
        if seulement_actifs:
            sql += " WHERE COALESCE(actif, 1) = 1"
        sql += " ORDER BY code_court"
        c.execute(sql)
        return c.fetchall()
    finally:
        conn.close()

def sauvegarder_departement(code_court, nom_long, actif=1, creation=True):
    conn = get_connection()
    try:
        c = conn.cursor()
        code_p = code_court.strip().upper()
        nom_p = nom_long.strip()
        if creation:
            c.execute("SELECT 1 FROM departements WHERE UPPER(code_court) = ?", (code_p,))
            if c.fetchone():
                return False, f"Le code département '{code_p}' existe déjà."
            c.execute("INSERT INTO departements (code_court, nom_long, actif) VALUES (?, ?, ?)",
                      (code_p, nom_p, int(actif)))
        else:
            c.execute("UPDATE departements SET nom_long = ?, actif = ? WHERE code_court = ?",
                      (nom_p, int(actif), code_court))
        conn.commit()
        return True, "Département enregistré avec succès."
    finally:
        conn.close()

def supprimer_departement(code_court):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM transferts WHERE UPPER(code_dept) = UPPER(?)", (code_court,))
        if itemgetter(0)(c.fetchone()) > 0:
            return False, "Suppression impossible : ce département possède déjà des consommations enregistrées."
        c.execute("DELETE FROM departements WHERE code_court = ?", (code_court,))
        conn.commit()
        return True, "Département supprimé."
    finally:
        conn.close()
# ==============================================================================
# MODULE 3 : ACHATS & MOUVEMENTS DE STOCK (BC, BR, RETOUR, ENTRÉE/SORTIE DIRECTE)
# ==============================================================================

def initialiser_achats_m3():
    conn = get_connection()
    try:
        c = conn.cursor()
        # Ajout du droit de réception d'achat sur les départements
        try:
            c.execute("ALTER TABLE departements ADD COLUMN autorise_achat INTEGER DEFAULT 0")
        except Exception:
            pass

        # Création automatique du département Économat (ECO) autorisé aux achats par défaut
        c.execute("""
            INSERT OR IGNORE INTO departements (code_court, nom_long, actif, autorise_achat)
            VALUES ('ECO', 'Économat Central', 1, 1)
        """)
        c.execute("UPDATE departements SET autorise_achat = 1 WHERE UPPER(code_court) = 'ECO'")

        # Entêtes des documents d'achats et mouvements
        c.execute("""CREATE TABLE IF NOT EXISTS doc_achats (
            numero TEXT PRIMARY KEY,
            type_doc TEXT NOT NULL,
            date_doc TEXT NOT NULL,
            code_fournisseur TEXT DEFAULT '',
            code_dept TEXT DEFAULT 'ECO',
            ref_externe TEXT DEFAULT '',
            observation TEXT DEFAULT '',
            statut TEXT DEFAULT 'VALIDÉ',
            total_ht REAL DEFAULT 0.0,
            total_ttc REAL DEFAULT 0.0
        )""")

        # Lignes d'articles des documents
        c.execute("""CREATE TABLE IF NOT EXISTS doc_achats_lignes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_doc TEXT NOT NULL,
            code_article TEXT NOT NULL,
            designation TEXT NOT NULL,
            unite TEXT DEFAULT 'Unité',
            qte REAL NOT NULL,
            prix_unitaire_ht REAL NOT NULL,
            tva REAL DEFAULT 19.0,
            montant_ht REAL NOT NULL,
            montant_ttc REAL NOT NULL,
            FOREIGN KEY (numero_doc) REFERENCES doc_achats(numero)
        )""")
        conn.commit()
        c.close()
    finally:
        conn.close()

def basculer_autorisation_achat_dept(code_court, autorise):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE departements SET autorise_achat = ? WHERE code_court = ?", (int(autorise), code_court))
        conn.commit()
    finally:
        conn.close()

def lister_depts_achats():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT code_court, nom_long
            FROM departements
            WHERE COALESCE(actif, 1) = 1 AND COALESCE(autorise_achat, 0) = 1
            ORDER BY CASE WHEN code_court = 'ECO' THEN 0 ELSE 1 END, code_court
        """)
        return c.fetchall()
    finally:
        conn.close()

def generer_numero_doc(prefixe, annee="2026"):
    conn = get_connection()
    try:
        c = conn.cursor()
        motif = f"{prefixe}-{annee}-%"
        c.execute("SELECT numero FROM doc_achats WHERE numero LIKE ? ORDER BY numero DESC LIMIT 1", (motif,))
        res = c.fetchone()
        if res:
            dernier_num = itemgetter(0)(res)
            try:
                seq = int(dernier_num.split("-").pop()) + 1
            except Exception:
                seq = 1
        else:
            seq = 1
        return f"{prefixe}-{annee}-{seq:04d}"
    finally:
        conn.close()

def enregistrer_document_achat(numero, type_doc, date_doc, code_fourn, code_dept, ref_ext, obs, lignes, creation=True):
    """
    lignes : liste de tuples (code_article, designation, unite, qte, prix_ht, tva)
    """
    if not lignes:
        return False, "Veuillez ajouter au moins une ligne d'article."
    conn = get_connection()
    try:
        c = conn.cursor()
        tot_ht = 0.0
        tot_ttc = 0.0
        lignes_calc = []
        for code_a, des_a, u_a, qte, pu_ht, tva in lignes:
            m_ht = round(float(qte) * float(pu_ht), 3)
            m_ttc = round(m_ht * (1.0 + float(tva) / 100.0), 3)
            tot_ht += m_ht
            tot_ttc += m_ttc
            lignes_calc.append((code_a, des_a, u_a, float(qte), float(pu_ht), float(tva), m_ht, m_ttc))

        if creation:
            c.execute("""INSERT INTO doc_achats
                (numero, type_doc, date_doc, code_fournisseur, code_dept, ref_externe, observation, statut, total_ht, total_ttc)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'VALIDÉ', ?, ?)""",
                (numero, type_doc, str(date_doc), code_fourn, code_dept, ref_ext.strip(), obs.strip(), round(tot_ht, 3), round(tot_ttc, 3)))
        else:
            c.execute("""UPDATE doc_achats
                SET date_doc = ?, code_fournisseur = ?, code_dept = ?, ref_externe = ?, observation = ?, total_ht = ?, total_ttc = ?
                WHERE numero = ?""",
                (str(date_doc), code_fourn, code_dept, ref_ext.strip(), obs.strip(), round(tot_ht, 3), round(tot_ttc, 3), numero))
            c.execute("DELETE FROM doc_achats_lignes WHERE numero_doc = ?", (numero,))

        for code_a, des_a, u_a, qte, pu_ht, tva, m_ht, m_ttc in lignes_calc:
            c.execute("""INSERT INTO doc_achats_lignes
                (numero_doc, code_article, designation, unite, qte, prix_unitaire_ht, tva, montant_ht, montant_ttc)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (numero, code_a, des_a, u_a, qte, pu_ht, tva, m_ht, m_ttc))
            # Mise à jour automatique du dernier prix d'achat lors d'un Bon de Réception (BR) ou Entrée Directe (ED)
            if type_doc in ("BR", "ED") and pu_ht > 0:
                c.execute("UPDATE articles SET prix_achat = ? WHERE code = ?", (pu_ht, code_a))

        conn.commit()
        return True, f"Document {numero} enregistré avec succès ({tot_ttc:,.3f} DT TTC)."
    finally:
        conn.close()

def lister_documents_achat(type_doc):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT d.numero, d.date_doc, COALESCE(f.nom, d.code_fournisseur), d.code_dept,
                   d.ref_externe, d.total_ht, d.total_ttc, d.observation, d.code_fournisseur
            FROM doc_achats d
            LEFT JOIN fournisseurs f ON d.code_fournisseur = f.code
            WHERE d.type_doc = ?
            ORDER BY d.date_doc DESC, d.numero DESC
        """, (type_doc,))
        return c.fetchall()
    finally:
        conn.close()

def lire_lignes_document(numero_doc):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT code_article, designation, unite, qte, prix_unitaire_ht, tva, montant_ht, montant_ttc
            FROM doc_achats_lignes
            WHERE numero_doc = ?
            ORDER BY id
        """, (numero_doc,))
        return c.fetchall()
    finally:
        conn.close()

def supprimer_document_achat(numero_doc):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM doc_achats_lignes WHERE numero_doc = ?", (numero_doc,))
        c.execute("DELETE FROM doc_achats WHERE numero = ?", (numero_doc,))
        conn.commit()
        return True, f"Document {numero_doc} supprimé."
    finally:
        conn.close()
