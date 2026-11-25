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

import json
import urllib.request

try:
    import streamlit as _st_turso
except Exception:
    _st_turso = None

def _get_turso_secrets():
    url, token = "", ""
    if _st_turso is not None:
        try:
            url = _st_turso.secrets.get("TURSO_DATABASE_URL", "")
            token = _st_turso.secrets.get("TURSO_AUTH_TOKEN", "")
        except Exception:
            pass
    if url.startswith("libsql://"):
        url = "https://" + url[len("libsql://"):]
    return url.rstrip("/"), token

class _TursoCursorWrapper:
    def __init__(self, local_cur, remote_exec_fn):
        self._cur = local_cur
        self._remote_exec = remote_exec_fn

    def execute(self, sql, params=()):
        res = self._cur.execute(sql, params)
        sql_up = sql.strip().upper()
        if sql_up.startswith(("INSERT", "UPDATE", "DELETE", "CREATE", "ALTER", "DROP", "REPLACE")):
            self._remote_exec(sql, params)
        return res

    def executemany(self, sql, seq_of_params):
        res = self._cur.executemany(sql, seq_of_params)
        for p in seq_of_params:
            self._remote_exec(sql, p)
        return res

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    @property
    def rowcount(self):
        return self._cur.rowcount

    def close(self):
        return self._cur.close()

class _TursoConnWrapper:
    def __init__(self, local_conn, url, token):
        self._conn = local_conn
        self._url = url
        self._token = token
        self._pending = []

    def _queue_remote(self, sql, params=()):
        args = []
        for p in (params or ()):
            if p is None:
                args.append({"type": "null"})
            elif isinstance(p, bool):
                args.append({"type": "integer", "value": "1" if p else "0"})
            elif isinstance(p, int):
                args.append({"type": "integer", "value": str(p)})
            elif isinstance(p, float):
                args.append({"type": "float", "value": p})
            else:
                args.append({"type": "text", "value": str(p)})
        self._pending.append({"type": "execute", "stmt": {"sql": sql, "args": args}})

    def cursor(self):
        return _TursoCursorWrapper(self._conn.cursor(), self._queue_remote)

    def execute(self, sql, params=()):
        cur = self.cursor()
        return cur.execute(sql, params)

    def commit(self):
        self._conn.commit()
        if self._pending and self._url and self._token:
            try:
                reqs = self._pending + [{"type": "close"}]
                payload = json.dumps({"requests": reqs}).encode("utf-8")
                req = urllib.request.Request(
                    f"{self._url}/v2/pipeline",
                    data=payload,
                    headers={
                        "Authorization": f"Bearer {self._token}",
                        "Content-Type": "application/json"
                    },
                    method="POST"
                )
                with urllib.request.urlopen(req, timeout=15) as resp:
                    resp.read()
            except Exception:
                pass
            finally:
                self._pending = []

    def close(self):
        self._conn.close()
def pousser_toute_la_base_vers_turso():
    """Envoie l'intégralité de la base locale vers Turso Cloud par lots rapides."""
    url, token = _get_turso_secrets()
    if not url or not token:
        return
    local_conn = sqlite3.connect(DB_NAME, timeout=30)
    try:
        c = local_conn.cursor()
        c.execute("SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
        tables = c.fetchall()

        for tbl_name, tbl_sql in tables:
            reqs = [
                {"type": "execute", "stmt": {"sql": f"DROP TABLE IF EXISTS {tbl_name}"}},
                {"type": "execute", "stmt": {"sql": tbl_sql}}
            ]
            c.execute(f"SELECT * FROM {tbl_name}")
            rows = c.fetchall()
            if rows:
                nb_cols = len(rows.pop(0)) if False else len(c.description)
                placeholders = ",".join(["?"] * nb_cols)
                sql_ins = f"INSERT INTO {tbl_name} VALUES ({placeholders})"
                for r in rows:
                    args = []
                    for val in r:
                        if val is None:
                            args.append({"type": "null"})
                        elif isinstance(val, int):
                            args.append({"type": "integer", "value": str(val)})
                        elif isinstance(val, float):
                            args.append({"type": "float", "value": val})
                        else:
                            args.append({"type": "text", "value": str(val)})
                    reqs.append({"type": "execute", "stmt": {"sql": sql_ins, "args": args}})
                    # Envoi par paquets de 400 lignes pour aller très vite
                    if len(reqs) >= 400:
                        reqs.append({"type": "close"})
                        req = urllib.request.Request(
                            f"{url}/v2/pipeline",
                            data=json.dumps({"requests": reqs}).encode("utf-8"),
                            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                            method="POST"
                        )
                        with urllib.request.urlopen(req, timeout=30) as resp:
                            resp.read()
                        reqs = []

            if reqs:
                reqs.append({"type": "close"})
                req = urllib.request.Request(
                    f"{url}/v2/pipeline",
                    data=json.dumps({"requests": reqs}).encode("utf-8"),
                    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                    method="POST"
                )
                with urllib.request.urlopen(req, timeout=30) as resp:
                    resp.read()
    finally:
        local_conn.close()

def synchroniser_depuis_turso():
    """Télécharge toutes les tables depuis Turso si la base locale est vide après un redémarrage Streamlit."""
    url, token = _get_turso_secrets()
    if not url or not token:
        return
    local_conn = sqlite3.connect(DB_NAME, timeout=30)
    try:
        c_loc = local_conn.cursor()
        try:
            c_loc.execute("SELECT COUNT(*) FROM articles")
            if itemgetter(0)(c_loc.fetchone()) > 0:
                return
        except Exception:
            pass

        reqs = [{"type": "execute", "stmt": {"sql": "SELECT name, sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"}}, {"type": "close"}]
        req = urllib.request.Request(
            f"{url}/v2/pipeline",
            data=json.dumps({"requests": reqs}).encode("utf-8"),
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=25) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        rows_tbl = data["results"]["response"]["result"]["rows"]
        for r in rows_tbl:
            tbl_name = r["value"]
            tbl_sql = r["value"]
            c_loc.execute(tbl_sql.replace("CREATE TABLE ", "CREATE TABLE IF NOT EXISTS "))
            req_data = urllib.request.Request(
                f"{url}/v2/pipeline",
                data=json.dumps({"requests": [{"type": "execute", "stmt": {"sql": f"SELECT * FROM {tbl_name}"}}, {"type": "close"}]}).encode("utf-8"),
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                method="POST"
            )
            with urllib.request.urlopen(req_data, timeout=30) as r_d:
                res_d = json.loads(r_d.read().decode("utf-8"))["results"]["response"]["result"]
            cols = [col["name"] for col in res_d["cols"]]
            if cols and res_d["rows"]:
                placeholders = ",".join(["?"] * len(cols))
                vals_list = [[None if cell["type"] == "null" else cell.get("value") for cell in row] for row in res_d["rows"]]
                c_loc.executemany(f"INSERT OR REPLACE INTO {tbl_name} ({','.join(cols)}) VALUES ({placeholders})", vals_list)
        local_conn.commit()
    except Exception:
        pass
    finally:
        local_conn.close()

def get_connection():
    local_conn = sqlite3.connect(DB_NAME, timeout=30, check_same_thread=False)
    url, token = _get_turso_secrets()
    if url and token:
        return _TursoConnWrapper(local_conn, url, token)
    return local_conn

def initialiser_base():
    synchroniser_depuis_turso()
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
# ==============================================================================
# MODULE 4 : TRANSFERTS INTER-DÉPARTEMENTS & CALCUL DU STOCK TEMPS RÉEL
# ==============================================================================

def initialiser_transferts_m4():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS doc_transferts (
            numero TEXT PRIMARY KEY,
            type_doc TEXT NOT NULL,
            date_doc TEXT NOT NULL,
            dept_source TEXT NOT NULL,
            dept_dest TEXT NOT NULL,
            statut TEXT NOT NULL,
            observation TEXT DEFAULT '',
            ref_demande TEXT DEFAULT ''
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS doc_transferts_lignes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_doc TEXT NOT NULL,
            code_article TEXT NOT NULL,
            designation TEXT NOT NULL,
            unite TEXT DEFAULT 'Unité',
            qte_demandee REAL NOT NULL,
            qte_validee REAL NOT NULL,
            prix_unitaire_ht REAL DEFAULT 0.0,
            montant_ht REAL DEFAULT 0.0,
            FOREIGN KEY (numero_doc) REFERENCES doc_transferts(numero)
        )""")
        conn.commit()
        c.close()
    finally:
        conn.close()

def calculer_stock_article_dept(code_article, code_dept, ignorer_doc_transfert=None):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT COALESCE(SUM(l.qte), 0.0)
            FROM doc_achats_lignes l
            JOIN doc_achats d ON l.numero_doc = d.numero
            WHERE UPPER(l.code_article) = UPPER(?)
              AND UPPER(d.code_dept) = UPPER(?)
              AND d.type_doc IN ('BR', 'ED')
        """, (code_article, code_dept))
        entrees_achats = float(itemgetter(0)(c.fetchone()))

        c.execute("""
            SELECT COALESCE(SUM(l.qte), 0.0)
            FROM doc_achats_lignes l
            JOIN doc_achats d ON l.numero_doc = d.numero
            WHERE UPPER(l.code_article) = UPPER(?)
              AND UPPER(d.code_dept) = UPPER(?)
              AND d.type_doc IN ('RF', 'SD')
        """, (code_article, code_dept))
        sorties_achats = float(itemgetter(0)(c.fetchone()))

        sql_in = """
            SELECT COALESCE(SUM(l.qte_validee), 0.0)
            FROM doc_transferts_lignes l
            JOIN doc_transferts d ON l.numero_doc = d.numero
            WHERE UPPER(l.code_article) = UPPER(?)
              AND UPPER(d.dept_dest) = UPPER(?)
              AND d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
        """
        params_in = [code_article, code_dept]
        if ignorer_doc_transfert:
            sql_in += " AND d.numero != ?"
            params_in.append(ignorer_doc_transfert)
        c.execute(sql_in, params_in)
        entrees_trans = float(itemgetter(0)(c.fetchone()))

        sql_out = """
            SELECT COALESCE(SUM(l.qte_validee), 0.0)
            FROM doc_transferts_lignes l
            JOIN doc_transferts d ON l.numero_doc = d.numero
            WHERE UPPER(l.code_article) = UPPER(?)
              AND UPPER(d.dept_source) = UPPER(?)
              AND d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
        """
        params_out = [code_article, code_dept]
        if ignorer_doc_transfert:
            sql_out += " AND d.numero != ?"
            params_out.append(ignorer_doc_transfert)
        c.execute(sql_out, params_out)
        sorties_trans = float(itemgetter(0)(c.fetchone()))

        stock = entrees_achats - sorties_achats + entrees_trans - sorties_trans
        return round(stock, 3)
    finally:
        conn.close()

def generer_numero_transfert(prefixe, annee="2026"):
    conn = get_connection()
    try:
        c = conn.cursor()
        motif = f"{prefixe}-{annee}-%"
        c.execute("SELECT numero FROM doc_transferts WHERE numero LIKE ? ORDER BY numero DESC LIMIT 1", (motif,))
        res = c.fetchone()
        if res:
            dernier = itemgetter(0)(res)
            try:
                seq = int(dernier.split("-").pop()) + 1
            except Exception:
                seq = 1
        else:
            seq = 1
        return f"{prefixe}-{annee}-{seq:04d}"
    finally:
        conn.close()
def enregistrer_transfert(numero, type_doc, date_doc, dept_source, dept_dest, statut, obs, lignes, creation=True, ref_demande=""):
    if dept_source == dept_dest:
        return False, "Le département source et le département destinataire doivent être différents."
    if not lignes:
        return False, "Veuillez ajouter au moins une ligne d'article."

    if type_doc == "BT" and statut == "VALIDÉ":
        ignorer = None if creation else numero
        for code_a, des_a, u_a, q_dem, q_val, pu in lignes:
            if float(q_val) > 0:
                dispo = calculer_stock_article_dept(code_a, dept_source, ignorer_doc_transfert=ignorer)
                if float(q_val) > dispo:
                    return False, f"Stock insuffisant pour '{code_a} - {des_a}' dans {dept_source} (Disponible : {dispo:,.3f} {u_a}, Demandé : {float(q_val):,.3f} {u_a})."

    conn = get_connection()
    try:
        c = conn.cursor()
        if creation:
            c.execute("""INSERT INTO doc_transferts
                (numero, type_doc, date_doc, dept_source, dept_dest, statut, observation, ref_demande)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (numero, type_doc, str(date_doc), dept_source, dept_dest, statut, obs.strip(), ref_demande))
        else:
            c.execute("""UPDATE doc_transferts
                SET date_doc = ?, dept_source = ?, dept_dest = ?, statut = ?, observation = ?
                WHERE numero = ?""",
                (str(date_doc), dept_source, dept_dest, statut, obs.strip(), numero))
            c.execute("DELETE FROM doc_transferts_lignes WHERE numero_doc = ?", (numero,))

        for code_a, des_a, u_a, q_dem, q_val, pu in lignes:
            m_ht = round(float(q_val) * float(pu), 3)
            c.execute("""INSERT INTO doc_transferts_lignes
                (numero_doc, code_article, designation, unite, qte_demandee, qte_validee, prix_unitaire_ht, montant_ht)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (numero, code_a, des_a, u_a, float(q_dem), float(q_val), float(pu), m_ht))

        conn.commit()
        return True, f"Document {numero} ({statut}) enregistré avec succès."
    finally:
        conn.close()

def valider_demande_et_generer_bt(num_demande, date_bt, lignes_validees, obs_validation=""):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT dept_source, dept_dest FROM doc_transferts WHERE numero = ?", (num_demande,))
        res = c.fetchone()
        if not res:
            return False, "Demande introuvable."
        dept_src, dept_dst = itemgetter(0)(res), itemgetter(1)(res)
    finally:
        conn.close()

    lignes_bt = []
    for code_a, des_a, u_a, q_dem, q_acc, pu in lignes_validees:
        if float(q_acc) > 0:
            dispo = calculer_stock_article_dept(code_a, dept_src)
            if float(q_acc) > dispo:
                return False, f"Blocage : Stock insuffisant dans {dept_src} pour '{code_a} - {des_a}' (Stock existant : {dispo:,.3f} {u_a}, Accordé : {float(q_acc):,.3f} {u_a})."
            lignes_bt.append((code_a, des_a, u_a, float(q_dem), float(q_acc), float(pu)))

    if not lignes_bt:
        return False, "Aucune quantité accordée supérieure à 0."

    num_bt = generer_numero_transfert("BT", str(date_bt)[:4])
    ok, msg = enregistrer_transfert(
        num_bt, "BT", date_bt, dept_src, dept_dst, "VALIDÉ",
        f"Généré depuis {num_demande}. {obs_validation}".strip(),
        lignes_bt, creation=True, ref_demande=num_demande
    )
    if not ok:
        return False, msg

    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE doc_transferts SET statut = 'GÉNÉRÉ' WHERE numero = ?", (num_demande,))
        conn.commit()
        return True, f"Demande {num_demande} validée ! Bon de transfert {num_bt} généré et stock mis à jour."
    finally:
        conn.close()

def changer_statut_demande(num_demande, nouveau_statut):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE doc_transferts SET statut = ? WHERE numero = ?", (nouveau_statut, num_demande))
        conn.commit()
        return True, f"Demande {num_demande} passée au statut '{nouveau_statut}'."
    finally:
        conn.close()

def lister_transferts(type_doc, statut_filtre=None):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = """
            SELECT d.numero, d.date_doc, d.dept_source, d.dept_dest, d.statut, d.observation, d.ref_demande,
                   COALESCE((SELECT SUM(montant_ht) FROM doc_transferts_lignes l WHERE l.numero_doc = d.numero), 0.0)
            FROM doc_transferts d
            WHERE d.type_doc = ?
        """
        params = [type_doc]
        if statut_filtre:
            sql += " AND d.statut = ?"
            params.append(statut_filtre)
        sql += " ORDER BY d.date_doc DESC, d.numero DESC"
        c.execute(sql, params)
        return c.fetchall()
    finally:
        conn.close()

def lire_lignes_transfert(numero_doc):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT code_article, designation, unite, qte_demandee, qte_validee, prix_unitaire_ht, montant_ht
            FROM doc_transferts_lignes
            WHERE numero_doc = ?
            ORDER BY id
        """, (numero_doc,))
        return c.fetchall()
    finally:
        conn.close()

def supprimer_transfert(numero_doc):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM doc_transferts_lignes WHERE numero_doc = ?", (numero_doc,))
        c.execute("DELETE FROM doc_transferts WHERE numero = ?", (numero_doc,))
        conn.commit()
        return True, f"Document {numero_doc} supprimé."
    finally:
        conn.close()
# ==============================================================================
# MODULE 5 : VISUALISATION DU STOCK & INVENTAIRE PHYSIQUE (CLÔTURE / DÉCLÔTURE)
# ==============================================================================

def initialiser_stock_inventaire_m5():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS inventaires (
            numero TEXT PRIMARY KEY,
            date_inv TEXT NOT NULL,
            code_dept TEXT NOT NULL,
            statut TEXT DEFAULT 'BROUILLON',
            observation TEXT DEFAULT ''
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS inventaires_lignes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_inv TEXT NOT NULL,
            code_article TEXT NOT NULL,
            designation TEXT NOT NULL,
            unite TEXT DEFAULT 'Unité',
            stock_theorique REAL DEFAULT 0.0,
            stock_physique REAL DEFAULT 0.0,
            ecart_qte REAL DEFAULT 0.0,
            prix_unitaire_ht REAL DEFAULT 0.0,
            valeur_ecart_ht REAL DEFAULT 0.0,
            FOREIGN KEY (numero_inv) REFERENCES inventaires(numero)
        )""")
        conn.commit()
        c.close()
    finally:
        conn.close()

def lister_etat_stock_global():
    """
    Calcule le stock actuel de tous les articles actifs sur tous les départements actifs
    en 4 requêtes SQL agrégées (temps d'exécution < 0,2s même avec 10 000+ articles).
    """
    from collections import defaultdict
    conn = get_connection()
    try:
        c = conn.cursor()

        # 1. Articles actifs + hiérarchie
        c.execute("""
            SELECT a.code, a.designation, COALESCE(a.sous_famille, 'DIVERS'),
                   COALESCE(sf.libelle, a.sous_famille, 'DIVERS'),
                   COALESCE(f.code, 'DIVERS'), COALESCE(f.libelle, 'DIVERS'),
                   COALESCE(cl.code, COALESCE(a.classe, 'DIVERS')), COALESCE(cl.libelle, COALESCE(a.classe, 'DIVERS')),
                   COALESCE(a.unite, 'Unité'), COALESCE(a.prix_achat, 0.0)
            FROM articles a
            LEFT JOIN sous_familles sf ON UPPER(a.sous_famille) = UPPER(sf.code)
            LEFT JOIN familles f ON sf.code_famille = f.code
            LEFT JOIN classes cl ON f.code_classe = cl.code
            WHERE COALESCE(a.actif, 1) = 1
            ORDER BY a.code
        """)
        articles = c.fetchall()

        # 2. Départements actifs
        c.execute("""
            SELECT code_court, nom_long
            FROM departements
            WHERE COALESCE(actif, 1) = 1
            ORDER BY code_court
        """)
        depts = c.fetchall()

        # 3. Cumul global des mouvements achats/sorties (BR, ED, RF, SD) par (article, dept)
        stock_map = defaultdict(float)
        c.execute("""
            SELECT UPPER(l.code_article), UPPER(d.code_dept),
                   SUM(CASE WHEN d.type_doc IN ('BR', 'ED') THEN l.qte ELSE -l.qte END)
            FROM doc_achats_lignes l
            JOIN doc_achats d ON l.numero_doc = d.numero
            WHERE d.type_doc IN ('BR', 'ED', 'RF', 'SD')
            GROUP BY UPPER(l.code_article), UPPER(d.code_dept)
        """)
        for ca_u, dp_u, delta in c.fetchall():
            stock_map[(ca_u, dp_u)] += float(delta or 0.0)

        # 4. Cumul global des transferts validés entrants et sortants par (article, dept)
        c.execute("""
            SELECT UPPER(l.code_article), UPPER(d.dept_dest), SUM(l.qte_validee)
            FROM doc_transferts_lignes l
            JOIN doc_transferts d ON l.numero_doc = d.numero
            WHERE d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
            GROUP BY UPPER(l.code_article), UPPER(d.dept_dest)
        """)
        for ca_u, dp_u, q_in in c.fetchall():
            stock_map[(ca_u, dp_u)] += float(q_in or 0.0)

        c.execute("""
            SELECT UPPER(l.code_article), UPPER(d.dept_source), SUM(l.qte_validee)
            FROM doc_transferts_lignes l
            JOIN doc_transferts d ON l.numero_doc = d.numero
            WHERE d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
            GROUP BY UPPER(l.code_article), UPPER(d.dept_source)
        """)
        for ca_u, dp_u, q_out in c.fetchall():
            stock_map[(ca_u, dp_u)] -= float(q_out or 0.0)

    finally:
        conn.close()

    lignes_stock = []
    for dp_code, dp_nom in depts:
        dp_u = dp_code.upper()
        nom_dp_aff = f"{dp_code} - {dp_nom}"
        for ca, da, sf_c, sf_l, f_c, f_l, cl_c, cl_l, unite, pu in articles:
            qte = round(stock_map.get((ca.upper(), dp_u), 0.0), 3)
            pu_f = float(pu)
            lignes_stock.append({
                "code_dept": dp_code,
                "nom_dept": nom_dp_aff,
                "classe": f"{cl_c} - {cl_l}" if cl_c != cl_l else cl_c,
                "famille": f"{f_c} - {f_l}" if f_c != f_l else f_c,
                "sous_famille": f"{sf_c} - {sf_l}" if sf_c != sf_l else sf_c,
                "code_article": ca,
                "designation": da,
                "unite": unite,
                "qte": qte,
                "pu_ht": pu_f,
                "valeur_ht": round(qte * pu_f, 3)
            })
    return lignes_stock
def generer_numero_inventaire(annee="2026"):
    conn = get_connection()
    try:
        c = conn.cursor()
        motif = f"INV-{annee}-%"
        c.execute("SELECT numero FROM inventaires WHERE numero LIKE ? ORDER BY numero DESC LIMIT 1", (motif,))
        res = c.fetchone()
        if res:
            dernier = itemgetter(0)(res)
            try:
                seq = int(dernier.split("-").pop()) + 1
            except Exception:
                seq = 1
        else:
            seq = 1
        return f"INV-{annee}-{seq:04d}"
    finally:
        conn.close()

def sauvegarder_inventaire(numero_inv, date_inv, code_dept, obs, lignes_pointage, creation=True):
    """
    lignes_pointage : liste de tuples (code_article, designation, unite, stock_theo, stock_phys, pu_ht)
    """
    if not lignes_pointage:
        return False, "Aucune ligne d'article à enregistrer."
    conn = get_connection()
    try:
        c = conn.cursor()
        if not creation:
            c.execute("SELECT statut FROM inventaires WHERE numero = ?", (numero_inv,))
            st_act = c.fetchone()
            if st_act and itemgetter(0)(st_act) == "CLÔTURÉ":
                return False, "Cet inventaire est clôturé. Déclôturez-le d'abord pour le modifier."
            c.execute("UPDATE inventaires SET date_inv = ?, observation = ? WHERE numero = ?",
                      (str(date_inv), obs.strip(), numero_inv))
            c.execute("DELETE FROM inventaires_lignes WHERE numero_inv = ?", (numero_inv,))
        else:
            c.execute("""INSERT INTO inventaires (numero, date_inv, code_dept, statut, observation)
                         VALUES (?, ?, ?, 'BROUILLON', ?)""",
                      (numero_inv, str(date_inv), code_dept, obs.strip()))

        for ca, da, ua, st_th, st_ph, pu in lignes_pointage:
            ecart = round(float(st_ph) - float(st_th), 3)
            val_ecart = round(ecart * float(pu), 3)
            c.execute("""INSERT INTO inventaires_lignes
                (numero_inv, code_article, designation, unite, stock_theorique, stock_physique, ecart_qte, prix_unitaire_ht, valeur_ecart_ht)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (numero_inv, ca, da, ua, float(st_th), float(st_ph), ecart, float(pu), val_ecart))

        conn.commit()
        return True, f"Pointage d'inventaire {numero_inv} enregistré (statut : BROUILLON)."
    finally:
        conn.close()

def cloturer_inventaire(numero_inv):
    """
    Clôture l'inventaire et aligne officiellement le stock théorique du département
    sur le stock physique compté via des écritures d'ajustement (ED ou SD).
    """
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT date_inv, code_dept, statut FROM inventaires WHERE numero = ?", (numero_inv,))
        res = c.fetchone()
        if not res:
            return False, "Inventaire introuvable."
        date_inv, code_dept, statut = res
        if statut == "CLÔTURÉ":
            return False, "Cet inventaire est déjà clôturé."

        c.execute("""SELECT code_article, designation, unite, stock_physique, prix_unitaire_ht
                     FROM inventaires_lignes WHERE numero_inv = ?""", (numero_inv,))
        lignes = c.fetchall()
    finally:
        conn.close()

    # Calcul de l'écart exact au moment de la clôture et génération des ajustements
    lignes_ed = []
    lignes_sd = []
    for ca, da, ua, st_phys, pu in lignes:
        stock_actuel = calculer_stock_article_dept(ca, code_dept)
        diff = round(float(st_phys) - stock_actuel, 3)
        if diff > 0:
            lignes_ed.append((ca, da, ua, diff, float(pu), 0.0))
        elif diff < 0:
            lignes_sd.append((ca, da, ua, abs(diff), float(pu), 0.0))

    if lignes_ed:
        enregistrer_document_achat(
            f"ADJ-E-{numero_inv}", "ED", date_inv, "", code_dept, numero_inv,
            f"Ajustement positif clôture {numero_inv}", lignes_ed, creation=True
        )
    if lignes_sd:
        enregistrer_document_achat(
            f"ADJ-S-{numero_inv}", "SD", date_inv, "", code_dept, numero_inv,
            f"Ajustement négatif clôture {numero_inv}", lignes_sd, creation=True
        )

    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE inventaires SET statut = 'CLÔTURÉ' WHERE numero = ?", (numero_inv,))
        conn.commit()
        return True, f"Inventaire {numero_inv} CLÔTURÉ ! Le stock théorique de {code_dept} est désormais aligné sur le physique."
    finally:
        conn.close()

def decloturer_inventaire(numero_inv):
    """
    Déclôture un inventaire : supprime les mouvements d'ajustement générés
    et repasse l'inventaire en statut BROUILLON pour modification.
    """
    supprimer_document_achat(f"ADJ-E-{numero_inv}")
    supprimer_document_achat(f"ADJ-S-{numero_inv}")
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("UPDATE inventaires SET statut = 'BROUILLON' WHERE numero = ?", (numero_inv,))
        conn.commit()
        return True, f"Inventaire {numero_inv} DÉCLÔTURÉ : vous pouvez à nouveau le modifier ou le supprimer."
    finally:
        conn.close()

def lister_inventaires():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT i.numero, i.date_inv, i.code_dept, i.statut, i.observation,
                   COALESCE((SELECT SUM(valeur_ecart_ht) FROM inventaires_lignes l WHERE l.numero_inv = i.numero), 0.0)
            FROM inventaires i
            ORDER BY i.date_inv DESC, i.numero DESC
        """)
        return c.fetchall()
    finally:
        conn.close()

def lire_lignes_inventaire(numero_inv):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT code_article, designation, unite, stock_theorique, stock_physique, ecart_qte, prix_unitaire_ht, valeur_ecart_ht
            FROM inventaires_lignes
            WHERE numero_inv = ?
            ORDER BY code_article
        """, (numero_inv,))
        return c.fetchall()
    finally:
        conn.close()

def supprimer_inventaire(numero_inv):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT statut FROM inventaires WHERE numero = ?", (numero_inv,))
        res = c.fetchone()
        if res and itemgetter(0)(res) == "CLÔTURÉ":
            return False, "Impossible de supprimer un inventaire clôturé. Déclôturez-le d'abord."
        c.execute("DELETE FROM inventaires_lignes WHERE numero_inv = ?", (numero_inv,))
        c.execute("DELETE FROM inventaires WHERE numero = ?", (numero_inv,))
        conn.commit()
        return True, f"Inventaire {numero_inv} supprimé."
    finally:
        conn.close()
# ==============================================================================
# MODULE 6 : GESTION & CENTRE DES COÛTS (VOLET ESTIMATIF & VOLET RÉEL INVENTAIRE)
# ==============================================================================

def calculer_gestion_achats_periode(date_deb, date_fin):
    """
    Synthèse des achats ERP saisis sur la période (BR + ED - RF).
    """
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT d.type_doc, d.date_doc, d.numero, COALESCE(f.nom, d.code_fournisseur, 'Interne'),
                   d.code_dept, l.code_article, l.designation,
                   COALESCE(sf.libelle, a.sous_famille, 'DIVERS'),
                   COALESCE(cl.libelle, COALESCE(a.classe, 'DIVERS')),
                   l.qte, l.montant_ht, l.montant_ttc
            FROM doc_achats_lignes l
            JOIN doc_achats d ON l.numero_doc = d.numero
            LEFT JOIN fournisseurs f ON d.code_fournisseur = f.code
            LEFT JOIN articles a ON UPPER(l.code_article) = UPPER(a.code)
            LEFT JOIN sous_familles sf ON UPPER(a.sous_famille) = UPPER(sf.code)
            LEFT JOIN familles fam ON sf.code_famille = fam.code
            LEFT JOIN classes cl ON fam.code_classe = cl.code
            WHERE d.type_doc IN ('BR', 'ED', 'RF')
              AND d.numero NOT LIKE 'ADJ-%'
              AND d.date_doc BETWEEN ? AND ?
            ORDER BY d.date_doc DESC, d.numero DESC
        """, (str(date_deb), str(date_fin)))
        return c.fetchall()
    finally:
        conn.close()

def calculer_couts_estimatifs(date_deb, date_fin, code_dept_filtre=None):
    """
    Volet 1 (Estimatif) : Basé sur les Bons de Transfert validés (BT) vers les départements
    consommateurs + les Sorties Directes (SD hors ajustements d'inventaire).
    """
    conn = get_connection()
    try:
        c = conn.cursor()
        # 1. Transferts validés reçus par les départements (hors retours vers ECO)
        sql_bt = """
            SELECT d.dept_dest, l.code_article, l.designation, l.unite,
                   COALESCE(a.sous_famille, 'DIVERS'), COALESCE(sf.libelle, a.sous_famille, 'DIVERS'),
                   COALESCE(cl.code, COALESCE(a.classe, 'DIVERS')), COALESCE(cl.libelle, COALESCE(a.classe, 'DIVERS')),
                   l.qte_validee, l.montant_ht
            FROM doc_transferts_lignes l
            JOIN doc_transferts d ON l.numero_doc = d.numero
            LEFT JOIN articles a ON UPPER(l.code_article) = UPPER(a.code)
            LEFT JOIN sous_familles sf ON UPPER(a.sous_famille) = UPPER(sf.code)
            LEFT JOIN familles fam ON sf.code_famille = fam.code
            LEFT JOIN classes cl ON fam.code_classe = cl.code
            WHERE d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
              AND UPPER(d.dept_dest) != 'ECO'
              AND d.date_doc BETWEEN ? AND ?
        """
        params_bt = [str(date_deb), str(date_fin)]
        if code_dept_filtre and code_dept_filtre != "TOUS":
            sql_bt += " AND UPPER(d.dept_dest) = UPPER(?)"
            params_bt.append(code_dept_filtre)
        c.execute(sql_bt, params_bt)
        rows_bt = c.fetchall()

        # 2. Sorties directes (SD hors ajustements d'inventaire ADJ-S-)
        sql_sd = """
            SELECT d.code_dept, l.code_article, l.designation, l.unite,
                   COALESCE(a.sous_famille, 'DIVERS'), COALESCE(sf.libelle, a.sous_famille, 'DIVERS'),
                   COALESCE(cl.code, COALESCE(a.classe, 'DIVERS')), COALESCE(cl.libelle, COALESCE(a.classe, 'DIVERS')),
                   l.qte, l.montant_ht
            FROM doc_achats_lignes l
            JOIN doc_achats d ON l.numero_doc = d.numero
            LEFT JOIN articles a ON UPPER(l.code_article) = UPPER(a.code)
            LEFT JOIN sous_familles sf ON UPPER(a.sous_famille) = UPPER(sf.code)
            LEFT JOIN familles fam ON sf.code_famille = fam.code
            LEFT JOIN classes cl ON fam.code_classe = cl.code
            WHERE d.type_doc = 'SD'
              AND d.numero NOT LIKE 'ADJ-%'
              AND d.date_doc BETWEEN ? AND ?
        """
        params_sd = [str(date_deb), str(date_fin)]
        if code_dept_filtre and code_dept_filtre != "TOUS":
            sql_sd += " AND UPPER(d.code_dept) = UPPER(?)"
            params_sd.append(code_dept_filtre)
        c.execute(sql_sd, params_sd)
        rows_sd = c.fetchall()

        resultats = []
        for r in (rows_bt + rows_sd):
            dp, ca, da, ua, sf_c, sf_l, cl_c, cl_l, qte, mht = r
            resultats.append({
                "dept": dp,
                "code_article": ca,
                "designation": da,
                "unite": ua,
                "sous_famille": f"{sf_c} - {sf_l}" if sf_c != sf_l else sf_c,
                "classe": f"{cl_c} - {cl_l}" if cl_c != cl_l else cl_c,
                "qte": float(qte),
                "cout_ht": float(mht)
            })
        return resultats
    finally:
        conn.close()
def lister_inventaires_clotures_dept(code_dept=None):
    conn = get_connection()
    try:
        c = conn.cursor()
        sql = "SELECT numero, date_inv, code_dept, observation FROM inventaires WHERE statut = 'CLÔTURÉ'"
        params = []
        if code_dept:
            sql += " AND UPPER(code_dept) = UPPER(?)"
            params.append(code_dept)
        sql += " ORDER BY date_inv ASC, numero ASC"
        c.execute(sql, params)
        return c.fetchall()
    finally:
        conn.close()

def calculer_couts_reels_inventaire(code_dept, num_inv_debut, num_inv_fin):
    """
    Volet 2 (Réel) : Consommation Réelle entre deux inventaires clôturés sur un département :
    Conso Réelle = Stock Physique (Inv Début) + Entrées reçues (BR + ED + BT entrants) - Sorties (RF + BT sortants) - Stock Physique (Inv Fin).
    Si num_inv_debut est None (premier inventaire), Stock Initial = 0.
    """
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT date_inv FROM inventaires WHERE numero = ?", (num_inv_fin,))
        res_fin = c.fetchone()
        if not res_fin:
            return []
        date_fin = itemgetter(0)(res_fin)

        stock_init = {}
        val_init = {}
        date_deb = "1900-01-01"
        if num_inv_debut:
            c.execute("SELECT date_inv FROM inventaires WHERE numero = ?", (num_inv_debut,))
            res_deb = c.fetchone()
            if res_deb:
                date_deb = itemgetter(0)(res_deb)
            c.execute("SELECT code_article, stock_physique, prix_unitaire_ht FROM inventaires_lignes WHERE numero_inv = ?", (num_inv_debut,))
            for ca, q_ph, pu in c.fetchall():
                stock_init[ca] = float(q_ph)
                val_init[ca] = float(q_ph) * float(pu)

        stock_fin = {}
        val_fin = {}
        c.execute("SELECT code_article, stock_physique, prix_unitaire_ht FROM inventaires_lignes WHERE numero_inv = ?", (num_inv_fin,))
        for ca, q_ph, pu in c.fetchall():
            stock_fin[ca] = float(q_ph)
            val_fin[ca] = float(q_ph) * float(pu)

        # Référentiel des articles actifs
        c.execute("""
            SELECT a.code, a.designation, COALESCE(a.unite, 'Unité'), COALESCE(a.prix_achat, 0.0),
                   COALESCE(a.sous_famille, 'DIVERS'), COALESCE(sf.libelle, a.sous_famille, 'DIVERS'),
                   COALESCE(cl.code, COALESCE(a.classe, 'DIVERS')), COALESCE(cl.libelle, COALESCE(a.classe, 'DIVERS'))
            FROM articles a
            LEFT JOIN sous_familles sf ON UPPER(a.sous_famille) = UPPER(sf.code)
            LEFT JOIN familles fam ON sf.code_famille = fam.code
            LEFT JOIN classes cl ON fam.code_classe = cl.code
            WHERE COALESCE(a.actif, 1) = 1
            ORDER BY a.code
        """)
        articles = c.fetchall()

        resultats = []
        for ca, da, ua, pu_act, sf_c, sf_l, cl_c, cl_l in articles:
            # Entrées achats (BR, ED hors ADJ-) entre date_deb et date_fin
            c.execute("""
                SELECT COALESCE(SUM(l.qte), 0.0), COALESCE(SUM(l.montant_ht), 0.0)
                FROM doc_achats_lignes l
                JOIN doc_achats d ON l.numero_doc = d.numero
                WHERE UPPER(l.code_article) = UPPER(?) AND UPPER(d.code_dept) = UPPER(?)
                  AND d.type_doc IN ('BR', 'ED') AND d.numero NOT LIKE 'ADJ-%'
                  AND d.date_doc > ? AND d.date_doc <= ?
            """, (ca, code_dept, date_deb, date_fin))
            q_in_a, v_in_a = [float(x) for x in c.fetchone()]

            # Retours fournisseurs (RF)
            c.execute("""
                SELECT COALESCE(SUM(l.qte), 0.0), COALESCE(SUM(l.montant_ht), 0.0)
                FROM doc_achats_lignes l
                JOIN doc_achats d ON l.numero_doc = d.numero
                WHERE UPPER(l.code_article) = UPPER(?) AND UPPER(d.code_dept) = UPPER(?)
                  AND d.type_doc = 'RF'
                  AND d.date_doc > ? AND d.date_doc <= ?
            """, (ca, code_dept, date_deb, date_fin))
            q_rf, v_rf = [float(x) for x in c.fetchone()]

            # Transferts reçus (BT)
            c.execute("""
                SELECT COALESCE(SUM(l.qte_validee), 0.0), COALESCE(SUM(l.montant_ht), 0.0)
                FROM doc_transferts_lignes l
                JOIN doc_transferts d ON l.numero_doc = d.numero
                WHERE UPPER(l.code_article) = UPPER(?) AND UPPER(d.dept_dest) = UPPER(?)
                  AND d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
                  AND d.date_doc > ? AND d.date_doc <= ?
            """, (ca, code_dept, date_deb, date_fin))
            q_in_t, v_in_t = [float(x) for x in c.fetchone()]

            # Transferts sortants (BT)
            c.execute("""
                SELECT COALESCE(SUM(l.qte_validee), 0.0), COALESCE(SUM(l.montant_ht), 0.0)
                FROM doc_transferts_lignes l
                JOIN doc_transferts d ON l.numero_doc = d.numero
                WHERE UPPER(l.code_article) = UPPER(?) AND UPPER(d.dept_source) = UPPER(?)
                  AND d.type_doc = 'BT' AND d.statut = 'VALIDÉ'
                  AND d.date_doc > ? AND d.date_doc <= ?
            """, (ca, code_dept, date_deb, date_fin))
            q_out_t, v_out_t = [float(x) for x in c.fetchone()]

            si_q = stock_init.get(ca, 0.0)
            sf_q = stock_fin.get(ca, 0.0)
            conso_qte = round(si_q + q_in_a - q_rf + q_in_t - q_out_t - sf_q, 3)

            si_v = val_init.get(ca, 0.0)
            sf_v = val_fin.get(ca, 0.0)
            conso_val = round(si_v + v_in_a - v_rf + v_in_t - v_out_t - sf_v, 3)

            if conso_qte != 0 or conso_val != 0 or si_q != 0 or sf_q != 0:
                resultats.append({
                    "dept": code_dept,
                    "code_article": ca,
                    "designation": da,
                    "unite": ua,
                    "sous_famille": f"{sf_c} - {sf_l}" if sf_c != sf_l else sf_c,
                    "classe": f"{cl_c} - {cl_l}" if cl_c != cl_l else cl_c,
                    "stock_init": si_q,
                    "entrees_nettes": round(q_in_a - q_rf + q_in_t - q_out_t, 3),
                    "stock_fin": sf_q,
                    "qte": conso_qte,
                    "cout_ht": conso_val
                })
        return resultats
    finally:
        conn.close()
# ==============================================================================
# MODULE SAUVEGARDE & RESTAURATION DE LA BASE DE DONNÉES
# ==============================================================================

def lire_fichier_base_octets():
    """
    Exporte une copie propre et cohérente de la base SQLite en mémoire (octets).
    """
    import os
    import tempfile
    conn_src = sqlite3.connect(DB_NAME, timeout=30)
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".db") as tmp:
            tmp_path = tmp.name
        conn_dst = sqlite3.connect(tmp_path)
        conn_src.backup(conn_dst)
        conn_dst.close()
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        conn_src.close()
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

def restaurer_base_depuis_octets(contenu_bytes):
    """
    Vérifie l'intégrité du fichier .db envoyé et remplace proprement la base active.
    """
    import os
    import tempfile
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".db") as tmp:
            tmp.write(contenu_bytes)
            tmp_path = tmp.name

        # Vérification que le fichier est bien une base SQLite valide
        conn_test = sqlite3.connect(tmp_path)
        c_test = conn_test.cursor()
        c_test.execute("PRAGMA integrity_check;")
        res = c_test.fetchone()
        if not res or itemgetter(0)(res).lower() != "ok":
            conn_test.close()
            return False, "Le fichier sélectionné n'est pas une base SQLite valide ou est corrompu."

        # Copie sécurisée vers la base active
                # Copie sécurisée vers la base active
        conn_active = sqlite3.connect(DB_NAME, timeout=30)
        conn_test.backup(conn_active)
        conn_active.close()
        conn_test.close()
        pousser_toute_la_base_vers_turso()
        return True, "Base de données restaurée avec succès ! Toutes vos données ont été récupérées."
    except Exception as e:
        return False, f"Erreur lors de la restauration : {e}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

def statistique_tables_base():
    """
    Retourne le nombre d'enregistrements dans chaque table principale pour contrôle.
    """
    tables = [
        ("Classes", "classes"),
        ("Familles", "familles"),
        ("Sous-Familles", "sous_familles"),
        ("Articles", "articles"),
        ("Fournisseurs", "fournisseurs"),
        ("Départements", "departements"),
        ("Bons Achats / Mouvements", "doc_achats"),
        ("Demandes & Bons Transfert", "doc_transferts"),
        ("Inventaires Physiques", "inventaires"),
        ("Lignes Achats (Import Excel)", "achats"),
        ("Lignes Conso (Import Excel)", "transferts")
    ]
    conn = get_connection()
    stats = []
    try:
        c = conn.cursor()
        for label, tbl in tables:
            try:
                c.execute(f"SELECT COUNT(*) FROM {tbl}")
                nb = itemgetter(0)(c.fetchone())
            except Exception:
                nb = 0
            stats.append((label, nb))
        return stats
    finally:
        conn.close()
# ==============================================================================
# MODULE NUITÉES : PAYANTS, ENFANTS, BÉBÉS, DAY USE (TARIF VARIABLE) & GRATUITS
# ==============================================================================

def initialiser_nuitees_m7():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS nuitees_detail (
            date_nuit TEXT PRIMARY KEY,
            payants INTEGER DEFAULT 0,
            enfants INTEGER DEFAULT 0,
            bebes INTEGER DEFAULT 0,
            day_use INTEGER DEFAULT 0,
            tarif_day_use REAL DEFAULT 0.0,
            gratuits INTEGER DEFAULT 0,
            observation TEXT DEFAULT ''
        )""")

        c.execute("""CREATE TABLE IF NOT EXISTS nuitees_coefficients (
            categorie TEXT PRIMARY KEY,
            coeff REAL NOT NULL
        )""")

        defaults = [
            ("payants", 1.00),
            ("enfants", 0.50),
            ("bebes", 0.00),
            ("day_use", 0.50),
            ("gratuits", 1.00)
        ]
        for cat, cf in defaults:
            c.execute("INSERT OR IGNORE INTO nuitees_coefficients (categorie, coeff) VALUES (?, ?)", (cat, cf))

        conn.commit()
        c.close()
    finally:
        conn.close()

def lire_coefficients_nuitees():
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("SELECT categorie, coeff FROM nuitees_coefficients")
        res = {itemgetter(0)(r): float(itemgetter(1)(r)) for r in c.fetchall()}
        return {
            "payants": res.get("payants", 1.0),
            "enfants": res.get("enfants", 0.5),
            "bebes": res.get("bebes", 0.0),
            "day_use": res.get("day_use", 0.5),
            "gratuits": res.get("gratuits", 1.0)
        }
    finally:
        conn.close()

def sauvegarder_coefficients_nuitees(coeffs_dict):
    conn = get_connection()
    try:
        c = conn.cursor()
        for cat, val in coeffs_dict.items():
            c.execute("INSERT OR REPLACE INTO nuitees_coefficients (categorie, coeff) VALUES (?, ?)", (cat, float(val)))
        conn.commit()
        return True, "Coefficients F&B mis à jour avec succès."
    finally:
        conn.close()

def sauvegarder_nuitee_jour(date_nuit, payants, enfants, bebes, day_use, tarif_du, gratuits, obs=""):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""INSERT OR REPLACE INTO nuitees_detail
            (date_nuit, payants, enfants, bebes, day_use, tarif_day_use, gratuits, observation)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (str(date_nuit), int(payants), int(enfants), int(bebes), int(day_use), float(tarif_du), int(gratuits), obs.strip()))
        conn.commit()
        return True, f"Nuitées du {date_nuit} enregistrées avec succès."
    finally:
        conn.close()

def supprimer_nuitee_jour(date_nuit):
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("DELETE FROM nuitees_detail WHERE date_nuit = ?", (str(date_nuit),))
        conn.commit()
        return True, f"Saisie du {date_nuit} supprimée."
    finally:
        conn.close()

def lister_nuitees_periode(date_deb, date_fin):
    coeffs = lire_coefficients_nuitees()
    conn = get_connection()
    try:
        c = conn.cursor()
        c.execute("""
            SELECT date_nuit, payants, enfants, bebes, day_use, tarif_day_use, gratuits, COALESCE(observation, '')
            FROM nuitees_detail
            WHERE date_nuit BETWEEN ? AND ?
            ORDER BY date_nuit DESC
        """, (str(date_deb), str(date_fin)))
        rows = c.fetchall()
    finally:
        conn.close()

    lignes = []
    tot_pay, tot_enf, tot_beb, tot_du, tot_rec_du, tot_grat = 0, 0, 0, 0, 0.0, 0
    tot_brut, tot_pondere = 0.0, 0.0

    for d_n, pay, enf, beb, du, t_du, grat, obs in rows:
        brut = pay + enf + beb + du + grat
        pond = round(
            pay * coeffs["payants"] +
            enf * coeffs["enfants"] +
            beb * coeffs["bebes"] +
            du * coeffs["day_use"] +
            grat * coeffs["gratuits"],
            2
        )
        rec_du = round(du * float(t_du), 3)

        tot_pay += pay
        tot_enf += enf
        tot_beb += beb
        tot_du += du
        tot_rec_du += rec_du
        tot_grat += grat
        tot_brut += brut
        tot_pondere += pond

        lignes.append({
            "date": d_n,
            "payants": pay,
            "enfants": enf,
            "bebes": beb,
            "day_use": du,
            "tarif_du": float(t_du),
            "recette_du": rec_du,
            "gratuits": grat,
            "total_brut": brut,
            "total_pondere": pond,
            "observation": obs
        })

    synthese = {
        "payants": tot_pay,
        "enfants": tot_enf,
        "bebes": tot_beb,
        "day_use": tot_du,
        "recette_day_use": round(tot_rec_du, 3),
        "gratuits": tot_grat,
        "total_brut": round(tot_brut, 2),
        "total_pondere": round(tot_pondere, 2)
    }
    return synthese, lignes
# ==============================================================================
# IMPORTATION INTELLIGENTE DU RÉFÉRENTIEL COMPLET (CLASSES, FAMILLES, SF, ARTICLES)
# ==============================================================================

def importer_base_articles_excel(chemin_excel, ecraser_existants=False):
    """
    Lit le fichier 'Base articles.xlsx' (entêtes ligne 5, données à partir ligne 6) :
    Col A (0): Code Article (11 chiffres : 2 cl + 2 fam + 2 sf + 5 seq)
    Col B (1): Désignation Article
    Col C (2): Libellé Classe
    Col D (3): Libellé Famille
    Col E (4): Libellé Sous-Famille
    Col F (5): Unité
    Col J (9): PMP (Prix Achat HT)
    """
    wb = openpyxl.load_workbook(chemin_excel, data_only=True, read_only=True)
    ws = wb.active
    conn = get_connection()
    nb_cl, nb_fam, nb_sf, nb_art = 0, 0, 0, 0
    try:
        c = conn.cursor()

        # 1. Pré-enregistrement des 14 classes officielles
        classes_std = [
            ("01", "BOISSONS"), ("02", "ALIMENTATION"), ("03", "FRTS BUREAUX"),
            ("04", "HYGIENE"), ("05", "FRTS CLIENTS"), ("06", "PME & PMC"),
            ("07", "MAINTENANCE"), ("08", "LINGE & UNIFORME"), ("09", "EMBALLAGE"),
            ("10", "PRODUITS PISCINES & SEL"), ("11", "EXTRA ARTICLES"),
            ("12", "ANIMATION ARTICLES"), ("13", "THALASSO"), ("14", "ACHATS DIVERS")
        ]
        for c_code, c_lib in classes_std:
            c.execute("INSERT OR IGNORE INTO classes (code, libelle, actif) VALUES (?, ?, 1)", (c_code, c_lib))
            if c.rowcount > 0:
                nb_cl += 1

        # Correspondance Libellé Sous-Famille -> Code court officiel (pour compatibilité historique Excel)
        map_lib_to_code_sf = {
            "WATR": "WATR", "BEER": "BEER", "WHIW": "WHIW", "ROSW": "ROSW", "REDW": "REDW",
            "SPAW": "SPAW", "INTS": "INTS", "NATS": "NATS", "SOFT": "SOFT",
            "MEAT": "MEAT", "FISH": "FISH", "POUL": "POUL", "SAUS": "SAUS",
            "DAIR": "DAIR", "EGGS": "EGGS", "THEC": "THEC", "MIXP": "MIXP",
            "BRKF": "BRKF", "PATI": "PATI", "ICEC": "ICEC",
            "VFRE": "VFRE", "VFRO": "VFRO", "VCAN": "VCAN", "FFRE": "FFRE",
            "ORAN": "ORAN", "FFRO": "FFRO", "FCAN": "FCAN", "DRYP": "DRYP",
            "PAST": "PAST", "BRED": "BRED", "OILF": "OILF", "SAUC": "SAUC",
            "FRTS BUREAUX": "FRTS BUREAUX", "IMPRIMES": "IMPRIMES",
            "CARTOUCHES,TON&RUBAN": "CARTOUCHES,TON&RUBAN",
            "MATERIELS D'HYG & NE": "MATERIELS D'HYG & NE",
            "PRODUITS D'HYG & NET": "PRODUITS D'HYG & NET",
            "PRODUITS BUANDERIE": "PRODUITS BUANDERIE",
            "FRTS GUESTS": "FRTS GUESTS",
            "PRODUITS PISCINES": "PRODUITS PISCINES"
        }

        for row in ws.iter_rows(min_row=6, values_only=True):
            raw_code = get_col(row, 0)
            desig = get_col(row, 1)
            if raw_code is None or desig is None:
                continue

            code_art = str(raw_code).strip()
            if code_art.endswith(".0"):
                code_art = code_art[:-2]
            des_p = str(desig).strip()
            if code_art == "" or des_p == "" or not code_art[0].isdigit():
                continue

            lib_cl = str(get_col(row, 2) or "").strip().upper()
            lib_fam = str(get_col(row, 3) or "").strip().upper()
            lib_sf = str(get_col(row, 4) or "").strip().upper()
            unite_p = str(get_col(row, 5) or "Unité").strip()
            pmp_p = nettoyer_nombre(get_col(row, 9))

            # Extraction hiérarchique depuis le code (ex: 11 11 11 11006)
            code_cl = code_art[:2] if len(code_art) >= 2 else "99"
            code_fam = code_art[:4] if len(code_art) >= 4 else f"{code_cl}01"
            code_sf_num = code_art[:6] if len(code_art) >= 6 else f"{code_fam}01"

            # Utilise le code court officiel si connu (ex: WATR, MEAT), sinon le libellé ou code 6 chiffres
            code_sf_final = map_lib_to_code_sf.get(lib_sf, lib_sf if lib_sf != "" else code_sf_num)

            if lib_cl != "":
                c.execute("INSERT OR IGNORE INTO classes (code, libelle, actif) VALUES (?, ?, 1)", (code_cl, lib_cl))
                if c.rowcount > 0:
                    nb_cl += 1

            if lib_fam != "":
                c.execute("INSERT OR IGNORE INTO familles (code, libelle, code_classe, actif) VALUES (?, ?, ?, 1)",
                          (code_fam, lib_fam, code_cl))
                if c.rowcount > 0:
                    nb_fam += 1

            if code_sf_final != "":
                c.execute("INSERT OR IGNORE INTO sous_familles (code, libelle, code_famille, actif) VALUES (?, ?, ?, 1)",
                          (code_sf_final, lib_sf or code_sf_final, code_fam))
                if c.rowcount > 0:
                    nb_sf += 1

            if ecraser_existants:
                c.execute("""INSERT OR REPLACE INTO articles
                    (code, designation, classe, sous_famille, unite, prix_achat, tva, actif)
                    VALUES (?, ?, ?, ?, ?, ?, 19.0, 1)""",
                    (code_art, des_p, code_cl, code_sf_final, unite_p, pmp_p))
                nb_art += 1
            else:
                c.execute("""INSERT OR IGNORE INTO articles
                    (code, designation, classe, sous_famille, unite, prix_achat, tva, actif)
                    VALUES (?, ?, ?, ?, ?, ?, 19.0, 1)""",
                    (code_art, des_p, code_cl, code_sf_final, unite_p, pmp_p))
                if c.rowcount > 0:
                    nb_art += 1

        conn.commit()
        return True, (nb_cl, nb_fam, nb_sf, nb_art)
    except Exception as e:
        return False, str(e)
    finally:
        wb.close()
        conn.close()
# ==============================================================================
# NETTOYAGE & FUSION DES DOUBLONS DE CLASSES (EX: 02 ET 22 -> 02 ALIMENTATION)
# ==============================================================================

def fusionner_doublons_classes():
    """
    Fusionne les doublons de classes (ex: 1 et 01 -> 01 BOISSONS, 2 et 22 -> 02 ALIMENTATION)
    en donnant la priorité absolue aux codes officiels 01..14, puis fusionne les familles en double.
    """
    import unicodedata

    def norm_txt(s):
        s = str(s or "").strip().upper()
        s = "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")
        return s.rstrip("S")

    conn = get_connection()
    nb_suppr_cl, nb_suppr_fam = 0, 0
    try:
        c = conn.cursor()

        # 1. Harmoniser les codes 1..9 en 01..09 si le code 01..09 existe ou non
        codes_officiels = {
            "01": "BOISSONS", "02": "ALIMENTATION", "03": "FRTS BUREAUX",
            "04": "HYGIENE", "05": "FRTS CLIENTS", "06": "PME & PMC",
            "07": "MAINTENANCE", "08": "LINGE & UNIFORME", "09": "EMBALLAGE",
            "10": "PRODUITS PISCINES & SEL", "11": "EXTRA ARTICLES",
            "12": "ANIMATION ARTICLES", "13": "THALASSO", "14": "ACHATS DIVERS"
        }

        c.execute("SELECT code, libelle FROM classes")
        toutes_cl = c.fetchall()

        # Priorité aux codes officiels 01..14
        lib_vers_maitre = {norm_txt(lib): code for code, lib in codes_officiels.items()}
        for code_off, lib_off in codes_officiels.items():
            c.execute("INSERT OR IGNORE INTO classes (code, libelle, actif) VALUES (?, ?, 1)", (code_off, lib_off))
            c.execute("UPDATE classes SET libelle = ? WHERE code = ?", (lib_off, code_off))

        for code_c, lib_c in toutes_cl:
            cle = norm_txt(lib_c)
            if cle in lib_vers_maitre:
                code_maitre = lib_vers_maitre[cle]
                if code_c != code_maitre:
                    c.execute("UPDATE familles SET code_classe = ? WHERE code_classe = ?", (code_maitre, code_c))
                    c.execute("UPDATE articles SET classe = ? WHERE classe = ?", (code_maitre, code_c))
                    c.execute("DELETE FROM classes WHERE code = ?", (code_c,))
                    nb_suppr_cl += 1
            else:
                lib_vers_maitre[cle] = code_c

        # 2. Fusionner également les familles ayant le même libellé dans la même classe
        c.execute("SELECT code, libelle, code_classe FROM familles ORDER BY LENGTH(code) ASC, code ASC")
        toutes_fam = c.fetchall()
        fam_vers_maitre = {}
        for code_f, lib_f, cl_f in toutes_fam:
            cle_f = (cl_f, norm_txt(lib_f))
            if cle_f not in fam_vers_maitre:
                fam_vers_maitre[cle_f] = code_f
            else:
                fam_maitre = fam_vers_maitre[cle_f]
                if code_f != fam_maitre:
                    c.execute("UPDATE sous_familles SET code_famille = ? WHERE code_famille = ?", (fam_maitre, code_f))
                    c.execute("DELETE FROM familles WHERE code = ?", (code_f,))
                    nb_suppr_fam += 1

        conn.commit()
        return True, f"Nettoyage réussi : {nb_suppr_cl} classe(s) et {nb_suppr_fam} famille(s) en double fusionnée(s) vers les codes officiels 01..14 !"
    except Exception as e:
        return False, f"Erreur lors de la fusion : {e}"
    finally:
        conn.close()
# ==============================================================================
# RECONSTRUCTION COMPLÈTE DU RÉFÉRENTIEL OFFICIEL (CLASSES, FAMILLES, SOUS-FAMILLES)
# ==============================================================================

def reconstruire_nomenclature_officielle():
    """
    Remplace proprement les tables classes, familles et sous_familles par la nomenclature
    officielle complète de l'hôtel (14 classes, 38 familles, 100+ sous-familles) et
    réaligne tous les articles existants sans les supprimer.
    """
    conn = get_connection()
    try:
        c = conn.cursor()

        # 1. Réinitialisation propre des tables de nomenclature uniquement (les articles et mouvements sont conservés)
        c.execute("DELETE FROM sous_familles")
        c.execute("DELETE FROM familles")
        c.execute("DELETE FROM classes")

        classes_off = [
            ("01", "BOISSONS"),
            ("02", "ALIMENTATION"),
            ("03", "FRTS BUREAUX"),
            ("04", "HYGIENE"),
            ("05", "FRTS CLIENTS"),
            ("06", "PME & PMC"),
            ("07", "MAINTENANCE"),
            ("08", "LINGE & UNIFORME"),
            ("09", "EMBALLAGE"),
            ("10", "PRODUITS PISCINES & SEL"),
            ("11", "EXTRA ARTICLES"),
            ("12", "ANIMATION ARTICLES"),
            ("13", "THALASSO"),
            ("14", "ACHATS DIVERS")
        ]
        for cl_c, cl_l in classes_off:
            c.execute("INSERT INTO classes (code, libelle, actif) VALUES (?, ?, 1)", (cl_c, cl_l))

        familles_off = [
            ("0101", "BOISSONS NON ALCOOLIS", "01"),
            ("0102", "BOISSONS ALCOOLISEES", "01"),
            ("0103", "LIQUEUR", "01"),
            ("0201", "VIAN,POISSO,POUL,CHAR", "02"),
            ("0202", "AUTRES NOURRITURES", "02"),
            ("0203", "PATISSERIE", "02"),
            ("0204", "LEG & FRUITS", "02"),
            ("0205", "FARINE", "02"),
            ("0206", "GRAISSE ET HUILE", "02"),
            ("0207", "EPICES", "02"),
            ("0301", "FRTS DES BUREAUX", "03"),
            ("0401", "MATERIELS D'HYG & NET", "04"),
            ("0402", "PRODUITS D'HYG & NET", "04"),
            ("0403", "PRODUITS BUANDERIE", "04"),
            ("0501", "FRTS GUESTS", "05"),
            ("0601", "PME", "06"),
            ("0602", "PMC", "06"),
            ("0701", "BUILDING & CONSTRUCT", "07"),
            ("0702", "ENGINEERING", "07"),
            ("0703", "HOTEL INVENTORY", "07"),
            ("0704", "JARDIN", "07"),
            ("0705", "MAIN ANIMAT MAT", "07"),
            ("0706", "SWIMMING POOL", "07"),
            ("0801", "LINGE", "08"),
            ("0802", "UNIFORME", "08"),
            ("0901", "EMBALLAGE", "09"),
            ("1001", "PRODUITS PISCINES", "10"),
            ("1002", "SEL INDUSTRIEL", "10"),
            ("1101", "EXTRA BOISSONS", "11"),
            ("1102", "EXTRA CHICHA", "11"),
            ("1103", "EXTRA NOURRITURES", "11"),
            ("1104", "MAGIC SCHOP ARTICLES", "11"),
            ("1201", "ANIMATION ARTICLES", "12"),
            ("1301", "BOUTIQUE THALASSO", "13"),
            ("1302", "CABINE THALASSO", "13"),
            ("1303", "HYGIENNE THALASSO", "13"),
            ("1401", "ACHATS DIVERS", "14")
        ]
        for fam_c, fam_l, cl_c in familles_off:
            c.execute("INSERT INTO familles (code, libelle, code_classe, actif) VALUES (?, ?, ?, 1)", (fam_c, fam_l, cl_c))
        sous_familles_off = [
            # 01 BOISSONS
            ("WATR", "WATR", "0101"), ("SOFT", "SOFT", "0101"),
            ("BEER", "BEER", "0102"), ("WHIW", "WHIW", "0102"), ("ROSW", "ROSW", "0102"), ("REDW", "REDW", "0102"),
            ("SPAW", "SPAW", "0103"), ("INTS", "INTS", "0103"), ("NATS", "NATS", "0103"),
            # 02 ALIMENTATION
            ("MEAT", "MEAT", "0201"), ("FISH", "FISH", "0201"), ("POUL", "POUL", "0201"), ("SAUS", "SAUS", "0201"),
            ("DAIR", "DAIR", "0202"), ("EGGS", "EGGS", "0202"), ("THEC", "THEC", "0202"), ("MIXP", "MIXP", "0202"),
            ("BRKF", "BRKF", "0203"), ("PATI", "PATI", "0203"), ("ICEC", "ICEC", "0203"),
            ("VFRE", "VFRE", "0204"), ("VFRO", "VFRO", "0204"), ("VCAN", "VCAN", "0204"),
            ("FFRE", "FFRE", "0204"), ("ORAN", "ORAN", "0204"), ("FFRO", "FFRO", "0204"), ("FCAN", "FCAN", "0204"), ("DRYP", "DRYP", "0204"),
            ("PAST", "PAST", "0205"), ("BRED", "BRED", "0205"),
            ("OILF", "OILF", "0206"),
            ("SAUC", "SAUC", "0207"),
            # 03 FRTS BUREAUX
            ("FRTS BUREAUX", "FRTS BUREAUX", "0301"), ("IMPRIMES", "IMPRIMES", "0301"), ("CARTOUCHES,TON&RUBAN", "CARTOUCHES,TON&RUBAN", "0301"),
            # 04 HYGIENE
            ("MATERIELS D'HYG & NE", "MATERIELS D'HYG & NE", "0401"),
            ("PRODUITS D'HYG & NET", "PRODUITS D'HYG & NET", "0402"),
            ("PRODUITS BUANDERIE", "PRODUITS BUANDERIE", "0403"),
            # 05 FRTS CLIENTS
            ("FRTS GUESTS", "FRTS GUESTS", "0501"),
            # 06 PME & PMC
            ("PORCELAINE", "PORCELAINE", "0601"), ("VERRERIE", "VERRERIE", "0601"), ("COUVERTS", "COUVERTS", "0601"),
            ("INOX &", "INOX &", "0601"), ("POTERIE", "POTERIE", "0601"),
            ("PMC", "PMC", "0602"), ("MELAMINE PMC", "MELAMINE PMC", "0602"),
            # 07 MAINTENANCE
            ("ENT REP AUTRE INSTAL", "ENT REP AUTRE INSTAL", "0701"), ("ENT REP BATIMENTS", "ENT REP BATIMENTS", "0701"), ("ENT REP INST CHAM CL", "ENT REP INST CHAM CL", "0701"),
            ("ENT REP CH-CLI-VENT", "ENT REP CH-CLI-VENT", "0702"), ("ENT REP ELEC ET MECA", "ENT REP ELEC ET MECA", "0702"),
            ("ENT REP EQUIP TECH", "ENT REP EQUIP TECH", "0702"), ("ENT REP INSTA BUANDE", "ENT REP INSTA BUANDE", "0702"), ("ENT REP PETIT MATERI", "ENT REP PETIT MATERI", "0702"),
            ("ENT REP INS REST BAR", "ENT REP INS REST BAR", "0703"), ("ENT REP INSTA CUISIN", "ENT REP INSTA CUISIN", "0703"),
            ("ENT REP MATE INFORMA", "ENT REP MATE INFORMA", "0703"), ("ENT REP TELEPHONE", "ENT REP TELEPHONE", "0703"), ("ENT REP MAINT TV RA", "ENT REP MAINT TV RA", "0703"),
            ("ENT REP JARDIN", "ENT REP JARDIN", "0704"),
            ("ENT REP SPORT NAUTQU", "ENT REP SPORT NAUTQU", "0705"),
            ("ENT REP PISCINE", "ENT REP PISCINE", "0706"),
            # 08 LINGE & UNIFORME
            ("LINGE F&B", "LINGE F&B", "0801"), ("LINGE ETAGE", "LINGE ETAGE", "0801"), ("LINGE THALASSO", "LINGE THALASSO", "0801"),
            ("UNIFORME REST & BAR", "UNIFORME REST & BAR", "0802"), ("UNIFORME ADMINIST", "UNIFORME ADMINIST", "0802"),
            ("UNIFORME ANIMATION", "UNIFORME ANIMATION", "0802"), ("UNIFORME CUISINE", "UNIFORME CUISINE", "0802"),
            ("UNIFORME ETAGE", "UNIFORME ETAGE", "0802"), ("UNIFORME RECEPTION", "UNIFORME RECEPTION", "0802"),
            ("UNIFORME THALASSO", "UNIFORME THALASSO", "0802"), ("UNIFORME ENTRETIEN", "UNIFORME ENTRETIEN", "0802"), ("UNIFORME SECURITE ET JAI", "UNIFORME SECURITE ET JAI", "0802"),
            # 09 EMBALLAGE
            ("EMBALLAGE", "EMBALLAGE", "0901"),
            # 10 PRODUITS PISCINES & SEL
            ("PRODUITS PISCINES", "PRODUITS PISCINES", "1001"),
            ("SEL GROS", "SEL GROS", "1002"),
            # 11 EXTRA ARTICLES
            ("EXTRA BIERRES", "EXTRA BIERRES", "1101"), ("EXTRA EAUX", "EXTRA EAUX", "1101"), ("EXTRA LIQUEURS", "EXTRA LIQUEURS", "1101"),
            ("EXTRA SODATS", "EXTRA SODATS", "1101"), ("EXTRA VINS", "EXTRA VINS", "1101"),
            ("EXTRA CHICHAS", "EXTRA CHICHAS", "1102"),
            ("EXTRA CAFE & THE & AUTRI", "EXTRA CAFE & THE & AUTRI", "1103"), ("EXTRA GLACES", "EXTRA GLACES", "1103"), ("EXTRA ORANGES", "EXTRA ORANGES", "1103"),
            ("VETEMENTS &", "VETEMENTS &", "1104"),
            # 12 ANIMATION ARTICLES
            ("ANIMATION ARTICLES", "ANIMATION ARTICLES", "1201"),
            # 13 THALASSO
            ("GUESTS THALASSO", "GUESTS THALASSO", "1301"), ("BOUTIQUE THALASSO", "BOUTIQUE THALASSO", "1301"),
            ("CABINE THALASSO", "CABINE THALASSO", "1302"),
            ("HYGIENNE THALASSO", "HYGIENNE THALASSO", "1303"),
            # 14 ACHATS DIVERS
            ("ACHATS DIVERS", "ACHATS DIVERS", "1401")
        ]
        for sf_c, sf_l, fam_c in sous_familles_off:
            c.execute("INSERT OR IGNORE INTO sous_familles (code, libelle, code_famille, actif) VALUES (?, ?, ?, 1)", (sf_c, sf_l, fam_c))

        # Réaligner la colonne 'classe' des articles existants selon leur sous-famille officielle
        c.execute("""
            UPDATE articles
            SET classe = COALESCE((
                SELECT f.code_classe
                FROM sous_familles sf
                JOIN familles f ON sf.code_famille = f.code
                WHERE UPPER(sf.code) = UPPER(articles.sous_famille)
            ), classe)
        """)

        conn.commit()
        return True, f"Nomenclature officielle reconstruite : {len(classes_off)} Classes, {len(familles_off)} Familles et {len(sous_familles_off)} Sous-Familles alignées !"
    except Exception as e:
        return False, f"Erreur lors de la reconstruction : {e}"
    finally:
        conn.close()
# ==============================================================================
# IMPORTATION AUTOMATIQUE DE LA BASE FOURNISSEURS EXCEL (1000+ FOURNISSEURS)
# ==============================================================================

def importer_fournisseurs_excel(chemin_excel, ecraser_existants=False):
    """
    Lit un fichier Excel de fournisseurs (entêtes ligne 1, données à partir ligne 2) :
    Col A (0): Code Fournisseur (conserve les zéros initiaux ex: 00000182)
    Col B (1): Nom
    Col C (2): Famille
    Col D (3): Identifiant 1 (Matricule Fiscal)
    Col E (4): Conventionné (true/false ou Oui/Non)
    Col F (5): Adresse
    """
    wb = openpyxl.load_workbook(chemin_excel, data_only=True, read_only=True)
    ws = wb.active
    conn = get_connection()
    nb_ajoutes, nb_maj = 0, 0
    try:
        c = conn.cursor()
        # Ajout doux des nouvelles colonnes si elles n'existent pas encore
        for col_sql in [
            "ALTER TABLE fournisseurs ADD COLUMN famille TEXT DEFAULT ''",
            "ALTER TABLE fournisseurs ADD COLUMN conventionne TEXT DEFAULT 'Non'",
            "ALTER TABLE fournisseurs ADD COLUMN adresse TEXT DEFAULT ''"
        ]:
            try:
                c.execute(col_sql)
            except Exception:
                pass

        for row in ws.iter_rows(min_row=2, values_only=True):
            raw_code = get_col(row, 0)
            raw_nom = get_col(row, 1)
            if raw_code is None or raw_nom is None:
                continue

            code_f = str(raw_code).strip()
            if code_f.endswith(".0"):
                code_f = code_f[:-2]
            nom_f = str(raw_nom).strip()

            if code_f == "" or nom_f == "" or code_f.upper() == "CODE FOURNISSEUR":
                continue

            fam_f = str(get_col(row, 2) or "").strip()
            mf_f = str(get_col(row, 3) or "").strip()
            raw_conv = str(get_col(row, 4) or "").strip().lower()
            conv_f = "Oui" if raw_conv in ("true", "1", "oui", "vrai") else "Non"
            adr_f = str(get_col(row, 5) or "").strip()

            if ecraser_existants:
                c.execute("SELECT 1 FROM fournisseurs WHERE code = ?", (code_f,))
                existe = c.fetchone() is not None
                c.execute("""INSERT OR REPLACE INTO fournisseurs
                    (code, nom, telephone, matricule_fiscal, email, actif, famille, conventionne, adresse)
                    VALUES (?, ?, COALESCE((SELECT telephone FROM fournisseurs WHERE code = ?), ''), ?,
                            COALESCE((SELECT email FROM fournisseurs WHERE code = ?), ''), 1, ?, ?, ?)""",
                    (code_f, nom_f, code_f, mf_f, code_f, fam_f, conv_f, adr_f))
                if existe:
                    nb_maj += 1
                else:
                    nb_ajoutes += 1
            else:
                c.execute("""INSERT OR IGNORE INTO fournisseurs
                    (code, nom, telephone, matricule_fiscal, email, actif, famille, conventionne, adresse)
                    VALUES (?, ?, '', ?, '', 1, ?, ?, ?)""",
                    (code_f, nom_f, mf_f, fam_f, conv_f, adr_f))
                if c.rowcount > 0:
                    nb_ajoutes += 1

        conn.commit()
        return True, (nb_ajoutes, nb_maj)
    except Exception as e:
        return False, str(e)
    finally:
        wb.close()
        conn.close()
# ==============================================================================
# INITIALISATION DES 30 DÉPARTEMENTS OFFICIELS DE L'HÔTEL
# ==============================================================================

def charger_departements_officiels():
    """
    Enregistre et active les 30 départements officiels (01 à 30) avec leur nature
    et leur suffixe comptable, sans supprimer d'éventuels codes existants.
    """
    conn = get_connection()
    try:
        c = conn.cursor()
        for col_sql in [
            "ALTER TABLE departements ADD COLUMN nature TEXT DEFAULT 'Points de ventes ordinaires'",
            "ALTER TABLE departements ADD COLUMN suffixe_comptable TEXT DEFAULT ''"
        ]:
            try:
                c.execute(col_sql)
            except Exception:
                pass

        depts_off = [
            ("01", "ECONOMAT", "Dépôts Principal", "32", 1),
            ("02", "CUISINE CENTRALE", "Cuisines", "08", 0),
            ("03", "MAGASIN TECHNIQUE", "Dépôts Principal", "40", 1),
            ("04", "CUISINE PERSONNEL", "Cuisines", "08", 0),
            ("05", "BOUCHERIE", "Cuisines", "08", 0),
            ("06", "CAFETERIE", "Cuisines", "08", 0),
            ("07", "CUISINE BEACH SNACK", "Cuisines", "06", 0),
            ("08", "SCE RESTAURATION", "Points de ventes ordinaires", "04", 0),
            ("09", "CAVE DU JOUR", "Points de ventes ordinaires", "06", 0),
            ("10", "LOBBY BAR", "Points de ventes ordinaires", "04", 0),
            ("11", "PATISSERIE", "Cuisines", "08", 0),
            ("12", "POOL BAR", "Points de ventes ordinaires", "04", 0),
            ("13", "CAVE BEACH SNACK", "Points de ventes ordinaires", "06", 0),
            ("14", "BEACH BAR", "Points de ventes ordinaires", "08", 0),
            ("15", "BAR EXTRA", "Points de ventes ordinaires", "04", 0),
            ("16", "CAFE MAURE", "Points de ventes ordinaires", "08", 0),
            ("17", "ETAGE", "Points de ventes ordinaires", "02", 0),
            ("18", "BUANDERIE", "Points de ventes ordinaires", "06", 0),
            ("19", "PME RESTAURANT BARS", "Points de ventes ordinaires", "32", 0),
            ("20", "PMC CUISINE", "Points de ventes ordinaires", "32", 0),
            ("21", "MAINTENANCE", "Points de ventes ordinaires", "40", 0),
            ("22", "LOCAL PISCINE", "Points de ventes ordinaires", "40", 0),
            ("23", "STIWARDING", "Points de ventes ordinaires", "09", 0),
            ("24", "FRONT OFFICE", "Points de ventes ordinaires", "01", 0),
            ("25", "ANIMATION", "Points de ventes ordinaires", "27", 0),
            ("26", "SERVICE HYGIÈNE", "Points de ventes ordinaires", "32", 0),
            ("27", "AQUA PARK", "Points de ventes ordinaires", "32", 0),
            ("28", "SERVICE PERSONNEL", "Points de ventes ordinaires", "32", 0),
            ("29", "SERVICE FINANCIER", "Points de ventes ordinaires", "32", 0),
            ("30", "CASSES & AVARIÉS (DLC)", "Dépôts de stockage uniquement", "32", 0)
        ]

        # Désactive l'ancien code provisoire ECO s'il n'a aucun mouvement, pour éviter le doublon avec 01 - ECONOMAT
        c.execute("SELECT COUNT(*) FROM doc_achats WHERE code_dept = 'ECO'")
        nb_a = itemgetter(0)(c.fetchone())
        c.execute("SELECT COUNT(*) FROM doc_transferts WHERE dept_source = 'ECO' OR dept_dest = 'ECO'")
        nb_t = itemgetter(0)(c.fetchone())
        if nb_a == 0 and nb_t == 0:
            c.execute("DELETE FROM departements WHERE code_court = 'ECO'")

        for code_d, nom_d, nat_d, suf_d, auth_achat in depts_off:
            c.execute("""
                INSERT INTO departements (code_court, nom_long, actif, autorise_achat, nature, suffixe_comptable)
                VALUES (?, ?, 1, ?, ?, ?)
                ON CONFLICT(code_court) DO UPDATE SET
                    nom_long = excluded.nom_long,
                    actif = 1,
                    nature = excluded.nature,
                    suffixe_comptable = excluded.suffixe_comptable
            """, (code_d, nom_d, auth_achat, nat_d, suf_d))

        conn.commit()
        return True, "Les 30 départements officiels (01 à 30) ont été chargés et activés avec succès !"
    except Exception as e:
        return False, f"Erreur : {e}"
    finally:
        conn.close()
