import io, os, json, webbrowser
from datetime import datetime
import requests
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
import msal

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ============================================================
#  CONFIGURATION
# ============================================================
SP_SITE        = "roseblanchetn.sharepoint.com"
SP_SITE_PATH   = "/sites/SDAHSESTPA"
SP_BASE_URL    = "https://roseblanchetn.sharepoint.com"
FILE_UNIQUE_ID = "0761FA65-3D84-4B10-B009-8CA5BF050C98"
CLIENT_ID      = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTHORITY      = "https://login.microsoftonline.com/organizations"
SCOPES         = [
    "https://roseblanchetn.sharepoint.com/AllSites.Read",
    "https://roseblanchetn.sharepoint.com/AllSites.Write"
]

MOIS_FR = {1:"Janvier",2:"Février",3:"Mars",4:"Avril",5:"Mai",6:"Juin",
           7:"Juillet",8:"Août",9:"Septembre",10:"Octobre",11:"Novembre",12:"Décembre"}

# ============================================================
#  CONFIGURATION DES TABLES
#
#  Chaque table est décrite par un dict contenant :
#  - sheet       : nom de la feuille Excel source
#  - output      : nom du fichier Excel de sortie
#  - col_date    : colonne Date
#  - col_lot     : colonne N° Lot
#  - col_etape   : colonne Étape
#  - col_echant  : colonne N° Échantillon
#  - col_notif   : colonne Résultat (Libération / Blocage)
#  - col_prob    : colonne Problème
#  - numeric     : liste de (col_excel, label, condition_lambda, cible_texte)
#  - string      : liste de (col_excel, label, valeur_cible_conforme)
# ============================================================

# ---------- Règles communes SSSE ----------
SSSE_NUMERIC = [
    ("Humidité (%)",               "Teneur en Eau",   lambda v: v <= 13 or v >= 14.5, "13 < x < 14.5"),
    ("AW",                         "AW",              lambda v: v >= 0.7,              "< 0.7"),
    ("Protéine Brut (%) (+/-0,7)", "Protéine Brute",  lambda v: v <= 10,               "> 10 %"),
    ("Protéine (%)/MS",            "Protéine/MS",     lambda v: v <= 12,               "> 12 %"),
    ("∑ >400µ",                    "G>400µ",          lambda v: v >= 10,               "< 10 %"),
    ("∑ 355;250",                  "G 355-250µ",      lambda v: v <= 40,               "> 40 %"),
    ("∑ < 200µ",                   "G<200µ",          lambda v: v >= 50,               "< 50 %"),
    ("G < 125µ",                   "G<125µ",          lambda v: v >= 10,               "< 10 %"),
    ("Gluten Humide",              "Gluten Humide",   lambda v: v <= 28,               "> 28 %"),
    ("Gluten Index",               "Gluten Index",    lambda v: v <= 65 or v >= 90,    "65 < x < 90"),
    ("Gluten Sec",                 "Gluten Sec",      lambda v: v <= 10,               "> 10 %"),
    ("Col. b",                     "Couleur b",       lambda v: v <= 18,               "> 18"),
    ("Piqûre Noir",                "Piqûre Noir",     lambda v: v >= 10,               "< 10"),
    ("Piqûre Brun",                "Piqûre Brun",     lambda v: v >= 100,              "< 100"),
    ("Cendres (%) (+/- 0,02)",     "Cendres",         lambda v: v >= 1,                "< 1 %"),
    ("T Chute",                    "Temps de Chute",  lambda v: v <= 250,              "> 250"),
]
SSSE_STRING = [
    ("Embalage (Etanchité,visuel...)", "Emballage",  "C"),
    ("C.Poids",                        "Poids",      "C"),
    ("C .Date",                        "Etiquetage", "C"),
]

# ---------- Règles PS-7 (T55) ----------
# Seuils issus du tableau de référence image
PS7_NUMERIC = [
    ("Humidité (%)",               "Teneur en Eau",   lambda v: v <= 13 or v >= 14.5, "13 < x < 14.5"),
    ("AW",                         "AW",              lambda v: v >= 0.7,              "< 0.7"),
    ("Protéine Brut (%) (+/-0,6)", "Protéine Brute",  lambda v: v <= 9.6,              "> 9.6 %"),
    ("Prot (%)/MS",                "Protéine/MS",     lambda v: v <= 11,               "> 11 %"),
    ("Tps de Chute",               "Temps de Chute",  lambda v: v <= 250,              "> 250"),
    ("Amidon End",                 "Amidon End (UCD)", lambda v: v <= 18 or v >= 25,   "18-25"),
    ("G 200µ",                     "G 200µ",          lambda v: v >= 2,                "< 2 %"),
    ("∑ 180;63",                   "∑ 180-63µ",       lambda v: v <= 63,               "> 63 %"),  # max2% for G200, sum target
    ("Gluten Humide",              "Gluten Humide",   lambda v: v <= 22,               "> 22 %"),
    ("Col. L",                     "Couleur L",       lambda v: v <= 90,               "> 90"),
    ("Cendres (%) (+/- 0,02)",     "Cendres",         lambda v: v >= 0.6,              "< 0.6 %"),
    ("Alvéo W",                    "Alvéo W",         lambda v: v <= 150,              ">= 150"),
    ("Alvéo P/L",                  "Alvéo P/L",       lambda v: v < 1 or v > 1.8,     "1 <= x <= 1.8"),
    ("Alvéo Ie",                   "Alvéo Ie",        lambda v: v <= 45,               "> 45"),
]
PS7_STRING = [
    ("Embalage (Etanchité,visuel...)", "Emballage",  "C"),
    ("C.Poids",                        "Poids",      "C"),
    ("C.Date",                         "Etiquetage", "C"),
]

# ---------- Règles PS (T65) ----------
PS_NUMERIC = [
    ("Humidité (%)",               "Teneur en Eau",   lambda v: v <= 13 or v >= 14.5, "13 < x < 14.5"),
    ("AW",                         "AW",              lambda v: v >= 0.7,              "< 0.7"),
    ("Protéine Brut (%) (+/-0,6)", "Protéine Brute",  lambda v: v <= 9.6,              "> 9.6 %"),
    ("Prot (%)/MS",                "Protéine/MS",     lambda v: v <= 11,               "> 11 %"),
    ("Tps de Chute",               "Temps de Chute",  lambda v: v <= 250,              "> 250"),
    ("Amidon End",                 "Amidon End (UCD)", lambda v: v <= 18 or v >= 25,   "18-25"),
    ("G 200µ",                     "G 200µ",          lambda v: v >= 2,                "< 2 %"),
    ("∑ 180;63",                   "∑ 180-63µ",       lambda v: v <= 63,               "> 63 %"),
    ("Gluten Humide",              "Gluten Humide",   lambda v: v <= 22,               "> 22 %"),
    ("Col. L",                     "Couleur L",       lambda v: v <= 89,               "> 89"),
    ("Cendres (%) (+/- 0,02)",     "Cendres",         lambda v: v >= 0.7,              "< 0.7 %"),
    ("Alvéo W",                    "Alvéo W",         lambda v: v <= 150,              ">= 150"),
    ("Alvéo P/L",                  "Alvéo P/L",       lambda v: v < 1 or v > 1.5,     "1 <= x <= 1.5"),
    ("Alvéo Ie",                   "Alvéo Ie",        lambda v: v <= 45,               "> 45"),
]
PS_STRING = [
    ("Emballage (Etanchité,visuel...)", "Emballage",  "C"),
    ("Poids",                           "Poids",      "C"),
    ("C.Date",                          "Etiquetage", "C"),
]

# ---------- Descripteurs des 3 tables ----------
TABLES = [
    {
        "name"      : "SSSE",
        "sheet"     : "Semoule SSSE",
        "output"    : "anomalies_ssse.xlsx",
        "col_date"  : "Date",
        "col_lot"   : "N°lot",
        "col_etape" : "Etape",
        "col_echant": "N° de l'échantillon",
        "col_notif" : "Notif",
        "col_prob"  : None,   # pas de colonne Problème dans SSSE
        "numeric"   : SSSE_NUMERIC,
        "string"    : SSSE_STRING,
    },
    {
        "name"      : "PS-7",
        "sheet"     : "PS-7",
        "output"    : "anomalies_ps7.xlsx",
        "col_date"  : "Date",
        "col_lot"   : "N°lot",
        "col_etape" : "Etape",
        "col_echant": "N° échantillon",
        "col_notif" : "Résultat",
        "col_prob"  : "Probléme",
        "numeric"   : PS7_NUMERIC,
        "string"    : PS7_STRING,
    },
    {
        "name"      : "PS",
        "sheet"     : "PS",
        "output"    : "anomalies_ps.xlsx",
        "col_date"  : "Date",
        "col_lot"   : "N°lot",
        "col_etape" : "Etape",
        "col_echant": "N° Echantillon",
        "col_notif" : "Résultat",
        "col_prob"  : "Problémes",
        "numeric"   : PS_NUMERIC,
        "string"    : PS_STRING,
    },
]

# ============================================================
#  COULEURS PAR PARAMÈTRE (communes aux 3 tables)
# ============================================================
PARAM_COLORS = {
    "Piqûre Brun"       : "FDECEA",
    "Piqûre Noir"       : "EDE7F6",
    "G>400µ"            : "FFF8E1",
    "G 355-250µ"        : "FFF8E1",
    "G<200µ"            : "FFF8E1",
    "G<125µ"            : "FFF8E1",
    "G 200µ"            : "FFF8E1",
    "∑ 180-63µ"         : "FFF8E1",
    "Couleur b"         : "E3F2FD",
    "Couleur L"         : "E3F2FD",
    "Teneur en Eau"     : "E8F5E9",
    "Gluten Humide"     : "FFF3E0",
    "Gluten Index"      : "FFF3E0",
    "Gluten Sec"        : "FFF3E0",
    "AW"                : "FDECEA",
    "Cendres"           : "F5F5F5",
    "Temps de Chute"    : "F5F5F5",
    "Alvéo W"           : "E8EAF6",
    "Alvéo P/L"         : "E8EAF6",
    "Alvéo Ie"          : "E8EAF6",
    "Protéine Brute"    : "F3E5F5",
    "Protéine/MS"       : "F3E5F5",
    "Amidon End (UCD)"  : "FFF8E1",
    "Emballage"         : "FCE4EC",
    "Poids"             : "FCE4EC",
    "Etiquetage"        : "FCE4EC",
}

# ============================================================
#  ETAPE 1 — Authentification
# ============================================================
def get_token():
    cache = msal.SerializableTokenCache()
    cache_file = ".token_cache.json"
    if os.path.exists(cache_file):
        with open(cache_file) as f:
            cache.deserialize(f.read())
    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if accounts:
        print(f"  Compte en cache : {accounts[0]['username']}")
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            print("  Token valide trouvé dans le cache")
            _save_cache(cache, cache_file)
            return result["access_token"]
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"Erreur : {flow}")
    print("\n" + "="*55)
    print("  CONNEXION REQUISE")
    print("="*55)
    print(f"\n  1. Ouvre : https://microsoft.com/devicelogin")
    print(f"  2. Entre le code : {flow['user_code']}")
    print(f"  3. Connecte-toi avec ton compte rose-blanche.com")
    print(f"\n  En attente...\n")
    try:
        webbrowser.open("https://microsoft.com/devicelogin")
    except:
        pass
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise Exception(f"Connexion échouée : {result.get('error_description', result)}")
    print("  Connecté avec succès !")
    _save_cache(cache, cache_file)
    return result["access_token"]

def _save_cache(cache, path):
    if cache.has_state_changed:
        with open(path, "w") as f:
            f.write(cache.serialize())

# ============================================================
#  ETAPE 2 — Lecture SharePoint
# ============================================================
def read_workbook(token):
    """Télécharge le fichier Excel SharePoint et retourne le contenu brut."""
    print("  Téléchargement direct via SharePoint...")
    download_url = (
        f"{SP_BASE_URL}/sites/SDAHSESTPA"
        f"/_layouts/15/download.aspx"
        f"?UniqueId={FILE_UNIQUE_ID}"
    )
    headers = {"Authorization": f"Bearer {token}", "Accept": "*/*", "User-Agent": "Mozilla/5.0"}
    resp = requests.get(download_url, headers=headers, timeout=60, allow_redirects=True)
    if resp.status_code != 200 or b"<!DOCTYPE" in resp.content[:200]:
        raise Exception(f"Erreur {resp.status_code} — supprime .token_cache.json et relance")
    print("  Fichier téléchargé avec succès.")
    return resp.content

def read_sheet(content: bytes, table_cfg: dict) -> pd.DataFrame:
    """Lit une feuille spécifique depuis le contenu binaire du classeur."""
    try:
        df = pd.read_excel(io.BytesIO(content), sheet_name=table_cfg["sheet"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        print(f"  [{table_cfg['name']}] {len(df)} lignes lues — feuille '{table_cfg['sheet']}'")
        return df
    except Exception as e:
        # Affiche les feuilles disponibles
        try:
            xls = pd.ExcelFile(io.BytesIO(content))
            available_sheets = xls.sheet_names
            print(f"  [ERREUR] Feuille '{table_cfg['sheet']}' introuvable!")
            print(f"  [INFO] Feuilles disponibles : {available_sheets}")
        except:
            pass
        raise

# ============================================================
#  ETAPE 3 — Détection des anomalies (générique)
# ============================================================
def _to_float(val):
    if val is None:
        return None
    if isinstance(val, float) and np.isnan(val):
        return None
    try:
        return float(str(val).replace(',', '.').replace(' ', '').strip())
    except (ValueError, TypeError):
        return None

def prepare_data(df: pd.DataFrame, table_cfg: dict):
    col_date   = table_cfg["col_date"]
    col_lot    = table_cfg["col_lot"]
    col_etape  = table_cfg["col_etape"]
    col_echant = table_cfg["col_echant"]
    col_notif  = table_cfg["col_notif"]
    col_prob   = table_cfg["col_prob"]

    df[col_date]  = pd.to_datetime(df[col_date], errors="coerce")
    df[col_etape] = df[col_etape].astype(str).str.strip().str.title()
    df["Année"]   = df[col_date].dt.year.astype("Int64").astype(str)
    df["Mois_num"]= df[col_date].dt.month.astype("Int64")

    rows_anom = []

    for _, r in df.iterrows():
        # --- Vérifications numériques ---
        for col, label, cond, cible in table_cfg["numeric"]:
            v = _to_float(r.get(col))
            if v is not None and cond(v):
                rows_anom.append({
                    "Date"          : r[col_date],
                    "Année"         : r["Année"],
                    "Mois_num"      : r["Mois_num"],
                    "N°lot"         : r.get(col_lot),
                    "N° Echantillon": r.get(col_echant),
                    "Etape"         : r[col_etape],
                    "Parametre"     : label,
                    "Valeur"        : v,
                    "Cible"         : cible,
                    "Notif"         : r.get(col_notif),
                    "Probleme"      : r.get(col_prob) if col_prob else "",
                    "Commentaires"  : r.get("Commentaire", r.get("Commentaires", "")),
                })

        # --- Vérifications texte ---
        for col, label, target in table_cfg["string"]:
            v = r.get(col)
            if v is not None and not (isinstance(v, float) and np.isnan(v)):
                if str(v).strip() != target:
                    rows_anom.append({
                        "Date"          : r[col_date],
                        "Année"         : r["Année"],
                        "Mois_num"      : r["Mois_num"],
                        "N°lot"         : r.get(col_lot),
                        "N° Echantillon": r.get(col_echant),
                        "Etape"         : r[col_etape],
                        "Parametre"     : label,
                        "Valeur"        : str(v).strip(),
                        "Cible"         : target,
                        "Notif"         : r.get(col_notif),
                        "Probleme"      : r.get(col_prob) if col_prob else "",
                        "Commentaires"  : r.get("Commentaire", r.get("Commentaires", "")),
                    })

    df_anom = pd.DataFrame(rows_anom)

    print(f"  [{table_cfg['name']}] Total lignes analysées : {len(df)}")
    print(f"  [{table_cfg['name']}] Total anomalies réelles: {len(df_anom)}")
    if len(df_anom) > 0:
        print(f"  Répartition par paramètre :")
        for param, cnt in df_anom["Parametre"].value_counts().items():
            print(f"    {param:<28} {cnt}")

    return df, df_anom

# ============================================================
#  ETAPE 4 — Génération Excel structuré (4 feuilles)
# ============================================================
def generate_excel(df_all: pd.DataFrame, df_anom: pd.DataFrame, table_cfg: dict) -> str:
    output_file = table_cfg["output"]
    col_notif   = table_cfg["col_notif"]
    has_prob    = table_cfg["col_prob"] is not None

    wb = Workbook()

    HEADER_FILL  = PatternFill("solid", fgColor="1E3A5F")
    HEADER_FONT  = Font(color="FFFFFF", bold=True, size=11, name="Arial")
    HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    BORDER = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0")
    )

    def style_header(ws, row, cols):
        for col in range(1, cols + 1):
            c = ws.cell(row=row, column=col)
            c.fill = HEADER_FILL; c.font = HEADER_FONT
            c.alignment = HEADER_ALIGN; c.border = BORDER

    def style_row(ws, row, cols, fill_hex=None):
        for col in range(1, cols + 1):
            c = ws.cell(row=row, column=col)
            if fill_hex:
                c.fill = PatternFill("solid", fgColor=fill_hex)
            c.border = BORDER
            c.alignment = Alignment(vertical="center")

    def set_widths(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # ---- Feuille 1 : Anomalies_Detail ----
    ws1 = wb.active
    ws1.title = "Anomalies_Detail"
    ws1.freeze_panes = "A2"
    ws1.row_dimensions[1].height = 35

    h1 = ["Date", "Année", "Mois", "Semaine", "N° Lot", "N° Échantillon",
          "Étape", "Paramètre", "Valeur mesurée", "Cible", "Résultat/Notif"]
    if has_prob:
        h1.append("Problème déclaré")
    h1.append("Commentaires")

    ws1.append(h1)
    style_header(ws1, 1, len(h1))

    for i, (_, r) in enumerate(df_anom.iterrows(), 2):
        d = r["Date"] if pd.notna(r["Date"]) else None
        row_data = [
            d,
            str(r["Année"])     if pd.notna(r.get("Année"))     else "",
            int(r["Mois_num"])  if pd.notna(r.get("Mois_num"))  else "",
            d.isocalendar()[1]  if d else "",
            str(r["N°lot"])         if pd.notna(r.get("N°lot"))         else "",
            str(r["N° Echantillon"])if pd.notna(r.get("N° Echantillon"))else "",
            str(r["Etape"]),
            str(r["Parametre"]),
            r["Valeur"],
            str(r["Cible"]),
            str(r["Notif"]) if pd.notna(r.get("Notif")) else "",
        ]
        if has_prob:
            row_data.append(str(r["Probleme"]) if pd.notna(r.get("Probleme")) else "")
        row_data.append(str(r["Commentaires"]) if pd.notna(r.get("Commentaires")) else "")

        ws1.append(row_data)
        if d:
            ws1.cell(row=i, column=1).number_format = "DD/MM/YYYY"
        fill = PARAM_COLORS.get(str(r["Parametre"]), "FFFFFF")
        style_row(ws1, i, len(h1), fill_hex=fill if i % 2 == 0 else None)

    widths1 = [13, 8, 7, 9, 10, 16, 16, 22, 15, 16, 14]
    if has_prob:
        widths1.append(22)
    widths1.append(30)
    set_widths(ws1, widths1)

    # ---- Feuille 2 : Resume_Mensuel ----
    ws2 = wb.create_sheet("Resume_Mensuel")
    ws2.freeze_panes = "A2"
    ws2.row_dimensions[1].height = 35

    df2 = df_anom.copy()
    df2["Mois_n"] = df2["Date"].dt.month
    df2["An"]     = df2["Date"].dt.year

    # Taux de notification : compte les valeurs contenant "Libération" ou "Oui"
    def count_notif(x):
        return sum(
            1 for v in x
            if pd.notna(v) and str(v).strip().lower() in ("oui", "libération", "liberation")
        )

    grp2 = df2.groupby(["An", "Mois_n"]).agg(
        Nb       =("Parametre", "count"),
        Notifiees=("Notif",     count_notif),
        Types    =("Parametre", "nunique")
    ).reset_index().sort_values(["An", "Mois_n"])

    h2 = ["Année", "Mois Num", "Mois", "Nb Anomalies", "Nb Résolues", "Types Distincts", "Taux Résolution (%)"]
    ws2.append(h2)
    style_header(ws2, 1, len(h2))

    for i, (_, r) in enumerate(grp2.iterrows(), 2):
        taux = round(r["Notifiees"] / r["Nb"] * 100, 1) if r["Nb"] > 0 else 0
        ws2.append([int(r["An"]), int(r["Mois_n"]), MOIS_FR.get(int(r["Mois_n"]), ""),
                    int(r["Nb"]), int(r["Notifiees"]), int(r["Types"]), taux])
        style_row(ws2, i, len(h2), fill_hex="F0F4FA" if i % 2 == 0 else None)
    set_widths(ws2, [10, 10, 14, 16, 16, 16, 22])

    if grp2.shape[0] > 0:
        chart2 = BarChart()
        chart2.type = "col"; chart2.title = f"Anomalies par mois — {table_cfg['name']}"
        chart2.y_axis.title = "Nb Anomalies"
        chart2.add_data(Reference(ws2, min_col=4, min_row=1, max_row=grp2.shape[0]+1), titles_from_data=True)
        chart2.set_categories(Reference(ws2, min_col=3, min_row=2, max_row=grp2.shape[0]+1))
        chart2.width = 20; chart2.height = 12
        ws2.add_chart(chart2, "I2")

    # ---- Feuille 3 : Resume_Parametre ----
    ws3 = wb.create_sheet("Resume_Parametre")
    ws3.row_dimensions[1].height = 35
    total = len(df_anom)

    grp3 = df_anom.groupby("Parametre").agg(
        Nb       =("Parametre", "count"),
        Notifiees=("Notif",     count_notif),
        Etapes   =("Etape",     "nunique")
    ).reset_index().sort_values("Nb", ascending=False)

    cible_map = {label: cible for (_, label, _, cible) in table_cfg["numeric"]}
    cible_map.update({label: target for (_, label, target) in table_cfg["string"]})

    h3 = ["Paramètre", "Cible", "Nb Anomalies", "% du Total", "Nb Résolues", "Taux Résolution (%)", "Étapes Concernées"]
    ws3.append(h3)
    style_header(ws3, 1, len(h3))

    for i, (_, r) in enumerate(grp3.iterrows(), 2):
        pct  = round(r["Nb"] / total * 100, 1) if total > 0 else 0
        taux = round(r["Notifiees"] / r["Nb"] * 100, 1) if r["Nb"] > 0 else 0
        ws3.append([
            str(r["Parametre"]),
            cible_map.get(str(r["Parametre"]), ""),
            int(r["Nb"]), pct,
            int(r["Notifiees"]), taux,
            int(r["Etapes"])
        ])
        fill = PARAM_COLORS.get(str(r["Parametre"]), "F0F4FA")
        style_row(ws3, i, len(h3), fill_hex=fill)
    set_widths(ws3, [25, 16, 14, 12, 14, 22, 18])

    if grp3.shape[0] > 0:
        pie3 = PieChart(); pie3.title = f"Répartition par paramètre — {table_cfg['name']}"
        pie3.add_data(Reference(ws3, min_col=3, min_row=1, max_row=grp3.shape[0]+1), titles_from_data=True)
        pie3.set_categories(Reference(ws3, min_col=1, min_row=2, max_row=grp3.shape[0]+1))
        pie3.width = 18; pie3.height = 12
        ws3.add_chart(pie3, "I2")

    # ---- Feuille 4 : Resume_Etape ----
    ws4 = wb.create_sheet("Resume_Etape")
    ws4.row_dimensions[1].height = 35

    grp4 = df_anom.groupby("Etape").agg(
        Nb       =("Parametre", "count"),
        Notifiees=("Notif",     count_notif),
        Types    =("Parametre", "nunique")
    ).reset_index().sort_values("Nb", ascending=False)

    h4 = ["Étape", "Nb Anomalies", "% du Total", "Nb Résolues", "Taux Résolution (%)", "Types Distincts"]
    ws4.append(h4)
    style_header(ws4, 1, len(h4))

    for i, (_, r) in enumerate(grp4.iterrows(), 2):
        pct  = round(r["Nb"] / total * 100, 1) if total > 0 else 0
        taux = round(r["Notifiees"] / r["Nb"] * 100, 1) if r["Nb"] > 0 else 0
        ws4.append([str(r["Etape"]), int(r["Nb"]), pct, int(r["Notifiees"]), taux, int(r["Types"])])
        style_row(ws4, i, len(h4), fill_hex="F0F4FA" if i % 2 == 0 else None)
    set_widths(ws4, [18, 14, 12, 14, 22, 16])

    if grp4.shape[0] > 0:
        chart4 = BarChart(); chart4.type = "bar"
        chart4.title = f"Anomalies par étape — {table_cfg['name']}"
        chart4.add_data(Reference(ws4, min_col=2, min_row=1, max_row=grp4.shape[0]+1), titles_from_data=True)
        chart4.set_categories(Reference(ws4, min_col=1, min_row=2, max_row=grp4.shape[0]+1))
        chart4.width = 18; chart4.height = 12
        ws4.add_chart(chart4, "H2")

    wb.save(output_file)
    print(f"  [{table_cfg['name']}] Fichier : {output_file}")
    print(f"  Feuilles : Anomalies_Detail ({len(df_anom)} lignes) | Resume_Mensuel | Resume_Parametre | Resume_Etape")
    return output_file

# ============================================================
#  MAIN
# ============================================================
if __name__ == "__main__":
    print("\n" + "="*55)
    print("  Export Anomalies SSSE / PS-7 / PS → Excel Power BI")
    print("="*55 + "\n")

    print("[1/4] Authentification Microsoft...")
    token = get_token()

    print("\n[2/4] Téléchargement fichier Excel SharePoint...")
    raw_content = read_workbook(token)

    generated_files = []
    total_lignes    = 0
    total_anomalies = 0
    errors = []

    print("\n[3/4] Détection des anomalies — 3 tables...")
    for table_cfg in TABLES:
        try:
            print(f"\n  --- {table_cfg['name']} ---")
            df_raw = read_sheet(raw_content, table_cfg)
            df_all, df_anom = prepare_data(df_raw, table_cfg)
            total_lignes    += len(df_all)
            total_anomalies += len(df_anom)

            print(f"\n[4/4] Génération Excel — {table_cfg['name']}...")
            path = generate_excel(df_all, df_anom, table_cfg)
            generated_files.append(path)
        except Exception as e:
            error_msg = f"ERREUR [{table_cfg['name']}]: {str(e)}"
            print(f"\n  ❌ {error_msg}")
            errors.append(error_msg)
            import traceback
            traceback.print_exc()
            continue
    
    if errors:
        print("\n" + "="*55)
        print("  ERREURS DÉTECTÉES")
        print("="*55)
        for err in errors:
            print(f"  ❌ {err}")
        print("\nLE SCRIPT CONTINUERA MALGRÉ LES ERREURS\n")

    # Commit & push
    import subprocess
    for path in generated_files:
        subprocess.run(["git", "add", path])
    subprocess.run(["git", "commit", "-m",
                    f"Update anomalies SSSE/PS7/PS {datetime.now().strftime('%Y-%m-%d')}"])
    subprocess.run(["git", "push"])

    print("\n" + "="*55)
    print("  TERMINÉ !")
    print("="*55)
    print(f"\n  Total lignes analysées : {total_lignes}")
    print(f"  Total anomalies réelles: {total_anomalies}")
    print(f"\n  Fichiers générés :")
    for p in generated_files:
        print(f"    → {p}")
    print(f"\n  Prochaine étape :")
    print(f"  → Télécharge les fichiers depuis Codespaces")
    print(f"  → Ouvre dans Power BI Desktop → Source Excel")
    print()