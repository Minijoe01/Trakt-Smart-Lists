import streamlit as st
import requests
import time
import qrcode
import io
import pandas as pd
import pytz
from datetime import datetime, timedelta
from streamlit_cookies_controller import CookieController
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from streamlit_echarts import st_echarts

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🎬", layout="wide", initial_sidebar_state="collapsed")

# ==================================================
# UTILITAIRES
# ==================================================

def format_duree(heures):
    if pd.isna(heures) or heures is None or heures <= 0:
        return "0h"
    total_min = round(heures * 60)
    ans = total_min // (365 * 24 * 60)
    mois = (total_min % (365 * 24 * 60)) // (30 * 24 * 60)
    sem = (total_min % (30 * 24 * 60)) // (7 * 24 * 60)
    jours = (total_min % (7 * 24 * 60)) // (24 * 60)
    h = (total_min % (24 * 60)) // 60
    parts = []
    if ans > 0:
        parts.append(f"{ans}an")
    if mois > 0:
        parts.append(f"{mois}mois")
    if sem > 0:
        parts.append(f"{sem}sem")
    if jours > 0:
        parts.append(f"{jours}j")
    if h > 0 or not parts:
        parts.append(f"{h}h")
    return " ".join(parts)

def format_minutes(minutes):
    if not minutes or minutes <= 0:
        return "inconnue"
    h = minutes // 60
    m = minutes % 60
    return f"{h}h{m:02d}" if h>0 else f"{m}min"

# ==================================================
# STYLE
# ==================================================

st.markdown("""
<style>
    :root {
        --am-green: #00A392;
        --am-green-aston: #00524B;
        --am-green-dark: #021412;
        --am-lime: #CEDC00;
        --am-bg-card: rgba(8, 55, 50, 0.75);
        --am-bg-card-hover: rgba(12, 75, 68, 0.85);
        --am-border: rgba(18, 90, 84, 0.5);
        --am-text: #F0FAF8;
        --am-text-muted: #9DC5BF;
    }

    /* On garde le ruban Streamlit visible, mais on l'harmonise et on ajoute de la marge
       pour que le scroll ne passe pas dessous */
    section[data-testid="stSidebar"] > div:first-child { padding-top: 0 !important; }
    footer {visibility: hidden;}
    .stApp { margin-top: 0; padding-top: 0; }
    .block-container {
        padding-top: 3.5rem !important;
    }

    /* FOND ASTON MARTIN 2026 - Diffusion radiale en "tache d'encre" depuis le centre/haut :
       - le centre est un peu plus clair (vert moyen mat)
       - les cotes gauche/droite et le bas s'assombrissent PROGRESSIVEMENT
       - pas de "point lumineux" vif comme un spot, une diffusion tres douce
       - les cotes sont plus sombres que le centre, le bas plus sombre que le haut */
    .stApp {
        background:
            /* assombrissement doux sur le bord gauche */
            radial-gradient(ellipse 50% 100% at 0% 50%, rgba(1,22,20,0.55) 0%, transparent 60%),
            /* assombrissement doux sur le bord droit */
            radial-gradient(ellipse 50% 100% at 100% 50%, rgba(1,22,20,0.55) 0%, transparent 60%),
            /* assombrissement par le bas */
            radial-gradient(ellipse 100% 50% at 50% 100%, rgba(1,22,20,0.85) 0%, rgba(1,22,20,0.3) 60%, transparent 100%),
            /* leger eclairage qui part du centre-haut, tres diffus, pas vif */
            radial-gradient(ellipse 100% 70% at 50% 0%, rgba(0,95,87,1) 0%, rgba(0,78,72,0.9) 30%, rgba(0,55,51,0.7) 60%, transparent 85%),
            /* couleur de base : vert moyen mat */
            #00554D !important;
        background-attachment: fixed !important;
        min-height: 100vh;
    }

    /* Badges obtenus */
    .badge-obtenu {
        background: linear-gradient(135deg, rgba(0,163,146,0.25) 0%, rgba(0,82,75,0.45) 100%) !important;
        border: 1px solid rgba(0,163,146,0.5) !important;
        backdrop-filter: blur(14px);
        border-radius: 16px;
        padding: 20px 16px;
        text-align: center;
        box-shadow: 0 0 25px rgba(0,163,146,0.15), 0 8px 24px rgba(0,0,0,0.25);
        transition: transform 0.25s ease;
        margin-bottom: 12px;
    }
    .badge-obtenu:hover { transform: translateY(-4px); }
    .badge-obtenu .emoji { font-size: 2.5em; margin-bottom: 8px; }
    .badge-obtenu .titre { font-size: 1.05em; font-weight: 700; color: #F0FAF8; margin-bottom: 6px; }
    .badge-obtenu .desc { font-size: 0.82em; color: #9DC5BF; line-height: 1.4; }

    /* Badges verrouilles */
    .badge-lock {
        background: rgba(4, 25, 22, 0.55) !important;
        border: 1px solid rgba(60, 80, 76, 0.4) !important;
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 18px 14px;
        text-align: center;
        opacity: 0.65;
        filter: grayscale(0.7);
        transition: all 0.25s ease;
        margin-bottom: 12px;
    }
    .badge-lock:hover { opacity: 0.9; filter: grayscale(0.2); transform: translateY(-2px); }
    .badge-lock .emoji { font-size: 2.2em; margin-bottom: 8px; filter: grayscale(1); opacity: 0.7; }
    .badge-lock .titre { font-size: 1em; font-weight: 600; color: #7EA8A0; margin-bottom: 6px; }
    .badge-lock .desc { font-size: 0.8em; color: #6B928C; line-height: 1.4; }
    .badge-lock .prog-badge {
        height: 6px;
        background: rgba(0,0,0,0.3);
        border-radius: 3px;
        margin-top: 10px;
        overflow: hidden;
    }
    .badge-lock .prog-badge-fill {
        height: 100%;
        background: linear-gradient(90deg, #00524B, #00A392);
        border-radius: 3px;
    }

    @media (max-width: 768px) {
        div[data-testid="stImage"] img {
            max-width: 80px !important;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.2em !important;
        }
        div[data-testid="stSlider"] {
            padding: 10px !important;
        }
    }

    /* Slider : texte ne depasse pas */
    div[data-testid="stSlider"] {
        padding: 16px !important;
    }
    div[data-testid="stSlider"] label {
        word-wrap: break-word !important;
        overflow: visible !important;
    }

    div[data-testid="stMetric"] {
        padding: 24px 16px !important;
        overflow: visible !important;
        min-height: 130px;
    }
    div[data-testid="stMetricValue"] {
        color: var(--am-text) !important;
        font-size: 1.5em !important;
        font-weight: 800 !important;
        white-space: normal !important;
        word-wrap: break-word !important;
    }
    div[data-testid="stMetricLabel"] p {
        color: var(--am-text-muted) !important;
        font-size: 0.9em !important;
    }

    div[data-testid="stMetric"],
    div.stAlert,
    div[data-testid="stContainer"],
    div.stButton > button,
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-testid="stDataFrame"],
    div[data-testid="stSlider"] > div,
    div[data-testid="stSelectSlider"] > div,
    div[data-testid="stDownloadButton"] > button {
        background-color: var(--am-bg-card) !important;
        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;
        border-radius: 16px !important;
        border: 1px solid var(--am-border) !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.35) !important;
    }
    /* Contraste supplementaire pour les selects/inputs pour qu'ils ne se fondent pas dans le fond */
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div {
        background: rgba(3, 30, 27, 0.88) !important;
        border: 1px solid rgba(0,163,146,0.45) !important;
        box-shadow: 0 4px 20px rgba(0,0,0,0.4) !important;
    }
    div[data-baseweb="select"] > div:hover,
    div[data-baseweb="input"] > div:hover {
        border-color: var(--am-green) !important;
    }

    /* Messages connexion : palette ASTON COHERENTE, plus de bleu/olive qui jurent.
       On utilise des selecteurs forts pour ecraser les styles inline Streamlit */
    div.stAlert, div[data-testid="stAlert"] {
        background-color: rgba(8, 55, 50, 0.75) !important;
        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;
        border-radius: 16px !important;
        border: 1px solid var(--am-border) !important;
        box-shadow: 0 8px 32px rgba(0,0,0,0.22) !important;
        color: var(--am-text) !important;
    }
    div.stAlert p, div[data-testid="stAlert"] p,
    div.stAlert span, div[data-testid="stAlert"] span,
    div.stAlert label, div[data-testid="stAlert"] label {
        color: var(--am-text) !important;
    }

    div.stInfo, div[data-testid="stAlert"][kind="info"] {
        background: linear-gradient(135deg, rgba(0,102,95,0.5) 0%, rgba(0,70,65,0.6) 100%) !important;
        border-left: 4px solid var(--am-green) !important;
        border: 1px solid rgba(0,163,146,0.4) !important;
    }
    div.stInfo svg, div[data-testid="stAlert"][kind="info"] svg { fill: var(--am-green) !important; }

    div.stSuccess, div[data-testid="stAlert"][kind="success"] {
        background: linear-gradient(135deg, rgba(0,163,146,0.22) 0%, rgba(0,82,75,0.38) 100%) !important;
        border-left: 4px solid var(--am-green) !important;
        border: 1px solid rgba(0,163,146,0.45) !important;
    }
    div.stSuccess svg, div[data-testid="stAlert"][kind="success"] svg { fill: var(--am-green) !important; }

    div.stWarning, div[data-testid="stAlert"][kind="warning"] {
        background: linear-gradient(135deg, rgba(206,220,0,0.12) 0%, rgba(110,120,0,0.22) 100%) !important;
        border-left: 4px solid var(--am-lime) !important;
        border: 1px solid rgba(206,220,0,0.4) !important;
    }
    div.stWarning svg, div[data-testid="stAlert"][kind="warning"] svg { fill: var(--am-lime) !important; }

    div.stError, div[data-testid="stAlert"][kind="error"] {
        background: rgba(237,34,36,0.12) !important;
        border-left: 4px solid #ED2224 !important;
        border: 1px solid rgba(237,34,36,0.35) !important;
    }
    div.stError svg, div[data-testid="stAlert"][kind="error"] svg { fill: #ED2224 !important; }

    .stButton > button {
        font-weight: 600; padding: 0.7em 1.3em; color: var(--am-text) !important;
        transition: all 0.25s ease;
        background: rgba(3, 30, 27, 0.85) !important;
        border: 1px solid rgba(0,163,146,0.4) !important;
        box-shadow: 0 4px 18px rgba(0,0,0,0.35) !important;
    }
    .stButton > button:hover {
        background: var(--am-bg-card-hover) !important;
        transform: translateY(-2px);
        box-shadow: 0 8px 26px rgba(0,163,146,0.3) !important;
        border-color: var(--am-green) !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--am-green) 0%, var(--am-green-aston) 100%) !important;
        border: none !important; font-weight: 700;
        box-shadow: 0 4px 18px rgba(0,163,146,0.35) !important;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 8px 28px rgba(0,163,146,0.5) !important;
    }

    div[data-testid="stDownloadButton"] > button {
        background: rgba(3, 30, 27, 0.85) !important;
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        border-radius: 16px !important;
        border: 1px solid rgba(0,163,146,0.4) !important;
        color: var(--am-text) !important;
        font-weight: 600;
        width: 100%;
        padding: 0.7em 1.3em;
        box-shadow: 0 4px 18px rgba(0,0,0,0.35) !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background: var(--am-bg-card-hover) !important;
        border-color: var(--am-green) !important;
        box-shadow: 0 8px 26px rgba(0,163,146,0.3) !important;
    }

    section[data-testid="stSidebar"] {
        background: rgba(2,20,18,0.96) !important;
        backdrop-filter: blur(22px) !important;
        border-right: 1px solid var(--am-border);
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] label {
        padding: 12px 16px !important;
        border-radius: 12px;
        gap: 12px !important;
        transition: all 0.2s ease;
        color: var(--am-text-muted) !important;
        font-weight: 500;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] label:hover {
        background: rgba(0,163,146,0.1) !important;
        color: var(--am-text) !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] label:has(input:checked) {
        background: linear-gradient(135deg, rgba(0,102,95,0.3) 0%, rgba(0,77,72,0.25) 100%) !important;
        color: var(--am-text) !important;
        font-weight: 700 !important;
        border: 1px solid rgba(0,163,146,0.4);
    }

    .section-menu-title { font-size:0.75em; font-weight:800; color: var(--am-lime); text-transform:uppercase; letter-spacing:1.5px; margin:20px 0 12px 0; }
    input[type="checkbox"]:checked { accent-color: var(--am-green); }
    hr { border-color: var(--am-border) !important; }
    p, li, label { color: var(--am-text) !important; }
    .stCaption { color: var(--am-text-muted) !important; }
    button[kind="header"] { background: var(--am-bg-card) !important; backdrop-filter: blur(14px); border-radius:12px !important; border: 1px solid var(--am-border) !important; }
    div[role="progressbar"] > div { background: linear-gradient(90deg, var(--am-green) 0%, var(--am-lime) 100%) !important; }

    .now-playing-card {
        background: linear-gradient(135deg, rgba(0,102,95,0.3) 0%, rgba(4,46,43,0.75) 100%);
        backdrop-filter: blur(16px);
        border-radius: 20px;
        padding: 24px;
        border: 1px solid rgba(0,163,146,0.4);
        box-shadow: 0 12px 40px rgba(0,0,0,0.3);
        margin-bottom: 24px;
    }

    .ghost-card {
        background: var(--am-bg-card);
        backdrop-filter: blur(14px);
        border-radius: 16px;
        padding: 18px 22px;
        margin-bottom: 14px;
        border-left: 4px solid var(--am-lime);
        transition: all 0.25s ease;
        box-shadow: 0 4px 16px rgba(0,0,0,0.15);
    }
    .ghost-card:hover { border-left:4px solid var(--am-green); transform:translateX(4px); background: var(--am-bg-card-hover); }
    .ghost-title { font-size:1.1em; font-weight:700; color: var(--am-text); margin-bottom:6px; }
    .ghost-meta { font-size:0.9em; color: var(--am-text-muted); margin-bottom:14px; }
    .progress-bar-container { width:100%; height:12px; background:rgba(6,59,55,0.8); border-radius:8px; overflow:hidden; }
    .progress-bar-fill { height:100%; border-radius:8px; transition: width 0.6s cubic-bezier(0.4,0,0.2,1); }
    .progress-low { background: linear-gradient(90deg, var(--am-green-aston) 0%, var(--am-green) 100%); }
    .progress-mid { background: linear-gradient(90deg, var(--am-green) 0%, var(--am-lime) 100%); }
    .progress-high { background: linear-gradient(90deg, var(--am-lime) 0%, #E8F064 100%); }
</style>
""", unsafe_allow_html=True)

cookies = CookieController()
CLIENT_ID = st.secrets["TRAKT_CLIENT_ID"]
CLIENT_SECRET = st.secrets["TRAKT_CLIENT_SECRET"]
TMDB_KEY = st.secrets.get("TMDB_API_KEY")

DEVICE_CODE_URL = "https://api.trakt.tv/oauth/device/code"
DEVICE_TOKEN_URL = "https://api.trakt.tv/oauth/device/token"
REFRESH_TOKEN_URL = "https://api.trakt.tv/oauth/token"

def formater_date(date_str, user_tz):
    if not date_str: return ""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z","+00:00"))
        dt_local = dt.astimezone(user_tz)
        offset = dt_local.strftime("%z")
        return dt_local.strftime("%d/%m/%Y %H:%M:%S") + f" ({offset[:3]}:{offset[3:]})"
    except Exception:
        return date_str

# ==================================================
# FONCTIONS TRAKT
# ==================================================

def demarrer_connexion():
    r = requests.post(DEVICE_CODE_URL, json={"client_id": CLIENT_ID}, timeout=15)
    r.raise_for_status()
    return r.json()

def verifier_connexion(dc):
    r = requests.post(DEVICE_TOKEN_URL, json={"code":dc,"client_id":CLIENT_ID,"client_secret":CLIENT_SECRET}, timeout=15)
    return r.json() if r.status_code == 200 else None

def rafraichir_token(rt):
    try:
        r = requests.post(REFRESH_TOKEN_URL, json={"refresh_token":rt,"client_id":CLIENT_ID,"client_secret":CLIENT_SECRET,"redirect_uri":"urn:ietf:wg:oauth:2.0:oob","grant_type":"refresh_token"}, timeout=10)
        return r.json() if r.status_code == 200 else None
    except Exception:
        return None

def sauvegarder_connexion(tokens):
    st.session_state["access_token"] = tokens["access_token"]
    st.session_state["refresh_token"] = tokens["refresh_token"]
    st.session_state["token_heure"] = time.time()
    try:
        cookies.set("trakt_rt", tokens["refresh_token"], expires=datetime.now() + timedelta(days=90))
    except Exception:
        pass
    time.sleep(0.3)

def oublier_connexion():
    try: cookies.remove("trakt_rt")
    except Exception: pass
    time.sleep(0.3)
    st.session_state.clear()

def entetes(at):
    return {"Content-Type":"application/json","trakt-api-version":"2","trakt-api-key":CLIENT_ID,"Authorization":f"Bearer {at}"}

def obtenir_infos(at):
    r = requests.get("https://api.trakt.tv/users/settings", headers=entetes(at), timeout=10)
    r.raise_for_status()
    d = r.json()
    tz = d["user"].get("timezone","Europe/Paris")
    try: utz = pytz.timezone(tz)
    except Exception: utz = pytz.timezone("Europe/Paris")
    return {"pseudo":d["user"]["username"],"tz":utz,"tz_name":tz}

def qrcode_img(url):
    img = qrcode.make(url)
    b = io.BytesIO()
    img.save(b, format="PNG")
    return b.getvalue()

def image_tmdb(tmdb_id, type_c="movie"):
    if not TMDB_KEY or not tmdb_id: return None
    try:
        r = requests.get(f"https://api.themoviedb.org/3/{type_c}/{tmdb_id}", params={"api_key":TMDB_KEY}, timeout=5)
        if r.status_code == 200:
            p = r.json().get("poster_path")
            return f"https://image.tmdb.org/t/p/w200{p}" if p else None
    except Exception:
        return None
    return None

def appliquer_filtres_periode(df, mt, periode):
    if periode == "Cette année":
        return df[df["date_dt"].dt.year == mt.year]
    elif periode == "12 derniers mois":
        return df[df["date_dt"] >= mt - pd.DateOffset(months=12)]
    elif periode == "6 derniers mois":
        return df[df["date_dt"] >= mt - pd.DateOffset(months=6)]
    elif periode == "Ce mois-ci":
        return df[(df["date_dt"].dt.year == mt.year) & (df["date_dt"].dt.month == mt.month)]
    elif periode == "Mois dernier":
        prem = mt.replace(day=1) - timedelta(days=1)
        return df[(df["date_dt"].dt.year == prem.year) & (df["date_dt"].dt.month == prem.month)]
    elif periode == "Aujourd'hui":
        return df[df["date_dt"].dt.date == mt.date()]
    return df

def recuperer_historique(at, barre=None):
    h = entetes(at)
    films, series, films_det, ep_det = {}, {}, [], []
    nf, ne = 0,0
    rp = requests.get("https://api.trakt.tv/users/me/history", headers=h, params={"page":1,"limit":100,"extended":"full"}, timeout=30)
    rp.raise_for_status()
    tp = int(rp.headers.get("X-Pagination-Page-Count",1))
    for p in range(1, tp+1):
        if barre: barre.progress(p/tp*0.6, text=f"Historique : page {p}/{tp}")
        r = requests.get("https://api.trakt.tv/users/me/history", headers=h, params={"page":p,"limit":100,"extended":"full"}, timeout=30)
        r.raise_for_status()
        for it in r.json():
            if it["type"] == "movie":
                nf +=1
                m = it["movie"]
                tid = m["ids"]["trakt"]
                films_det.append({"titre":m["title"],"annee":m.get("year"),"genre":", ".join(m.get("genres",[])) if m.get("genres") else "Inconnu","duree":m.get("runtime",0) or 0,"note":m.get("rating",0) or 0,"date":it["watched_at"],"id":tid})
                if tid not in films:
                    films[tid] = {"titre":m["title"],"annee":m.get("year"),"vues":1,"dernier":it["watched_at"]}
                else:
                    films[tid]["vues"] +=1
            elif it["type"] == "episode":
                ne +=1
                s = it["show"]
                ep = it["episode"]
                sid = s["ids"]["trakt"]
                ep_det.append({"serie":s["title"],"titre":ep["title"],"saison":ep["season"],"episode":ep["number"],"annee":s.get("year"),"genre":", ".join(s.get("genres",[])) if s.get("genres") else "Inconnu","duree":ep.get("runtime",0) or s.get("runtime",40) or 40,"note":s.get("rating",0) or 0,"date":it["watched_at"],"id":sid,"network":s.get("network","Inconnu")})
                if sid not in series:
                    series[sid] = {"titre":s["title"],"annee":s.get("year"),"vues":1,"dernier":it["watched_at"]}
                else:
                    series[sid]["vues"] +=1
    return {"films":films,"series":series,"films_det":films_det,"ep_det":ep_det,"nb_films":len(films),"nb_series":len(series),"nb_vf":nf,"nb_ep":ne}

def recuperer_listes(at):
    r = requests.get("https://api.trakt.tv/users/me/lists", headers=entetes(at), timeout=15)
    r.raise_for_status()
    return r.json()

def recuperer_contenu_liste(at, lid):
    h = entetes(at)
    items, p = [],1
    while True:
        r = requests.get(f"https://api.trakt.tv/users/me/lists/{lid}/items", headers=h, params={"page":p,"limit":100,"extended":"full"}, timeout=15)
        r.raise_for_status()
        d = r.json()
        if not d: break
        for it in d:
            it["_listed_at"] = it.get("listed_at")
        items.extend(d)
        p +=1
    return items

def recuperer_watchlist(at):
    h = entetes(at)
    items, p = [],1
    while True:
        r = requests.get("https://api.trakt.tv/users/me/watchlist", headers=h, params={"page":p,"limit":100,"extended":"full"}, timeout=15)
        r.raise_for_status()
        d = r.json()
        if not d: break
        for it in d:
            it["_listed_at"] = it.get("listed_at")
        items.extend(d)
        p +=1
    return items

def recuperer_lecture(at):
    try:
        r = requests.get("https://api.trakt.tv/users/me/watching", headers=entetes(at), timeout=8)
        if r.status_code == 204: return None
        return r.json() if r.status_code == 200 else None
    except Exception:
        return None

def ct(items):
    return sum(1 for i in items if i["type"]=="movie"), sum(1 for i in items if i["type"]=="show")

def comparer(items, histo):
    res = []
    for it in items:
        if it["type"] == "movie":
            m = it["movie"]
            tid = m["ids"]["trakt"]
            if tid in histo["films"]:
                v = histo["films"][tid]
                ajoute_apres = False
                if it.get("_listed_at") and v.get("dernier"):
                    try:
                        d_list = datetime.fromisoformat(it["_listed_at"].replace("Z","+00:00"))
                        d_vue = datetime.fromisoformat(v["dernier"].replace("Z","+00:00"))
                        ajoute_apres = d_list > d_vue
                    except Exception:
                        pass
                res.append({"type":"Film","titre":m["title"],"annee":m.get("year"),"vues":v["vues"],"dernier":v["dernier"],"tid":tid,"tmdb":m["ids"].get("tmdb"),"ajoute_apres":ajoute_apres})
        elif it["type"] == "show":
            s = it["show"]
            sid = s["ids"]["trakt"]
            if sid in histo["series"]:
                v = histo["series"][sid]
                ajoute_apres = False
                if it.get("_listed_at") and v.get("dernier"):
                    try:
                        d_list = datetime.fromisoformat(it["_listed_at"].replace("Z","+00:00"))
                        d_vue = datetime.fromisoformat(v["dernier"].replace("Z","+00:00"))
                        ajoute_apres = d_list > d_vue
                    except Exception:
                        pass
                res.append({"type":"Série","titre":s["title"],"annee":s.get("year"),"vues":v["vues"],"dernier":v["dernier"],"tid":sid,"tmdb":s["ids"].get("tmdb"),"ajoute_apres":ajoute_apres})
    return res

def analyser(at, histo, barre=None):
    res, stats, app = [], [], {}
    def aj(it, nom, lid):
        if it["type"] == "movie":
            med, t = it["movie"], "Film"
        elif it["type"] == "show":
            med, t = it["show"], "Série"
        else: return
        tid = med["ids"]["trakt"]
        cle = (t, tid)
        if cle not in app:
            app[cle] = {"titre":med["title"],"annee":med.get("year"),"type":t,"tid":tid,"tmdb":med["ids"].get("tmdb"),"dans":[]}
        app[cle]["dans"].append({"nom":nom,"lid":lid})
    if barre: barre.progress(0.6, text="Analyse liste de suivi...")
    wl = recuperer_watchlist(at)
    for it in wl: aj(it, "Liste de suivi", "watchlist")
    m = comparer(wl, histo)
    for x in m:
        x["liste"] = "Liste de suivi"
        x["lid"] = "watchlist"
    res.extend(m)
    nf, ns = ct(wl)
    stats.append({"nom":"Liste de suivi","nf":nf,"ns":ns,"total":len(wl),"vus":len(m)})
    listes = recuperer_listes(at)
    for i,l in enumerate(listes):
        if barre: barre.progress(0.6 + (i+1)/max(len(listes),1)*0.3, text=f"Analyse : {l['name']}")
        items = recuperer_contenu_liste(at, l["ids"]["trakt"])
        for it in items: aj(it, l["name"], l["ids"]["trakt"])
        m = comparer(items, histo)
        for x in m:
            x["liste"] = l["name"]
            x["lid"] = l["ids"]["trakt"]
        res.extend(m)
        nf, ns = ct(items)
        stats.append({"nom":l["name"],"nf":nf,"ns":ns,"total":len(items),"vus":len(m)})
    doublons, doublons_det = [], []
    for info in app.values():
        if len(info["dans"])>=2:
            doublons.append({"type":info["type"],"titre":info["titre"],"annee":info["annee"],"tmdb":info["tmdb"],"nb_listes":len(info["dans"]),"listes":", ".join(v["nom"] for v in info["dans"])})
            for v in info["dans"]:
                doublons_det.append({"type":info["type"],"titre":info["titre"],"annee":info["annee"],"tid":info["tid"],"liste":v["nom"],"lid":v["lid"]})
    return res, stats, doublons, doublons_det

def recuperer_playback(at, barre=None):
    if barre: barre.progress(0.95, text="Recherche des fantômes...")
    r = requests.get("https://api.trakt.tv/sync/playback", headers=entetes(at), timeout=15)
    r.raise_for_status()
    res = []
    for it in r.json():
        if it["type"] == "movie" and it.get("movie"):
            t = it["movie"]["title"]
            a = it["movie"].get("year")
            ty = "Film"
            duree = it["movie"].get("runtime",0)
            tmdb = it["movie"]["ids"].get("tmdb")
        elif it["type"] == "episode" and it.get("show") and it.get("episode"):
            ep = it["episode"]
            t = f"{it['show']['title']} — S{ep['season']:02d}E{ep['number']:02d}"
            a = it["show"].get("year")
            ty = "Épisode"
            duree = ep.get("runtime",0) or it["show"].get("runtime",0)
            tmdb = it["show"]["ids"].get("tmdb")
        else: continue
        prog = round(it.get("progress",0))
        res.append({"type":ty,"titre":t,"annee":a,"prog":prog,"dernier":it["paused_at"],"pid":it["id"],"duree":duree,"tmdb":tmdb})
    res.sort(key=lambda x: x["dernier"])
    return res

def lancer_analyse(rafraichir=False, page_suivante="🏠 Tableau de bord"):
    barre = st.progress(0, text="Démarrage...")
    if rafraichir or "historique" not in st.session_state:
        st.session_state["historique"] = recuperer_historique(st.session_state["access_token"], barre)
    res, stats, doub, doub_det = analyser(st.session_state["access_token"], st.session_state["historique"], barre)
    pb = recuperer_playback(st.session_state["access_token"], barre)
    np = recuperer_lecture(st.session_state["access_token"])
    st.session_state["res"] = res
    st.session_state["stats"] = stats
    st.session_state["doub"] = doub
    st.session_state["doub_det"] = doub_det
    st.session_state["pb"] = pb
    st.session_state["np"] = np
    st.session_state["page_active"] = page_suivante
    barre.empty()
    st.rerun()

# ==================================================
# SUPPRESSION
# ==================================================

def sup_liste(at, lid, items):
    corps = {"movies":[],"shows":[]}
    for it in items:
        c = corps["movies"] if it["type"]=="Film" else corps["shows"]
        c.append({"ids":{"trakt":it["tid"]}})
    url = "https://api.trakt.tv/sync/watchlist/remove" if lid == "watchlist" else f"https://api.trakt.tv/users/me/lists/{lid}/items/remove"
    requests.post(url, headers=entetes(at), json=corps, timeout=15).raise_for_status()

def sup_selection(at, items):
    pl = {}
    for it in items:
        pl.setdefault(it["lid"], []).append(it)
    for lid, its in pl.items():
        sup_liste(at, lid, its)
        time.sleep(0.7)

def sup_playback(at, items):
    for it in items:
        requests.delete(f"https://api.trakt.tv/sync/playback/{it['pid']}", headers=entetes(at), timeout=10).raise_for_status()
        time.sleep(0.3)

# ==================================================
# EXCEL
# ==================================================

def ajuster(ws):
    for col in ws.columns:
        l = 0
        lettre = get_column_letter(col[0].column)
        for c in col:
            try:
                if len(str(c.value)) > l:
                    l = len(str(c.value))
            except: pass
        ws.column_dimensions[lettre].width = min(l+4, 50)

def forme(ws, coul="00524B"):
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        t = Table(displayName=f"Tab_{ws.title.replace(' ','_')}", ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(t)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color=coul, end_color=coul, fill_type="solid")
        c.alignment = Alignment(horizontal="center")
    ajuster(ws)

def generer_excel(pseudo, histo, res, stats, doub, pb, utz):
    th = sum(m["duree"] for m in histo["films_det"])/60 + sum(e["duree"] for e in histo["ep_det"])/60
    df_sum = pd.DataFrame([
        ["Compte",pseudo],["Fuseau",utz.zone],
        ["Films",histo["nb_films"]],["Séries",histo["nb_series"]],
        ["Épisodes",histo["nb_ep"]],["Temps total",format_duree(th)],
        ["Listes",len(stats)-1],["Total contenus",sum(s["total"] for s in stats)],
        ["Déjà vus",len(res)],["Doublons",len(doub)],["Fantômes",len(pb)]
    ], columns=["Statistique","Valeur"])
    df_res = pd.DataFrame(res)
    if not df_res.empty:
        df_res = df_res[["liste","type","titre","annee","vues","dernier","ajoute_apres","tmdb"]].copy()
        df_res["dernier"] = pd.to_datetime(df_res["dernier"]).dt.tz_convert(utz).dt.strftime("%d/%m/%Y %H:%M")
        df_res["ajoute_apres"] = df_res["ajoute_apres"].map({True:"Oui", False:"Non"})
        df_res.columns = ["Liste","Type","Titre","Année","Vues","Dernier","Ajouté après visionnage","TMDB"]
    else:
        df_res = pd.DataFrame(columns=["Liste","Type","Titre","Année","Vues","Dernier","Ajouté après visionnage","TMDB"])
    df_d = pd.DataFrame(doub)
    if not df_d.empty:
        df_d = df_d[["type","titre","annee","tmdb","nb_listes","listes"]].copy()
        df_d.columns = ["Type","Titre","Année","TMDB","Nb listes","Dans"]
    else:
        df_d = pd.DataFrame(columns=["Type","Titre","Année","TMDB","Nb listes","Dans"])
    df_sl = pd.DataFrame(stats)
    df_sl["% nettoyage"] = (df_sl["vus"]/df_sl["total"].replace(0,1)*100).round(1)
    df_sl = df_sl[["nom","nf","ns","total","vus","% nettoyage"]]
    df_sl.columns = ["Liste","Films","Séries","Total","Déjà vus","% nettoyage"]
    df_pb = pd.DataFrame(pb)
    if not df_pb.empty:
        df_pb = df_pb[["type","titre","annee","prog","dernier"]].copy()
        df_pb["dernier"] = pd.to_datetime(df_pb["dernier"]).dt.tz_convert(utz).dt.strftime("%d/%m/%Y %H:%M")
        df_pb.columns = ["Type","Titre","Année","Progression","Dernier"]
    else:
        df_pb = pd.DataFrame(columns=["Type","Titre","Année","Progression","Dernier"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df_sum.to_excel(wr, sheet_name="Résumé", index=False)
        df_res.to_excel(wr, sheet_name="À nettoyer", index=False)
        df_d.to_excel(wr, sheet_name="Doublons", index=False)
        df_pb.to_excel(wr, sheet_name="Fantômes", index=False)
        df_sl.to_excel(wr, sheet_name="Listes", index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    for sh in wb: forme(sh)
    ws = wb["Listes"]
    cp = None
    for c in ws[1]:
        if c.value == "% nettoyage":
            cp = c.column
    if cp:
        l = get_column_letter(cp)
        ws.conditional_formatting.add(f"{l}2:{l}{ws.max_row}", ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="F8696B"))
    buf_f = io.BytesIO()
    wb.save(buf_f)
    buf_f.seek(0)
    return buf_f.getvalue()

# ==================================================
# NAVIGATION
# ==================================================

def naviguer():
    PAGES = [
        "🏠 Tableau de bord",
        "▶️ En cours de lecture",
        "👻 Progression Fantôme",
        "🧹 Nettoyage des listes",
        "🔍 Recherche de doublons",
        "🎯 Que regarder ?",
        "📊 Statistiques",
        "📅 Calendrier des sorties",
        "🎬 Rendez-vous annuel",
        "📤 Sauvegarde",
        "🏆 Succès",
    ]
    if "page_active" not in st.session_state:
        st.session_state["page_active"] = PAGES[0]
    with st.sidebar:
        st.markdown('<p class="section-menu-title">Menu</p>', unsafe_allow_html=True)
        page = st.radio("Navigation", PAGES, index=PAGES.index(st.session_state["page_active"]), label_visibility="collapsed", key="nav")
    st.session_state["page_active"] = page
    return page

# ==================================================
# ENTETE
# ==================================================

def entete():
    # En-tete COMPACT : logo + titre + connexion sur une seule ligne
    cl, ct, ci, cd = st.columns([0.05, 0.40, 0.42, 0.13])
    with cl:
        try: st.image("trakt-logo.svg", width=36)
        except: pass
    with ct:
        st.markdown("<h3 style='margin:0; padding:4px 0 0 0; color:#F0FAF8; font-weight:800; font-size:1.3em;'>Trakt Smart Lists</h3>", unsafe_allow_html=True)
    if "access_token" not in st.session_state:
        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
        return None
    if "token_heure" in st.session_state and (time.time() - st.session_state["token_heure"]) > 7*86400:
        nouveau = rafraichir_token(st.session_state["refresh_token"])
        if nouveau:
            sauvegarder_connexion(nouveau)
    if "infos" not in st.session_state or (time.time() - st.session_state.get("infos_h",0)) > 3600:
        st.session_state["infos"] = obtenir_infos(st.session_state["access_token"])
        st.session_state["infos_h"] = time.time()
    infos = st.session_state["infos"]
    pseudo, utz = infos["pseudo"], infos["tz"]
    with ci:
        # Bandeau de connexion THEMATISE, plus de st.info bleu canard
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, rgba(0,102,95,0.5) 0%, rgba(0,70,65,0.6) 100%);
                    border:1px solid rgba(0,163,146,0.4); border-radius:12px;
                    padding:10px 16px; margin-top:2px; color:#F0FAF8;
                    font-size:0.92em; backdrop-filter: blur(12px);">
            👤 Connecté en tant que <b>{pseudo}</b> • 🕒 <span style="color:#9DC5BF;">{infos['tz_name']}</span>
        </div>
        """, unsafe_allow_html=True)
    with cd:
        if st.button("🚪", use_container_width=True, help="Déconnexion"):
            oublier_connexion()
            st.rerun()
    st.markdown("<div style='height:2px;'></div>", unsafe_allow_html=True)
    if "res" in st.session_state:
        h = st.session_state["historique"]
        res = st.session_state["res"]
        stats = st.session_state["stats"]
        doub = st.session_state["doub"]
        pb = st.session_state["pb"]
        xl = generer_excel(pseudo, h, res, stats, doub, pb, utz)
        c1,c2,c3 = st.columns(3)
        with c1:
            if st.button("🔄 Analyse rapide", use_container_width=True):
                for k in ["res","stats","doub","doub_det","pb","np"]:
                    st.session_state.pop(k, None)
                lancer_analyse(False, st.session_state["page_active"])
        with c2:
            if st.button("🔃 Rafraîchir tout", use_container_width=True):
                for k in ["historique","res","stats","doub","doub_det","pb","np","infos"]:
                    st.session_state.pop(k, None)
                st.rerun()
        with c3:
            st.download_button("📥 Rapport Excel", data=xl, file_name=f"trakt_{pseudo}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.markdown("<hr style='margin:0.4rem 0 1rem 0; border-color: rgba(18,90,84,0.4);'/>", unsafe_allow_html=True)
    return utz

def bloc_lancement():
    if "res" in st.session_state:
        return False
    if "historique" in st.session_state:
        st.markdown("""
        <div style="background: linear-gradient(135deg, rgba(0,102,95,0.5) 0%, rgba(0,70,65,0.6) 100%);
                    border:1px solid rgba(0,163,146,0.4); border-radius:14px;
                    padding:14px 18px; color:#F0FAF8; margin-bottom:14px;">
            ℹ️ Ton historique est déjà chargé, l'analyse sera rapide.
        </div>""", unsafe_allow_html=True)
        txt = "🔄 Lancer l'analyse rapide"
    else:
        st.markdown("""
        <div style="background: linear-gradient(135deg, rgba(0,102,95,0.5) 0%, rgba(0,70,65,0.6) 100%);
                    border:1px solid rgba(0,163,146,0.4); border-radius:14px;
                    padding:14px 18px; color:#F0FAF8; margin-bottom:14px;">
            ℹ️ Lance l'analyse pour accéder à tous les outils.
        </div>""", unsafe_allow_html=True)
        txt = "🔍 Lancer l'analyse complète"
    if st.button(txt, type="primary", use_container_width=True):
        lancer_analyse()
    return True

# ==================================================
# PAGES
# ==================================================

def page_connexion():
    if "dc" not in st.session_state:
        st.write("Connecte ton compte Trakt pour commencer.")
        if st.button("🚀 Se connecter à Trakt", type="primary"):
            infos = demarrer_connexion()
            st.session_state["dc"] = infos["device_code"]
            st.session_state["uc"] = infos["user_code"]
            st.session_state["vu"] = infos["verification_url"]
            st.session_state["exp"] = infos["expires_in"]
            st.session_state["iv"] = infos["interval"]
            st.rerun()
    else:
        url = f"{st.session_state['vu']}/{st.session_state['uc']}"
        cg, cd = st.columns(2)
        with cg:
            st.markdown(f'<a href="{url}" target="_blank" style="display:inline-block; background:linear-gradient(135deg,#00A392,#00524B); color:white; padding:0.9em 1.7em; border-radius:12px; text-decoration:none; font-weight:700;">Autoriser l\'accès</a>', unsafe_allow_html=True)
            st.caption("Sur n'importe quel appareil.")
            st.info(f"Code : **{st.session_state['uc']}**")
        with cd:
            st.image(qrcode_img(url), width=160)
            st.caption("Ou scanne le QR code.")
        st.caption("La page se met à jour automatiquement.")
        with st.spinner("Attente de l'autorisation..."):
            t=0
            while t < st.session_state["exp"]:
                time.sleep(st.session_state["iv"])
                t += st.session_state["iv"]
                tok = verifier_connexion(st.session_state["dc"])
                if tok:
                    sauvegarder_connexion(tok)
                    del st.session_state["dc"]
                    st.rerun()
        st.error("Délai expiré.")
        if st.button("Réessayer"): st.rerun()

def page_lecture(utz):
    if bloc_lancement(): return
    st.subheader("▶️ En cours de lecture")
    np = st.session_state.get("np")
    if not np:
        st.info("🎬 Aucun contenu en lecture actuellement.")
        return
    if np["type"] == "movie":
        med = np["movie"]
        titre = med["title"]
        annee = med.get("year")
        tc = "Film"
        duree = med.get("runtime", 0)
        tmdb = med["ids"].get("tmdb")
    else:
        med = np["show"]
        ep = np["episode"]
        titre = f"{med['title']} — S{ep['season']:02d}E{ep['number']:02d}"
        annee = med.get("year")
        tc = "Épisode"
        duree = ep.get("runtime",0) or med.get("runtime",0)
        tmdb = med["ids"].get("tmdb")
    prog = round(np.get("progress",0))
    debut = datetime.fromisoformat(np["started_at"].replace("Z","+00:00")).astimezone(utz)
    fin = debut + timedelta(minutes=duree) if duree>0 else None
    img = image_tmdb(tmdb, "movie" if tc=="Film" else "tv")
    ci, cd = st.columns([0.2,0.8])
    with ci:
        if img: st.image(img, use_container_width=True)
        else: st.markdown("🎬")
    with cd:
        st.markdown(f"""
        <div class="now-playing-card">
            <div style="font-size:0.85em; color:#CEDC00; font-weight:700; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">▶️ EN LECTURE</div>
            <div style="font-size:1.8em; font-weight:800; color:#F0FAF8; margin-bottom:8px;">{titre}</div>
            <div style="font-size:1em; color:#9DC5BF; margin-bottom:16px;">{tc} {f'({annee})' if annee else ''}</div>
            <div class="progress-bar-container" style="height:14px; margin-bottom:16px;">
                <div class="progress-bar-fill progress-high" style="width:{prog}%"></div>
            </div>
            <div style="display:grid; grid-template-columns: repeat(2,1fr); gap:14px;">
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase;">Début</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{debut.strftime('%H:%M:%S')}</div>
                </div>
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase;">Fin estimée</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{fin.strftime('%H:%M') if fin else 'Inconnue'}</div>
                </div>
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase;">Durée</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{format_minutes(duree)}</div>
                </div>
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase;">Progression</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{prog}%</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

def page_dashboard(utz):
    if bloc_lancement(): return
    h = st.session_state["historique"]
    res = st.session_state["res"]
    stats = st.session_state["stats"]
    doub = st.session_state["doub"]
    pb = st.session_state["pb"]
    th = sum(m["duree"] for m in h["films_det"])/60 + sum(e["duree"] for e in h["ep_det"])/60
    total_items = sum(s["total"] for s in stats)

    st.subheader("📊 Vue d'ensemble")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("🎬 Films", h["nb_films"])
    c2.metric("📺 Séries", h["nb_series"])
    c3.metric("🎞️ Épisodes", h["nb_ep"])
    c4.metric("⏱️ Temps total", format_duree(th))

    st.divider()
    st.subheader("⚠️ État du nettoyage")
    c5,c6,c7,c8 = st.columns(4)
    with c5:
        with st.container(border=True):
            st.markdown("#### 👻 Fantômes")
            if len(pb) > 0:
                pct = round(len(pb)/max(len(pb)+h["nb_vf"]+h["nb_ep"],1)*100,1)
                st.metric("Nombre", len(pb), delta=f"{pct}%")
                st.warning(f"{len(pb)} fantôme(s) à nettoyer")
            else:
                st.metric("Nombre",0)
                st.success("✅ Rien à nettoyer")
    with c6:
        with st.container(border=True):
            st.markdown("#### 🔁 Doublons")
            if len(doub) > 0:
                pct = round(len(doub)/max(total_items,1)*100,1)
                st.metric("Nombre", len(doub), delta=f"{pct}%")
                st.warning(f"{len(doub)} doublon(s)")
            else:
                st.metric("Nombre",0)
                st.success("✅ Aucun doublon")
    with c7:
        with st.container(border=True):
            st.markdown("#### 🧹 Déjà vus")
            if len(res) > 0:
                pct = round(len(res)/max(total_items,1)*100,1)
                st.metric("Nombre", len(res), delta=f"{pct}%")
                st.warning(f"{len(res)} contenu(s) vus dans les listes")
            else:
                st.metric("Nombre",0)
                st.success("✅ Listes à jour")
    with c8:
        with st.container(border=True):
            st.markdown("#### 🚀 Nettoyage auto")
            st.write("Tout nettoyer en 1 clic")
            if st.button("🧹 Tout nettoyer", type="primary", use_container_width=True):
                st.session_state["conf_tout"] = True
                st.rerun()
            if st.session_state.get("conf_tout"):
                st.warning("⚠️ Supprime tous les déjà vus et tous les fantômes.")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Confirmer", type="primary"):
                        with st.spinner("Nettoyage..."):
                            sup_selection(st.session_state["access_token"], res)
                            sup_playback(st.session_state["access_token"], pb)
                        st.success(f"✅ Nettoyage terminé")
                        del st.session_state["conf_tout"]
                        time.sleep(2)
                        for k in ["res","stats","doub","doub_det","pb","np"]:
                            st.session_state.pop(k, None)
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        del st.session_state["conf_tout"]
                        st.rerun()

def page_nettoyage(utz):
    if bloc_lancement(): return
    res = st.session_state["res"]
    msg = "msg_vus"
    if st.session_state.get(msg):
        st.success(st.session_state[msg])
        del st.session_state[msg]
    st.subheader("🧹 Nettoyage des listes")
    st.caption("Retire les contenus déjà vus. La colonne **Ajouté après visionnage** identifie les contenus que tu as ajoutés *après* les avoir vus pour les revoir.")
    if not res:
        st.success("Tes listes sont à jour ! 🎉")
    else:
        st.write(f"**{len(res)}** contenu(s) déjà vu(s) détecté(s).")
        tab = pd.DataFrame(res)
        ta = tab[["type","titre","annee","vues","dernier","ajoute_apres","liste"]].copy()
        ta["dernier"] = pd.to_datetime(ta["dernier"]).dt.tz_convert(utz).dt.strftime("%d/%m/%Y %H:%M")
        ta.insert(0,"Sel",False)
        ta.columns = ["Sel","Type","Titre","Année","Vues","Dernier","Ajouté après visionnage","Liste"]
        ed = st.data_editor(ta, use_container_width=True, hide_index=True, disabled=["Type","Titre","Année","Vues","Dernier","Ajouté après visionnage","Liste"], key="ed_vus")
        nb = int(ed["Sel"].sum())
        if nb:
            conf = "conf_vus"
            if not st.session_state.get(conf, False):
                if st.button(f"🗑️ Supprimer {nb} élément(s)", type="primary"):
                    st.session_state[conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer la suppression ?")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Oui"):
                        idx = ed[ed["Sel"]].index
                        items = [res[i] for i in idx]
                        with st.spinner("Suppression..."):
                            sup_selection(st.session_state["access_token"], items)
                        st.session_state[conf] = False
                        st.session_state[msg] = f"✅ {len(items)} élément(s) supprimé(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[conf] = False
                        st.rerun()
    st.divider()
    st.subheader("Pourcentage de nettoyage par liste")
    df = pd.DataFrame(st.session_state["stats"])
    df["% nettoyable"] = (df["vus"]/df["total"].replace(0,1)*100).round(1)
    st.bar_chart(df.set_index("nom")["% nettoyable"], color="#CEDC00")

def page_doublons(utz):
    if bloc_lancement(): return
    dd = st.session_state["doub_det"]
    msg = "msg_d"
    if st.session_state.get(msg):
        st.success(st.session_state[msg])
        del st.session_state[msg]
    st.subheader("🔍 Recherche de doublons")
    st.caption("Trouve les contenus présents dans plusieurs listes.")
    if not dd:
        st.success("Aucun doublon !")
    else:
        st.write(f"**{len(st.session_state['doub'])}** doublon(s).")
        tab = pd.DataFrame(dd)
        ta = tab[["type","titre","annee","liste"]].copy()
        ta.insert(0,"Sel",False)
        ta.columns = ["Sel","Type","Titre","Année","Liste"]
        ed = st.data_editor(ta, use_container_width=True, hide_index=True, disabled=["Type","Titre","Année","Liste"], key="ed_d")
        nb = int(ed["Sel"].sum())
        if nb:
            conf = "conf_d"
            if not st.session_state.get(conf, False):
                if st.button(f"🗑️ Retirer {nb} élément(s)", type="primary"):
                    st.session_state[conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer ?")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Oui"):
                        idx = ed[ed["Sel"]].index
                        items = [dd[i] for i in idx]
                        with st.spinner("Suppression..."):
                            sup_selection(st.session_state["access_token"], items)
                        st.session_state[conf] = False
                        st.session_state[msg] = f"✅ {len(items)} retiré(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[conf] = False
                        st.rerun()

def page_fantomes(utz):
    if bloc_lancement(): return
    pb = st.session_state["pb"]
    msg = "msg_pb"
    if st.session_state.get(msg):
        st.success(st.session_state[msg])
        del st.session_state[msg]
    st.subheader("👻 Progression Fantôme")
    st.caption("Supprime les entrées bloquées dans 'Continuer à regarder'.")
    st.divider()
    if not pb:
        st.success("Aucune progression en cours.")
    else:
        tout = st.checkbox("Tout sélectionner")
        sels = {}
        for it in pb:
            p = it["prog"]
            cls = "progress-low" if p<30 else "progress-mid" if p<80 else "progress-high"
            df = formater_date(it["dernier"], utz)
            ic = "🎬" if it["type"]=="Film" else "📺"
            img = image_tmdb(it.get("tmdb"), "movie" if it["type"]=="Film" else "tv")
            with st.container():
                cc, cimg, cd = st.columns([0.05, 0.12, 0.83])
                with cc:
                    sels[it["pid"]] = st.checkbox("", value=tout, key=f"c_{it['pid']}", label_visibility="collapsed")
                with cimg:
                    if img:
                        st.image(img, use_container_width=True)
                with cd:
                    st.markdown(f"""
                    <div class="ghost-card">
                        <div class="ghost-title">{ic} {it['titre']} {f'({it["annee"]})' if it["annee"] else ''}</div>
                        <div class="ghost-meta">{it['type']} • {p}% visionné • 🕒 {df}</div>
                        <div class="progress-bar-container">
                            <div class="progress-bar-fill {cls}" style="width:{p}%"></div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        st.divider()
        ids = [pid for pid,s in sels.items() if s]
        if ids:
            conf = "conf_pb"
            if not st.session_state.get(conf, False):
                if st.button(f"🗑️ Supprimer {len(ids)} progression(s)", type="primary"):
                    st.session_state[conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer ?")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Oui"):
                        items = [p for p in pb if p["pid"] in ids]
                        with st.spinner("Suppression..."):
                            sup_playback(st.session_state["access_token"], items)
                        st.session_state[conf] = False
                        st.session_state[msg] = f"✅ {len(items)} fantôme(s) supprimé(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[conf] = False
                        st.rerun()
        else:
            st.info("Coche les éléments à supprimer.")

def page_calendrier(utz):
    if bloc_lancement(): return
    st.subheader("📅 Calendrier des sorties")
    st.info("🚧 Prochainement.")

def page_succes(utz):
    if bloc_lancement(): return
    st.subheader("🏆 Succès")
    st.caption("Tes badges de grand fan de cinéma et de séries. Tu débloques des badges automatiquement au fil de ton visionnage.")
    h = st.session_state["historique"]
    total_h = sum(m["duree"] for m in h["films_det"])/60 + sum(e["duree"] for e in h["ep_det"])/60
    total_jours = total_h / 24
    total_films = h["nb_films"]
    total_eps = h["nb_ep"]
    total_vues = h["nb_vf"] + h["nb_ep"]

    # Calculs pour badges de diversite
    films_df = pd.DataFrame(h["films_det"])
    eps_df = pd.DataFrame(h["ep_det"])
    genres_diff = set()
    annees_diff = set()
    if not films_df.empty:
        for g in films_df["genre"].str.split(", "):
            genres_diff.update([x for x in g if x != "Inconnu"])
        for a in films_df["annee"].dropna():
            try: annees_diff.add(int(a))
            except: pass
    if not eps_df.empty:
        for g in eps_df["genre"].str.split(", "):
            genres_diff.update([x for x in g if x != "Inconnu"])
        for a in eps_df["annee"].dropna():
            try: annees_diff.add(int(a))
            except: pass
    nb_genres = len(genres_diff)
    nb_annees = len(annees_diff)

    # Marathons detectes
    rec_jour = 0
    if not eps_df.empty:
        ep = eps_df.copy()
        ep["date_dt"] = pd.to_datetime(ep["date"], utc=True).dt.tz_convert(utz)
        ep["jour"] = ep["date_dt"].dt.date
        rec_jour = ep.groupby(["jour","serie"]).size().max() or 0

    # Nocturne : visionnage entre 00h et 5h
    vues_nuit = 0
    for df_tmp in [films_df, eps_df]:
        if not df_tmp.empty:
            dt = pd.to_datetime(df_tmp["date"], utc=True).dt.tz_convert(utz)
            vues_nuit += ((dt.dt.hour >= 0) & (dt.dt.hour < 5)).sum()

    # Series avec au moins 10 episodes vus
    series_10ep = 0
    series_50ep = 0
    series_100ep = 0
    series_200ep = 0
    nb_jours_diff = 0
    note_coup_coeur = False
    nuit_blanche = False
    if not eps_df.empty:
        par_serie = eps_df.groupby("serie").size()
        series_10ep = int((par_serie >= 10).sum())
        series_50ep = int((par_serie >= 50).sum())
        series_100ep = int((par_serie >= 100).sum())
        series_200ep = int((par_serie >= 200).sum())
    # Jours differents de visionnage
    toutes_dt = []
    if not films_df.empty:
        toutes_dt.extend(pd.to_datetime(films_df["date"], utc=True).dt.tz_convert(utz).tolist())
    if not eps_df.empty:
        toutes_dt.extend(pd.to_datetime(eps_df["date"], utc=True).dt.tz_convert(utz).tolist())
    if toutes_dt:
        s_dt = pd.Series(toutes_dt)
        nb_jours_diff = s_dt.dt.date.nunique()
        # Nuit blanche : plus de 6h entre 0h et 6h sur une meme date (nuit)
        dn = s_dt[(s_dt.dt.hour >= 0) & (s_dt.dt.hour < 6)]
        if not dn.empty:
            nuit_blanche = bool((dn.groupby([dn.dt.date]).count() >= 3).any())  # au moins 3 contenus dans une nuit
        # Coup de coeur : note >= 9
        if not films_df.empty and (films_df["note"] >= 9).any():
            note_coup_coeur = True
        if not eps_df.empty and (eps_df["note"] >= 9).any():
            note_coup_coeur = True

    # Liste complete des badges : (id, emoji, titre, desc, condition bool, progression pct pour les lock)
    badges = [
        # -- Paliers de temps --
        ("h1",    "🌱", "Première heure",       "Tu as regardé ton premier contenu",                                  total_h >= 1,        min(total_h/1*100,100)),
        ("h10",   "⏳", "Dix heures",           "10 heures de visionnage cumulées",                                    total_h >= 10,       min(total_h/10*100,100)),
        ("h24",   "⏰", "Un jour complet",      "24h passées devant des films et séries",                              total_h >= 24,       min(total_h/24*100,100)),
        ("h168",  "📅", "Une semaine entière",  "Tu as passé une semaine entière de visionnage (168h)",                total_h >= 168,      min(total_h/168*100,100)),
        ("h720",  "🗓️", "Un mois de binge",     "30 jours complets de visionnage (720h)",                              total_h >= 720,      min(total_h/720*100,100)),
        ("h2160", "🏁", "Trimestre sur écran",  "3 mois entiers à regarder des contenus (2160h)",                      total_h >= 2160,     min(total_h/2160*100,100)),
        ("h8760", "👑", "Une année d'écran",    "1 AN de visionnage cumulé (8760h) — statut de légende",               total_h >= 8760,     min(total_h/8760*100,100)),
        ("h26k",  "⚜️", "Empereur du canapé",    "3 ANS de visionnage — tu vis sur Trakt (26 280h)",                    total_h >= 26280,    min(total_h/26280*100,100)),
        ("h43k",  "🧙", "Archiviste ultime",     "5 ANS entiers de visionnage — tu as vu presque tout (43 800h)",       total_h >= 43800,    min(total_h/43800*100,100)),

        # -- Films --
        ("f1",    "🎬", "Premier film",         "Ton premier film vu",                                                 total_films >= 1,    min(total_films/1*100,100)),
        ("f10",   "🎞️", "Dix films",            "10 films différents vus",                                             total_films >= 10,   min(total_films/10*100,100)),
        ("f50",   "🎥", "Cinéphile",            "50 films différents vus",                                             total_films >= 50,   min(total_films/50*100,100)),
        ("f100",  "🍿", "Cent films",           "100 films vus !",                                                     total_films >= 100,  min(total_films/100*100,100)),
        ("f250",  "🏅", "Amoureux du 7ème art", "250 films vus — une belle cinémathèque",                              total_films >= 250,  min(total_films/250*100,100)),
        ("f500",  "🎭", "Véritable cinéphile",  "500 films différents vus",                                            total_films >= 500,  min(total_films/500*100,100)),
        ("f1000", "🎪", "Maître du grand écran","1000 films, impressionnant !",                                        total_films >= 1000, min(total_films/1000*100,100)),
        ("f2000", "🏛️", "Bibliothèque vivante", "2000 films — ta culture ciné est immense",                            total_films >= 2000, min(total_films/2000*100,100)),
        ("f5000", "🧠", "Encyclopédie du cinéma","5000 films différents — tu devrais écrire un blog",                   total_films >= 5000, min(total_films/5000*100,100)),

        # -- Series --
        ("s1",    "📺", "Premier épisode",      "Ton tout premier épisode vu",                                         total_eps >= 1,      min(total_eps/1*100,100)),
        ("s10",   "📡", "Dix épisodes",         "10 épisodes vus",                                                     total_eps >= 10,     min(total_eps/10*100,100)),
        ("s100",  "📶", "Cent épisodes",        "100 épisodes vus",                                                    total_eps >= 100,    min(total_eps/100*100,100)),
        ("s500",  "💻", "Accro aux séries",     "500 épisodes — les séries n'ont plus de secrets pour toi",            total_eps >= 500,    min(total_eps/500*100,100)),
        ("s1000", "🔥", "Mille épisodes",       "1000 épisodes ! Une belle performance",                               total_eps >= 1000,   min(total_eps/1000*100,100)),
        ("s2500", "🚀", "Marathonien TV",       "2500 épisodes — tu vis littéralement devant les séries",              total_eps >= 2500,   min(total_eps/2500*100,100)),
        ("s5000", "🏯", "Forteresse de canapé", "5000 épisodes — rien ne t'arrête",                                    total_eps >= 5000,   min(total_eps/5000*100,100)),
        ("s10k",  "🌋", "Dix mille épisodes",   "10 000 épisodes. Juste... wow.",                                      total_eps >= 10000,  min(total_eps/10000*100,100)),
        ("s25k",  "🌌", "Univers télévisuel",   "25 000 épisodes — tu as plus vu de séries que la plupart des gens",   total_eps >= 25000,  min(total_eps/25000*100,100)),

        # -- Series suivies --
        ("sv1",   "✅", "Une série terminée",   "Au moins 10 épisodes vus d'une même série",                           series_10ep >= 1,    min(series_10ep/1*100,100)),
        ("sv5",   "💪", "Cinq séries suivies",  "Tu as vu 10+ épisodes de 5 séries différentes",                       series_10ep >= 5,    min(series_10ep/5*100,100)),
        ("sv10",  "📚", "Dix séries suivies",   "10 séries dont tu as vu plus de 10 épisodes",                         series_10ep >= 10,   min(series_10ep/10*100,100)),
        ("sv25",  "🗂️", "Collectionneur",      "25 séries différentes avec 10+ épisodes chacune",                     series_10ep >= 25,   min(series_10ep/25*100,100)),
        ("sv50",  "💎", "Fan inconditionnel",   "Une série avec plus de 50 épisodes vus",                              series_50ep >= 1,    min(series_50ep/1*100,100)),
        ("sv100", "💍", "Relation sérieuse",    "Une série avec plus de 100 épisodes vus — un investissement",         series_100ep >= 1,   min(series_100ep/1*100,100)),
        ("sv200", "👑", "Série culte",          "Une série avec plus de 200 épisodes — un compagnon de vie",           series_200ep >= 1,   min(series_200ep/1*100,100)),

        # -- Marathons --
        ("mar4",  "🏃", "Marathonien",          "4+ épisodes d'une même série en 1 jour",                              rec_jour >= 4,       min(rec_jour/4*100,100)),
        ("mar8",  "⚡", "Marathon éclair",      "8+ épisodes en une seule journée",                                    rec_jour >= 8,       min(rec_jour/8*100,100)),
        ("mar12", "🚄", "Train fou",            "12+ épisodes en 1 jour — ça c'est du binge !",                        rec_jour >= 12,      min(rec_jour/12*100,100)),
        ("mar20", "🏁", "Journée sans sortir",  "20+ épisodes en 1 jour — tu n'as pas vu le soleil",                   rec_jour >= 20,      min(rec_jour/20*100,100)),

        # -- Diversite --
        ("divg",  "🌈", "Explorateur de genres","Tu as touché à au moins 10 genres différents",                        nb_genres >= 10,     min(nb_genres/10*100,100)),
        ("divg2", "🎨", "Palette complète",     "20 genres différents explorés",                                       nb_genres >= 20,     min(nb_genres/20*100,100)),
        ("diva",  "🕰️", "Voyageur temporel",    "Tu as vu des contenus de 20 années de sortie différentes",            nb_annees >= 20,     min(nb_annees/20*100,100)),
        ("diva3", "🗿", "Amateur de classiques","Des contenus de 40 années différentes — du vieux au neuf !",          nb_annees >= 40,     min(nb_annees/40*100,100)),
        ("diva6", "🏛️", "Passé et présent",     "60 années de cinéma/séries — des années 60 à aujourd'hui",            nb_annees >= 60,     min(nb_annees/60*100,100)),

        # -- Nocturne --
        ("nuit",  "🌙", "Oiseau de nuit",       "Plus de 20 visionnages entre minuit et 5h du matin",                  vues_nuit >= 20,     min(vues_nuit/20*100,100)),
        ("nuit2", "🦉", "Chouette cinéphile",   "Plus de 100 visionnages nocturnes",                                   vues_nuit >= 100,    min(vues_nuit/100*100,100)),
        ("nuit3", "🦇", "Créature de la nuit",  "Plus de 500 visionnages entre minuit et 5h",                          vues_nuit >= 500,    min(vues_nuit/500*100,100)),
        ("nuitb", "🌃", "Nuit blanche",         "Plus de 3 visionnages entre minuit et 6h sur une même nuit",           nuit_blanche,       100 if nuit_blanche else 0),

        # -- Global --
        ("all1",  "👶", "Nouveau venu",         "Ton tout premier visionnage sur Trakt",                               total_vues >= 1,     min(total_vues/1*100,100)),
        ("all100","⭐", "Cent visionnages",     "100 visionnages au total (films + épisodes)",                         total_vues >= 100,   min(total_vues/100*100,100)),
        ("all1k", "🌟", "Mille visionnages",    "1000 visionnages, belle courbe de progression !",                     total_vues >= 1000,  min(total_vues/1000*100,100)),
        ("all5k", "💫", "Cinq mille",           "5000 visionnages, une véritable habitude",                            total_vues >= 5000,  min(total_vues/5000*100,100)),
        ("all10k","🪽", "Dix mille",            "10 000 visionnages — c'est de la passion à ce niveau",                total_vues >= 10000, min(total_vues/10000*100,100)),
        ("all25k","🔱", "25 000",               "25 000 visionnages, tu es un abonné historique",                      total_vues >= 25000, min(total_vues/25000*100,100)),
        ("all50k","🌠", "50 000",               "50 000 visionnages, la légende est en marche",                        total_vues >= 50000, min(total_vues/50000*100,100)),

        # -- Rythme --
        ("ryth",  "📆", "Un an de fidélité",    "Visionnages répartis sur au moins 365 jours différents",              nb_jours_diff >= 365, min(nb_jours_diff/365*100,100)),
        ("ryth2", "🗓️", "Deux ans de fidélité", "Contenus vus sur plus de 730 jours différents",                       nb_jours_diff >= 730, min(nb_jours_diff/730*100,100)),
        ("note9", "💯", "Critique exigeant",    "Au moins un contenu noté 9 ou 10 — tu as eu un coup de cœur",        note_coup_coeur,    100 if note_coup_coeur else 0),
    ]

    obtenus = [b for b in badges if b[4]]
    locks   = [b for b in badges if not b[4]]

    st.markdown(f"#### 🎖️ Badges obtenus ({len(obtenus)}/{len(badges)})")
    if obtenus:
        cols = st.columns(min(4, len(obtenus)))
        for i,(_,em,titre,desc,_,_) in enumerate(obtenus):
            with cols[i%4]:
                st.markdown(f"""
                <div class="badge-obtenu">
                    <div class="emoji">{em}</div>
                    <div class="titre">{titre}</div>
                    <div class="desc">{desc}</div>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.info("Continue de regarder des contenus pour gagner tes premiers badges !")

    if locks:
        st.divider()
        st.markdown(f"#### 🔒 Prochains badges à décrocher ({len(locks)})")
        st.caption("Voici les badges que tu n'as pas encore, triés avec les plus proches en premier.")
        # Trier les locks par progression decroissante (les plus proches d'etre debloques d'abord)
        locks = sorted(locks, key=lambda b: -b[5])
        cols = st.columns(min(4, len(locks)))
        for i,(_,em,titre,desc,_,prog) in enumerate(locks):
            with cols[i%4]:
                st.markdown(f"""
                <div class="badge-lock">
                    <div class="emoji">{em}</div>
                    <div class="titre">{titre}</div>
                    <div class="desc">{desc}</div>
                    <div class="prog-badge"><div class="prog-badge-fill" style="width:{round(prog,1)}%"></div></div>
                </div>
                """, unsafe_allow_html=True)


def page_stats(utz):
    if bloc_lancement(): return
    st.subheader("📊 Statistiques détaillées")
    st.caption("Toutes tes données de visionnage. Les valeurs sont exprimées en **heures** sauf indication contraire. Note : un contenu avec plusieurs genres est compté dans chaque genre, la somme peut donc être supérieure au total.")
    h = st.session_state["historique"]
    films = pd.DataFrame(h["films_det"])
    eps = pd.DataFrame(h["ep_det"])

    f1,f2,f3 = st.columns(3)
    with f1:
        tc = st.selectbox("Type de contenu", ["Tous","Films","Séries"])
    with f2:
        periode = st.selectbox("Période", ["Tout","Cette année","12 derniers mois","6 derniers mois","Ce mois-ci","Mois dernier","Aujourd'hui","Période personnalisée"], index=0)
    with f3:
        genres = set()
        if tc in ["Tous","Films"] and not films.empty:
            for g in films["genre"].str.split(", "):
                genres.update([x for x in g if x != "Inconnu"])
        if tc in ["Tous","Séries"] and not eps.empty:
            for g in eps["genre"].str.split(", "):
                genres.update([x for x in g if x != "Inconnu"])
        genre = st.selectbox("Genre", ["Tous"] + sorted(genres))

    toutes_dates = []
    for df_tmp in [films, eps]:
        if not df_tmp.empty:
            toutes_dates.extend(pd.to_datetime(df_tmp["date"], utc=True).dt.tz_convert(utz).tolist())
    if periode == "Période personnalisée" and toutes_dates:
        dates_df = pd.DataFrame({"date": toutes_dates})
        dates_df["ma"] = dates_df["date"].dt.strftime("%m-%Y")
        mois_dispo = sorted(dates_df["ma"].unique(), key=lambda x: (int(x.split("-")[1]), int(x.split("-")[0])))
        periode_mois = st.select_slider("Sélectionne la période (mois)", options=mois_dispo, value=(mois_dispo[0], mois_dispo[-1]))

    dfs = []
    if tc in ["Tous","Films"] and not films.empty:
        df = films.copy()
        df["type"] = "Film"
        df["date_dt"] = pd.to_datetime(df["date"], utc=True).dt.tz_convert(utz)
        dfs.append(df)
    if tc in ["Tous","Séries"] and not eps.empty:
        df = eps.copy()
        df["type"] = "Épisode"
        df["titre"] = df["serie"]
        df["date_dt"] = pd.to_datetime(df["date"], utc=True).dt.tz_convert(utz)
        dfs.append(df)
    if not dfs:
        st.info("Aucune donnée.")
        return
    df = pd.concat(dfs, ignore_index=True)
    mt = datetime.now(utz)

    if periode == "Période personnalisée" and toutes_dates:
        d_deb = datetime.strptime(periode_mois[0], "%m-%Y").replace(tzinfo=utz)
        d_fin = datetime.strptime(periode_mois[1], "%m-%Y").replace(day=28) + timedelta(days=4)
        d_fin = d_fin.replace(tzinfo=utz)
        df = df[(df["date_dt"] >= d_deb) & (df["date_dt"] <= d_fin)]
    else:
        df = appliquer_filtres_periode(df, mt, periode)
    if genre != "Tous":
        df = df[df["genre"].str.contains(genre, na=False)]
    if df.empty:
        st.warning("Aucun résultat.")
        return

    df["duree_h"] = df["duree"].fillna(40)/60
    th = df["duree_h"].sum()

    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Nombre de visionnages", len(df))
    m2.metric("Temps de visionnage", format_duree(th))
    nm = df[df["note"]>0]["note"].mean()
    m3.metric("Note moyenne /10", f"{round(nm,1)}" if pd.notna(nm) else "-")
    nb_jours = max((df["date_dt"].max() - df["date_dt"].min()).days +1, 1)
    m4.metric("Moyenne par jour", f"{round(len(df)/nb_jours,1)}")
    marathon = df.groupby(df["date_dt"].dt.date).size().max()
    m5.metric("Record en 1 jour", f"{marathon}")

    # Marathons
    marathons = pd.DataFrame()
    if tc in ["Tous","Séries"] and not eps.empty:
        ej = eps.copy()
        ej["date_dt"] = pd.to_datetime(ej["date"], utc=True).dt.tz_convert(utz)
        if periode == "Période personnalisée" and toutes_dates:
            ej = ej[(ej["date_dt"] >= d_deb) & (ej["date_dt"] <= d_fin)]
        else:
            ej = appliquer_filtres_periode(ej, mt, periode)
        if genre != "Tous":
            ej = ej[ej["genre"].str.contains(genre, na=False)]
        if not ej.empty:
            ej["jour"] = ej["date_dt"].dt.date
            cmp = ej.groupby(["jour","serie"]).size().reset_index(name="nb")
            marathons = cmp[cmp["nb"] >=4].sort_values("nb", ascending=False)
    if not marathons.empty:
        st.divider()
        with st.container(border=True):
            st.markdown("#### 🏆 Marathons (4+ épisodes en 1 jour)")
            for _, row in marathons.head(5).iterrows():
                st.write(f"📅 **{row['jour'].strftime('%d/%m/%Y')}** : {row['nb']} épisodes de **{row['serie']}**")

    # Plateformes
    if tc in ["Tous","Séries"] and not eps.empty:
        ep_n = eps.copy()
        ep_n["date_dt"] = pd.to_datetime(ep_n["date"], utc=True).dt.tz_convert(utz)
        if periode == "Période personnalisée" and toutes_dates:
            ep_n = ep_n[(ep_n["date_dt"] >= d_deb) & (ep_n["date_dt"] <= d_fin)]
        else:
            ep_n = appliquer_filtres_periode(ep_n, mt, periode)
        if genre != "Tous":
            ep_n = ep_n[ep_n["genre"].str.contains(genre, na=False)]
        if not ep_n.empty and "network" in ep_n.columns:
            st.divider()
            with st.container(border=True):
                st.markdown("#### 📺 Plateformes les plus regardées")
                pf = ep_n["network"].value_counts().head(5)
                cols = st.columns(min(len(pf),5))
                for i,(n,nb_pf) in enumerate(pf.items()):
                    if n != "Inconnu":
                        cols[i].metric(n, f"{nb_pf} ép.")

    st.divider()

    # Heures par mois
    df["mois"] = df["date_dt"].dt.strftime("%m-%Y")
    h_mois = df.groupby("mois")["duree_h"].sum().round(1).sort_index()
    opt_m = {"title":{"text":"Heures par mois","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis","formatter":"{b} : {c}h"},"backgroundColor":"transparent","textStyle":{"color":"#F0FAF8"},"xAxis":{"type":"category","data":list(h_mois.index),"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(h_mois.values),"type":"line","smooth":True,"lineStyle":{"color":"#CEDC00","width":3},"areaStyle":{"color":"rgba(206,220,0,0.1)"},"itemStyle":{"color":"#CEDC00"}}]}
    st_echarts(opt_m, height="350px")

    g1,g2 = st.columns(2)
    with g1:
        # Genres : par nombre de contenus pour éviter incohérences de double comptage
        genres_n = {}
        for lg in df["genre"].str.split(", "):
            for g in lg:
                if g and g != "Inconnu":
                    genres_n[g] = genres_n.get(g,0) + 1
        opt_g = {"title":{"text":"Genres les plus regardés (nombre de contenus)","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"item"},"backgroundColor":"transparent","legend":{"bottom":0,"textStyle":{"color":"#9DC5BF"}},"series":[{"type":"pie","radius":["40%","70%"],"data":[{"name":k,"value":v} for k,v in sorted(genres_n.items(), key=lambda x:-x[1])[:8]],"itemStyle":{"borderRadius":8,"borderColor":"#042E2B","borderWidth":2},"label":{"color":"#F0FAF8"}}],"color":["#00A392","#CEDC00","#00C7B3","#A3B300","#00524B","#869400","#125A54","#E8F064"]}
        st_echarts(opt_g, height="400px")
    with g2:
        df["h"] = df["date_dt"].dt.hour
        hh = df.groupby("h")["duree_h"].sum().reindex(range(24), fill_value=0).round(1)
        opt_h = {"title":{"text":"Par heure de la journée","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis","formatter":"{b} : {c}h"},"backgroundColor":"transparent","xAxis":{"type":"category","data":[f"{h}h" for h in range(24)],"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(hh.values),"type":"bar","itemStyle":{"color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,"colorStops":[{"offset":0,"color":"#00A392"},{"offset":1,"color":"#00524B"}]},"borderRadius":[4,4,0,0]}}]}
        st_echarts(opt_h, height="400px")

    g3,g4 = st.columns(2)
    with g3:
        jours = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"]
        df["jsem"] = df["date_dt"].dt.weekday
        hj = df.groupby("jsem")["duree_h"].sum().reindex(range(7), fill_value=0).round(1)
        opt_j = {"title":{"text":"Par jour de la semaine","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis","formatter":"{b} : {c}h"},"backgroundColor":"transparent","xAxis":{"type":"category","data":jours,"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(hj.values),"type":"bar","itemStyle":{"color":"#CEDC00","borderRadius":[4,4,0,0]}}]}
        st_echarts(opt_j, height="400px")
    with g4:
        # Années de sortie
        if "annee" in df.columns:
            df_annees = df.dropna(subset=["annee"])
            df_annees["annee"] = df_annees["annee"].astype(int)
            annees = df_annees.groupby("annee")["duree_h"].sum().round(1).sort_index()
            opt_a = {"title":{"text":"Par année de sortie","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis","formatter":"{b} : {c}h"},"backgroundColor":"transparent","xAxis":{"type":"category","data":list(annees.index.astype(str)),"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(annees.values),"type":"bar","itemStyle":{"color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,"colorStops":[{"offset":0,"color":"#00A392"},{"offset":1,"color":"#00524B"}]},"borderRadius":[4,4,0,0]}}]}
            st_echarts(opt_a, height="400px")

    g5,g6 = st.columns(2)
    with g5:
        rt = df.groupby("type")["duree_h"].sum().round(1)
        opt_t = {"title":{"text":"Films vs Séries","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"item","formatter":"{b} : {c}h ({d}%)"},"backgroundColor":"transparent","legend":{"bottom":0,"textStyle":{"color":"#9DC5BF"}},"series":[{"type":"pie","radius":["40%","70%"],"data":[{"value":v,"name":k} for k,v in rt.items()],"itemStyle":{"borderRadius":8,"borderColor":"#042E2B","borderWidth":2},"label":{"color":"#F0FAF8"}}],"color":["#00A392","#CEDC00"]}
        st_echarts(opt_t, height="400px")

    with st.expander("📋 Détail des visionnages"):
        df_aff = df[["date_dt","type","titre","annee","genre","duree","note"]].copy()
        df_aff["date_dt"] = df_aff["date_dt"].dt.strftime("%d/%m/%Y %H:%M")
        df_aff["duree"] = df_aff["duree"].apply(lambda x: format_minutes(x) if x>0 else "-")
        df_aff.columns = ["Date","Type","Titre","Année","Genres","Durée","Note"]
        st.dataframe(df_aff, use_container_width=True, hide_index=True)

def construire_profil(histo, utz):
    """Construit un profil de gouts depuis l'historique : genres preferes, reseaux, notes."""
    films = pd.DataFrame(histo["films_det"])
    eps = pd.DataFrame(histo["ep_det"])
    genres_score = {}
    reseaux_score = {}
    decennies_score = {}
    notes_par_genre = {}
    total_duree = 0.0
    # Films
    if not films.empty:
        for _, r in films.iterrows():
            d = r.get("duree", 0) or 0
            if not d: d = 100
            total_duree += d
            note = r.get("note", 0) or 0
            for g in str(r.get("genre","")).split(", "):
                if g and g != "Inconnu":
                    genres_score[g] = genres_score.get(g,0) + d
                    if note > 0:
                        notes_par_genre[g] = notes_par_genre.get(g, []); notes_par_genre[g].append(note)
            try:
                an = int(r.get("annee")) if r.get("annee") else None
                if an:
                    dec = (an//10)*10
                    decennies_score[dec] = decennies_score.get(dec,0) + d
            except: pass
    # Series
    if not eps.empty:
        for _, r in eps.iterrows():
            d = r.get("duree", 0) or 0
            if not d: d = 40
            total_duree += d
            note = r.get("note", 0) or 0
            for g in str(r.get("genre","")).split(", "):
                if g and g != "Inconnu":
                    genres_score[g] = genres_score.get(g,0) + d
                    if note > 0:
                        notes_par_genre[g] = notes_par_genre.get(g, []); notes_par_genre[g].append(note)
            net = r.get("network")
            if net and net != "Inconnu":
                reseaux_score[net] = reseaux_score.get(net,0) + d
            try:
                an = int(r.get("annee")) if r.get("annee") else None
                if an:
                    dec = (an//10)*10
                    decennies_score[dec] = decennies_score.get(dec,0) + d
            except: pass
    # Normaliser
    def normaliser(d):
        if not d: return {}
        m = max(d.values()) if d else 1
        return {k: v/m*100 for k,v in d.items()}
    note_moy_genre = {k: sum(v)/len(v) for k,v in notes_par_genre.items() if v}
    return {
        "genres": normaliser(genres_score),
        "reseaux": normaliser(reseaux_score),
        "decennies": normaliser(decennies_score),
        "note_genre": note_moy_genre,
        "total_h": total_duree/60,
        "date_plus_recent": pd.Timestamp.now(tz=utz)
    }

def evaluer_contenu(item, profil, maintenant_tz):
    """
    Retourne un score 0-100 de recommandation, et des tags/alertes.
    """
    if item["type"] == "movie":
        med = item["movie"]
        duree = (med.get("runtime") or 0)
        genres = med.get("genres") or []
        annee = med.get("year")
        note = med.get("rating") or 0
        titre = med.get("title","")
        votes = med.get("votes") or 0
        status_txt = "ended"
        nb_aired = 1
    elif item["type"] == "show":
        med = item["show"]
        ep_dur = med.get("runtime") or 40
        nb_aired = med.get("aired_episodes") or 0
        status_txt = med.get("status") or ""
        if nb_aired > 0:
            duree = ep_dur * nb_aired
        else:
            duree = ep_dur * 50
        genres = med.get("genres") or []
        annee = med.get("year")
        note = med.get("rating") or 0
        titre = med.get("title","")
        votes = med.get("votes") or 0
    else:
        return None

    score = 0.0
    raisons = []
    points_noirs = []

    # 1. Correspondance avec les genres preferes (40 points max)
    if genres:
        g_match = sum(profil["genres"].get(g,0) for g in genres) / max(len(genres),1)
        score += min(g_match * 0.4, 40)
        if any(profil["genres"].get(g,0) > 60 for g in genres):
            raisons.append("Genre que tu adores")
        elif all(profil["genres"].get(g,0) < 10 for g in genres):
            points_noirs.append("Genre que tu regardes rarement")

    # 2. Note moyenne du contenu (25 points max)
    if note > 0:
        score += min((note/10)*25, 25)
        if note >= 9.0:
            raisons.append(f"Pépite critique ({note:.1f}/10)")
        elif note >= 8.0:
            raisons.append(f"Très bien noté ({note:.1f}/10)")
        elif note < 5.0:
            points_noirs.append(f"Note faible ({note:.1f}/10)")
    if votes >= 100000:
        raisons.append("Très populaire")
    elif votes >= 10000:
        raisons.append("Apprécié du public")

    # 3. Recence / classiques
    if annee:
        age = maintenant_tz.year - annee
        if age <= 1:
            score += 18
            raisons.append("Toute dernière sortie")
        elif age <= 2:
            score += 15
            raisons.append("Sortie récente")
        elif age <= 10:
            score += 8
        elif age >= 40:
            if profil["decennies"].get((annee//10)*10,0) > 50:
                score += 12
                raisons.append("Classique qui correspond à tes goûts")
            else:
                score += 1
        if age >= 30 and note >= 7.5:
            raisons.append("Classique incontournable")

    # 4. Duree
    if item["type"] == "movie":
        if duree and duree <= 90:
            score += 12
            raisons.append("Film très court (< 1h30)")
        elif duree and duree <= 100:
            score += 10
            raisons.append("Film rapide (< 1h40)")
        elif duree and duree <= 120:
            score += 5
        elif duree and duree >= 200:
            points_noirs.append(f"Film très long ({format_minutes(duree)})")
            score -= 8
        elif duree and duree >= 160:
            points_noirs.append(f"Film long ({format_minutes(duree)})")
            score -= 3
    else:
        if nb_aired <= 6:
            score += 10
            raisons.append("Mini-série rapide à finir")
        elif nb_aired <= 13:
            score += 7
            raisons.append("Saison courte (1 saison)")
        elif nb_aired <= 25:
            score += 5
            raisons.append("Série courte")
        elif nb_aired >= 300:
            score -= 12
            points_noirs.append(f"Gros engagement ({nb_aired} épisodes)")
        elif nb_aired >= 150:
            score -= 6
            points_noirs.append(f"Série longue ({nb_aired} épisodes)")
        elif nb_aired >= 200:
            score -= 4
            points_noirs.append(f"Engagement important ({nb_aired} épisodes)")
        if status_txt == "ended":
            raisons.append("Série terminée (pas d'attente)")
        elif status_txt == "returning" or status_txt == "in production":
            raisons.append("Série en cours de diffusion")
        elif status_txt == "canceled":
            points_noirs.append("Série annulée")

    # 5. Ajout dans la liste
    listed_at = item.get("_listed_at")
    anciennete_jours = None
    if listed_at:
        try:
            dt_l = pd.to_datetime(listed_at, utc=True).tz_convert(maintenant_tz.tzinfo)
            delta = maintenant_tz - dt_l.to_pydatetime()
            anciennete_jours = delta.days
            if anciennete_jours <= 7:
                score += 12
            elif anciennete_jours <= 14:
                score += 10
            elif anciennete_jours > 730:
                score -= 25
                points_noirs.append("Ajouté il y a plus de 2 ans, tu as probablement oublié")
            elif anciennete_jours > 365:
                score -= 20
                points_noirs.append("Ajouté il y a plus d'un an")
            elif anciennete_jours > 180:
                score -= 10
                points_noirs.append("Ajouté il y a longtemps")
        except:
            pass

    # Ne correspond pas a mon profil si score bas OU points noirs importants
    pas_pour_moi = (score < 35) or (len(points_noirs) >= 2)

    # Format duree
    if item["type"] == "movie":
        temps_necessaire = format_minutes(duree) if duree else "inconnu"
    else:
        heures = duree/60
        temps_necessaire = format_duree(heures) if heures > 0 else "inconnu"

    return {
        "type": "Film" if item["type"]=="movie" else "Série",
        "titre": titre,
        "annee": annee,
        "note": round(note,1) if note else None,
        "genres": ", ".join(genres) if genres else "Inconnu",
        "genres_liste": genres,
        "temps": temps_necessaire,
        "duree_min": duree,
        "score": max(0, min(round(score,1), 100)),
        "raisons": raisons,
        "averti": points_noirs,
        "pas_pour_moi": pas_pour_moi,
        "tmdb": med["ids"].get("tmdb"),
        "ajout": anciennete_jours,
        "votes": votes,
        "nb_episodes": nb_aired if item["type"] == "show" else 0,
        "status": status_txt,
    }


def page_quoi_regarder(utz):
    if bloc_lancement(): return
    st.subheader("🎯 Que regarder ?")
    st.caption("Sélectionne une liste, applique tes filtres et laisse-moi te recommander le prochain contenu à regarder selon TES goûts. Fini le scroll infini !")

    h = st.session_state["historique"]
    profil = construire_profil(h, utz)

    # Construire la liste des listes disponibles
    listes_dispo = [("👀 Liste de suivi", "watchlist")]
    for s in st.session_state["stats"]:
        if s["nom"] != "Liste de suivi":
            listes_dispo.append((f"📋 {s['nom']}", s["nom"]))

    col_l, col_t = st.columns([1,1])
    with col_l:
        choix_label = st.selectbox("📋 Liste à explorer", [l[0] for l in listes_dispo])
    with col_t:
        type_f = st.selectbox("🎞️ Type de contenu", ["Tous", "Films seulement", "Séries seulement"])
    lid_nom = dict(listes_dispo)[choix_label]

    # Recuperer les items
    at = st.session_state["access_token"]
    with st.spinner("Analyse intelligente de la liste..."):
        if lid_nom == "watchlist":
            items = recuperer_watchlist(at)
        else:
            l_id = None
            for l in recuperer_listes(at):
                if l["name"] == lid_nom:
                    l_id = l["ids"]["trakt"]; break
            if not l_id:
                st.warning("Liste introuvable."); return
            items = recuperer_contenu_liste(at, l_id)

        deja_vus_tids = set()
        for r in st.session_state["res"]:
            if not r.get("ajoute_apres", False):
                deja_vus_tids.add((r["type"], r["tid"]))

        mt = datetime.now(utz)
        resultats = []
        for it in items:
            ev = evaluer_contenu(it, profil, mt)
            if not ev: continue
            cle_type = ev["type"]
            cle_tid = it["movie"]["ids"]["trakt"] if it["type"]=="movie" else it["show"]["ids"]["trakt"]
            if (cle_type, cle_tid) in deja_vus_tids:
                continue
            ev["_raw"] = it
            resultats.append(ev)

    if not resultats:
        st.info("Aucun contenu à évaluer dans cette liste.")
        return

    # FILTRES
    tous_genres = set()
    for r in resultats:
        tous_genres.update(r["genres_liste"])
    tous_genres.discard("Inconnu"); tous_genres.discard("")

    st.markdown("#### 🔎 Filtres")
    cf1, cf2, cf3, cf4 = st.columns(4)
    with cf1:
        f_genre = st.selectbox("🎭 Genre", ["Tous"] + sorted(tous_genres))
    with cf2:
        f_note_min = st.select_slider("⭐ Note minimum", options=[0,5,6,7,7.5,8,8.5,9], value=0)
    with cf3:
        f_temps_max = st.selectbox("⏱️ Temps max à investir", ["Aucune limite", "Moins d'1h30 (film)", "Moins de 2h", "Moins de 3h", "Soirée (< 10h)", "Week-end (< 24h)"])
    with cf4:
        f_tri = st.selectbox("🔀 Trier par", ["✨ Pour moi (recommandé)", "⭐ Meilleures notes", "⏱️ Plus rapide à regarder", "🔥 Populaires", "🆕 Ajouté récemment", "📅 Nouveautés (sorties)", "🎬 Films d'abord", "📺 Séries d'abord", "🤔 Ne correspond pas à mon profil"])

    # Appliquer les filtres
    def limite_temps_ok(r):
        if f_temps_max == "Aucune limite": return True
        m = r["duree_min"]
        if f_temps_max == "Moins d'1h30 (film)" and r["type"]=="Film": return m <= 90
        if f_temps_max == "Moins de 2h" and r["type"]=="Film": return m <= 120
        if f_temps_max == "Moins de 3h": return m <= 180
        if f_temps_max == "Soirée (< 10h)": return m/60 <= 10
        if f_temps_max == "Week-end (< 24h)": return m/60 <= 24
        # si filtre film mais contenu serie => on exclut
        if f_temps_max in ["Moins d'1h30 (film)", "Moins de 2h"] and r["type"]=="Série":
            return False
        return True

    filtrés = []
    for r in resultats:
        if type_f == "Films seulement" and r["type"] != "Film": continue
        if type_f == "Séries seulement" and r["type"] != "Série": continue
        if f_genre != "Tous" and f_genre not in r["genres_liste"]: continue
        if r["note"] is not None and r["note"] < f_note_min: continue
        if not limite_temps_ok(r): continue
        filtrés.append(r)

    if not filtrés:
        st.warning("Aucun contenu ne correspond à tes filtres.")
        return

    st.markdown(f"**{len(filtrés)}** contenus évalués.")

    # Tris
    if f_tri == "✨ Pour moi (recommandé)":
        filtrés.sort(key=lambda x: -x["score"])
        top = [r for r in filtrés if r["score"] >= 50 and not r["pas_pour_moi"]]
        bof = [r for r in filtrés if 30 <= r["score"] < 50 and not r["pas_pour_moi"]]
        bad = [r for r in filtrés if r["pas_pour_moi"]]
        sections = [("✨ Recommandations personnalisées", top, "rec"),
                    ("🤔 Pourquoi pas", bof, "mid"),
                    ("🙅 Ne correspond pas à mon profil", bad, "bad")]
    elif f_tri == "⭐ Meilleures notes":
        filtrés.sort(key=lambda x: -(x["note"] or 0))
        sections = [("⭐ Par note décroissante", filtrés, "mid")]
    elif f_tri == "⏱️ Plus rapide à regarder":
        filtrés.sort(key=lambda x: x["duree_min"])
        sections = [("⏱️ Du plus rapide au plus long", filtrés, "mid")]
    elif f_tri == "🔥 Populaires":
        filtrés.sort(key=lambda x: -x["votes"])
        sections = [("🔥 Les plus populaires", filtrés, "mid")]
    elif f_tri == "🆕 Ajouté récemment":
        filtrés.sort(key=lambda x: 999999 if x["ajout"] is None else x["ajout"])
        sections = [("🆕 Derniers ajouts dans la liste", filtrés, "mid")]
    elif f_tri == "📅 Nouveautés (sorties)":
        filtrés.sort(key=lambda x: -(x["annee"] or 0))
        sections = [("📅 Sorties les plus récentes", filtrés, "mid")]
    elif f_tri == "🎬 Films d'abord":
        films = sorted([r for r in filtrés if r["type"]=="Film"], key=lambda x: -x["score"])
        series = sorted([r for r in filtrés if r["type"]=="Série"], key=lambda x: -x["score"])
        sections = [("🎬 Films", films, "rec"), ("📺 Séries", series, "mid")]
    elif f_tri == "📺 Séries d'abord":
        series = sorted([r for r in filtrés if r["type"]=="Série"], key=lambda x: -x["score"])
        films = sorted([r for r in filtrés if r["type"]=="Film"], key=lambda x: -x["score"])
        sections = [("📺 Séries", series, "rec"), ("🎬 Films", films, "mid")]
    else:
        filtrés.sort(key=lambda x: x["score"])
        sections = [("🙅 Contenus qui ne correspondent pas à mon profil", [r for r in filtrés if r["pas_pour_moi"]], "bad")]

    st.markdown("""
    <style>
    .rec-card-rec { background: linear-gradient(135deg, rgba(0,163,146,0.20) 0%, rgba(0,82,75,0.35) 100%); border:1px solid rgba(0,163,146,0.4); border-radius:16px; padding:16px; margin-bottom:14px; backdrop-filter: blur(12px); }
    .rec-card-mid { background: rgba(8,55,50,0.55); border:1px solid rgba(18,90,84,0.4); border-radius:16px; padding:16px; margin-bottom:14px; backdrop-filter: blur(12px); }
    .rec-card-bad { background: rgba(80,35,20,0.35); border:1px solid rgba(180,110,40,0.35); border-radius:16px; padding:16px; margin-bottom:14px; backdrop-filter: blur(12px); }
    .rec-titre { font-size:1.15em; font-weight:700; color:#F0FAF8; margin-bottom:4px; }
    .rec-meta { font-size:0.88em; color:#9DC5BF; margin-bottom:8px; }
    .rec-score { font-size:1.6em; font-weight:800; color:#CEDC00; line-height:1; }
    .rec-tag-ok { display:inline-block; padding:4px 11px; margin:3px 4px 3px 0; border-radius:12px; background:rgba(0,163,146,0.25); color:#7EE0D3; font-size:0.82em; font-weight:600; }
    .rec-tag-warn { display:inline-block; padding:4px 11px; margin:3px 4px 3px 0; border-radius:12px; background:rgba(210,130,40,0.22); color:#F2B670; font-size:0.82em; font-weight:600; }
    .rec-score-bar { height:8px; background:rgba(0,0,0,0.3); border-radius:4px; margin-top:8px; overflow:hidden; }
    .rec-score-fill { height:100%; border-radius:4px; background: linear-gradient(90deg, #00524B, #00A392, #CEDC00); }
    </style>
    """, unsafe_allow_html=True)

    for (nom_titre, groupe, cls) in sections:
        if not groupe: continue
        st.divider()
        st.markdown(f"### {nom_titre} ({len(groupe)})")
        for r in groupe:
            img = image_tmdb(r.get("tmdb"), "movie" if r["type"]=="Film" else "tv")
            cimg, cc = st.columns([0.09, 0.91])
            with cimg:
                if img:
                    st.image(img, use_container_width=True)
                else:
                    st.markdown("🎬" if r["type"]=="Film" else "📺")
            with cc:
                an_part = f"({r['annee']})" if r.get('annee') else ""
                aj_part = f"· 📥 Ajouté il y a {r['ajout']}j" if r['ajout'] is not None else ""
                note_part = f"{r['note']}" if r['note'] else "?"
                ep_part = f"· 📺 {r['nb_episodes']} ép." if r["type"]=="Série" and r["nb_episodes"]>0 else ""
                st.markdown(f"""
                <div class="rec-card-{cls}">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">
                        <div style="flex:1;">
                            <div class="rec-titre">{r['type']} — {r['titre']} {an_part}</div>
                            <div class="rec-meta">⭐ {note_part}<b>/10</b> · ⏱️ {r['temps']} · 🎭 {r['genres']} {ep_part} {aj_part}</div>
                        </div>
                        <div style="text-align:center; min-width:60px;">
                            <div class="rec-score">{r['score']}</div>
                            <div style="font-size:0.7em; color:#9DC5BF;">/100</div>
                        </div>
                    </div>
                    <div class="rec-score-bar"><div class="rec-score-fill" style="width:{r['score']}%"></div></div>
                    <div style="margin-top:10px;">
                        {''.join(f'<span class="rec-tag-ok">✅ {x}</span>' for x in r['raisons'])}
                        {''.join(f'<span class="rec-tag-warn">⚠️ {x}</span>' for x in r['averti'])}
                    </div>
                </div>
                """, unsafe_allow_html=True)


def page_wrapped():
    st.subheader("🎬 Rendez-vous annuel")
    st.info("🚧 Bientôt : récapitulatif annuel façon Spotify Wrapped.")

def page_sauvegarde():
    st.subheader("📤 Sauvegarde et restauration")
    st.info("🚧 Bientôt : export/import de tes données.")

# ==================================================
# RECONNEXION AUTO
# ==================================================

if "access_token" not in st.session_state:
    rt = cookies.get("trakt_rt")
    if rt:
        tok = rafraichir_token(rt)
        if tok:
            sauvegarder_connexion(tok)
        else:
            try: cookies.remove("trakt_rt")
            except: pass

utz = entete()
if "access_token" not in st.session_state:
    page_connexion()
else:
    p = naviguer()
    if p == "🏠 Tableau de bord": page_dashboard(utz)
    elif p == "▶️ En cours de lecture": page_lecture(utz)
    elif p == "👻 Progression Fantôme": page_fantomes(utz)
    elif p == "🧹 Nettoyage des listes": page_nettoyage(utz)
    elif p == "🔍 Recherche de doublons": page_doublons(utz)
    elif p == "🎯 Que regarder ?": page_quoi_regarder(utz)
    elif p == "📊 Statistiques": page_stats(utz)
    elif p == "📅 Calendrier des sorties": page_calendrier(utz)
    elif p == "🎬 Rendez-vous annuel": page_wrapped()
    elif p == "📤 Sauvegarde": page_sauvegarde()
    elif p == "🏆 Succès": page_succes(utz)
