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

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🏎️", layout="wide", initial_sidebar_state="collapsed")

# ==================================================
# UTILITAIRES
# ==================================================

def format_duree(heures):
    """Convertit des heures en format très lisible : années/mois/semaines/jours/heures"""
    if pd.isna(heures) or heures is None or heures <= 0:
        return "0h"
    total_min = round(heures * 60)
    # Calcul des unités
    ans = total_min // (365 * 24 * 60)
    mois = (total_min % (365 * 24 * 60)) // (30 * 24 * 60)
    sem = (total_min % (30 * 24 * 60)) // (7 * 24 * 60)
    jours = (total_min % (7 * 24 * 60)) // (24 * 60)
    h = (total_min % (24 * 60)) // 60
    parts = []
    if ans > 0:
        parts.append(f"{ans} an{'s' if ans > 1 else ''}")
    if mois > 0:
        parts.append(f"{mois} mois")
    if sem > 0:
        parts.append(f"{sem} sem{'s' if sem > 1 else ''}")
    if jours > 0:
        parts.append(f"{jours}j")
    if h > 0 or not parts:
        parts.append(f"{h}h")
    return " ".join(parts)

# ==================================================
# STYLE ASTON MARTIN x CINOPSYS
# ==================================================

st.markdown("""
<style>
    :root {
        --am-green: #00A392;
        --am-green-aston: #00665F;
        --am-green-dark: #004D48;
        --am-lime: #CEDC00;
        --am-bg-card: rgba(8, 68, 63, 0.85);
        --am-bg-card-hover: rgba(12, 88, 81, 0.9);
        --am-border: rgba(18, 90, 84, 0.6);
        --am-text: #F0FAF8;
        --am-text-muted: #9DC5BF;
    }

    /* DEGRADE FOND avec VERT ASTON OFFICIEL #00665F */
    .stApp {
        background: linear-gradient(180deg, #042E2B 0%, #00665F 50%, #042E2B 100%) !important;
        background-attachment: fixed !important;
    }

    /* Effet verre */
    div[data-testid="stMetric"],
    div.stAlert,
    div[data-testid="stContainer"],
    div.stButton > button,
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-testid="stDataFrame"] {
        background-color: rgba(4, 46, 43, 0.75) !important;
        backdrop-filter: blur(12px) !important;
        -webkit-backdrop-filter: blur(12px) !important;
        border-radius: 16px !important;
        border: 1px solid var(--am-border) !important;
        box-shadow: 0 8px 32px rgba(0,0,0,0.25) !important;
    }

    /* Métriques - plus de largeur pour les temps longs */
    div[data-testid="stMetric"] {
        padding: 20px !important;
        overflow: visible !important;
    }
    div[data-testid="stMetricValue"] {
        color: var(--am-text) !important;
        font-size: 1.7em !important;
        font-weight: 800 !important;
        white-space: nowrap !important;
        overflow: visible !important;
    }
    div[data-testid="stMetricLabel"] p {
        color: var(--am-text-muted) !important;
        font-size: 0.9em !important;
        font-weight: 500;
    }

    /* Messages uniformisés */
    div.stSuccess {
        background-color: rgba(0, 163, 146, 0.15) !important;
        border-left: 4px solid var(--am-green) !important;
        border: 1px solid rgba(0,163,146,0.3) !important;
        color: var(--am-text) !important;
    }
    div.stSuccess svg { fill: var(--am-green) !important; }
    div.stInfo {
        background-color: rgba(4, 46, 43, 0.75) !important;
        border-left: 4px solid var(--am-green) !important;
        border: 1px solid var(--am-border) !important;
        color: var(--am-text) !important;
    }
    div.stInfo svg { fill: var(--am-green) !important; }
    div.stWarning {
        background-color: rgba(206, 220, 0, 0.1) !important;
        border-left: 4px solid var(--am-lime) !important;
        border: 1px solid rgba(206,220,0,0.3) !important;
        color: var(--am-text) !important;
    }
    div.stWarning svg { fill: var(--am-lime) !important; }
    div.stError {
        background-color: rgba(237,34,36,0.1) !important;
        border-left: 4px solid #ED2224 !important;
        border: 1px solid rgba(237,34,36,0.3) !important;
    }

    /* Boutons */
    .stButton > button {
        font-weight: 600;
        padding: 0.7em 1.3em;
        color: var(--am-text) !important;
        transition: all 0.25s ease;
    }
    .stButton > button:hover {
        background-color: var(--am-bg-card-hover) !important;
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(0, 163, 146, 0.25);
        border-color: var(--am-green) !important;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--am-green) 0%, var(--am-green-aston) 100%) !important;
        border: none !important;
        font-weight: 700;
    }
    div[data-testid="stDownloadButton"] > button {
        background-color: rgba(4, 46, 43, 0.75) !important;
        backdrop-filter: blur(12px);
        border-radius: 16px !important;
        border: 1px solid var(--am-border) !important;
        color: var(--am-text) !important;
        font-weight: 600;
        width: 100%;
        padding: 0.7em 1.3em;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background-color: var(--am-bg-card-hover) !important;
        border-color: var(--am-green) !important;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: rgba(2, 20, 19, 0.95) !important;
        backdrop-filter: blur(20px) !important;
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
        background-color: rgba(0, 163, 146, 0.1) !important;
        color: var(--am-text) !important;
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] label:has(input:checked) {
        background: linear-gradient(135deg, rgba(0,102,95,0.3) 0%, rgba(0,77,72,0.25) 100%) !important;
        color: var(--am-text) !important;
        font-weight: 700 !important;
        border: 1px solid rgba(0,163,146,0.4);
    }
    section[data-testid="stSidebar"] div[data-testid="stRadio"] label:has(input:checked) span {
        color: var(--am-lime) !important;
    }

    .section-menu-title {
        font-size: 0.75em;
        font-weight: 800;
        color: var(--am-lime);
        text-transform: uppercase;
        letter-spacing: 1.5px;
        margin: 20px 0 12px 0;
    }

    input[type="checkbox"]:checked { accent-color: var(--am-green); }
    hr { border-color: var(--am-border) !important; }
    p, li, label { color: var(--am-text) !important; }
    .stCaption { color: var(--am-text-muted) !important; }
    button[kind="header"] {
        background-color: rgba(4, 46, 43, 0.75) !important;
        backdrop-filter: blur(12px);
        border-radius: 12px !important;
        border: 1px solid var(--am-border) !important;
    }
    div[role="progressbar"] > div {
        background: linear-gradient(90deg, var(--am-green) 0%, var(--am-lime) 100%) !important;
    }

    /* Carte lecture en cours */
    .now-playing-card {
        background: linear-gradient(135deg, rgba(0,102,95,0.3) 0%, rgba(4,46,43,0.8) 100%);
        backdrop-filter: blur(16px);
        border-radius: 20px;
        padding: 24px;
        border: 1px solid rgba(0,163,146,0.4);
        box-shadow: 0 12px 40px rgba(0,0,0,0.3);
        margin-bottom: 24px;
    }

    /* Cartes fantômes */
    .ghost-card {
        background-color: rgba(4, 46, 43, 0.75);
        backdrop-filter: blur(12px);
        border-radius: 16px;
        padding: 20px 24px;
        margin-bottom: 14px;
        border-left: 4px solid var(--am-lime);
        transition: all 0.25s ease;
        box-shadow: 0 4px 16px rgba(0,0,0,0.15);
    }
    .ghost-card:hover {
        border-left: 4px solid var(--am-green);
        transform: translateX(4px);
        background-color: var(--am-bg-card-hover);
    }
    .ghost-title {
        font-size: 1.1em;
        font-weight: 700;
        color: var(--am-text);
        margin-bottom: 6px;
    }
    .ghost-meta {
        font-size: 0.9em;
        color: var(--am-text-muted);
        margin-bottom: 14px;
    }
    .progress-bar-container {
        width: 100%;
        height: 12px;
        background-color: rgba(6, 59, 55, 0.8);
        border-radius: 8px;
        overflow: hidden;
    }
    .progress-bar-fill {
        height: 100%;
        border-radius: 8px;
        transition: width 0.6s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .progress-low { background: linear-gradient(90deg, #ED2224 0%, #F8696B 100%); }
    .progress-mid { background: linear-gradient(90deg, var(--am-lime) 0%, #E8F064 100%); }
    .progress-high { background: linear-gradient(90deg, var(--am-green) 0%, #00C7B3 100%); }
</style>
""", unsafe_allow_html=True)

cookies = CookieController()

# ==================================================
# CONFIGURATION
# ==================================================

CLIENT_ID = st.secrets["TRAKT_CLIENT_ID"]
CLIENT_SECRET = st.secrets["TRAKT_CLIENT_SECRET"]

DEVICE_CODE_URL = "https://api.trakt.tv/oauth/device/code"
DEVICE_TOKEN_URL = "https://api.trakt.tv/oauth/device/token"
REFRESH_TOKEN_URL = "https://api.trakt.tv/oauth/token"

def formater_date(date_str, user_tz):
    if not date_str:
        return ""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        dt_local = dt.astimezone(user_tz)
        offset = dt_local.strftime("%z")
        offset_f = f"{offset[:3]}:{offset[3:]}"
        return dt_local.strftime("%d/%m/%Y %H:%M:%S") + f" ({offset_f})"
    except Exception:
        return date_str

def formater_heure_duree(minutes):
    """Formatte une durée en minutes en h/min"""
    if not minutes or minutes <=0:
        return "inconnue"
    h = minutes // 60
    m = minutes % 60
    if h >0:
        return f"{h}h{m:02d}"
    return f"{m}min"

# ==================================================
# FONCTIONS TRAKT
# ==================================================

def demarrer_connexion():
    r = requests.post(DEVICE_CODE_URL, json={"client_id": CLIENT_ID})
    r.raise_for_status()
    return r.json()

def verifier_connexion(dc):
    r = requests.post(DEVICE_TOKEN_URL, json={"code": dc, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET})
    if r.status_code == 200:
        return r.json()
    return None

def rafraichir_token(rt):
    try:
        r = requests.post(REFRESH_TOKEN_URL, json={"refresh_token": rt, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET, "redirect_uri": "urn:ietf:wg:oauth:2.0:oob", "grant_type": "refresh_token"}, timeout=10)
        if r.status_code == 200:
            return r.json()
    except Exception:
        return None
    return None

def sauvegarder_connexion(tokens):
    st.session_state["access_token"] = tokens["access_token"]
    st.session_state["refresh_token"] = tokens["refresh_token"]
    st.session_state["token_saved_at"] = time.time()
    try:
        cookies.set("trakt_refresh_token", tokens["refresh_token"], expires_at=datetime.now() + timedelta(days=30))
    except Exception:
        pass
    time.sleep(0.3)

def oublier_connexion():
    try:
        cookies.remove("trakt_refresh_token")
    except Exception:
        pass
    time.sleep(0.3)
    st.session_state.clear()

def entetes(at):
    return {"Content-Type": "application/json", "trakt-api-version": "2", "trakt-api-key": CLIENT_ID, "Authorization": f"Bearer {at}"}

def obtenir_infos(at):
    r = requests.get("https://api.trakt.tv/users/settings", headers=entetes(at), timeout=10)
    r.raise_for_status()
    d = r.json()
    tz = d["user"].get("timezone", "Europe/Paris")
    try:
        utz = pytz.timezone(tz)
    except Exception:
        utz = pytz.timezone("Europe/Paris")
    return {"pseudo": d["user"]["username"], "tz": utz, "tz_name": tz, "pays": d["user"].get("locale", "fr")}

def qrcode_img(url):
    img = qrcode.make(url)
    b = io.BytesIO()
    img.save(b, format="PNG")
    return b.getvalue()

def obtenir_image_tmdb(tmdb_id, type_contenu="movie"):
    """Récupère l'affiche du contenu depuis TMDB (si on a la clé, sinon rien)"""
    if not tmdb_id:
        return None
    try:
        TMDB_KEY = st.secrets.get("TMDB_API_KEY")
        if not TMDB_KEY:
            return None
        r = requests.get(f"https://api.themoviedb.org/3/{type_contenu}/{tmdb_id}", params={"api_key": TMDB_KEY}, timeout=5)
        if r.status_code == 200:
            poster = r.json().get("poster_path")
            if poster:
                return f"https://image.tmdb.org/t/p/w200{poster}"
    except Exception:
        return None
    return None

def recuperer_historique(at, barre=None):
    h = entetes(at)
    films, series, films_det, ep_det = {}, {}, [], []
    nf, ne = 0, 0
    rp = requests.get("https://api.trakt.tv/users/me/history", headers=h, params={"page":1,"limit":100,"extended":"full"}, timeout=30)
    rp.raise_for_status()
    tp = int(rp.headers.get("X-Pagination-Page-Count", 1))
    for p in range(1, tp+1):
        if barre:
            barre.progress(p/tp*0.6, text=f"Historique : page {p}/{tp}")
        r = requests.get("https://api.trakt.tv/users/me/history", headers=h, params={"page":p,"limit":100,"extended":"full"}, timeout=30)
        r.raise_for_status()
        for it in r.json():
            if it["type"] == "movie":
                nf +=1
                m = it["movie"]
                tid = m["ids"]["trakt"]
                films_det.append({"titre":m["title"],"annee":m["year"],"genre":", ".join(m.get("genres",[])) if m.get("genres") else "Inconnu","duree":m.get("runtime",0) or 0,"note":m.get("rating",0) or 0,"date":it["watched_at"],"id":tid,"studio":m.get("certification", "")})
                if tid not in films:
                    films[tid] = {"titre":m["title"],"annee":m["year"],"vues":1,"dernier":it["watched_at"]}
                else:
                    films[tid]["vues"] +=1
            elif it["type"] == "episode":
                ne +=1
                s = it["show"]
                ep = it["episode"]
                sid = s["ids"]["trakt"]
                ep_det.append({"serie":s["title"],"titre":ep["title"],"saison":ep["season"],"episode":ep["number"],"annee":s["year"],"genre":", ".join(s.get("genres",[])) if s.get("genres") else "Inconnu","duree":ep.get("runtime",0) or s.get("runtime",40) or 40,"note":s.get("rating",0) or 0,"date":it["watched_at"],"id":sid,"network":s.get("network", "Inconnu")})
                if sid not in series:
                    series[sid] = {"titre":s["title"],"annee":s["year"],"vues":1,"dernier":it["watched_at"]}
                else:
                    series[sid]["vues"] +=1
    return {"films":films,"series":series,"films_det":films_det,"ep_det":ep_det,"nb_films":len(films),"nb_series":len(series),"nb_vf":nf,"nb_ep":ne}

def recuperer_listes(at):
    r = requests.get("https://api.trakt.tv/users/me/lists", headers=entetes(at), timeout=15)
    r.raise_for_status()
    return r.json()

def recuperer_contenu_liste(at, lid):
    h = entetes(at)
    items, p = [], 1
    while True:
        r = requests.get(f"https://api.trakt.tv/users/me/lists/{lid}/items", headers=h, params={"page":p,"limit":100}, timeout=15)
        r.raise_for_status()
        d = r.json()
        if not d: break
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
        items.extend(d)
        p +=1
    return items

def recuperer_lecture_en_cours(at):
    """Récupère le contenu actuellement en lecture"""
    try:
        r = requests.get("https://api.trakt.tv/users/me/watching", headers=entetes(at), timeout=10)
        if r.status_code == 204:
            return None
        if r.status_code != 200:
            return None
        return r.json()
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
                res.append({"type":"Film","titre":m["title"],"annee":m["year"],"vues":v["vues"],"dernier":v["dernier"],"tid":tid,"tmdb":m["ids"].get("tmdb")})
        elif it["type"] == "show":
            s = it["show"]
            sid = s["ids"]["trakt"]
            if sid in histo["series"]:
                v = histo["series"][sid]
                res.append({"type":"Série","titre":s["title"],"annee":s["year"],"vues":v["vues"],"dernier":v["dernier"],"tid":sid,"tmdb":s["ids"].get("tmdb")})
    return res

def analyser(at, histo, barre=None):
    res, stats, app = [], [], {}
    def aj(it, nom, lid):
        if it["type"] == "movie":
            med, t = it["movie"], "Film"
        elif it["type"] == "show":
            med, t = it["show"], "Série"
        else:
            return
        tid = med["ids"]["trakt"]
        cle = (t, tid)
        if cle not in app:
            app[cle] = {"titre":med["title"],"annee":med["year"],"type":t,"tid":tid,"tmdb":med["ids"].get("tmdb"),"dans":[]}
        app[cle]["dans"].append({"nom":nom,"lid":lid})
    if barre:
        barre.progress(0.6, text="Analyse liste de suivi...")
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
    for i, l in enumerate(listes):
        if barre:
            barre.progress(0.6 + (i+1)/max(len(listes),1)*0.3, text=f"Analyse : {l['name']}")
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
    if barre:
        barre.progress(0.95, text="Recherche des fantômes...")
    r = requests.get("https://api.trakt.tv/sync/playback", headers=entetes(at), timeout=15)
    r.raise_for_status()
    res = []
    for it in r.json():
        if it["type"] == "movie" and it.get("movie"):
            t, a = it["movie"]["title"], it["movie"].get("year")
            ty, duree_totale = "Film", it["movie"].get("runtime", 0)
            tmdb = it["movie"]["ids"].get("tmdb")
        elif it["type"] == "episode" and it.get("show") and it.get("episode"):
            ep = it["episode"]
            t = f"{it['show']['title']} — S{ep['season']:02d}E{ep['number']:02d}" if ep.get("season") and ep.get("number") else it["show"]["title"]
            a = it["show"].get("year")
            ty = "Épisode"
            duree_totale = ep.get("runtime", 0) or it["show"].get("runtime", 0)
            tmdb = it["show"]["ids"].get("tmdb")
        else:
            continue
        progression = round(it.get("progress",0))
        res.append({"type":ty,"titre":t,"annee":a,"prog":progression,"dernier":it["paused_at"],"pid":it["id"], "duree_totale": duree_totale, "tmdb": tmdb})
    res.sort(key=lambda x: x["dernier"])
    return res

def lancer_analyse(rafraichir_histo=False):
    barre = st.progress(0, text="Démarrage...")
    if rafraichir_histo or "historique" not in st.session_state:
        st.session_state["historique"] = recuperer_historique(st.session_state["access_token"], barre)
    res, stats, doub, doub_det = analyser(st.session_state["access_token"], st.session_state["historique"], barre)
    pb = recuperer_playback(st.session_state["access_token"], barre)
    np = recuperer_lecture_en_cours(st.session_state["access_token"])
    st.session_state["res"] = res
    st.session_state["stats"] = stats
    st.session_state["doub"] = doub
    st.session_state["doub_det"] = doub_det
    st.session_state["pb"] = pb
    st.session_state["np"] = np
    barre.empty()
    # Retour accueil après analyse
    st.session_state["page_active"] = "🏠 Tableau de bord"
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
            except:
                pass
        ws.column_dimensions[lettre].width = min(l+4, 40)

def forme(ws, coul="#00665F"):
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
        ["Compte", pseudo], ["Fuseau horaire", utz.zone],
        ["Films vus", histo["nb_films"]], ["Séries vues", histo["nb_series"]],
        ["Épisodes vus", histo["nb_ep"]], ["Temps total de visionnage", format_duree(th)],
        ["Listes personnalisées", len(stats)-1], ["Total de contenus", sum(s["total"] for s in stats)],
        ["Contenus déjà vus", len(res)], ["Doublons", len(doub)], ["Progressions fantômes", len(pb)]
    ], columns=["Statistique", "Valeur"])
    df_res = pd.DataFrame(res)
    if not df_res.empty:
        df_res = df_res[["liste","type","titre","annee","vues","dernier","tmdb"]].copy()
        df_res["dernier"] = pd.to_datetime(df_res["dernier"]).dt.tz_convert(utz).dt.strftime("%d/%m/%Y %H:%M")
        df_res.columns = ["Liste","Type","Titre","Année","Nombre de vues","Dernier visionnage","ID TMDB"]
    else:
        df_res = pd.DataFrame(columns=["Liste","Type","Titre","Année","Nombre de vues","Dernier visionnage","ID TMDB"])
    df_d = pd.DataFrame(doub)
    if not df_d.empty:
        df_d = df_d[["type","titre","annee","tmdb","nb_listes","listes"]].copy()
        df_d.columns = ["Type","Titre","Année","ID TMDB","Nombre de listes","Présent dans"]
    else:
        df_d = pd.DataFrame(columns=["Type","Titre","Année","ID TMDB","Nombre de listes","Présent dans"])
    df_sl = pd.DataFrame(stats)
    df_sl["% nettoyage"] = (df_sl["vus"]/df_sl["total"].replace(0,1)*100).round(1)
    df_sl = df_sl[["nom","nf","ns","total","vus","% nettoyage"]]
    df_sl.columns = ["Liste","Films","Séries","Total","Déjà vus","% nettoyage"]
    df_pb = pd.DataFrame(pb)
    if not df_pb.empty:
        df_pb = df_pb[["type","titre","annee","prog","dernier"]].copy()
        df_pb["dernier"] = pd.to_datetime(df_pb["dernier"]).dt.tz_convert(utz).dt.strftime("%d/%m/%Y %H:%M")
        df_pb.columns = ["Type","Titre","Année","Progression (%)","Dernier visionnage"]
    else:
        df_pb = pd.DataFrame(columns=["Type","Titre","Année","Progression (%)","Dernier visionnage"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df_sum.to_excel(wr, sheet_name="Résumé", index=False)
        df_res.to_excel(wr, sheet_name="À nettoyer", index=False)
        df_d.to_excel(wr, sheet_name="Doublons", index=False)
        df_pb.to_excel(wr, sheet_name="Progressions fantômes", index=False)
        df_sl.to_excel(wr, sheet_name="Analyse des listes", index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    for sh in wb: forme(sh)
    ws = wb["Analyse des listes"]
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

def aller_a(page):
    st.session_state["page_active"] = page
    st.rerun()

def naviguer():
    PAGES = [
        "🏠 Tableau de bord",
        "▶️ En cours de lecture",
        "👻 Progression Fantôme",
        "🧹 Nettoyage des listes",
        "🔍 Recherche de doublons",
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
        page = st.radio(
            "Navigation",
            PAGES,
            index=PAGES.index(st.session_state["page_active"]),
            label_visibility="collapsed",
            key="nav_principale"
        )
    st.session_state["page_active"] = page
    return page

# ==================================================
# ENTÊTE
# ==================================================

def entete():
    cl, ct = st.columns([0.08, 0.92])
    with cl:
        try: st.image("trakt-logo.svg", width=60)
        except: pass
    with ct:
        st.title("Trakt Smart Lists")
    if "access_token" not in st.session_state:
        return None
    # Rafraîchissement token automatique si besoin
    if "token_saved_at" in st.session_state and (time.time() - st.session_state["token_saved_at"]) > 7*24*3600:
        nouveau = rafraichir_token(st.session_state["refresh_token"])
        if nouveau:
            sauvegarder_connexion(nouveau)
    if "infos" not in st.session_state or (time.time() - st.session_state.get("infos_heure", 0)) > 3600:
        st.session_state["infos"] = obtenir_infos(st.session_state["access_token"])
        st.session_state["infos_heure"] = time.time()
    infos = st.session_state["infos"]
    pseudo, utz = infos["pseudo"], infos["tz"]
    ci, cd = st.columns([4,1])
    with ci:
        st.info(f"👤 Connecté en tant que **{pseudo}** • 🕒 Fuseau : `{infos['tz_name']}`")
    with cd:
        if st.button("🚪 Déconnexion", use_container_width=True):
            oublier_connexion()
            st.rerun()
    st.divider()
    if "res" in st.session_state:
        h = st.session_state["historique"]
        res = st.session_state["res"]
        stats = st.session_state["stats"]
        doub = st.session_state["doub"]
        pb = st.session_state["pb"]
        xl = generer_excel(pseudo, h, res, stats, doub, pb, utz)
        c1,c2,c3 = st.columns(3)
        with c1:
            if st.button("🔄 Analyse rapide", use_container_width=True, help="Garde l'historique en mémoire"):
                for k in ["res","stats","doub","doub_det","pb","np"]:
                    st.session_state.pop(k, None)
                lancer_analyse(False)
        with c2:
            if st.button("🔃 Rafraîchir tout", use_container_width=True, help="Récupère tout l'historique à nouveau"):
                for k in ["historique","res","stats","doub","doub_det","pb","np","infos"]:
                    st.session_state.pop(k, None)
                st.rerun()
        with c3:
            st.download_button("📥 Rapport Excel", data=xl, file_name=f"trakt_{pseudo}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.divider()
    return utz

def bloc_lancement():
    if "res" in st.session_state:
        return False
    if "historique" in st.session_state:
        st.info("ℹ️ Ton historique est déjà chargé, l'analyse sera rapide.")
        txt = "🔄 Lancer l'analyse rapide"
    else:
        st.info("ℹ️ Lance l'analyse pour accéder à tous les outils.")
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
            st.markdown(f'<a href="{url}" target="_blank" style="display:inline-block; background:linear-gradient(135deg,#00A392,#00665F); color:white; padding:0.9em 1.7em; border-radius:12px; text-decoration:none; font-weight:700;">Autoriser l\'accès à Trakt</a>', unsafe_allow_html=True)
            st.caption("Tu peux ouvrir ce lien sur n'importe quel appareil.")
            st.info(f"Code à entrer si demandé : **{st.session_state['uc']}**")
        with cd:
            st.image(qrcode_img(url), width=160)
            st.caption("Ou scanne ce QR code avec ton téléphone.")
        st.caption("La page se met à jour automatiquement une fois l'accès autorisé.")
        with st.spinner("Attente de l'autorisation..."):
            t = 0
            while t < st.session_state["exp"]:
                time.sleep(st.session_state["iv"])
                t += st.session_state["iv"]
                tok = verifier_connexion(st.session_state["dc"])
                if tok:
                    sauvegarder_connexion(tok)
                    del st.session_state["dc"]
                    st.rerun()
        st.error("Le délai d'attente est écoulé.")
        if st.button("Réessayer"):
            st.rerun()

def page_lecture_en_cours(utz):
    if bloc_lancement():
        return
    st.subheader("▶️ En cours de lecture")
    np = st.session_state.get("np")
    if not np:
        st.info("🎬 Aucun contenu en cours de lecture pour le moment.")
        return
    # Récupération infos
    if np["type"] == "movie":
        media = np["movie"]
        titre = media["title"]
        annee = media.get("year")
        type_c = "Film"
        duree = media.get("runtime", 0)
        tmdb = media["ids"].get("tmdb")
    else:
        media = np["show"]
        ep = np["episode"]
        titre = f"{media['title']} — S{ep['season']:02d}E{ep['number']:02d}"
        annee = media.get("year")
        type_c = "Épisode"
        duree = ep.get("runtime", 0) or media.get("runtime", 0)
        tmdb = media["ids"].get("tmdb")
    progression = round(np.get("progress", 0))
    debut = datetime.fromisoformat(np["started_at"].replace("Z", "+00:00")).astimezone(utz)
    fin_estimee = debut + timedelta(minutes=duree) if duree > 0 else None
    # Affiche
    img_url = obtenir_image_tmdb(tmdb, "movie" if type_c == "Film" else "tv")
    col_img, col_infos = st.columns([0.2, 0.8])
    with col_img:
        if img_url:
            st.image(img_url, use_container_width=True)
        else:
            st.markdown("🎬")
    with col_infos:
        st.markdown(f"""
        <div class="now-playing-card">
            <div style="font-size:0.85em; color: #CEDC00; font-weight:700; text-transform:uppercase; letter-spacing:1px; margin-bottom:8px;">▶️ EN LECTURE</div>
            <div style="font-size:1.8em; font-weight:800; color:#F0FAF8; margin-bottom:8px;">{titre}</div>
            <div style="font-size:1em; color:#9DC5BF; margin-bottom:16px;">{type_c} {f'({annee})' if annee else ''}</div>
            <div class="progress-bar-container" style="height:14px; margin-bottom:16px;">
                <div class="progress-bar-fill progress-high" style="width:{progression}%"></div>
            </div>
            <div style="display:grid; grid-template-columns: repeat(2, 1fr); gap:12px;">
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase; letter-spacing:0.5px;">Débuté à</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{debut.strftime('%H:%M:%S')}</div>
                </div>
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase; letter-spacing:0.5px;">Fin estimée à</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{fin_estimee.strftime('%H:%M') if fin_estimee else 'Inconnue'}</div>
                </div>
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase; letter-spacing:0.5px;">Durée totale</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{formater_heure_duree(duree)}</div>
                </div>
                <div>
                    <div style="font-size:0.8em; color:#9DC5BF; text-transform:uppercase; letter-spacing:0.5px;">Progression</div>
                    <div style="font-size:1.1em; font-weight:600; color:#F0FAF8;">{progression}%</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

def page_dashboard(utz):
    if bloc_lancement():
        return
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
    st.subheader("⚠️ Actions de nettoyage")
    c5,c6,c7,c8 = st.columns(4)
    with c5:
        with st.container(border=True):
            st.markdown("#### 👻 Fantômes")
            if len(pb) > 0:
                pct_pb = round(len(pb)/max(len(pb)+h["nb_vf"]+h["nb_ep"], 1)*100, 1)
                st.metric("Nombre", len(pb), delta=f"{pct_pb}%")
                st.warning(f"{len(pb)} fantôme(s) à nettoyer")
                if st.button("Nettoyer →", key="b_pb", use_container_width=True):
                    aller_a("👻 Progression Fantôme")
            else:
                st.metric("Nombre", 0)
                st.success("✅ Rien à nettoyer")
    with c6:
        with st.container(border=True):
            st.markdown("#### 🔁 Doublons")
            if len(doub) > 0:
                pct_d = round(len(doub)/max(total_items,1)*100, 1)
                st.metric("Nombre", len(doub), delta=f"{pct_d}%")
                st.warning(f"{len(doub)} doublon(s)")
                if st.button("Voir →", key="b_d", use_container_width=True):
                    aller_a("🔍 Recherche de doublons")
            else:
                st.metric("Nombre", 0)
                st.success("✅ Aucun doublon")
    with c7:
        with st.container(border=True):
            st.markdown("#### 🧹 Déjà vus")
            if len(res) > 0:
                pct_r = round(len(res)/max(total_items,1)*100,1)
                st.metric("Nombre", len(res), delta=f"{pct_r}%")
                st.warning(f"{len(res)} contenu(s) vus")
                if st.button("Nettoyer →", key="b_r", use_container_width=True):
                    aller_a("🧹 Nettoyage des listes")
            else:
                st.metric("Nombre", 0)
                st.success("✅ Listes à jour")
    with c8:
        with st.container(border=True):
            st.markdown("#### 🚀 Nettoyage auto")
            st.write("Tout nettoyer en 1 clic")
            if st.button("🧹 Tout nettoyer", type="primary", use_container_width=True):
                st.session_state["confirmer_nettoyage_complet"] = True
                st.rerun()
            if st.session_state.get("confirmer_nettoyage_complet"):
                st.warning("⚠️ Attention : cela va supprimer TOUS les contenus déjà vus de tes listes ET TOUS les fantômes. Cette action est irréversible.")
                co, cn = st.columns(2)
                with co:
                    if st.button("✅ Confirmer", type="primary"):
                        with st.spinner("Nettoyage complet en cours..."):
                            sup_selection(st.session_state["access_token"], res)
                            sup_playback(st.session_state["access_token"], pb)
                        st.success(f"✅ {len(res)} contenus et {len(pb)} fantômes supprimés !")
                        del st.session_state["confirmer_nettoyage_complet"]
                        time.sleep(2)
                        for k in ["res","stats","doub","doub_det","pb","np"]:
                            st.session_state.pop(k, None)
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        del st.session_state["confirmer_nettoyage_complet"]
                        st.rerun()

def page_nettoyage(utz):
    if bloc_lancement():
        return
    res = st.session_state["res"]
    msg = "msg_sup_vus"
    if st.session_state.get(msg):
        st.success(st.session_state[msg])
        del st.session_state[msg]
    st.subheader("🧹 Nettoyage des listes")
    st.caption("Retire les films et épisodes déjà vus de ta liste de suivi et de tes listes personnalisées.")
    if not res:
        st.success("Tes listes sont parfaitement à jour ! 🎉")
    else:
        st.write(f"**{len(res)}** contenu(s) déjà vu(s) détecté(s). Coche ceux que tu veux supprimer :")
        tab = pd.DataFrame(res)
        ta = tab[["type","titre","annee","vues","dernier","liste"]].copy()
        ta["dernier"] = pd.to_datetime(ta["dernier"]).dt.tz_convert(utz).dt.strftime("%d/%m/%Y %H:%M")
        ta.insert(0, "Sel", False)
        ta.columns = ["Sel","Type","Titre","Année","Vues","Dernier visionnage","Liste"]
        ed = st.data_editor(ta, use_container_width=True, hide_index=True, disabled=["Type","Titre","Année","Vues","Dernier visionnage","Liste"], key="ed_vus")
        nb = int(ed["Sel"].sum())
        if nb:
            conf = "conf_vus"
            if not st.session_state.get(conf, False):
                if st.button(f"🗑️ Supprimer les {nb} élément(s) sélectionné(s)", type="primary"):
                    st.session_state[conf] = True
                    st.rerun()
            else:
                st.warning(f"Confirmer la suppression de {nb} élément(s) ? Cette action est irréversible.")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Oui, supprimer"):
                        idx = ed[ed["Sel"]].index
                        items = [res[i] for i in idx]
                        with st.spinner("Suppression en cours..."):
                            sup_selection(st.session_state["access_token"], items)
                        st.session_state[conf] = False
                        st.session_state[msg] = f"✅ {len(items)} élément(s) supprimé(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[conf] = False
                        st.rerun()
    st.divider()
    st.subheader("Taux de contenu à nettoyer par liste")
    df = pd.DataFrame(st.session_state["stats"])
    df["% nettoyable"] = (df["vus"]/df["total"].replace(0,1)*100).round(1)
    st.bar_chart(df.set_index("nom")["% nettoyable"], color="#CEDC00")
    st.caption("Ce graphique te permet de voir rapidement quelle liste a besoin d'être nettoyée en priorité.")

def page_doublons(utz):
    if bloc_lancement():
        return
    dd = st.session_state["doub_det"]
    msg = "msg_sup_d"
    if st.session_state.get(msg):
        st.success(st.session_state[msg])
        del st.session_state[msg]
    st.subheader("🔍 Recherche de doublons")
    st.caption("Trouve les contenus qui sont présents dans plusieurs listes à la fois.")
    if not dd:
        st.success("Aucun doublon dans tes listes !")
    else:
        st.write(f"**{len(st.session_state['doub'])}** doublon(s) détecté(s). Coche les lignes à retirer :")
        tab = pd.DataFrame(dd)
        ta = tab[["type","titre","annee","liste"]].copy()
        ta.insert(0, "Sel", False)
        ta.columns = ["Sel","Type","Titre","Année","Liste"]
        ed = st.data_editor(ta, use_container_width=True, hide_index=True, disabled=["Type","Titre","Année","Liste"], key="ed_d")
        nb = int(ed["Sel"].sum())
        if nb:
            conf = "conf_d"
            if not st.session_state.get(conf, False):
                if st.button(f"🗑️ Retirer les {nb} élément(s) sélectionné(s)", type="primary"):
                    st.session_state[conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer le retrait ?")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Oui, retirer"):
                        idx = ed[ed["Sel"]].index
                        items = [dd[i] for i in idx]
                        with st.spinner("Suppression..."):
                            sup_selection(st.session_state["access_token"], items)
                        st.session_state[conf] = False
                        st.session_state[msg] = f"✅ {len(items)} élément(s) retiré(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[conf] = False
                        st.rerun()

def page_fantomes(utz):
    if bloc_lancement():
        return
    pb = st.session_state["pb"]
    msg = "msg_sup_pb"
    if st.session_state.get(msg):
        st.success(st.session_state[msg])
        del st.session_state[msg]
    st.subheader("👻 Progression Fantôme")
    st.caption("Gère tes vidéos en pause : supprime les entrées obsolètes qui restent bloquées dans 'Continuer à regarder'.")
    st.divider()
    if not pb:
        st.success("Aucune progression en cours. Tout est propre !")
    else:
        tout = st.checkbox("Tout sélectionner")
        sels = {}
        for it in pb:
            p = it["prog"]
            cls = "progress-low" if p<30 else "progress-mid" if p<80 else "progress-high"
            df = formater_date(it["dernier"], utz)
            ic = "🎬" if it["type"] == "Film" else "📺"
            with st.container():
                cc, cd = st.columns([0.05, 0.95])
                with cc:
                    sels[it["pid"]] = st.checkbox("", value=tout, key=f"c_{it['pid']}", label_visibility="collapsed")
                with cd:
                    st.markdown(f"""
                    <div class="ghost-card">
                        <div class="ghost-title">{ic} {it['titre']} {f'({it["annee"]})' if it["annee"] else ''}</div>
                        <div class="ghost-meta">{it['type']} • {p}% visionné • 🕒 {df}</div>
                        <div class="progress-bar-container">
                            <div class="progress-bar-fill {cls}" style="width: {p}%"></div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        st.divider()
        ids = [pid for pid, s in sels.items() if s]
        if ids:
            conf = "conf_pb"
            if not st.session_state.get(conf, False):
                if st.button(f"🗑️ Supprimer les {len(ids)} progression(s) sélectionnée(s)", type="primary"):
                    st.session_state[conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer la suppression ?")
                co,cn = st.columns(2)
                with co:
                    if st.button("✅ Oui, supprimer"):
                        items = [p for p in pb if p["pid"] in ids]
                        with st.spinner("Suppression en cours..."):
                            sup_playback(st.session_state["access_token"], items)
                        st.session_state[conf] = False
                        st.session_state[msg] = f"✅ {len(items)} fantôme(s) supprimé(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[conf] = False
                        st.rerun()
        else:
            st.info("Coche les éléments que tu veux supprimer.")

def page_calendrier(utz):
    if bloc_lancement():
        return
    st.subheader("📅 Calendrier des sorties")
    st.info("🚧 Prochainement : les prochaines dates de sortie des films et séries présents dans ta liste de suivi.")

def page_stats(utz):
    if bloc_lancement():
        return
    st.subheader("📊 Statistiques détaillées")
    st.caption("Toutes tes données de visionnage détaillées.")
    h = st.session_state["historique"]
    films = pd.DataFrame(h["films_det"])
    eps = pd.DataFrame(h["ep_det"])

    # Filtres
    f1,f2,f3 = st.columns(3)
    with f1:
        tc = st.selectbox("Type de contenu", ["Tous", "Films", "Séries"])
    with f2:
        # Liste des mois disponibles
        toutes_dates = []
        for df in [films, eps]:
            if not df.empty:
                dates = pd.to_datetime(df["date"], utc=True).dt.tz_convert(utz)
                toutes_dates.extend(dates.tolist())
        if toutes_dates:
            dates_df = pd.DataFrame({"date": toutes_dates})
            dates_df["mois_annee"] = dates_df["date"].dt.strftime("%m-%Y")
            mois_dispo = sorted(dates_df["mois_annee"].unique(), key=lambda x: (int(x.split("-")[1]), int(x.split("-")[0])))
            mois = st.select_slider("Mois", options=mois_dispo, value=(mois_dispo[0], mois_dispo[-1]))
            d_deb = datetime.strptime(mois[0], "%m-%Y").replace(tzinfo=utz)
            d_fin = datetime.strptime(mois[1], "%m-%Y").replace(day=28) + timedelta(days=4)
            d_fin = d_fin.replace(tzinfo=utz)
    with f3:
        genres = set()
        if tc in ["Tous","Films"] and not films.empty:
            for g in films["genre"].str.split(", "):
                genres.update([x for x in g if x != "Inconnu"])
        if tc in ["Tous","Séries"] and not eps.empty:
            for g in eps["genre"].str.split(", "):
                genres.update([x for x in g if x != "Inconnu"])
        genre = st.selectbox("Genre", ["Tous"] + sorted(genres))

    # Construction données
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
        st.info("Aucune donnée disponible.")
        return
    df = pd.concat(dfs, ignore_index=True)

    # Filtre mois
    df = df[(df["date_dt"] >= d_deb) & (df["date_dt"] <= d_fin)]
    # Filtre genre
    if genre != "Tous":
        df = df[df["genre"].str.contains(genre, na=False)]
    if df.empty:
        st.warning("Aucun visionnage ne correspond à ces filtres.")
        return

    df["duree_h"] = df["duree"].fillna(40)/60
    th = df["duree_h"].sum()

    # Métriques
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Total visionnages", len(df))
    m2.metric("Temps total", format_duree(th))
    nm = df[df["note"]>0]["note"].mean()
    m3.metric("Note moyenne", f"{round(nm,1)}/10" if pd.notna(nm) else "-")
    nb_jours = max((df["date_dt"].max() - df["date_dt"].min()).days +1, 1)
    m4.metric("Moyenne / jour", f"{round(len(df)/nb_jours,1)}")
    marathon = df.groupby(df["date_dt"].dt.date).size().max()
    m5.metric("Record en 1 jour", f"{marathon}")

    # Détection marathons (FILTRÉE selon la période !)
    marathons = pd.DataFrame()
    if tc in ["Tous", "Séries"] and not eps.empty:
        ej = eps.copy()
        ej["date_dt"] = pd.to_datetime(ej["date"], utc=True).dt.tz_convert(utz)
        ej = ej[(ej["date_dt"] >= d_deb) & (ej["date_dt"] <= d_fin)]
        if genre != "Tous":
            ej = ej[ej["genre"].str.contains(genre, na=False)]
        if not ej.empty:
            ej["jour"] = ej["date_dt"].dt.date
            comptage = ej.groupby(["jour", "serie"]).size().reset_index(name="nb")
            marathons = comptage[comptage["nb"] >=4].sort_values("nb", ascending=False)
    if not marathons.empty:
        st.divider()
        with st.container(border=True):
            st.markdown("#### 🏆 Marathons détectés")
            st.caption("Jours où tu as regardé 4 épisodes ou plus de la même série.")
            for _, row in marathons.head(5).iterrows():
                st.write(f"📅 **{row['jour'].strftime('%d/%m/%Y')}** : {row['nb']} épisodes de **{row['serie']}**")

    # Réseaux/studios préférés
    if tc in ["Tous", "Séries"] and not eps.empty:
        ej_n = eps.copy()
        ej_n["date_dt"] = pd.to_datetime(ej_n["date"], utc=True).dt.tz_convert(utz)
        ej_n = ej_n[(ej_n["date_dt"] >= d_deb) & (ej_n["date_dt"] <= d_fin)]
        if not ej_n.empty and "network" in ej_n.columns:
            st.divider()
            with st.container(border=True):
                st.markdown("#### 📺 Plateformes de streaming / chaînes préférées")
                plateformes = ej_n["network"].value_counts().head(5)
                cols_pf = st.columns(len(plateformes))
                for i, (pf, nb) in enumerate(plateformes.items()):
                    if pf != "Inconnu":
                        cols_pf[i].metric(pf, f"{nb} ép.")

    st.divider()

    # Graphique heures par mois
    df["mois"] = df["date_dt"].dt.strftime("%m-%Y")
    h_mois = df.groupby("mois")["duree_h"].sum().round(1).sort_index()
    opt_m = {"title":{"text":"Heures de visionnage par mois","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis"},"backgroundColor":"transparent","textStyle":{"color":"#F0FAF8"},"xAxis":{"type":"category","data":list(h_mois.index),"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(h_mois.values),"type":"line","smooth":True,"lineStyle":{"color":"#CEDC00","width":3},"areaStyle":{"color":"rgba(206,220,0,0.1)"},"itemStyle":{"color":"#CEDC00"}}]}
    st_echarts(opt_m, height="350px")

    g1,g2 = st.columns(2)
    with g1:
        genres = {}
        for lg in df["genre"].str.split(", "):
            for g in lg:
                if g and g != "Inconnu":
                    genres[g] = genres.get(g,0)+1
        opt_g = {"title":{"text":"Répartition par genre","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"item"},"backgroundColor":"transparent","legend":{"bottom":0,"textStyle":{"color":"#9DC5BF"}},"series":[{"type":"pie","radius":["40%","70%"],"data":[{"name":k,"value":v} for k,v in sorted(genres.items(), key=lambda x:-x[1])[:8]],"itemStyle":{"borderRadius":8,"borderColor":"#042E2B","borderWidth":2},"label":{"color":"#F0FAF8"}}],"color":["#00A392","#CEDC00","#00C7B3","#A3B300","#00665F","#869400","#125A54","#E8F064"]}
        st_echarts(opt_g, height="400px")
    with g2:
        df["h"] = df["date_dt"].dt.hour
        hh = df.groupby("h")["duree_h"].sum().reindex(range(24), fill_value=0).round(1)
        opt_h = {"title":{"text":"Par heure de la journée","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis"},"backgroundColor":"transparent","xAxis":{"type":"category","data":[f"{h}h" for h in range(24)],"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(hh.values),"type":"bar","itemStyle":{"color":{"type":"linear","x":0,"y":0,"x2":0,"y2":1,"colorStops":[{"offset":0,"color":"#00A392"},{"offset":1,"color":"#00665F"}]},"borderRadius":[4,4,0,0]}}]}
        st_echarts(opt_h, height="400px")

    g3,g4 = st.columns(2)
    with g3:
        jours = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"]
        df["jsem"] = df["date_dt"].dt.weekday
        hj = df.groupby("jsem")["duree_h"].sum().reindex(range(7), fill_value=0).round(1)
        opt_j = {"title":{"text":"Par jour de la semaine","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"axis"},"backgroundColor":"transparent","xAxis":{"type":"category","data":jours,"axisLabel":{"color":"#9DC5BF"}},"yAxis":{"type":"value","name":"Heures","axisLabel":{"color":"#9DC5BF"},"splitLine":{"lineStyle":{"color":"rgba(18,90,84,0.4)"}}},"series":[{"data":list(hj.values),"type":"bar","itemStyle":{"color":"#CEDC00","borderRadius":[4,4,0,0]}}]}
        st_echarts(opt_j, height="400px")
    with g4:
        rt = df.groupby("type")["duree_h"].sum().round(1)
        opt_t = {"title":{"text":"Films vs Séries","left":"center","textStyle":{"color":"#F0FAF8"}},"tooltip":{"trigger":"item"},"backgroundColor":"transparent","legend":{"bottom":0,"textStyle":{"color":"#9DC5BF"}},"series":[{"type":"pie","radius":["40%","70%"],"data":[{"value":v,"name":k} for k,v in rt.items()],"itemStyle":{"borderRadius":8,"borderColor":"#042E2B","borderWidth":2},"label":{"color":"#F0FAF8"}}],"color":["#00A392","#CEDC00"]}
        st_echarts(opt_t, height="400px")

    # Tableau détail
    with st.expander("📋 Voir le détail des visionnages filtrés"):
        df_aff = df[["date_dt", "type", "titre", "annee", "genre", "duree", "note"]].copy()
        df_aff["date_dt"] = df_aff["date_dt"].dt.strftime("%d/%m/%Y %H:%M")
        df_aff["duree"] = df_aff["duree"].apply(lambda x: formater_heure_duree(x) if x >0 else "-")
        df_aff.columns = ["Date", "Type", "Titre", "Année", "Genres", "Durée", "Note"]
        st.dataframe(df_aff, use_container_width=True, hide_index=True)

def page_wrapped():
    st.subheader("🎬 Rendez-vous annuel")
    st.info("🚧 Bientôt disponible : ton récapitulatif annuel façon Spotify Wrapped avec carte partageable.")

def page_sauvegarde():
    st.subheader("📤 Sauvegarde et restauration")
    st.info("🚧 Bientôt : export et import complet de tes listes et données.")

def page_succes():
    st.subheader("🏆 Succès")
    st.info("🚧 Bientôt : badges et objectifs de visionnage.")

# ==================================================
# RECONNEXION AUTO COOKIES
# ==================================================

if "access_token" not in st.session_state:
    rt = cookies.get("trakt_refresh_token")
    if rt:
        tok = rafraichir_token(rt)
        if tok:
            sauvegarder_connexion(tok)
        else:
            try:
                cookies.remove("trakt_refresh_token")
            except Exception:
                pass

# ==================================================
# LANCEMENT APPLI
# ==================================================

utz = entete()
if "access_token" not in st.session_state:
    page_connexion()
else:
    p = naviguer()
    if p == "🏠 Tableau de bord":
        page_dashboard(utz)
    elif p == "▶️ En cours de lecture":
        page_lecture_en_cours(utz)
    elif p == "👻 Progression Fantôme":
        page_fantomes(utz)
    elif p == "🧹 Nettoyage des listes":
        page_nettoyage(utz)
    elif p == "🔍 Recherche de doublons":
        page_doublons(utz)
    elif p == "📊 Statistiques":
        page_stats(utz)
    elif p == "📅 Calendrier des sorties":
        page_calendrier(utz)
    elif p == "🎬 Rendez-vous annuel":
        page_wrapped()
    elif p == "📤 Sauvegarde":
        page_sauvegarde()
    elif p == "🏆 Succès":
        page_succes()
