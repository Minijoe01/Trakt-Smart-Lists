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
    """Convertit automatiquement des heures en semaines / jours / heures pour plus de lisibilité."""
    if heures is None or heures != heures:  # Gère NaN
        return "0h"
    heures = round(heures, 1)
    total_minutes = round(heures * 60)
    sem = total_minutes // (7 * 24 * 60)
    jours = (total_minutes % (7 * 24 * 60)) // (24 * 60)
    h = (total_minutes % (24 * 60)) // 60
    parts = []
    if sem > 0:
        parts.append(f"{sem} semaine{'s' if sem > 1 else ''}")
    if jours > 0:
        parts.append(f"{jours} jour{'s' if jours > 1 else ''}")
    if h > 0 or not parts:
        parts.append(f"{h}h")
    return " ".join(parts)

# Injection CSS thème Aston Martin vert profond
st.markdown("""
<style>
    :root {
        --am-green: #00A392;
        --am-green-dark: #007D70;
        --am-green-darker: #04332F;
        --am-lime: #CEDC00;
        --am-bg-card: #08443F;
        --am-border: #125A54;
        --am-text-light: #F0FAF8;
        --am-text-muted: #9DC5BF;
    }

    /* Fond global VERT ASTON MARTIN, pas de noir */
    .stApp {
        background: linear-gradient(180deg, #04332F 0%, #032825 100%) !important;
    }

    /* Cartes, métriques, alertes */
    div[data-testid="stMetric"], div.stAlert, div[data-testid="stContainer"] {
        background-color: var(--am-bg-card) !important;
        border-radius: 14px !important;
        border: 1px solid var(--am-border) !important;
        padding: 18px !important;
        box-shadow: 0 4px 14px rgba(0,0,0,0.2) !important;
    }

    /* Barres de progression */
    div[role="progressbar"] > div {
        background-color: var(--am-lime) !important;
    }

    /* Cartes fantômes */
    .ghost-card {
        background-color: var(--am-bg-card);
        border-radius: 12px;
        padding: 18px 22px;
        margin-bottom: 14px;
        border-left: 4px solid var(--am-lime);
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 2px 8px rgba(0,0,0,0.12);
    }
    .ghost-card:hover {
        border-left: 4px solid var(--am-green);
        transform: translateX(3px);
        box-shadow: 0 6px 16px rgba(0, 163, 146, 0.2);
    }
    .ghost-title {
        font-size: 1.08em;
        font-weight: 600;
        color: var(--am-text-light);
        margin-bottom: 6px;
    }
    .ghost-meta {
        font-size: 0.9em;
        color: var(--am-text-muted);
        margin-bottom: 12px;
    }
    .progress-bar-container {
        width: 100%;
        height: 10px;
        background-color: #063B37;
        border-radius: 6px;
        overflow: hidden;
    }
    .progress-bar-fill {
        height: 100%;
        border-radius: 6px;
        transition: width 0.6s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .progress-low { background: linear-gradient(90deg, #ED2224 0%, #F8696B 100%); }
    .progress-mid { background: linear-gradient(90deg, #CEDC00 0%, #E8F064 100%); }
    .progress-high { background: linear-gradient(90deg, #00A392 0%, #00C7B3 100%); }

    /* Boutons moderne */
    .stButton > button {
        border-radius: 10px;
        border: 0;
        font-weight: 500;
        padding: 0.6em 1.2em;
        transition: all 0.2s ease;
        border: 1px solid var(--am-border);
        background-color: var(--am-bg-card);
        color: var(--am-text-light);
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0, 163, 146, 0.3);
        border-color: var(--am-green);
        background-color: var(--am-green-dark);
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--am-green) 0%, var(--am-green-dark) 100%);
        border: 0;
        color: white;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 6px 20px rgba(0, 163, 146, 0.4);
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #032825 !important;
        border-right: 1px solid var(--am-border);
    }

    /* Menu titre */
    .section-menu-title {
        font-size: 0.78em;
        font-weight: 700;
        color: var(--am-lime);
        text-transform: uppercase;
        letter-spacing: 1.2px;
        margin: 18px 0 10px 0;
    }

    /* Dataframes */
    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid var(--am-border);
    }

    hr {
        border-color: var(--am-border) !important;
    }

    /* Textes */
    p, li, label {
        color: var(--am-text-light) !important;
    }
    .stCaption {
        color: var(--am-text-muted) !important;
    }

    /* Focus */
    *:focus {
        outline: 2px solid var(--am-green) !important;
        outline-offset: 2px;
    }

    /* Metrics */
    div[data-testid="stMetricValue"] {
        color: var(--am-text-light) !important;
        font-size: 1.7em !important;
        font-weight: 700;
    }
    div[data-testid="stMetricLabel"] p {
        color: var(--am-text-muted) !important;
        font-size: 0.9em !important;
    }
    div[data-testid="stMetricDelta"] {
        font-size: 0.85em;
    }

    /* Selectbox, inputs */
    div[data-baseweb="select"] > div {
        background-color: var(--am-bg-card) !important;
        border-color: var(--am-border) !important;
    }
    div[data-baseweb="input"] > div {
        background-color: var(--am-bg-card) !important;
        border-color: var(--am-border) !important;
    }

    div[data-testid="stSidebarCollapsedControl"] {
        background-color: var(--am-bg-card) !important;
        border-radius: 8px;
        border: 1px solid var(--am-border);
    }
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
    """Formate une date Trakt (UTC) dans le fuseau horaire de l'utilisateur."""
    if not date_str:
        return ""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        dt_local = dt.astimezone(user_tz)
        offset = dt_local.strftime("%z")
        offset_formate = f"{offset[:3]}:{offset[3:]}"
        return dt_local.strftime("%Y-%m-%d %H:%M:%S") + f" ({offset_formate})"
    except Exception:
        return date_str

# ==================================================
# FONCTIONS TRAKT
# ==================================================

def demarrer_connexion():
    response = requests.post(DEVICE_CODE_URL, json={"client_id": CLIENT_ID})
    response.raise_for_status()
    return response.json()

def verifier_connexion(device_code):
    payload = {"code": device_code, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET}
    response = requests.post(DEVICE_TOKEN_URL, json=payload)
    if response.status_code == 200:
        return response.json()
    erreurs = {404: "Code invalide.", 409: "Code déjà utilisé.", 410: "Code expiré.", 418: "Connexion refusée."}
    if response.status_code in erreurs:
        raise Exception(erreurs[response.status_code])
    return None

def rafraichir_token(refresh_token):
    payload = {"refresh_token": refresh_token, "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET, "redirect_uri": "urn:ietf:wg:oauth:2.0:oob", "grant_type": "refresh_token"}
    response = requests.post(REFRESH_TOKEN_URL, json=payload)
    return response.json() if response.status_code == 200 else None

def sauvegarder_connexion(tokens):
    st.session_state["access_token"] = tokens["access_token"]
    st.session_state["refresh_token"] = tokens["refresh_token"]
    cookies.set("trakt_refresh_token", tokens["refresh_token"])
    time.sleep(0.4)

def oublier_connexion():
    cookies.remove("trakt_refresh_token")
    time.sleep(0.4)
    st.session_state.clear()

def entetes_trakt(access_token):
    return {"Content-Type": "application/json", "trakt-api-version": "2", "trakt-api-key": CLIENT_ID, "Authorization": f"Bearer {access_token}"}

def obtenir_infos_utilisateur(access_token):
    response = requests.get("https://api.trakt.tv/users/settings", headers=entetes_trakt(access_token))
    response.raise_for_status()
    data = response.json()
    tz_str = data["user"].get("timezone", "Europe/Paris")
    try:
        user_tz = pytz.timezone(tz_str)
    except Exception:
        user_tz = pytz.timezone("Europe/Paris")
    return {"pseudo": data["user"]["username"], "timezone": user_tz, "tz_name": tz_str, "joined_at": data["user"].get("joined_at")}

def generer_qr_code(url):
    image = qrcode.make(url)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()

def recuperer_historique(access_token, barre=None):
    headers = entetes_trakt(access_token)
    films = {}
    series = {}
    films_details = []
    episodes_details = []
    ratings = {}
    nb_visionnages_films = 0
    nb_episodes = 0

    premiere_page = requests.get("https://api.trakt.tv/users/me/history", headers=headers, params={"page": 1, "limit": 100, "extended": "full"})
    premiere_page.raise_for_status()
    total_pages = int(premiere_page.headers.get("X-Pagination-Page-Count", 1))

    for page in range(1, total_pages + 1):
        if barre:
            barre.progress(page / total_pages * 0.6, text=f"Récupération historique : page {page}/{total_pages}")
        reponse = requests.get("https://api.trakt.tv/users/me/history", headers=headers, params={"page": page, "limit": 100, "extended": "full"})
        reponse.raise_for_status()
        for item in reponse.json():
            if item["type"] == "movie":
                nb_visionnages_films += 1
                film = item["movie"]
                identifiant = film["ids"]["trakt"]
                note = item.get("rated_at")
                rating_value = item.get("rating")
                if rating_value:
                    ratings[("movie", identifiant)] = rating_value
                films_details.append({
                    "titre": film["title"], "annee": film["year"],
                    "genre": ", ".join(film.get("genres", [])) if film.get("genres") else "Inconnu",
                    "duree": film.get("runtime", 0) or 0, "note": rating_value or film.get("rating", 0) or 0,
                    "date_visionnage": item["watched_at"], "id": identifiant
                })
                if identifiant not in films:
                    films[identifiant] = {"titre": film["title"], "annee": film["year"], "vues": 1, "dernier_visionnage": item["watched_at"]}
                else:
                    films[identifiant]["vues"] += 1
            elif item["type"] == "episode":
                nb_episodes += 1
                serie = item["show"]
                episode = item["episode"]
                identifiant = serie["ids"]["trakt"]
                rating_value = item.get("rating")
                if rating_value:
                    ratings[("episode", episode["ids"]["trakt"])] = rating_value
                episodes_details.append({
                    "serie": serie["title"], "titre_episode": episode["title"],
                    "saison": episode["season"], "episode": episode["number"], "annee": serie["year"],
                    "genre": ", ".join(serie.get("genres", [])) if serie.get("genres") else "Inconnu",
                    "duree": (episode.get("runtime", 0) or serie.get("runtime", 0) or 40),
                    "note": rating_value or serie.get("rating", 0) or 0,
                    "date_visionnage": item["watched_at"], "id": identifiant
                })
                if identifiant not in series:
                    series[identifiant] = {"titre": serie["title"], "annee": serie["year"], "vues": 1, "dernier_visionnage": item["watched_at"]}
                else:
                    series[identifiant]["vues"] += 1
    return {
        "films": films, "series": series, "films_details": films_details, "episodes_details": episodes_details,
        "nb_films": len(films), "nb_series": len(series),
        "nb_visionnages_films": nb_visionnages_films, "nb_episodes": nb_episodes,
        "ratings": ratings
    }

def recuperer_listes(access_token):
    reponse = requests.get("https://api.trakt.tv/users/me/lists", headers=entetes_trakt(access_token))
    reponse.raise_for_status()
    return reponse.json()

def recuperer_contenu_liste(access_token, list_id):
    headers = entetes_trakt(access_token)
    items_total = []
    page = 1
    while True:
        reponse = requests.get(f"https://api.trakt.tv/users/me/lists/{list_id}/items", headers=headers, params={"page": page, "limit": 100})
        reponse.raise_for_status()
        items = reponse.json()
        if not items:
            break
        items_total.extend(items)
        page += 1
    return items_total

def recuperer_watchlist(access_token):
    headers = entetes_trakt(access_token)
    items_total = []
    page = 1
    while True:
        reponse = requests.get("https://api.trakt.tv/users/me/watchlist", headers=headers, params={"page": page, "limit": 100})
        reponse.raise_for_status()
        items = reponse.json()
        if not items:
            break
        items_total.extend(items)
        page += 1
    return items_total

def compter_types(items):
    return sum(1 for i in items if i["type"] == "movie"), sum(1 for i in items if i["type"] == "show")

def comparer_items_avec_historique(items, historique):
    resultats = []
    for item in items:
        if item["type"] == "movie":
            film = item["movie"]
            identifiant = film["ids"]["trakt"]
            if identifiant in historique["films"]:
                vu = historique["films"][identifiant]
                resultats.append({"type": "Film", "titre": film["title"], "annee": film["year"], "vues": vu["vues"], "dernier_visionnage": vu["dernier_visionnage"], "trakt_id": identifiant, "tmdb_id": film["ids"].get("tmdb")})
        elif item["type"] == "show":
            serie = item["show"]
            identifiant = serie["ids"]["trakt"]
            if identifiant in historique["series"]:
                vu = historique["series"][identifiant]
                resultats.append({"type": "Série", "titre": serie["title"], "annee": serie["year"], "vues": vu["vues"], "dernier_visionnage": vu["dernier_visionnage"], "trakt_id": identifiant, "tmdb_id": serie["ids"].get("tmdb")})
    return resultats

def analyser_toutes_les_donnees(access_token, historique, barre=None):
    resultats = []
    stats_listes = []
    apparitions = {}

    def enregistrer_apparition(item, nom_liste, liste_id):
        if item["type"] == "movie":
            media, type_affiche = item["movie"], "Film"
        elif item["type"] == "show":
            media, type_affiche = item["show"], "Série"
        else:
            return
        trakt_id = media["ids"]["trakt"]
        cle = (type_affiche, trakt_id)
        if cle not in apparitions:
            apparitions[cle] = {"titre": media["title"], "annee": media["year"], "type": type_affiche, "trakt_id": trakt_id, "tmdb_id": media["ids"].get("tmdb"), "vu_dans": []}
        apparitions[cle]["vu_dans"].append({"nom_liste": nom_liste, "liste_id": liste_id})

    if barre:
        barre.progress(0.6, text="Analyse de la liste de suivi...")
    watchlist = recuperer_watchlist(access_token)
    for item in watchlist:
        enregistrer_apparition(item, "Liste de suivi", "watchlist")
    matches = comparer_items_avec_historique(watchlist, historique)
    for m in matches:
        m["liste"] = "Liste de suivi"
        m["liste_id"] = "watchlist"
    resultats.extend(matches)
    nb_films, nb_series = compter_types(watchlist)
    stats_listes.append({"nom": "Liste de suivi (officielle)", "nb_films": nb_films, "nb_series": nb_series, "total": len(watchlist), "deja_vus": len(matches)})

    listes = recuperer_listes(access_token)
    for i, liste in enumerate(listes):
        if barre:
            barre.progress(0.6 + (i+1)/max(len(listes),1)*0.3, text=f"Analyse : {liste['name']}")
        items = recuperer_contenu_liste(access_token, liste["ids"]["trakt"])
        for item in items:
            enregistrer_apparition(item, liste["name"], liste["ids"]["trakt"])
        matches = comparer_items_avec_historique(items, historique)
        for m in matches:
            m["liste"] = liste["name"]
            m["liste_id"] = liste["ids"]["trakt"]
        resultats.extend(matches)
        nb_films, nb_series = compter_types(items)
        stats_listes.append({"nom": liste["name"], "nb_films": nb_films, "nb_series": nb_series, "total": len(items), "deja_vus": len(matches)})

    doublons, doublons_detail = [], []
    for info in apparitions.values():
        if len(info["vu_dans"]) >= 2:
            doublons.append({"type": info["type"], "titre": info["titre"], "annee": info["annee"], "tmdb_id": info["tmdb_id"], "nombre_listes": len(info["vu_dans"]), "listes": ", ".join(v["nom_liste"] for v in info["vu_dans"])})
            for v in info["vu_dans"]:
                doublons_detail.append({"type": info["type"], "titre": info["titre"], "annee": info["annee"], "trakt_id": info["trakt_id"], "liste": v["nom_liste"], "liste_id": v["liste_id"]})
    return resultats, stats_listes, doublons, doublons_detail

def recuperer_progressions(access_token, barre=None):
    if barre:
        barre.progress(0.95, text="Recherche des progressions fantômes...")
    reponse = requests.get("https://api.trakt.tv/sync/playback", headers=entetes_trakt(access_token))
    reponse.raise_for_status()
    resultats = []
    for item in reponse.json():
        if item["type"] == "movie" and item.get("movie"):
            media, titre, annee = item["movie"], item["movie"]["title"], item["movie"].get("year")
        elif item["type"] == "episode" and item.get("show") and item.get("episode"):
            ep = item["episode"]
            titre = f"{item['show']['title']} — S{ep['season']:02d}E{ep['number']:02d}" if ep.get("season") and ep.get("number") else item['show']['title']
            annee = item["show"].get("year")
        else:
            continue
        resultats.append({"type": "Film" if item["type"] == "movie" else "Épisode", "titre": titre, "annee": annee, "progression": round(item.get("progress",0)), "dernier_visionnage": item["paused_at"], "playback_id": item["id"]})
    resultats.sort(key=lambda x: x["dernier_visionnage"])
    return resultats

def lancer_analyse_complete(raffraichir_historique=False):
    barre = st.progress(0, text="Démarrage...")
    if raffraichir_historique or "historique" not in st.session_state:
        st.session_state["historique"] = recuperer_historique(st.session_state["access_token"], barre)
    resultats, stats_listes, doublons, doublons_detail = analyser_toutes_les_donnees(st.session_state["access_token"], st.session_state["historique"], barre)
    playback = recuperer_progressions(st.session_state["access_token"], barre)
    st.session_state["resultats"] = resultats
    st.session_state["stats_listes"] = stats_listes
    st.session_state["doublons"] = doublons
    st.session_state["doublons_detail"] = doublons_detail
    st.session_state["playback"] = playback
    barre.empty()
    st.rerun()

# ==================================================
# FONCTIONS SUPPRESSION
# ==================================================

def supprimer_de_liste(access_token, liste_id, items_a_supprimer):
    corps = {"movies": [], "shows": []}
    for item in items_a_supprimer:
        cible = corps["movies"] if item["type"] == "Film" else corps["shows"]
        cible.append({"ids": {"trakt": item["trakt_id"]}})
    url = "https://api.trakt.tv/sync/watchlist/remove" if liste_id == "watchlist" else f"https://api.trakt.tv/users/me/lists/{liste_id}/items/remove"
    requests.post(url, headers=entetes_trakt(access_token), json=corps).raise_for_status()

def supprimer_selection(access_token, items_selectionnes):
    par_liste = {}
    for item in items_selectionnes:
        par_liste.setdefault(item["liste_id"], []).append(item)
    for liste_id, items in par_liste.items():
        supprimer_de_liste(access_token, liste_id, items)
        time.sleep(0.8)

def supprimer_progressions(access_token, items_selectionnes):
    for item in items_selectionnes:
        requests.delete(f"https://api.trakt.tv/sync/playback/{item['playback_id']}", headers=entetes_trakt(access_token)).raise_for_status()
        time.sleep(0.4)

# ==================================================
# EXPORT EXCEL
# ==================================================

def ajuster_largeur_colonnes(ws):
    for colonne in ws.columns:
        longueur = 0
        lettre = get_column_letter(colonne[0].column)
        for cellule in colonne:
            try:
                if len(str(cellule.value)) > longueur:
                    longueur = len(str(cellule.value))
            except Exception:
                pass
        ws.column_dimensions[lettre].width = longueur + 4

def mettre_en_forme_feuille(ws, couleur_entete="00665F"):
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        table = Table(displayName=f"Table_{ws.title.replace(' ', '_')}", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(table)
    for cellule in ws[1]:
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = PatternFill(start_color=couleur_entete, end_color=couleur_entete, fill_type="solid")
        cellule.alignment = Alignment(horizontal="center")
    ajuster_largeur_colonnes(ws)

def generer_rapport_excel(pseudo, historique, resultats, stats_listes, doublons, playback, user_tz):
    total_heures = sum(m["duree"] for m in historique["films_details"]) / 60 + sum(e["duree"] for e in historique["episodes_details"]) /60
    df_resume = pd.DataFrame([
        ["Compte Trakt", pseudo], ["Fuseau horaire", user_tz.zone],
        ["Films vus (uniques)", historique["nb_films"]], ["Séries vues (uniques)", historique["nb_series"]],
        ["Épisodes vus", historique["nb_episodes"]], ["Temps total de visionnage", format_duree(total_heures)],
        ["Listes personnalisées", len(stats_listes) -1], ["Contenus dans listes + suivi", sum(s["total"] for s in stats_listes)],
        ["Contenus déjà vus", len(resultats)], ["Doublons", len(doublons)], ["Progressions fantômes", len(playback)]
    ], columns=["Statistique", "Valeur"])
    df_resultats = pd.DataFrame(resultats)
    if not df_resultats.empty:
        df_resultats = df_resultats[["liste", "type", "titre", "annee", "vues", "dernier_visionnage", "tmdb_id"]].copy()
        df_resultats["dernier_visionnage"] = pd.to_datetime(df_resultats["dernier_visionnage"]).dt.tz_convert(user_tz).dt.strftime("%d/%m/%Y %H:%M")
        df_resultats.columns = ["Liste", "Type", "Titre", "Année", "Vues", "Dernier visionnage", "ID TMDB"]
    else:
        df_resultats = pd.DataFrame(columns=["Liste", "Type", "Titre", "Année", "Vues", "Dernier visionnage", "ID TMDB"])
    df_doublons = pd.DataFrame(doublons)
    if not df_doublons.empty:
        df_doublons = df_doublons[["type", "titre", "annee", "tmdb_id", "nombre_listes", "listes"]].copy()
        df_doublons.columns = ["Type", "Titre", "Année", "ID TMDB", "Nombre de listes", "Présent dans"]
    else:
        df_doublons = pd.DataFrame(columns=["Type", "Titre", "Année", "ID TMDB", "Nombre de listes", "Présent dans"])
    df_listes = pd.DataFrame(stats_listes)
    df_listes["% nettoyage possible"] = (df_listes["deja_vus"] / df_listes["total"].replace(0,1) * 100).round(1)
    df_listes = df_listes[["nom", "nb_films", "nb_series", "total", "deja_vus", "% nettoyage possible"]]
    df_listes.columns = ["Liste", "Films", "Séries", "Nombre de contenus", "Déjà vus", "% nettoyage possible"]
    df_fantomes = pd.DataFrame(playback)
    if not df_fantomes.empty:
        df_fantomes = df_fantomes[["type", "titre", "annee", "progression", "dernier_visionnage"]].copy()
        df_fantomes["dernier_visionnage"] = pd.to_datetime(df_fantomes["dernier_visionnage"]).dt.tz_convert(user_tz).dt.strftime("%d/%m/%Y %H:%M")
        df_fantomes.columns = ["Type", "Titre", "Année", "Progression (%)", "Dernier visionnage"]
    else:
        df_fantomes = pd.DataFrame(columns=["Type", "Titre", "Année", "Progression (%)", "Dernier visionnage"])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_resume.to_excel(writer, sheet_name="Résumé", index=False)
        df_resultats.to_excel(writer, sheet_name="À nettoyer", index=False)
        df_doublons.to_excel(writer, sheet_name="Doublons", index=False)
        df_fantomes.to_excel(writer, sheet_name="Progressions fantômes", index=False)
        df_listes.to_excel(writer, sheet_name="Analyse par liste", index=False)
    buffer.seek(0)
    classeur = load_workbook(buffer)
    for feuille in classeur:
        mettre_en_forme_feuille(feuille)
    # Color scale pour pourcentage
    ws_listes = classeur["Analyse par liste"]
    colonne_pct = None
    for c in ws_listes[1]:
        if c.value == "% nettoyage possible":
            colonne_pct = c.column
    if colonne_pct:
        lettre = get_column_letter(colonne_pct)
        ws_listes.conditional_formatting.add(f"{lettre}2:{lettre}{ws_listes.max_row}", ColorScaleRule(start_type="min", start_color="63BE7B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="F8696B"))
    buffer_final = io.BytesIO()
    classeur.save(buffer_final)
    buffer_final.seek(0)
    return buffer_final.getvalue()

# ==================================================
# ENTÊTE COMMUN
# ==================================================

def afficher_entete():
    col_logo, col_titre = st.columns([0.08, 0.92])
    with col_logo:
        try:
            st.image("trakt-logo.svg", width=60)
        except Exception:
            pass
    with col_titre:
        st.title("Trakt Smart Lists")

    if "access_token" not in st.session_state:
        return None

    if "infos_utilisateur" not in st.session_state:
        st.session_state["infos_utilisateur"] = obtenir_infos_utilisateur(st.session_state["access_token"])
    infos = st.session_state["infos_utilisateur"]
    pseudo = infos["pseudo"]
    user_tz = infos["timezone"]

    colonne_info, colonne_deco = st.columns([4,1])
    with colonne_info:
        st.success(f"Connecté en tant que **{pseudo}** • Fuseau : `{infos['tz_name']}`", icon="👤")
    with colonne_deco:
        if st.button("🚪 Se déconnecter", use_container_width=True):
            oublier_connexion()
            st.rerun()
    st.divider()

    # Boutons globaux sur TOUTES les pages
    if "resultats" in st.session_state:
        historique = st.session_state["historique"]
        resultats = st.session_state["resultats"]
        stats_listes = st.session_state["stats_listes"]
        doublons = st.session_state["doublons"]
        playback = st.session_state["playback"]
        excel_bytes = generer_rapport_excel(pseudo, historique, resultats, stats_listes, doublons, playback, user_tz)
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("🔄 Analyse rapide", use_container_width=True, help="Garde l'historique"):
                for cle in ["resultats", "stats_listes", "doublons", "doublons_detail", "playback"]:
                    st.session_state.pop(cle, None)
                lancer_analyse_complete(False)
        with col2:
            if st.button("🔃 Rafraîchir tout", use_container_width=True, help="Récupère tout l'historique à nouveau"):
                for cle in ["historique", "resultats", "stats_listes", "doublons", "doublons_detail", "playback", "infos_utilisateur"]:
                    st.session_state.pop(cle, None)
                st.rerun()
        with col3:
            st.download_button("📥 Rapport Excel", data=excel_bytes, file_name=f"trakt_{pseudo}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.divider()
    return user_tz

def afficher_bloc_lancement():
    if "resultats" in st.session_state:
        return False
    if "historique" in st.session_state:
        st.info("ℹ️ Ton historique est déjà chargé — l'analyse sera rapide.")
        texte = "🔄 Lancer l'analyse rapide"
    else:
        st.info("ℹ️ Lance l'analyse pour accéder à tous les outils.")
        texte = "🔍 Lancer l'analyse complète"
    if st.button(texte, type="primary", use_container_width=True):
        lancer_analyse_complete()
    return True

def navigation():
    """Affiche la navigation et gère la redirection des boutons, en fermant bien le menu sur mobile."""
    PAGES = [
        "🏠 Tableau de bord",
        "👻 Progression Fantôme",
        "🧹 Nettoyage des listes",
        "🔍 Recherche de doublons",
        "📊 Statistiques détaillées",
        "🎬 Rendez-vous annuel",
        "📤 Sauvegarde / Restauration",
        "🏆 Succès",
    ]
    with st.sidebar:
        st.markdown('<p class="section-menu-title">Menu</p>', unsafe_allow_html=True)
        page = st.radio(
            "Sélectionne une page",
            PAGES,
            index=PAGES.index(st.session_state.get("page_active", PAGES[0])),
            label_visibility="collapsed",
            key="navigation_radio",
        )
        st.session_state["page_active"] = page
    return page

# ==================================================
# PAGES
# ==================================================

def page_connexion():
    if "device_code" not in st.session_state:
        st.write("Connecte ton compte Trakt pour commencer.")
        if st.button("🚀 Se connecter à Trakt", type="primary"):
            infos = demarrer_connexion()
            st.session_state["device_code"] = infos["device_code"]
            st.session_state["user_code"] = infos["user_code"]
            st.session_state["verification_url"] = infos["verification_url"]
            st.session_state["expires_in"] = infos["expires_in"]
            st.session_state["interval"] = infos["interval"]
            st.rerun()
    else:
        url_complete = f"{st.session_state['verification_url']}/{st.session_state['user_code']}"
        col_g, col_d = st.columns(2)
        with col_g:
            st.markdown(f'<a href="{url_complete}" target="_blank" style="display:inline-block; background-color:#00A392; color:white; padding:0.8em 1.6em; border-radius:10px; text-decoration:none; font-weight:600;">Autoriser l\'accès</a>', unsafe_allow_html=True)
            st.caption("Sur cet appareil ou un autre.")
            st.info(f"Code : **{st.session_state['user_code']}**")
        with col_d:
            st.image(generer_qr_code(url_complete), width=160)
            st.caption("Ou scanne le QR code.")
        st.caption("La page se met à jour automatiquement.")
        with st.spinner("Attente de l'autorisation..."):
            tps = 0
            while tps < st.session_state["expires_in"]:
                time.sleep(st.session_state["interval"])
                tps += st.session_state["interval"]
                try:
                    tokens = verifier_connexion(st.session_state["device_code"])
                except Exception as e:
                    st.error(str(e))
                    break
                if tokens:
                    sauvegarder_connexion(tokens)
                    del st.session_state["device_code"]
                    st.rerun()
        if st.button("Réessayer"):
            st.rerun()

def page_tableau_de_bord(user_tz):
    if afficher_bloc_lancement():
        return
    histo = st.session_state["historique"]
    resultats = st.session_state["resultats"]
    stats_listes = st.session_state["stats_listes"]
    doublons = st.session_state["doublons"]
    playback = st.session_state["playback"]

    # Calcul temps total
    total_heures = sum(m["duree"] for m in histo["films_details"]) / 60 + sum(e["duree"] for e in histo["episodes_details"]) / 60

    st.subheader("📊 Vue d'ensemble")
    # Première ligne de métriques
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🎬 Films", histo["nb_films"])
    c2.metric("📺 Séries", histo["nb_series"])
    c3.metric("🎞️ Épisodes", histo["nb_episodes"])
    c4.metric("⏱️ Temps total", format_duree(total_heures))

    st.divider()
    st.subheader("⚠️ État du nettoyage")
    c5, c6, c7 = st.columns(3)

    # Fantômes
    with c5:
        with st.container(border=True):
            st.markdown("#### 👻 Progressions fantômes")
            st.metric("Nombre", len(playback))
            if len(playback) > 0:
                st.warning(f"{len(playback)} fantôme(s) à nettoyer")
                if st.button("Aller au nettoyage →", key="btn_fantomes"):
                    st.session_state["page_active"] = "👻 Progression Fantôme"
                    st.rerun()
            else:
                st.success("✅ 0 fantôme, c'est propre !")
    # Doublons
    with c6:
        with st.container(border=True):
            st.markdown("#### 🔁 Doublons")
            st.metric("Nombre", len(doublons))
            if len(doublons) > 0:
                st.warning(f"{len(doublons)} doublon(s) détecté(s)")
                if st.button("Voir les doublons →", key="btn_doublons"):
                    st.session_state["page_active"] = "🔍 Recherche de doublons"
                    st.rerun()
            else:
                st.success("✅ Aucun doublon !")
    # Contenus déjà vus
    with c7:
        with st.container(border=True):
            st.markdown("#### 🧹 Contenus déjà vus")
            total_items = sum(s["total"] for s in stats_listes)
            pct = round(len(resultats)/total_items*100, 1) if total_items > 0 else 0
            st.metric("Nombre", len(resultats), delta=f"{pct}% des listes")
            if len(resultats) > 0:
                st.warning(f"{len(resultats)} contenu(s) déjà vu(s)")
                if st.button("Nettoyer les listes →", key="btn_vus"):
                    st.session_state["page_active"] = "🧹 Nettoyage des listes"
                    st.rerun()
            else:
                st.success("✅ Listes à jour !")

def page_nettoyage_listes(user_tz):
    if afficher_bloc_lancement():
        return
    resultats = st.session_state["resultats"]
    stats_listes = st.session_state["stats_listes"]
    cle_msg = "msg_suppression_vus"
    if st.session_state.get(cle_msg):
        st.success(st.session_state[cle_msg])
        del st.session_state[cle_msg]
    st.subheader("🧹 Nettoyage des listes")
    st.caption("Retire les films/séries que tu as déjà vus de ta liste de suivi.")
    if not resultats:
        st.success("Aucun contenu déjà vu dans tes listes. C'est parfait ! 🎉")
    else:
        st.write(f"**{len(resultats)}** contenu(s) déjà vu(s) trouvé(s). Coche ceux à supprimer :")
        tab = pd.DataFrame(resultats)
        tab_aff = tab[["type", "titre", "annee", "vues", "dernier_visionnage", "liste"]].copy()
        tab_aff["dernier_visionnage"] = pd.to_datetime(tab_aff["dernier_visionnage"]).dt.tz_convert(user_tz).dt.strftime("%d/%m/%Y %H:%M")
        tab_aff.insert(0, "Sélectionner", False)
        tab_aff.columns = ["Sélectionner", "Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"]
        edite = st.data_editor(tab_aff, use_container_width=True, hide_index=True, disabled=["Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"], key="edit_vus")
        nb_sel = int(edite["Sélectionner"].sum())
        if nb_sel:
            cle_conf = "conf_vus"
            if not st.session_state.get(cle_conf, False):
                if st.button(f"🗑️ Supprimer {nb_sel} élément(s)", type="primary"):
                    st.session_state[cle_conf] = True
                    st.rerun()
            else:
                st.warning(f"Confirmer la suppression ?")
                co, cn = st.columns(2)
                with co:
                    if st.button("✅ Oui"):
                        idx = edite[edite["Sélectionner"]].index
                        items = [resultats[i] for i in idx]
                        with st.spinner("Suppression..."):
                            supprimer_selection(st.session_state["access_token"], items)
                        st.session_state[cle_conf] = False
                        st.session_state[cle_msg] = f"✅ {len(items)} élément(s) supprimé(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[cle_conf] = False
                        st.rerun()
    st.divider()
    st.subheader("Taux de contenu à nettoyer par liste")
    df_sl = pd.DataFrame(stats_listes)
    df_sl["% nettoyable"] = (df_sl["deja_vus"] / df_sl["total"].replace(0, 1)*100).round(1)
    st.bar_chart(df_sl.set_index("nom")["% nettoyable"], color="#CEDC00")
    st.caption("Ce graphique te permet de voir rapidement quelle liste a le plus grand besoin de nettoyage.")

def page_doublons(user_tz):
    if afficher_bloc_lancement():
        return
    doublons_detail = st.session_state["doublons_detail"]
    cle_msg = "msg_suppression_doublons"
    if st.session_state.get(cle_msg):
        st.success(st.session_state[cle_msg])
        del st.session_state[cle_msg]
    st.subheader("🔍 Recherche de doublons")
    st.caption("Trouve les contenus présents dans plusieurs listes à la fois.")
    if not doublons_detail:
        st.success("Aucun doublon.")
    else:
        st.write(f"**{len(st.session_state['doublons'])}** doublon(s) détecté(s).")
        tab = pd.DataFrame(doublons_detail)
        tab_aff = tab[["type", "titre", "annee", "liste"]].copy()
        tab_aff.insert(0, "Sélectionner", False)
        tab_aff.columns = ["Sélectionner", "Type", "Titre", "Année", "Liste"]
        edite = st.data_editor(tab_aff, use_container_width=True, hide_index=True, disabled=["Type", "Titre", "Année", "Liste"], key="edit_doublons")
        nb_sel = int(edite["Sélectionner"].sum())
        if nb_sel:
            cle_conf = "conf_doublons"
            if not st.session_state.get(cle_conf, False):
                if st.button(f"🗑️ Retirer {nb_sel} élément(s)", type="primary"):
                    st.session_state[cle_conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer ?")
                co, cn = st.columns(2)
                with co:
                    if st.button("✅ Oui"):
                        idx = edite[edite["Sélectionner"]].index
                        items = [doublons_detail[i] for i in idx]
                        with st.spinner("Suppression..."):
                            supprimer_selection(st.session_state["access_token"], items)
                        st.session_state[cle_conf] = False
                        st.session_state[cle_msg] = f"✅ {len(items)} élément(s) retiré(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[cle_conf] = False
                        st.rerun()

def page_fantomes(user_tz):
    if afficher_bloc_lancement():
        return
    playback = st.session_state["playback"]
    cle_msg = "msg_suppression_fantomes"
    if st.session_state.get(cle_msg):
        st.success(st.session_state[cle_msg])
        del st.session_state[cle_msg]
    st.subheader("👻 Progression Fantôme")
    st.caption("Supprime les entrées bloquées dans 'Continuer à regarder'.")
    st.divider()
    if not playback:
        st.success("Aucune progression en cours.")
    else:
        tout = st.checkbox("Tout sélectionner")
        selections = {}
        for item in playback:
            pct = item["progression"]
            cls = "progress-low" if pct < 30 else "progress-mid" if pct < 80 else "progress-high"
            date_f = formater_date(item["dernier_visionnage"], user_tz)
            icone = "🎬" if item["type"] == "Film" else "📺"
            with st.container():
                cc, cd = st.columns([0.05, 0.95])
                with cc:
                    selections[item["playback_id"]] = st.checkbox("", value=tout, key=f"chk_{item['playback_id']}", label_visibility="collapsed")
                with cd:
                    st.markdown(f"""
                    <div class="ghost-card">
                        <div class="ghost-title">{icone} {item['titre']} {f'({item["annee"]})' if item["annee"] else ''}</div>
                        <div class="ghost-meta">{item['type']} • {pct}% • 🕒 {date_f}</div>
                        <div class="progress-bar-container">
                            <div class="progress-bar-fill {cls}" style="width: {pct}%"></div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
        st.divider()
        ids_sel = [pid for pid, s in selections.items() if s]
        if ids_sel:
            cle_conf = "conf_fantomes"
            if not st.session_state.get(cle_conf, False):
                if st.button(f"🗑️ Supprimer {len(ids_sel)} progression(s)", type="primary"):
                    st.session_state[cle_conf] = True
                    st.rerun()
            else:
                st.warning("Confirmer ?")
                co, cn = st.columns(2)
                with co:
                    if st.button("✅ Oui"):
                        items = [p for p in playback if p["playback_id"] in ids_sel]
                        with st.spinner("Suppression..."):
                            supprimer_progressions(st.session_state["access_token"], items)
                        st.session_state[cle_conf] = False
                        st.session_state[cle_msg] = f"✅ {len(items)} fantôme(s) supprimé(s)."
                        st.rerun()
                with cn:
                    if st.button("❌ Annuler"):
                        st.session_state[cle_conf] = False
                        st.rerun()
        else:
            st.info("Coche les éléments à supprimer.")

def page_statistiques(user_tz):
    if afficher_bloc_lancement():
        return
    st.subheader("📊 Statistiques détaillées")
    st.caption("Toutes tes données de visionnage, comme Trakt VIP / Cinopsys.")

    histo = st.session_state["historique"]
    films = pd.DataFrame(histo["films_details"])
    episodes = pd.DataFrame(histo["episodes_details"])

    # Filtres
    f1, f2, f3 = st.columns(3)
    with f1:
        type_c = st.selectbox("Type de contenu", ["Tous", "Films", "Séries"])
    with f2:
        periode = st.selectbox("Période", ["Tout", "Cette année", "12 derniers mois", "Ce mois-ci", "Aujourd'hui"])
    with f3:
        genres_dispo = set()
        if type_c in ["Tous", "Films"] and not films.empty:
            for g in films["genre"].str.split(", "):
                genres_dispo.update([x for x in g if x != "Inconnu"])
        if type_c in ["Tous", "Séries"] and not episodes.empty:
            for g in episodes["genre"].str.split(", "):
                genres_dispo.update([x for x in g if x != "Inconnu"])
        genre = st.selectbox("Genre", ["Tous"] + sorted(genres_dispo))

    # Construction df complet
    dfs = []
    if type_c in ["Tous", "Films"] and not films.empty:
        df = films.copy()
        df["type"] = "Film"
        df["date"] = pd.to_datetime(df["date_visionnage"], utc=True).dt.tz_convert(user_tz)
        dfs.append(df)
    if type_c in ["Tous", "Séries"] and not episodes.empty:
        df = episodes.copy()
        df["type"] = "Épisode"
        df["titre"] = df["serie"]
        df["date"] = pd.to_datetime(df["date_visionnage"], utc=True).dt.tz_convert(user_tz)
        dfs.append(df)
    if not dfs:
        st.info("Aucune donnée.")
        return
    df = pd.concat(dfs, ignore_index=True)
    mt = datetime.now(user_tz)
    if periode == "Cette année":
        df = df[df["date"].dt.year == mt.year]
    elif periode == "12 derniers mois":
        df = df[df["date"] >= mt - pd.DateOffset(months=12)]
    elif periode == "Ce mois-ci":
        df = df[(df["date"].dt.year == mt.year) & (df["date"].dt.month == mt.month)]
    elif periode == "Aujourd'hui":
        df = df[(df["date"].dt.date == mt.date())]
    if genre != "Tous":
        df = df[df["genre"].str.contains(genre, na=False)]
    if df.empty:
        st.warning("Aucun résultat pour ces filtres.")
        return

    df["duree_h"] = df["duree"].fillna(40) / 60
    total_h = df["duree_h"].sum()

    # Métriques
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Visionnages", len(df))
    m2.metric("Temps total", format_duree(total_h))
    note_moy = df[df["note"] > 0]["note"].mean() if "note" in df.columns else 0
    m3.metric("Note moyenne", f"{round(note_moy,1)}/10" if note_moy == note_moy else "-")
    moy_vue_par_jour = len(df) / max((df["date"].max() - df["date"].min()).days + 1, 1)
    m4.metric("Moyenne / jour", f"{round(moy_vue_par_jour,1)}")
    marathon = df.sort_values("date").groupby(df["date"].dt.date).size().max()
    m5.metric("Record en 1 jour", f"{marathon} contenu{'s' if marathon>1 else ''}")

    st.divider()

    # Graphiques
    # Par mois
    df["mois"] = df["date"].dt.strftime("%Y-%m")
    h_par_mois = df.groupby("mois")["duree_h"].sum().round(1)
    opt_mois = {"title": {"text": "Heures par mois", "textStyle": {"color": "#F0FAF8"}}, "tooltip": {"trigger": "axis"}, "backgroundColor": "transparent", "textStyle": {"color": "#F0FAF8"}, "xAxis": {"type": "category", "data": list(h_par_mois.index), "axisLabel": {"color": "#9DC5BF"}}, "yAxis": {"type": "value", "axisLabel": {"color": "#9DC5BF"}, "splitLine": {"lineStyle": {"color": "#125A54"}}}, "series": [{"data": list(h_par_mois.values), "type": "line", "smooth": True, "lineStyle": {"color": "#CEDC00", "width":3}, "areaStyle": {"color": "rgba(206,220,0,0.1)"}, "itemStyle": {"color": "#CEDC00"}}]}
    st_echarts(opt_mois, height="350px")

    g1, g2 = st.columns(2)
    # Genres
    with g1:
        genres = {}
        for liste_g in df["genre"].str.split(", "):
            for g in liste_g:
                if g and g != "Inconnu":
                    genres[g] = genres.get(g, 0) +1
        opt_genres = {"title": {"text": "Genres les plus regardés", "left": "center", "textStyle": {"color": "#F0FAF8"}}, "tooltip": {"trigger": "item"}, "backgroundColor": "transparent", "textStyle": {"color": "#F0FAF8"}, "legend": {"bottom": 0, "textStyle": {"color": "#9DC5BF"}}, "series": [{"type": "pie", "radius": ["40%", "70%"], "data": [{"name":k, "value":v} for k,v in sorted(genres.items(), key=lambda x:-x[1])[:8]], "itemStyle": {"borderRadius": 8, "borderColor": "#04332F", "borderWidth":2}, "label": {"color": "#F0FAF8"}}], "color": ["#00A392", "#CEDC00", "#00C7B3", "#A3B300", "#007D70", "#869400", "#125A54", "#E8F064"]}
        st_echarts(opt_genres, height="400px")
    # Heures de la journée
    with g2:
        df["h"] = df["date"].dt.hour
        h_par_h = df.groupby("h")["duree_h"].sum().reindex(range(24), fill_value=0).round(1)
        opt_heures = {"title": {"text": "Par heure de la journée", "left": "center", "textStyle": {"color": "#F0FAF8"}}, "tooltip": {"trigger": "axis"}, "backgroundColor": "transparent", "textStyle": {"color": "#F0FAF8"}, "xAxis": {"type": "category", "data": [f"{h}h" for h in range(24)], "axisLabel": {"color": "#9DC5BF"}}, "yAxis": {"type": "value", "axisLabel": {"color": "#9DC5BF"}, "splitLine": {"lineStyle": {"color": "#125A54"}}}, "series": [{"data": list(h_par_h.values), "type": "bar", "itemStyle": {"color": {"type":"linear", "x":0, "y":0, "x2":0, "y2":1, "colorStops": [{"offset":0, "color":"#00A392"}, {"offset":1, "color":"#007D70"}]}, "borderRadius": [4,4,0,0]}}]}
        st_echarts(opt_heures, height="400px")

    g3, g4 = st.columns(2)
    # Par jour de la semaine
    with g3:
        jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
        df["jour_semaine"] = df["date"].dt.weekday
        h_par_jour = df.groupby("jour_semaine")["duree_h"].sum().reindex(range(7), fill_value=0).round(1)
        opt_jours = {"title": {"text": "Par jour de la semaine", "left": "center", "textStyle": {"color": "#F0FAF8"}}, "tooltip": {"trigger": "axis"}, "backgroundColor": "transparent", "textStyle": {"color": "#F0FAF8"}, "xAxis": {"type": "category", "data": jours, "axisLabel": {"color": "#9DC5BF"}}, "yAxis": {"type": "value", "axisLabel": {"color": "#9DC5BF"}, "splitLine": {"lineStyle": {"color": "#125A54"}}}, "series": [{"data": list(h_par_jour.values), "type": "bar", "itemStyle": {"color": "#CEDC00", "borderRadius": [4,4,0,0]}}]}
        st_echarts(opt_jours, height="400px")
    # Répartition films / séries
    with g4:
        rep_type = df.groupby("type")["duree_h"].sum().round(1)
        opt_type = {"title": {"text": "Films vs Séries", "left": "center", "textStyle": {"color": "#F0FAF8"}}, "tooltip": {"trigger": "item"}, "backgroundColor": "transparent", "legend": {"bottom": 0, "textStyle": {"color": "#9DC5BF"}}, "series": [{"type": "pie", "radius": ["40%", "70%"], "data": [{"value": v, "name": k} for k,v in rep_type.items()], "itemStyle": {"borderRadius":8, "borderColor":"#04332F", "borderWidth":2}, "label": {"color":"#F0FAF8"}}], "color": ["#00A392", "#CEDC00"]}
        st_echarts(opt_type, height="400px")

def page_wrapped():
    st.subheader("🎬 Rendez-vous annuel")
    st.info("🚧 Bientôt disponible : ton récap annuel façon Spotify Wrapped avec tes tops, statistiques et carte partageable !")

def page_sauvegarde():
    st.subheader("📤 Sauvegarde et restauration")
    st.info("🚧 Bientôt : export complet de toutes tes données / listes / historique en JSON/Excel, et restauration.")

def page_succes():
    st.subheader("🏆 Succès")
    st.info("🚧 Bientôt : badges et objectifs : marathon du week-end, 100 films vus, série la plus longue, etc.")

# ==================================================
# RECONNEXION AUTO
# ==================================================

if "access_token" not in st.session_state:
    rt = cookies.get("trakt_refresh_token")
    if rt:
        tokens = rafraichir_token(rt)
        if tokens:
            sauvegarder_connexion(tokens)
        else:
            cookies.remove("trakt_refresh_token")

# ==================================================
# LANCEMENT
# ==================================================

user_tz = afficher_entete()

if "access_token" not in st.session_state:
    page_connexion()
else:
    page = navigation()
    if page == "🏠 Tableau de bord":
        page_tableau_de_bord(user_tz)
    elif page == "👻 Progression Fantôme":
        page_fantomes(user_tz)
    elif page == "🧹 Nettoyage des listes":
        page_nettoyage_listes(user_tz)
    elif page == "🔍 Recherche de doublons":
        page_doublons(user_tz)
    elif page == "📊 Statistiques détaillées":
        page_statistiques(user_tz)
    elif page == "🎬 Rendez-vous annuel":
        page_wrapped()
    elif page == "📤 Sauvegarde / Restauration":
        page_sauvegarde()
    elif page == "🏆 Succès":
        page_succes()
