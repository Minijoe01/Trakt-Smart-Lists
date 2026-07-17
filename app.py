import streamlit as st
import requests
import time
import qrcode
import io
import pandas as pd
import pytz
from datetime import datetime
from streamlit_cookies_controller import CookieController
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule
from streamlit_echarts import st_echarts

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🏎️", layout="wide", initial_sidebar_state="auto")

# Injection de styles CSS custom + JS pour fermer automatiquement le menu sur mobile
st.markdown("""
<style>
    /* Couleurs globales thème Aston Martin F1 2026 - Moderne futuriste */
    :root {
        --am-green: #008778;
        --am-green-dark: #00665F;
        --am-lime: #CEDC00;
        --am-bg-dark: #071816;
        --am-bg-card: #0F2B28;
        --am-border: #1A443F;
        --am-text-light: #F0F7F6;
        --am-text-muted: #8FA8A4;
    }

    /* Style global moderne */
    .stApp {
        background: linear-gradient(180deg, var(--am-bg-dark) 0%, #051210 100%);
    }

    /* Style des cartes - ombre douce, coins arrondis */
    div[data-testid="stMetric"], div.stAlert, div[data-testid="stContainer"] {
        background-color: var(--am-bg-card) !important;
        border-radius: 12px !important;
        border: 1px solid var(--am-border) !important;
        padding: 16px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
    }

    /* Barres de progression thème */
    div[role="progressbar"] > div {
        background-color: var(--am-lime) !important;
    }

    /* Style des cartes fantômes */
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
        box-shadow: 0 6px 16px rgba(0, 135, 120, 0.15);
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
        background-color: #1A3330;
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
    .progress-high { background: linear-gradient(90deg, #008778 0%, #00B3A0 100%); }

    /* Style des boutons moderne */
    .stButton > button {
        border-radius: 10px;
        border: 0;
        font-weight: 500;
        padding: 0.6em 1.2em;
        transition: all 0.2s ease;
        border: 1px solid var(--am-border);
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0, 135, 120, 0.25);
        border-color: var(--am-green);
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--am-green) 0%, var(--am-green-dark) 100%);
        border: 0;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 6px 20px rgba(0, 135, 120, 0.4);
    }

    /* Style sidebar */
    section[data-testid="stSidebar"] {
        background-color: #05100F !important;
        border-right: 1px solid var(--am-border);
    }

    /* Titres sections menu */
    .section-menu-title {
        font-size: 0.78em;
        font-weight: 700;
        color: var(--am-lime);
        text-transform: uppercase;
        letter-spacing: 1.2px;
        margin: 18px 0 10px 0;
    }

    /* Style dataframes */
    div[data-testid="stDataFrame"] {
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid var(--am-border);
    }

    /* Style dividers */
    hr {
        border-color: var(--am-border) !important;
    }

    /* Amélioration lisibilité textes */
    p, li, label {
        color: var(--am-text-light) !important;
    }
    .stCaption {
        color: var(--am-text-muted) !important;
    }

    /* Supprimer les outlines moches */
    *:focus {
        outline: 2px solid var(--am-green) !important;
        outline-offset: 2px;
    }

    /* Style metric */
    div[data-testid="stMetricValue"] {
        color: var(--am-text-light) !important;
        font-size: 1.8em !important;
        font-weight: 700;
    }
    div[data-testid="stMetricLabel"] p {
        color: var(--am-text-muted) !important;
        font-size: 0.9em !important;
    }
</style>
<script>
// Ferme automatiquement la sidebar sur mobile quand on choisit une page
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(function() {
        const radioButtons = document.querySelectorAll('section[data-testid="stSidebar"] input[type="radio"]');
        radioButtons.forEach(radio => {
            radio.addEventListener('change', function() {
                // Ferme le menu seulement sur petits écrans (mobile)
                if (window.innerWidth < 768) {
                    const closeButton = document.querySelector('button[kind="header"]');
                    if (closeButton) closeButton.click();
                }
            });
        });
    }, 1000);
});
</script>
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
    """Formate une date Trakt (UTC) dans le fuseau horaire de l'utilisateur, récupéré depuis ses paramètres Trakt."""
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
# FONCTIONS TRAKT — CONNEXION
# ==================================================

def demarrer_connexion():
    response = requests.post(DEVICE_CODE_URL, json={"client_id": CLIENT_ID})
    response.raise_for_status()
    return response.json()

def verifier_connexion(device_code):
    payload = {
        "code": device_code,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
    }
    response = requests.post(DEVICE_TOKEN_URL, json=payload)
    if response.status_code == 200:
        return response.json()
    erreurs = {
        404: "Code invalide, recommence depuis le début.",
        409: "Ce code a déjà été utilisé.",
        410: "Le code a expiré, recommence depuis le début.",
        418: "La connexion a été refusée.",
    }
    if response.status_code in erreurs:
        raise Exception(erreurs[response.status_code])
    return None

def rafraichir_token(refresh_token):
    payload = {
        "refresh_token": refresh_token,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": "urn:ietf:wg:oauth:2.0:oob",
        "grant_type": "refresh_token",
    }
    response = requests.post(REFRESH_TOKEN_URL, json=payload)
    if response.status_code != 200:
        return None
    return response.json()

def sauvegarder_connexion(tokens):
    st.session_state["access_token"] = tokens["access_token"]
    st.session_state["refresh_token"] = tokens["refresh_token"]
    cookies.set("trakt_refresh_token", tokens["refresh_token"])
    time.sleep(0.5)

def oublier_connexion():
    cookies.remove("trakt_refresh_token")
    time.sleep(0.5)
    st.session_state.clear()

def entetes_trakt(access_token):
    return {
        "Content-Type": "application/json",
        "trakt-api-version": "2",
        "trakt-api-key": CLIENT_ID,
        "Authorization": f"Bearer {access_token}",
    }

def obtenir_infos_utilisateur(access_token):
    """Récupère pseudo ET fuseau horaire de l'utilisateur depuis ses paramètres Trakt."""
    response = requests.get("https://api.trakt.tv/users/settings", headers=entetes_trakt(access_token))
    response.raise_for_status()
    data = response.json()
    tz_str = data["user"].get("timezone", "Europe/Paris")
    try:
        user_tz = pytz.timezone(tz_str)
    except Exception:
        user_tz = pytz.timezone("Europe/Paris")
    return {
        "pseudo": data["user"]["username"],
        "timezone": user_tz,
        "tz_name": tz_str
    }

def generer_qr_code(url):
    image = qrcode.make(url)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()

# ==================================================
# FONCTIONS TRAKT — RÉCUPÉRATION DONNÉES
# ==================================================

def recuperer_historique(access_token, barre=None):
    headers = entetes_trakt(access_token)
    films = {}
    series = {}
    films_details = []
    episodes_details = []
    nb_visionnages_films = 0
    nb_episodes = 0

    premiere_page = requests.get(
        "https://api.trakt.tv/users/me/history",
        headers=headers,
        params={"page": 1, "limit": 100, "extended": "full"},
    )
    premiere_page.raise_for_status()
    total_pages = int(premiere_page.headers.get("X-Pagination-Page-Count", 1))

    for page in range(1, total_pages + 1):
        if barre:
            barre.progress(page / total_pages * 0.6, text=f"Récupération de l'historique : page {page}/{total_pages}")
        reponse = requests.get(
            "https://api.trakt.tv/users/me/history",
            headers=headers,
            params={"page": page, "limit": 100, "extended": "full"},
        )
        reponse.raise_for_status()
        for item in reponse.json():
            if item["type"] == "movie":
                nb_visionnages_films += 1
                film = item["movie"]
                identifiant = film["ids"]["trakt"]
                films_details.append({
                    "titre": film["title"],
                    "annee": film["year"],
                    "genre": ", ".join(film.get("genres", [])) if film.get("genres") else "Inconnu",
                    "duree": film.get("runtime", 0) or 0,
                    "note": film.get("rating", 0) or 0,
                    "date_visionnage": item["watched_at"],
                })
                if identifiant not in films:
                    films[identifiant] = {
                        "titre": film["title"],
                        "annee": film["year"],
                        "vues": 1,
                        "dernier_visionnage": item["watched_at"],
                    }
                else:
                    films[identifiant]["vues"] += 1
            elif item["type"] == "episode":
                nb_episodes += 1
                serie = item["show"]
                episode = item["episode"]
                identifiant = serie["ids"]["trakt"]
                episodes_details.append({
                    "serie": serie["title"],
                    "titre_episode": episode["title"],
                    "saison": episode["season"],
                    "episode": episode["number"],
                    "annee": serie["year"],
                    "genre": ", ".join(serie.get("genres", [])) if serie.get("genres") else "Inconnu",
                    "duree": (episode.get("runtime", 0) or serie.get("runtime", 0) or 0),
                    "date_visionnage": item["watched_at"],
                })
                if identifiant not in series:
                    series[identifiant] = {
                        "titre": serie["title"],
                        "annee": serie["year"],
                        "vues": 1,
                        "dernier_visionnage": item["watched_at"],
                    }
                else:
                    series[identifiant]["vues"] += 1
    return {
        "films": films,
        "series": series,
        "films_details": films_details,
        "episodes_details": episodes_details,
        "nb_films": len(films),
        "nb_series": len(series),
        "nb_visionnages_films": nb_visionnages_films,
        "nb_episodes": nb_episodes,
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
        reponse = requests.get(
            f"https://api.trakt.tv/users/me/lists/{list_id}/items",
            headers=headers,
            params={"page": page, "limit": 100},
        )
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
        reponse = requests.get(
            "https://api.trakt.tv/users/me/watchlist",
            headers=headers,
            params={"page": page, "limit": 100},
        )
        reponse.raise_for_status()
        items = reponse.json()
        if not items:
            break
        items_total.extend(items)
        page += 1
    return items_total

def compter_types(items):
    nb_films = sum(1 for i in items if i["type"] == "movie")
    nb_series = sum(1 for i in items if i["type"] == "show")
    return nb_films, nb_series

def comparer_items_avec_historique(items, historique):
    resultats = []
    for item in items:
        if item["type"] == "movie":
            film = item["movie"]
            identifiant = film["ids"]["trakt"]
            if identifiant in historique["films"]:
                vu = historique["films"][identifiant]
                resultats.append({
                    "type": "Film",
                    "titre": film["title"],
                    "annee": film["year"],
                    "vues": vu["vues"],
                    "dernier_visionnage": vu["dernier_visionnage"],
                    "trakt_id": identifiant,
                    "tmdb_id": film["ids"].get("tmdb"),
                })
        elif item["type"] == "show":
            serie = item["show"]
            identifiant = serie["ids"]["trakt"]
            if identifiant in historique["series"]:
                vu = historique["series"][identifiant]
                resultats.append({
                    "type": "Série",
                    "titre": serie["title"],
                    "annee": serie["year"],
                    "vues": vu["vues"],
                    "dernier_visionnage": vu["dernier_visionnage"],
                    "trakt_id": identifiant,
                    "tmdb_id": serie["ids"].get("tmdb"),
                })
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
            apparitions[cle] = {
                "titre": media["title"],
                "annee": media["year"],
                "type": type_affiche,
                "trakt_id": trakt_id,
                "tmdb_id": media["ids"].get("tmdb"),
                "vu_dans": [],
            }
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
    stats_listes.append({
        "nom": "Liste de suivi (officielle)",
        "nb_films": nb_films,
        "nb_series": nb_series,
        "total": len(watchlist),
        "deja_vus": len(matches),
    })
    listes = recuperer_listes(access_token)
    for i, liste in enumerate(listes):
        if barre:
            barre.progress(0.6 + (i + 1) / max(len(listes), 1) * 0.3, text=f"Analyse de la liste : {liste['name']}")
        items = recuperer_contenu_liste(access_token, liste["ids"]["trakt"])
        for item in items:
            enregistrer_apparition(item, liste["name"], liste["ids"]["trakt"])
        matches = comparer_items_avec_historique(items, historique)
        for m in matches:
            m["liste"] = liste["name"]
            m["liste_id"] = liste["ids"]["trakt"]
        resultats.extend(matches)
        nb_films, nb_series = compter_types(items)
        stats_listes.append({
            "nom": liste["name"],
            "nb_films": nb_films,
            "nb_series": nb_series,
            "total": len(items),
            "deja_vus": len(matches),
        })
    doublons = []
    doublons_detail = []
    for info in apparitions.values():
        if len(info["vu_dans"]) >= 2:
            doublons.append({
                "type": info["type"],
                "titre": info["titre"],
                "annee": info["annee"],
                "tmdb_id": info["tmdb_id"],
                "nombre_listes": len(info["vu_dans"]),
                "listes": ", ".join(v["nom_liste"] for v in info["vu_dans"]),
            })
            for v in info["vu_dans"]:
                doublons_detail.append({
                    "type": info["type"],
                    "titre": info["titre"],
                    "annee": info["annee"],
                    "trakt_id": info["trakt_id"],
                    "liste": v["nom_liste"],
                    "liste_id": v["liste_id"],
                })
    return resultats, stats_listes, doublons, doublons_detail

def recuperer_progressions(access_token, barre=None):
    if barre:
        barre.progress(0.95, text="Recherche des progressions fantômes...")
    reponse = requests.get("https://api.trakt.tv/sync/playback", headers=entetes_trakt(access_token))
    reponse.raise_for_status()
    resultats = []
    for item in reponse.json():
        if item["type"] == "movie" and item.get("movie"):
            media = item["movie"]
            titre = media["title"]
            annee = media.get("year")
        elif item["type"] == "episode" and item.get("show") and item.get("episode"):
            episode = item["episode"]
            saison = episode.get("season")
            numero = episode.get("number")
            titre = f"{item['show']['title']} — S{saison:02d}E{numero:02d}" if saison and numero else item['show']['title']
            annee = item["show"].get("year")
        else:
            continue
        resultats.append({
            "type": "Film" if item["type"] == "movie" else "Épisode",
            "titre": titre,
            "annee": annee,
            "progression": round(item.get("progress", 0)),
            "dernier_visionnage": item["paused_at"],
            "playback_id": item["id"],
        })
    resultats.sort(key=lambda x: x["dernier_visionnage"])
    return resultats

def lancer_analyse_complete(raffraichir_historique=False):
    """Lance l'analyse complète et stocke tous les résultats en session."""
    barre = st.progress(0, text="Démarrage...")
    if raffraichir_historique or "historique" not in st.session_state:
        st.session_state["historique"] = recuperer_historique(st.session_state["access_token"], barre)
    resultats, stats_listes, doublons, doublons_detail = analyser_toutes_les_donnees(
        st.session_state["access_token"], st.session_state["historique"], barre
    )
    playback = recuperer_progressions(st.session_state["access_token"], barre)
    st.session_state["resultats"] = resultats
    st.session_state["stats_listes"] = stats_listes
    st.session_state["doublons"] = doublons
    st.session_state["doublons_detail"] = doublons_detail
    st.session_state["playback"] = playback
    barre.empty()
    st.rerun()

# ==================================================
# FONCTIONS TRAKT — SUPPRESSION
# ==================================================

def supprimer_de_liste(access_token, liste_id, items_a_supprimer):
    corps = {"movies": [], "shows": []}
    for item in items_a_supprimer:
        cible = corps["movies"] if item["type"] == "Film" else corps["shows"]
        cible.append({"ids": {"trakt": item["trakt_id"]}})
    url = "https://api.trakt.tv/sync/watchlist/remove" if liste_id == "watchlist" else f"https://api.trakt.tv/users/me/lists/{liste_id}/items/remove"
    reponse = requests.post(url, headers=entetes_trakt(access_token), json=corps)
    reponse.raise_for_status()

def supprimer_selection(access_token, items_selectionnes):
    par_liste = {}
    for item in items_selectionnes:
        par_liste.setdefault(item["liste_id"], []).append(item)
    for liste_id, items in par_liste.items():
        supprimer_de_liste(access_token, liste_id, items)
        time.sleep(1)

def supprimer_progressions(access_token, items_selectionnes):
    for item in items_selectionnes:
        reponse = requests.delete(
            f"https://api.trakt.tv/sync/playback/{item['playback_id']}",
            headers=entetes_trakt(access_token),
        )
        reponse.raise_for_status()
        time.sleep(0.5)

# ==================================================
# FONCTIONS — EXPORT EXPORT_EXCEL
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
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)
    for cellule in ws[1]:
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = PatternFill(start_color=couleur_entete, end_color=couleur_entete, fill_type="solid")
        cellule.alignment = Alignment(horizontal="center")
    ajuster_largeur_colonnes(ws)

def generer_rapport_excel(pseudo, historique, resultats, stats_listes, doublons, playback, user_tz):
    df_resume = pd.DataFrame([
        ["Compte Trakt", pseudo],
        ["Fuseau horaire", user_tz.zone],
        ["Films vus", historique["nb_films"]],
        ["Séries vues", historique["nb_series"]],
        ["Épisodes vus", historique["nb_episodes"]],
        ["Listes personnalisées", len(stats_listes) - 1],
        ["Contenus dans tes listes + suivi", sum(s["total"] for s in stats_listes)],
        ["Contenus déjà vus à nettoyer", len(resultats)],
        ["Doublons entre listes", len(doublons)],
        ["Progressions fantômes", len(playback)],
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
    df_listes["% nettoyage possible"] = (df_listes["deja_vus"] / df_listes["total"].replace(0, 1) * 100).round(1)
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
    colonne_pourcentage = None
    ws_listes = classeur["Analyse par liste"]
    for cellule in ws_listes[1]:
        if cellule.value == "% nettoyage possible":
            colonne_pourcentage = cellule.column
            break
    if colonne_pourcentage:
        lettre = get_column_letter(colonne_pourcentage)
        ws_listes.conditional_formatting.add(
            f"{lettre}2:{lettre}{ws_listes.max_row}",
            ColorScaleRule(
                start_type="min", start_color="63BE7B",
                mid_type="percentile", mid_value=50, mid_color="FFEB84",
                end_type="max", end_color="F8696B",
            ),
        )
    buffer_final = io.BytesIO()
    classeur.save(buffer_final)
    buffer_final.seek(0)
    return buffer_final.getvalue()

# ==================================================
# INTERFACE — ENTÊTE COMMUN
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
        return

    if "infos_utilisateur" not in st.session_state:
        st.session_state["infos_utilisateur"] = obtenir_infos_utilisateur(st.session_state["access_token"])
    infos = st.session_state["infos_utilisateur"]
    pseudo = infos["pseudo"]
    user_tz = infos["timezone"]

    colonne_info, colonne_deco = st.columns([4, 1])
    with colonne_info:
        st.success(f"Connecté en tant que **{pseudo}** • Fuseau horaire : `{infos['tz_name']}`", icon="👤")
    with colonne_deco:
        if st.button("🚪 Se déconnecter", use_container_width=True):
            oublier_connexion()
            st.rerun()
    st.divider()

    # Boutons globaux : présents sur TOUTES les pages
    if "resultats" in st.session_state:
        historique = st.session_state["historique"]
        resultats = st.session_state["resultats"]
        stats_listes = st.session_state["stats_listes"]
        doublons = st.session_state["doublons"]
        playback = st.session_state["playback"]
        excel_bytes = generer_rapport_excel(pseudo, historique, resultats, stats_listes, doublons, playback, user_tz)
        col_relance_rapide, col_relance_totale, col_excel = st.columns(3)
        with col_relance_rapide:
            if st.button("🔄 Relancer l'analyse (rapide)", use_container_width=True, help="Garde l'historique en mémoire, ré-analyse seulement les listes"):
                for cle in ["resultats", "stats_listes", "doublons", "doublons_detail", "playback"]:
                    st.session_state.pop(cle, None)
                lancer_analyse_complete(raffraichir_historique=False)
        with col_relance_totale:
            if st.button("🔃 Tout rafraîchir (historique inclus)", use_container_width=True, help="Récupère à nouveau tout ton historique depuis Trakt"):
                for cle in ["historique", "resultats", "stats_listes", "doublons", "doublons_detail", "playback", "infos_utilisateur"]:
                    st.session_state.pop(cle, None)
                st.rerun()
        with col_excel:
            st.download_button(
                "📥 Télécharger le rapport Excel",
                data=excel_bytes,
                file_name=f"trakt_smart_lists_{pseudo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.divider()
    return user_tz

def afficher_bloc_lancement_analyse():
    """Affiche le bouton de lancement d'analyse, disponible sur TOUTES les pages."""
    if "resultats" in st.session_state:
        return False
    if "historique" in st.session_state:
        st.info("ℹ️ Ton historique est déjà en mémoire — l'analyse sera rapide.")
        texte_bouton = "🔄 Lancer l'analyse rapide"
    else:
        st.info("ℹ️ Lance l'analyse pour charger tes données et accéder à tous les outils. La première analyse peut prendre quelques instants.")
        texte_bouton = "🔍 Lancer l'analyse complète"
    if st.button(texte_bouton, type="primary", use_container_width=True):
        lancer_analyse_complete()
    return True

# ==================================================
# INTERFACE — PAGES
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
        url = st.session_state["verification_url"]
        code = st.session_state["user_code"]
        url_complete = f"{url}/{code}"
        st.write("Connecte-toi à Trakt avec l'une de ces options :")
        colonne_gauche, colonne_droite = st.columns(2)
        with colonne_gauche:
            st.markdown(
                f'<a href="{url_complete}" target="_blank" '
                f'style="display:inline-block; background-color:#008778; color:white; '
                f'padding:0.8em 1.6em; border-radius:10px; text-decoration:none; '
                f'font-weight:600;">Ouvrir la page d\'autorisation</a>',
                unsafe_allow_html=True,
            )
            st.caption("Sur cet appareil, ou un autre.")
            st.info(f"Code à entrer si demandé : **{code}**")
        with colonne_droite:
            st.image(generer_qr_code(url_complete), width=160)
            st.caption("Ou scanne avec ton téléphone.")
        st.caption("Garde cette page ouverte : elle se met à jour automatiquement dès que tu autorises l'accès.")
        with st.spinner("En attente de ton autorisation sur Trakt..."):
            temps_ecoule = 0
            interval = st.session_state["interval"]
            expiration = st.session_state["expires_in"]
            erreur = None
            while temps_ecoule < expiration:
                time.sleep(interval)
                temps_ecoule += interval
                try:
                    tokens = verifier_connexion(st.session_state["device_code"])
                except Exception as e:
                    erreur = str(e)
                    break
                if tokens:
                    sauvegarder_connexion(tokens)
                    del st.session_state["device_code"]
                    st.rerun()
        if erreur:
            st.error(erreur)
        else:
            st.error("Le temps est écoulé.")
        st.session_state.pop("device_code", None)
        if st.button("Réessayer"):
            st.rerun()

def page_tableau_de_bord(user_tz):
    if afficher_bloc_lancement_analyse():
        return
    historique = st.session_state["historique"]
    resultats = st.session_state["resultats"]
    stats_listes = st.session_state["stats_listes"]
    doublons = st.session_state["doublons"]
    playback = st.session_state["playback"]

    st.subheader("📊 Vue d'ensemble")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🎬 Films vus", historique["nb_films"])
    col2.metric("📺 Séries vues", historique["nb_series"])
    col3.metric("🎞️ Épisodes vus", historique["nb_episodes"])
    col4.metric("📋 Listes", len(stats_listes))

    total_items = sum(s["total"] for s in stats_listes)
    total_deja_vus = sum(s["deja_vus"] for s in stats_listes)
    pourcentage_global = round(total_deja_vus / total_items * 100, 1) if total_items else 0
    col5, col6, col7, col8 = st.columns(4)
    col5.metric("Total contenus", total_items)
    col6.metric("🔴 Déjà vus", total_deja_vus)
    col7.metric("🔁 Doublons", len(doublons))
    col8.metric("👻 Fantômes", len(playback), delta=f"{pourcentage_global}% nettoyable")

    st.divider()
    st.subheader("⚠️ Actions rapides")
    col_warn1, col_warn2, col_warn3 = st.columns(3)
    with col_warn1:
        with st.container(border=True):
            st.markdown("#### 👻 Progressions fantômes")
            if len(playback) > 0:
                st.warning(f"{len(playback)} progression(s) non terminée(s)")
                if st.button("Nettoyer les fantômes", key="goto_fantomes"):
                    st.session_state["page_selectionnee"] = "👻 Progression Fantôme"
                    st.rerun()
            else:
                st.success("Aucun fantôme, c'est propre !")
    with col_warn2:
        with st.container(border=True):
            st.markdown("#### 🔁 Doublons")
            if len(doublons) > 0:
                st.warning(f"{len(doublons)} doublon(s) détecté(s)")
                if st.button("Voir les doublons", key="goto_doublons"):
                    st.session_state["page_selectionnee"] = "🔍 Recherche de doublons"
                    st.rerun()
            else:
                st.success("Aucun doublon !")
    with col_warn3:
        with st.container(border=True):
            st.markdown("#### 🧹 Contenus déjà vus")
            if len(resultats) > 0:
                st.warning(f"{len(resultats)} contenu(s) déjà vu(s) dans tes listes")
                if st.button("Nettoyer les listes", key="goto_vus"):
                    st.session_state["page_selectionnee"] = "🧹 Nettoyage des listes"
                    st.rerun()
            else:
                st.success("Tes listes sont à jour !")

    st.divider()
    st.subheader("Taux de contenu à nettoyer par liste")
    df_stats_listes = pd.DataFrame(stats_listes)
    df_stats_listes["% nettoyable"] = (df_stats_listes["deja_vus"] / df_stats_listes["total"].replace(0, 1) * 100).round(1)
    st.bar_chart(df_stats_listes.set_index("nom")["% nettoyable"], color="#CEDC00")

def page_nettoyage_listes(user_tz):
    if afficher_bloc_lancement_analyse():
        return
    resultats = st.session_state["resultats"]
    cle_message = "message_suppression_deja_vus"
    if st.session_state.get(cle_message):
        st.success(st.session_state[cle_message])
        del st.session_state[cle_message]
    st.subheader("🧹 Nettoyage des listes")
    st.caption("Retire automatiquement les films et séries que tu as déjà vus de ta liste de suivi et de tes listes personnalisées.")
    if not resultats:
        st.success("Aucun contenu déjà vu trouvé. Tout est propre ! 🎉")
        return
    st.write(f"**{len(resultats)}** contenu(s) déjà vu(s) trouvé(s). Coche ceux à supprimer :")
    tableau = pd.DataFrame(resultats)
    tableau_affichage = tableau[["type", "titre", "annee", "vues", "dernier_visionnage", "liste"]].copy()
    tableau_affichage["dernier_visionnage"] = pd.to_datetime(tableau_affichage["dernier_visionnage"]).dt.tz_convert(user_tz).dt.strftime("%d/%m/%Y %H:%M")
    tableau_affichage.insert(0, "Sélectionner", False)
    tableau_affichage.columns = ["Sélectionner", "Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"]
    edite = st.data_editor(tableau_affichage, use_container_width=True, hide_index=True, disabled=["Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"], key="editeur_deja_vus")
    nb_selectionnes = int(edite["Sélectionner"].sum())
    if nb_selectionnes == 0:
        return
    cle_confirmation = "confirmation_deja_vus"
    if not st.session_state.get(cle_confirmation, False):
        if st.button(f"🗑️ Supprimer les {nb_selectionnes} élément(s) sélectionné(s)", type="primary", key="bouton_deja_vus"):
            st.session_state[cle_confirmation] = True
            st.rerun()
        return
    st.warning(f"Confirmer la suppression de {nb_selectionnes} élément(s) ? Cette action est irréversible.")
    col_oui, col_non = st.columns(2)
    with col_oui:
        if st.button("✅ Oui, supprimer", key="oui_deja_vus"):
            indices = edite[edite["Sélectionner"]].index
            items_a_supprimer = [resultats[i] for i in indices]
            with st.spinner("Suppression en cours..."):
                supprimer_selection(st.session_state["access_token"], items_a_supprimer)
            st.session_state[cle_confirmation] = False
            st.session_state[cle_message] = f"✅ {len(items_a_supprimer)} élément(s) supprimé(s). Relance l'analyse pour voir les données à jour."
            st.rerun()
    with col_non:
        if st.button("❌ Annuler", key="non_deja_vus"):
            st.session_state[cle_confirmation] = False
            st.rerun()

def page_recherche_doublons(user_tz):
    if afficher_bloc_lancement_analyse():
        return
    doublons_detail = st.session_state["doublons_detail"]
    cle_message = "message_suppression_doublons"
    if st.session_state.get(cle_message):
        st.success(st.session_state[cle_message])
        del st.session_state[cle_message]
    st.subheader("🔍 Recherche de doublons")
    st.caption("Trouve les contenus qui sont présents dans plusieurs listes à la fois.")
    if not doublons_detail:
        st.success("Aucun doublon trouvé entre tes listes.")
        return
    st.write(f"**{len(st.session_state['doublons'])}** contenu(s) présents dans plusieurs listes. Coche les lignes à retirer :")
    tableau = pd.DataFrame(doublons_detail)
    tableau_affichage = tableau[["type", "titre", "annee", "liste"]].copy()
    tableau_affichage.insert(0, "Sélectionner", False)
    tableau_affichage.columns = ["Sélectionner", "Type", "Titre", "Année", "Liste"]
    edite = st.data_editor(tableau_affichage, use_container_width=True, hide_index=True, disabled=["Type", "Titre", "Année", "Liste"], key="editeur_doublons")
    nb_selectionnes = int(edite["Sélectionner"].sum())
    if nb_selectionnes == 0:
        return
    cle_confirmation = "confirmation_doublons"
    if not st.session_state.get(cle_confirmation, False):
        if st.button(f"🗑️ Retirer les {nb_selectionnes} élément(s) sélectionné(s)", type="primary", key="bouton_doublons"):
            st.session_state[cle_confirmation] = True
            st.rerun()
        return
    st.warning(f"Confirmer le retrait de {nb_selectionnes} élément(s) ?")
    col_oui, col_non = st.columns(2)
    with col_oui:
        if st.button("✅ Oui, retirer", key="oui_doublons"):
            indices = edite[edite["Sélectionner"]].index
            items_a_supprimer = [doublons_detail[i] for i in indices]
            with st.spinner("Suppression en cours..."):
                supprimer_selection(st.session_state["access_token"], items_a_supprimer)
            st.session_state[cle_confirmation] = False
            st.session_state[cle_message] = f"✅ {len(items_a_supprimer)} élément(s) retiré(s)."
            st.rerun()
    with col_non:
        if st.button("❌ Annuler", key="non_doublons"):
            st.session_state[cle_confirmation] = False
            st.rerun()

def page_progressions_fantomes(user_tz):
    if afficher_bloc_lancement_analyse():
        return
    playback = st.session_state["playback"]
    cle_message = "message_suppression_fantomes"
    if st.session_state.get(cle_message):
        st.success(st.session_state[cle_message])
        del st.session_state[cle_message]
    st.subheader("👻 Progression Fantôme")
    st.caption("Gère tes vidéos en pause : supprime les entrées obsolètes qui restent bloquées dans ta section 'Continuer à regarder'.")
    st.divider()
    if not playback:
        st.success("Aucune progression en cours trouvée. Tout est propre ! 🎉")
        return
    col_tout_selectionner, _ = st.columns([1, 4])
    with col_tout_selectionner:
        tout_selectionner = st.checkbox("Tout sélectionner", key="select_tout_fantomes")
    selections = {}
    for item in playback:
        progress = item["progression"]
        progress_class = "progress-low" if progress < 30 else "progress-mid" if progress < 80 else "progress-high"
        date_formatee = formater_date(item["dernier_visionnage"], user_tz)
        icone_type = "🎬" if item["type"] == "Film" else "📺"
        with st.container():
            col_check, col_content = st.columns([0.05, 0.95])
            with col_check:
                selections[item["playback_id"]] = st.checkbox("", value=tout_selectionner, key=f"check_{item['playback_id']}", label_visibility="collapsed")
            with col_content:
                st.markdown(f"""
                <div class="ghost-card">
                    <div class="ghost-title">{icone_type} {item['titre']} {f'({item["annee"]})' if item['annee'] else ''}</div>
                    <div class="ghost-meta">{item['type']} • {item['progression']}% visionné • 🕒 {date_formatee}</div>
                    <div class="progress-bar-container">
                        <div class="progress-bar-fill {progress_class}" style="width: {item['progression']}%"></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    st.divider()
    ids_selectionnes = [pid for pid, sel in selections.items() if sel]
    if not ids_selectionnes:
        st.info("Coche les cartes à supprimer.")
        return
    cle_confirmation = "confirmation_fantomes"
    if not st.session_state.get(cle_confirmation, False):
        if st.button(f"🗑️ Supprimer les {len(ids_selectionnes)} progression(s) sélectionnée(s)", type="primary", key="bouton_fantomes"):
            st.session_state[cle_confirmation] = True
            st.rerun()
        return
    st.warning(f"Confirmer la suppression de {len(ids_selectionnes)} progression(s) ? Cette action est irréversible.")
    col_oui, col_non = st.columns(2)
    with col_oui:
        if st.button("✅ Oui, supprimer", key="oui_fantomes"):
            items_a_supprimer = [p for p in playback if p["playback_id"] in ids_selectionnes]
            with st.spinner("Suppression en cours..."):
                supprimer_progressions(st.session_state["access_token"], items_a_supprimer)
            st.session_state[cle_confirmation] = False
            st.session_state[cle_message] = f"✅ {len(items_a_supprimer)} progression(s) fantôme(s) supprimée(s). Relance l'analyse pour voir les changements."
            st.rerun()
    with col_non:
        if st.button("❌ Annuler", key="non_fantomes"):
            st.session_state[cle_confirmation] = False
            st.rerun()

def page_tableau_de_bord_statistiques(user_tz):
    if afficher_bloc_lancement_analyse():
        return
    st.subheader("📊 Tableau de bord statistiques")
    st.caption("Toutes tes statistiques de visionnage détaillées, avec graphiques interactifs.")

    historique = st.session_state["historique"]
    films = pd.DataFrame(historique["films_details"])
    episodes = pd.DataFrame(historique["episodes_details"])

    # Filtres
    col1, col2, col3 = st.columns(3)
    with col1:
        type_contenu = st.selectbox("Type de contenu", ["Tous", "Films", "Séries"], index=0)
    with col2:
        periode = st.selectbox("Période", ["Tout l'historique", "Cette année", "Les 12 derniers mois", "Ce mois-ci"], index=0)
    with col3:
        if type_contenu in ["Tous", "Séries"] and not episodes.empty:
            tous_genres = sorted(list(set(g for liste in episodes["genre"].str.split(", ") for g in liste if g != "Inconnu")))
        else:
            tous_genres = sorted(list(set(g for liste in films["genre"].str.split(", ") for g in liste if g != "Inconnu"))) if not films.empty else []
        genre = st.selectbox("Genre", ["Tous"] + tous_genres, index=0)

    # Préparation des données
    tous_visionnages = []
    if type_contenu in ["Tous", "Films"] and not films.empty:
        df_f = films.copy()
        df_f["type"] = "Film"
        df_f["date"] = pd.to_datetime(df_f["date_visionnage"], utc=True).dt.tz_convert(user_tz)
        tous_visionnages.append(df_f)
    if type_contenu in ["Tous", "Séries"] and not episodes.empty:
        df_e = episodes.copy()
        df_e["type"] = "Épisode"
        df_e["date"] = pd.to_datetime(df_e["date_visionnage"], utc=True).dt.tz_convert(user_tz)
        df_e["duree"] = df_e["duree"].fillna(40)  # Durée moyenne épisode si inconnue
        tous_visionnages.append(df_e)

    if not tous_visionnages:
        st.info("Aucune donnée pour les filtres sélectionnés.")
        return

    df = pd.concat(tous_visionnages, ignore_index=True)

    # Filtre période
    maintenant = datetime.now(user_tz)
    if periode == "Cette année":
        df = df[df["date"].dt.year == maintenant.year]
    elif periode == "Les 12 derniers mois":
        df = df[df["date"] >= maintenant - pd.DateOffset(months=12)]
    elif periode == "Ce mois-ci":
        df = df[(df["date"].dt.year == maintenant.year) & (df["date"].dt.month == maintenant.month)]

    # Filtre genre
    if genre != "Tous":
        df = df[df["genre"].str.contains(genre, na=False)]

    if df.empty:
        st.warning("Aucun visionnage ne correspond à tes filtres.")
        return

    # Calcul des heures de visionnage
    df["duree_heures"] = df["duree"] / 60
    total_heures = df["duree_heures"].sum().round(1)

    # Metrics
    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total visionnages", len(df))
    m2.metric("Heures de visionnage", f"{total_heures}h")
    if "note" in df.columns:
        note_moyenne = df[df["note"] > 0]["note"].mean().round(1)
        m3.metric("Note moyenne", f"{note_moyenne}/10" if note_moyenne == note_moyenne else "-")
    if "annee" in df.columns:
        annee_moyenne = round(df["annee"].mean())
        m4.metric("Année moyenne des contenus", annee_moyenne if annee_moyenne == annee_moyenne else "-")

    st.divider()

    # Graphique 1 : Heures de visionnage par mois
    df["mois"] = df["date"].dt.strftime("%Y-%m")
    heures_par_mois = df.groupby("mois")["duree_heures"].sum().round(1)
    options_mois = {
        "title": {"text": "Heures de visionnage par mois", "textStyle": {"color": "#F0F7F6"}},
        "tooltip": {"trigger": "axis"},
        "backgroundColor": "transparent",
        "textStyle": {"color": "#F0F7F6"},
        "xAxis": {"type": "category", "data": list(heures_par_mois.index), "axisLabel": {"color": "#8FA8A4"}},
        "yAxis": {"type": "value", "axisLabel": {"color": "#8FA8A4"}, "splitLine": {"lineStyle": {"color": "#1A443F"}}},
        "series": [{"data": list(heures_par_mois.values), "type": "line", "smooth": True, "lineStyle": {"color": "#CEDC00", "width": 3}, "areaStyle": {"color": "rgba(206, 220, 0, 0.1)"}, "itemStyle": {"color": "#CEDC00"}}],
    }
    st_echarts(options=options_mois, height="400px")

    col_g1, col_g2 = st.columns(2)

    # Graphique 2 : Répartition par genres (camembert)
    with col_g1:
        genres = {}
        for liste_genres in df["genre"].str.split(", "):
            for g in liste_genres:
                if g and g != "Inconnu":
                    genres[g] = genres.get(g, 0) + 1
        options_genres = {
            "title": {"text": "Répartition par genre", "left": "center", "textStyle": {"color": "#F0F7F6"}},
            "tooltip": {"trigger": "item"},
            "backgroundColor": "transparent",
            "textStyle": {"color": "#F0F7F6"},
            "legend": {"bottom": 0, "textStyle": {"color": "#8FA8A4"}},
            "series": [{"type": "pie", "radius": ["40%", "70%"], "data": [{"name": k, "value": v} for k, v in sorted(genres.items(), key=lambda x: -x[1])[:8]], "itemStyle": {"borderRadius": 8, "borderColor": "#071816", "borderWidth": 2}, "label": {"color": "#F0F7F6"}}],
            "color": ["#008778", "#CEDC00", "#00B3A0", "#889900", "#00574F", "#9DB000", "#1A443F", "#E8F064"]
        }
        st_echarts(options=options_genres, height="400px")

    # Graphique 3 : Heures de visionnage par heure de la journée
    with col_g2:
        df["heure"] = df["date"].dt.hour
        heures = df.groupby("heure")["duree_heures"].sum()
        # Remplit toutes les heures de 0 à 23
        heures = heures.reindex(range(24), fill_value=0)
        options_heures = {
            "title": {"text": "Heures de visionnage dans la journée", "left": "center", "textStyle": {"color": "#F0F7F6"}},
            "tooltip": {"trigger": "axis"},
            "backgroundColor": "transparent",
            "textStyle": {"color": "#F0F7F6"},
            "xAxis": {"type": "category", "data": [f"{h}h" for h in range(24)], "axisLabel": {"color": "#8FA8A4"}},
            "yAxis": {"type": "value", "axisLabel": {"color": "#8FA8A4"}, "splitLine": {"lineStyle": {"color": "#1A443F"}}},
            "series": [{"data": list(heures.values.round(1)), "type": "bar", "itemStyle": {"color": {"type": "linear", "x": 0, "y": 0, "x2": 0, "y2": 1, "colorStops": [{"offset": 0, "color": "#008778"}, {"offset": 1, "color": "#00574F"}]}, "borderRadius": [4,4,0,0]}}],
        }
        st_echarts(options=options_heures, height="400px")

def page_rendez_vous_annuel():
    st.subheader("🎬 Rendez-vous annuel (Wrapped)")
    st.info("🚧 Bientôt disponible : ton récapitulatif annuel façon Spotify Wrapped pour tes visionnages !")

def page_sauvegarde():
    st.subheader("📤 Sauvegarde et restauration")
    st.info("🚧 Bientôt disponible : export et restauration complète de tes données Trakt en JSON/Excel.")

def page_succes():
    st.subheader("🏆 Succès et badges")
    st.info("🚧 Bientôt disponible : badges et objectifs de visionnage personnalisés.")

# ==================================================
# RECONNEXION AUTOMATIQUE
# ==================================================

if "access_token" not in st.session_state:
    refresh_token_sauvegarde = cookies.get("trakt_refresh_token")
    if refresh_token_sauvegarde:
        tokens = rafraichir_token(refresh_token_sauvegarde)
        if tokens:
            sauvegarder_connexion(tokens)
        else:
            cookies.remove("trakt_refresh_token")

# ==================================================
# STRUCTURE PRINCIPALE
# ==================================================

user_tz = afficher_entete()

if "access_token" not in st.session_state:
    page_connexion()
else:
    # Menu de navigation dans la sidebar
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
        st.markdown('<p class="section-menu-title">Navigation</p>', unsafe_allow_html=True)
        page_defaut = PAGES.index(st.session_state.get("page_selectionnee", PAGES[0]))
        page = st.radio(
            "Choisis un outil",
            PAGES,
            index=page_defaut,
            label_visibility="collapsed",
            key="menu_navigation",
        )
        st.session_state["page_selectionnee"] = page

    # Routage
    if page == "🏠 Tableau de bord":
        page_tableau_de_bord(user_tz)
    elif page == "👻 Progression Fantôme":
        page_progressions_fantomes(user_tz)
    elif page == "🧹 Nettoyage des listes":
        page_nettoyage_listes(user_tz)
    elif page == "🔍 Recherche de doublons":
        page_recherche_doublons(user_tz)
    elif page == "📊 Statistiques détaillées":
        page_tableau_de_bord_statistiques(user_tz)
    elif page == "🎬 Rendez-vous annuel":
        page_rendez_vous_annuel()
    elif page == "📤 Sauvegarde / Restauration":
        page_sauvegarde()
    elif page == "🏆 Succès":
        page_succes()
