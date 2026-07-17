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

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🏎️", layout="wide")

# Injection de styles CSS custom pour le thème Aston Martin F1
st.markdown("""
<style>
    /* Couleurs globales thème Aston Martin F1 */
    :root {
        --am-green: #00665F;
        --am-lime: #CEDC00;
        --am-black: #0A0F0D;
        --am-dark: #151B19;
        --am-light-gray: #2A3330;
    }

    /* Style des cartes fantômes */
    .ghost-card {
        background-color: var(--am-dark);
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 14px;
        border-left: 4px solid var(--am-lime);
        transition: all 0.2s ease;
    }
    .ghost-card:hover {
        border-left: 4px solid var(--am-green);
        transform: translateX(2px);
    }
    .ghost-title {
        font-size: 1.05em;
        font-weight: 600;
        color: white;
        margin-bottom: 6px;
    }
    .ghost-meta {
        font-size: 0.9em;
        color: #b0b8b6;
        margin-bottom: 10px;
    }
    .progress-bar-container {
        width: 100%;
        height: 8px;
        background-color: var(--am-light-gray);
        border-radius: 4px;
        overflow: hidden;
    }
    .progress-bar-fill {
        height: 100%;
        border-radius: 4px;
        transition: width 0.5s ease;
    }
    .progress-low { background-color: #ED2224; }
    .progress-mid { background-color: var(--am-lime); }
    .progress-high { background-color: var(--am-green); }

    /* Style des boutons */
    .stButton > button {
        border-radius: 8px;
        border: 0;
        font-weight: 500;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(0, 102, 95, 0.3);
    }
    div[data-testid="stSidebarNav"] {
        padding-top: 20px;
    }
    .section-menu-title {
        font-size: 0.8em;
        font-weight: 700;
        color: var(--am-lime);
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 18px 0 8px 0;
    }
</style>
""", unsafe_allow_html=True)

cookies = CookieController()

# ==================================================
# CONFIGURATION (lue depuis les secrets Streamlit)
# ==================================================

CLIENT_ID = st.secrets["TRAKT_CLIENT_ID"]
CLIENT_SECRET = st.secrets["TRAKT_CLIENT_SECRET"]

DEVICE_CODE_URL = "https://api.trakt.tv/oauth/device/code"
DEVICE_TOKEN_URL = "https://api.trakt.tv/oauth/device/token"
REFRESH_TOKEN_URL = "https://api.trakt.tv/oauth/token"
TZ_PARIS = pytz.timezone("Europe/Paris")

def format_date_trakt(date_str):
    """Formate une date Trakt (UTC) en heure locale Paris, avec heure exacte comme sur TPPM."""
    if not date_str:
        return ""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        dt_local = dt.astimezone(TZ_PARIS)
        return dt_local.strftime("%Y-%m-%d %H:%M:%S") + " (+02:00)"
    except Exception:
        return date_str

# ==================================================
# FONCTIONS TRAKT — CONNEXION
# ==================================================

def demarrer_connexion():
    """Demande à Trakt un code d'activation à usage unique."""
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

def obtenir_pseudo_trakt(access_token):
    response = requests.get("https://api.trakt.tv/users/settings", headers=entetes_trakt(access_token))
    response.raise_for_status()
    return response.json()["user"]["username"]

def generer_qr_code(url):
    image = qrcode.make(url)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()

# ==================================================
# FONCTIONS TRAKT — DONNÉES
# ==================================================

def obtenir_historique(access_token, barre=None):
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
                    "duree": film.get("runtime", 0),
                    "note": film.get("rating", 0),
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
                    "duree": episode.get("runtime", serie.get("runtime", 0)),
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

def obtenir_listes(access_token):
    reponse = requests.get("https://api.trakt.tv/users/me/lists", headers=entetes_trakt(access_token))
    reponse.raise_for_status()
    return reponse.json()

def obtenir_contenu_liste(access_token, list_id):
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

def obtenir_watchlist(access_token):
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

def analyser_tout(access_token, historique, barre=None):
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
        barre.progress(0.6, text="Analyse de la watchlist...")
    watchlist = obtenir_watchlist(access_token)
    for item in watchlist:
        enregistrer_apparition(item, "Watchlist", "watchlist")
    matches = comparer_items_avec_historique(watchlist, historique)
    for m in matches:
        m["liste"] = "Watchlist"
        m["liste_id"] = "watchlist"
    resultats.extend(matches)
    nb_films, nb_series = compter_types(watchlist)
    stats_listes.append({
        "nom": "Watchlist (officielle)",
        "nb_films": nb_films,
        "nb_series": nb_series,
        "total": len(watchlist),
        "deja_vus": len(matches),
    })
    listes = obtenir_listes(access_token)
    for i, liste in enumerate(listes):
        if barre:
            barre.progress(0.6 + (i + 1) / max(len(listes), 1) * 0.3, text=f"Analyse de la liste : {liste['name']}")
        items = obtenir_contenu_liste(access_token, liste["ids"]["trakt"])
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

def obtenir_playback(access_token, barre=None):
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

def supprimer_playback(access_token, items_selectionnes):
    for item in items_selectionnes:
        reponse = requests.delete(
            f"https://api.trakt.tv/sync/playback/{item['playback_id']}",
            headers=entetes_trakt(access_token),
        )
        reponse.raise_for_status()
        time.sleep(0.5)

# ==================================================
# FONCTIONS — EXPORT EXCEL
# ==================================================

def auto_ajuster_colonnes(ws):
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

def mettre_en_forme_feuille(ws, header_color="00665F"):
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
        cellule.fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        cellule.alignment = Alignment(horizontal="center")
    auto_ajuster_colonnes(ws)

def generer_excel(pseudo, historique, resultats, stats_listes, doublons, playback):
    df_resume = pd.DataFrame([
        ["Compte Trakt", pseudo],
        ["Films vus", historique["nb_films"]],
        ["Séries vues", historique["nb_series"]],
        ["Épisodes vus", historique["nb_episodes"]],
        ["Listes personnalisées", len(stats_listes) - 1],
        ["Contenus dans tes listes + watchlist", sum(s["total"] for s in stats_listes)],
        ["Contenus déjà vus à nettoyer", len(resultats)],
        ["Doublons entre listes", len(doublons)],
        ["Progressions fantômes (Continue Watching)", len(playback)],
    ], columns=["Statistique", "Valeur"])
    df_resultats = pd.DataFrame(resultats)
    if not df_resultats.empty:
        df_resultats = df_resultats[["liste", "type", "titre", "annee", "vues", "dernier_visionnage", "tmdb_id"]].copy()
        df_resultats["dernier_visionnage"] = pd.to_datetime(df_resultats["dernier_visionnage"]).dt.strftime("%d/%m/%Y %H:%M")
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
    df_listes.columns = ["Liste", "Film", "Série", "Nombre de contenus", "Déjà vus", "% nettoyage possible"]
    # Feuille fantômes
    df_fantomes = pd.DataFrame(playback)
    if not df_fantomes.empty:
        df_fantomes = df_fantomes[["type", "titre", "annee", "progression", "dernier_visionnage"]].copy()
        df_fantomes["dernier_visionnage"] = pd.to_datetime(df_fantomes["dernier_visionnage"]).dt.strftime("%d/%m/%Y %H:%M")
        df_fantomes.columns = ["Type", "Titre", "Année", "Progression (%)", "Dernier visionnage"]
    else:
        df_fantomes = pd.DataFrame(columns=["Type", "Titre", "Année", "Progression (%)", "Dernier visionnage"])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_resume.to_excel(writer, sheet_name="Résumé", index=False)
        df_resultats.to_excel(writer, sheet_name="À nettoyer", index=False)
        df_doublons.to_excel(writer, sheet_name="Doublons entre listes", index=False)
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
# INTERFACE — ENTÊTE COMMUN À TOUTES LES PAGES
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
    pseudo = obtenir_pseudo_trakt(st.session_state["access_token"])
    colonne_info, colonne_deco = st.columns([4, 1])
    with colonne_info:
        st.success(f"Connecté à Trakt en tant que **{pseudo}** ✅", icon="👤")
    with colonne_deco:
        if st.button("🚪 Se déconnecter", use_container_width=True):
            oublier_connexion()
            st.rerun()
    st.divider()
    # Boutons d'action globaux, présents sur toutes les pages
    if "resultats" in st.session_state:
        historique = st.session_state["historique"]
        resultats = st.session_state["resultats"]
        stats_listes = st.session_state["stats_listes"]
        doublons = st.session_state["doublons"]
        playback = st.session_state["playback"]
        excel_bytes = generer_excel(pseudo, historique, resultats, stats_listes, doublons, playback)
        col_relance_rapide, col_relance_totale, col_excel = st.columns(3)
        with col_relance_rapide:
            if st.button("🔄 Relancer l'analyse (rapide)", use_container_width=True):
                for cle in ["resultats", "stats_listes", "doublons", "doublons_detail", "playback"]:
                    st.session_state.pop(cle, None)
                st.rerun()
        with col_relance_totale:
            if st.button("🔃 Tout rafraîchir (historique inclus)", use_container_width=True):
                for cle in ["historique", "resultats", "stats_listes", "doublons", "doublons_detail", "playback"]:
                    st.session_state.pop(cle, None)
                st.rerun()
        with col_excel:
            st.download_button(
                "📥 Télécharger le rapport Excel complet",
                data=excel_bytes,
                file_name=f"trakt_smart_lists_{pseudo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.divider()

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
                f'style="display:inline-block; background-color:#00665F; color:white; '
                f'padding:0.8em 1.6em; border-radius:8px; text-decoration:none; '
                f'font-weight:600;">Ouvrir la page d\'autorisation</a>',
                unsafe_allow_html=True,
            )
            st.caption("Sur cet appareil, ou un autre.")
            st.info(f"Code à entrer si demandé : **{code}**")
        with colonne_droite:
            st.image(generer_qr_code(url_complete), width=160)
            st.caption("Ou scanne avec ton téléphone.")
        st.caption("Garde cette page ouverte : elle se met à jour toute seule dès que c'est approuvé.")
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

def page_tableau_de_bord():
    if "resultats" not in st.session_state:
        if "historique" in st.session_state:
            st.info("ℹ️ Ton historique est déjà en mémoire pour cette session — l'analyse sera rapide.")
        else:
            st.info("ℹ️ Lance ta première analyse pour charger ton historique et toutes tes données. La première fois peut prendre quelques instants.")
        if st.button("🔍 Lancer l'analyse complète", type="primary", use_container_width=True):
            barre = st.progress(0, text="Démarrage...")
            if "historique" not in st.session_state:
                st.session_state["historique"] = obtenir_historique(st.session_state["access_token"], barre)
            resultats, stats_listes, doublons, doublons_detail = analyser_tout(st.session_state["access_token"], st.session_state["historique"], barre)
            playback = obtenir_playback(st.session_state["access_token"], barre)
            st.session_state["resultats"] = resultats
            st.session_state["stats_listes"] = stats_listes
            st.session_state["doublons"] = doublons
            st.session_state["doublons_detail"] = doublons_detail
            st.session_state["playback"] = playback
            barre.empty()
            st.rerun()
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
    col6.metric("🔴 Déjà vus à nettoyer", total_deja_vus)
    col7.metric("🔁 Doublons", len(doublons))
    col8.metric("👻 Progressions fantômes", len(playback), delta=f"{pourcentage_global}% nettoyable")
    st.divider()
    st.subheader("⚠️ Actions rapides")
    col_warn1, col_warn2, col_warn3 = st.columns(3)
    with col_warn1:
        with st.container(border=True):
            st.markdown("#### 👻 Fantômes")
            if len(playback) > 0:
                st.warning(f"{len(playback)} progression(s) non terminée(s)")
                if st.button("Aller au nettoyage", key="goto_fantomes"):
                    st.session_state["page"] = "Progression Fantôme"
                    st.rerun()
            else:
                st.success("Aucun fantôme, c'est propre !")
    with col_warn2:
        with st.container(border=True):
            st.markdown("#### 🔁 Doublons")
            if len(doublons) > 0:
                st.warning(f"{len(doublons)} doublon(s) détecté(s)")
                if st.button("Aller aux doublons", key="goto_doublons"):
                    st.session_state["page"] = "Duplicate Finder"
                    st.rerun()
            else:
                st.success("Aucun doublon !")
    with col_warn3:
        with st.container(border=True):
            st.markdown("#### 🧹 Contenus vus")
            if len(resultats) > 0:
                st.warning(f"{len(resultats)} contenu(s) déjà vu(s) dans tes listes")
                if st.button("Aller au nettoyage", key="goto_vus"):
                    st.session_state["page"] = "List Cleaner"
                    st.rerun()
            else:
                st.success("Tes listes sont à jour !")
    st.divider()
    st.subheader("% de contenu nettoyable par liste")
    df_stats_listes = pd.DataFrame(stats_listes)
    df_stats_listes["% nettoyable"] = (df_stats_listes["deja_vus"] / df_stats_listes["total"].replace(0, 1) * 100).round(1)
    st.bar_chart(df_stats_listes.set_index("nom")["% nettoyable"], color="#CEDC00")

def page_list_cleaner():
    if "resultats" not in st.session_state:
        st.info("Lance d'abord une analyse depuis le Tableau de bord.")
        return
    resultats = st.session_state["resultats"]
    cle_message = "message_suppression_deja_vus"
    if st.session_state.get(cle_message):
        st.success(st.session_state[cle_message])
        del st.session_state[cle_message]
    st.subheader("🧹 List Cleaner : Contenus déjà vus")
    st.caption("Retire automatiquement les films et séries que tu as déjà vus de ta watchlist et tes listes.")
    if not resultats:
        st.success("Aucun contenu déjà vu trouvé. Tout est propre ! 🎉")
        return
    st.write(f"**{len(resultats)}** contenu(s) déjà vu(s) trouvé(s). Coche ceux à supprimer :")
    tableau = pd.DataFrame(resultats)
    tableau_affichage = tableau[["type", "titre", "annee", "vues", "dernier_visionnage", "liste"]].copy()
    tableau_affichage["dernier_visionnage"] = pd.to_datetime(tableau_affichage["dernier_visionnage"]).dt.strftime("%d/%m/%Y")
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

def page_duplicate_finder():
    if "resultats" not in st.session_state:
        st.info("Lance d'abord une analyse depuis le Tableau de bord.")
        return
    doublons_detail = st.session_state["doublons_detail"]
    cle_message = "message_suppression_doublons"
    if st.session_state.get(cle_message):
        st.success(st.session_state[cle_message])
        del st.session_state[cle_message]
    st.subheader("🔍 Duplicate Finder : Doublons entre listes")
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

def page_fantomes():
    if "resultats" not in st.session_state:
        st.info("Lance d'abord une analyse depuis le Tableau de bord.")
        return
    playback = st.session_state["playback"]
    cle_message = "message_suppression_fantomes"
    if st.session_state.get(cle_message):
        st.success(st.session_state[cle_message])
        del st.session_state[cle_message]
    st.subheader("👻 Progression Fantôme (Ghost Progress)")
    st.caption("Gère tes vidéos en pause : supprime les entrées obsolètes qui restent bloquées dans ta section 'Continuer à regarder'.")
    st.divider()
    if not playback:
        st.success("Aucune progression en cours trouvée. Tout est propre ! 🎉")
        return
    # Bouton suppression globale
    col_tout_selectionner, col_tout_supprimer = st.columns([1, 4])
    with col_tout_selectionner:
        tout_selectionner = st.checkbox("Tout sélectionner", key="select_tout_fantomes")
    selections = {}
    # Affichage en cartes comme TPPM
    for item in playback:
        progress = item["progression"]
        progress_class = "progress-low" if progress < 30 else "progress-mid" if progress < 80 else "progress-high"
        date_formatee = format_date_trakt(item["dernier_visionnage"])
        icone_type = "🎬" if item["type"] == "Film" else "📺"
        with st.container():
            col_check, col_content, col_btn = st.columns([0.05, 0.8, 0.15])
            with col_check:
                selections[item["playback_id"]] = st.checkbox("", value=tout_selectionner, key=f"check_{item['playback_id']}", label_visibility="collapsed")
            with col_content:
                st.markdown(f"""
                <div class="ghost-card">
                    <div class="ghost-title">{icone_type} {item['titre']} {f'({item["annee"]})' if item['annee'] else ''}</div>
                    <div class="ghost-meta">Saison/Épisode : {item['type']} • {item['progression']}% • 🕒 {date_formatee}</div>
                    <div class="progress-bar-container">
                        <div class="progress-bar-fill {progress_class}" style="width: {item['progression']}%"></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    # Séparation et bouton de suppression
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
                supprimer_playback(st.session_state["access_token"], items_a_supprimer)
            st.session_state[cle_confirmation] = False
            st.session_state[cle_message] = f"✅ {len(items_a_supprimer)} progression(s) fantôme(s) supprimée(s). Relance l'analyse pour voir les changements."
            st.rerun()
    with col_non:
        if st.button("❌ Annuler", key="non_fantomes"):
            st.session_state[cle_confirmation] = False
            st.rerun()

def page_stats_dashboard():
    if "resultats" not in st.session_state:
        st.info("Lance d'abord une analyse depuis le Tableau de bord pour accéder aux statistiques.")
        return
    st.subheader("📊 Stats Dashboard (bientôt disponible)")
    st.info("🚧 Ce dashboard est en cours de développement : il inclura les graphiques interactifs que tu as vu sur la démo Stockpeers, avec les statistiques d'heures de visionnage, genres, années, heures de la journée, et tous les filtres que tu souhaites !")
    st.write("Prochainement disponible :")
    st.write("✅ Nombre d'heures par jour/semaine/mois/année")
    st.write("✅ Répartition par genre, année, note")
    st.write("✅ Graphiques interactifs filtrables avec ECharts")
    st.write("✅ Segmentation Films / Séries")
    st.write("✅ Statistiques par heure de la journée")

def page_wrapped():
    st.subheader("🎬 Ton Wrapped 2026")
    st.info("🚧 Bientôt disponible : ton récapitulatif annuel façon Spotify Wrapped pour tes visionnages 2026 !")

def page_backup():
    st.subheader("📤 Backup / Restore")
    st.info("🚧 Bientôt disponible : export et restauration complète de tes données Trakt en JSON/Excel.")

def page_achievements():
    st.subheader("🏆 Achievements")
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
# STRUCTURE PRINCIPALE AVEC NAVIGATION
# ==================================================

# Entête commun toujours affiché en haut
afficher_entete()

if "access_token" not in st.session_state:
    page_connexion()
else:
    # Menu de navigation dans la sidebar
    with st.sidebar:
        st.markdown('<p class="section-menu-title">Navigation</p>', unsafe_allow_html=True)
        page = st.radio(
            "Sélectionne un outil",
            [
                "🏠 Tableau de bord",
                "👻 Progression Fantôme",
                "🧹 List Cleaner",
                "🔍 Duplicate Finder",
                "📊 Stats Dashboard",
                "🎬 Wrapped",
                "📤 Backup / Restore",
                "🏆 Achievements",
            ],
            label_visibility="collapsed",
            key="page_radio",
        )
        st.session_state["page"] = page
    # Routage des pages
    if page == "🏠 Tableau de bord":
        page_tableau_de_bord()
    elif page == "👻 Progression Fantôme":
        page_fantomes()
    elif page == "🧹 List Cleaner":
        page_list_cleaner()
    elif page == "🔍 Duplicate Finder":
        page_duplicate_finder()
    elif page == "📊 Stats Dashboard":
        page_stats_dashboard()
    elif page == "🎬 Wrapped":
        page_wrapped()
    elif page == "📤 Backup / Restore":
        page_backup()
    elif page == "🏆 Achievements":
        page_achievements()
