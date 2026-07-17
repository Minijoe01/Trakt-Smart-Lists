import streamlit as st
import requests
import time
import qrcode
import io
import pandas as pd
from streamlit_cookies_controller import CookieController
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🎬", layout="wide")

cookies = CookieController()

# ==================================================
# CONFIGURATION (lue depuis les secrets Streamlit)
# ==================================================

CLIENT_ID = st.secrets["TRAKT_CLIENT_ID"]
CLIENT_SECRET = st.secrets["TRAKT_CLIENT_SECRET"]

DEVICE_CODE_URL = "https://api.trakt.tv/oauth/device/code"
DEVICE_TOKEN_URL = "https://api.trakt.tv/oauth/device/token"
REFRESH_TOKEN_URL = "https://api.trakt.tv/oauth/token"


# ==================================================
# FONCTIONS TRAKT — CONNEXION
# ==================================================

def demarrer_connexion():
    """Demande à Trakt un code d'activation à usage unique."""

    response = requests.post(DEVICE_CODE_URL, json={"client_id": CLIENT_ID})
    response.raise_for_status()

    return response.json()


def verifier_connexion(device_code):
    """Vérifie une fois si le code a été approuvé.
    Retourne les tokens si oui, None si toujours en attente."""

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

    return None  # toujours en attente


def rafraichir_token(refresh_token):
    """Utilise un refresh_token sauvegardé pour obtenir un nouvel access_token,
    sans repasser par tout le processus de connexion. Retourne None si ça échoue."""

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
    """Enregistre les tokens en mémoire ET dans un cookie,
    pour que la connexion survive à un rechargement de page."""
    st.session_state["access_token"] = tokens["access_token"]
    st.session_state["refresh_token"] = tokens["refresh_token"]
    cookies.set("trakt_refresh_token", tokens["refresh_token"])
    time.sleep(0.5)


def oublier_connexion():
    """Supprime le cookie et la session, proprement (avec le même délai que pour l'écriture)."""
    cookies.remove("trakt_refresh_token")
    time.sleep(0.5)
    st.session_state.clear()


def entetes_trakt(access_token):
    """Fabrique les en-têtes nécessaires pour parler à l'API Trakt."""

    return {
        "Content-Type": "application/json",
        "trakt-api-version": "2",
        "trakt-api-key": CLIENT_ID,
        "Authorization": f"Bearer {access_token}",
    }


def obtenir_pseudo_trakt(access_token):
    """Récupère le pseudo de l'utilisateur connecté, pour confirmer que ça fonctionne."""

    response = requests.get("https://api.trakt.tv/users/settings", headers=entetes_trakt(access_token))
    response.raise_for_status()

    return response.json()["user"]["username"]


def generer_qr_code(url):
    """Génère un QR code en mémoire (aucun fichier écrit sur disque)."""

    image = qrcode.make(url)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")

    return buffer.getvalue()


# ==================================================
# FONCTIONS TRAKT — HISTORIQUE, LISTES, WATCHLIST
# ==================================================

def obtenir_historique(access_token, barre=None):
    """Récupère tout l'historique de visionnage, avec quelques statistiques au passage."""

    headers = entetes_trakt(access_token)
    films = {}
    series = {}
    nb_visionnages_films = 0
    nb_episodes = 0

    premiere_page = requests.get(
        "https://api.trakt.tv/users/me/history",
        headers=headers,
        params={"page": 1, "limit": 100},
    )
    premiere_page.raise_for_status()
    total_pages = int(premiere_page.headers.get("X-Pagination-Page-Count", 1))

    for page in range(1, total_pages + 1):

        if barre:
            barre.progress(page / total_pages * 0.7, text=f"Récupération de l'historique : page {page}/{total_pages}")

        reponse = requests.get(
            "https://api.trakt.tv/users/me/history",
            headers=headers,
            params={"page": page, "limit": 100},
        )
        reponse.raise_for_status()

        for item in reponse.json():

            if item["type"] == "movie":
                nb_visionnages_films += 1
                film = item["movie"]
                identifiant = film["ids"]["trakt"]

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
                identifiant = serie["ids"]["trakt"]

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
        "nb_films": len(films),
        "nb_series": len(series),
        "nb_visionnages_films": nb_visionnages_films,
        "nb_episodes": nb_episodes,
    }


def obtenir_listes(access_token):
    """Récupère les listes personnalisées de l'utilisateur (hors watchlist officielle)."""

    reponse = requests.get("https://api.trakt.tv/users/me/lists", headers=entetes_trakt(access_token))
    reponse.raise_for_status()

    return reponse.json()


def obtenir_contenu_liste(access_token, list_id):
    """Récupère tous les éléments d'une liste personnalisée (avec pagination)."""

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
    """Récupère la watchlist officielle (avec pagination)."""

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
    """Compte combien d'items Trakt (bruts) sont des films vs des séries."""

    nb_films = sum(1 for i in items if i["type"] == "movie")
    nb_series = sum(1 for i in items if i["type"] == "show")

    return nb_films, nb_series


def comparer_items_avec_historique(items, historique):
    """Compare une liste d'items (peu importe leur source) avec l'historique de visionnage."""

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
    """Compare la watchlist + toutes les listes personnalisées avec l'historique.
    Retourne : contenus déjà vus, stats par liste, doublons (résumé), doublons (détail par liste)."""

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
        cle = (type_affiche, trakt_id)  # type + id : évite qu'un film et une série au même numéro se mélangent

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
        barre.progress(0.7, text="Analyse de la watchlist...")

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
            barre.progress(0.7 + (i + 1) / max(len(listes), 1) * 0.3, text=f"Analyse de la liste : {liste['name']}")

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


# ==================================================
# FONCTIONS TRAKT — SUPPRESSION
# ==================================================

def supprimer_de_liste(access_token, liste_id, items_a_supprimer):
    """Supprime les films/séries sélectionnés d'une liste précise (ou de la watchlist)."""

    corps = {"movies": [], "shows": []}

    for item in items_a_supprimer:
        cible = corps["movies"] if item["type"] == "Film" else corps["shows"]
        cible.append({"ids": {"trakt": item["trakt_id"]}})

    if liste_id == "watchlist":
        url = "https://api.trakt.tv/sync/watchlist/remove"
    else:
        url = f"https://api.trakt.tv/users/me/lists/{liste_id}/items/remove"

    reponse = requests.post(url, headers=entetes_trakt(access_token), json=corps)
    reponse.raise_for_status()


def supprimer_selection(access_token, items_selectionnes):
    """Supprime les éléments sélectionnés, liste par liste (un appel par liste concernée)."""

    par_liste = {}
    for item in items_selectionnes:
        par_liste.setdefault(item["liste_id"], []).append(item)

    for liste_id, items in par_liste.items():
        supprimer_de_liste(access_token, liste_id, items)
        time.sleep(1)


# ==================================================
# FONCTIONS — EXPORT EXCEL
# ==================================================

def auto_ajuster_colonnes(ws):
    """Ajuste automatiquement la largeur des colonnes selon leur contenu."""

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


def mettre_en_forme_feuille(ws):
    """Mise en forme générale : en-têtes stylées, figées, tableau filtrable, colonnes ajustées."""

    ws.freeze_panes = "A2"

    if ws.max_row > 1:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        table = Table(displayName=f"Table_{ws.title.replace(' ', '_')}", ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)

    for cellule in ws[1]:
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = PatternFill(start_color="ED2224", end_color="ED2224", fill_type="solid")
        cellule.alignment = Alignment(horizontal="center")

    auto_ajuster_colonnes(ws)


def generer_excel(pseudo, historique, resultats, stats_listes, doublons):
    """Génère le fichier Excel complet, en mémoire (aucun fichier écrit sur disque)."""

    df_resume = pd.DataFrame([
        ["Compte Trakt", pseudo],
        ["Films vus", historique["nb_films"]],
        ["Séries vues", historique["nb_series"]],
        ["Épisodes vus", historique["nb_episodes"]],
        ["Listes personnalisées", len(stats_listes) - 1],
        ["Contenus dans tes listes + watchlist", sum(s["total"] for s in stats_listes)],
        ["Contenus déjà vus à nettoyer", len(resultats)],
        ["Doublons entre listes", len(doublons)],
    ], columns=["Statistique", "Valeur"])

    df_resultats = pd.DataFrame(resultats)
    if not df_resultats.empty:
        df_resultats = df_resultats[["liste", "type", "titre", "annee", "vues", "dernier_visionnage", "tmdb_id"]].copy()
        df_resultats["dernier_visionnage"] = pd.to_datetime(df_resultats["dernier_visionnage"]).dt.strftime("%d/%m/%Y")
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
    df_listes["% nettoyage possible"] = (
        df_listes["deja_vus"] / df_listes["total"].replace(0, 1) * 100
    ).round(1)
    df_listes = df_listes[["nom", "nb_films", "nb_series", "total", "deja_vus", "% nettoyage possible"]]
    df_listes.columns = ["Liste", "Film", "Série", "Nombre de contenus", "Déjà vus", "% nettoyage possible"]

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_resume.to_excel(writer, sheet_name="Résumé", index=False)
        df_resultats.to_excel(writer, sheet_name="À nettoyer", index=False)
        df_doublons.to_excel(writer, sheet_name="Doublons entre listes", index=False)
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
# INTERFACE — TABLEAU SÉLECTIONNABLE RÉUTILISABLE
# ==================================================

def afficher_tableau_selectionnable(items, colonnes, noms_colonnes, cle, access_token):
    """Affiche un tableau avec cases à cocher + suppression confirmée.
    'items' est une liste de dicts contenant au moins : type, trakt_id, liste_id."""

    tableau = pd.DataFrame(items)
    tableau_affichage = tableau[colonnes].copy()

    if "dernier_visionnage" in colonnes:
        tableau_affichage["dernier_visionnage"] = pd.to_datetime(tableau_affichage["dernier_visionnage"]).dt.strftime("%d/%m/%Y")

    tableau_affichage.insert(0, "Sélectionner", False)
    tableau_affichage.columns = ["Sélectionner"] + noms_colonnes

    edite = st.data_editor(
        tableau_affichage,
        use_container_width=True,
        hide_index=True,
        disabled=noms_colonnes,
        key=f"editeur_{cle}",
    )

    nb_selectionnes = int(edite["Sélectionner"].sum())

    if nb_selectionnes == 0:
        return

    cle_confirmation = f"confirmation_{cle}"

    if not st.session_state.get(cle_confirmation, False):
        if st.button(f"🗑️ Supprimer les {nb_selectionnes} élément(s) sélectionné(s)", key=f"bouton_{cle}"):
            st.session_state[cle_confirmation] = True
            st.rerun()
        return

    st.warning(f"Confirmer la suppression de {nb_selectionnes} élément(s) de tes listes Trakt ? Cette action est irréversible.")

    col_oui, col_non = st.columns(2)

    with col_oui:
        if st.button("✅ Oui, supprimer", key=f"oui_{cle}"):
            indices = edite[edite["Sélectionner"]].index
            items_a_supprimer = [items[i] for i in indices]

            with st.spinner("Suppression en cours..."):
                supprimer_selection(access_token, items_a_supprimer)

            st.session_state[cle_confirmation] = False
            st.session_state["message_suppression"] = (
                f"{len(items_a_supprimer)} élément(s) supprimé(s) de tes listes Trakt. "
                f"Relance l'analyse (en haut) pour voir les données à jour."
            )
            st.rerun()

    with col_non:
        if st.button("❌ Annuler", key=f"non_{cle}"):
            st.session_state[cle_confirmation] = False
            st.rerun()


# ==================================================
# RECONNEXION AUTOMATIQUE (si déjà connecté avant)
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
# INTERFACE
# ==================================================

try:
    st.image("trakt-logo.svg", width=150)
except Exception:
    pass

st.title("Trakt Smart Lists")

if "access_token" not in st.session_state:

    if "device_code" not in st.session_state:

        st.write("Connecte ton compte Trakt pour commencer.")

        if st.button("Se connecter à Trakt"):
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
                f'style="display:inline-block; background-color:#ED2224; color:white; '
                f'padding:0.6em 1.4em; border-radius:8px; text-decoration:none; '
                f'font-weight:600;">Ouvrir la page d\'autorisation</a>',
                unsafe_allow_html=True,
            )
            st.caption("Sur cet appareil, ou un autre.")

        with colonne_droite:
            st.image(generer_qr_code(url_complete), width=150)
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

else:

    pseudo = obtenir_pseudo_trakt(st.session_state["access_token"])

    colonne_titre, colonne_bouton = st.columns([4, 1])
    with colonne_titre:
        st.success(f"Connecté à Trakt en tant que **{pseudo}** ✅")
    with colonne_bouton:
        if st.button("Se déconnecter"):
            oublier_connexion()
            st.rerun()

    st.divider()

    if "resultats" not in st.session_state:

        if "historique" in st.session_state:
            st.write("Ton historique est déjà en mémoire pour cette session — cette analyse sera rapide.")
        else:
            st.write("Compare ta watchlist et tes listes à ton historique pour repérer ce que tu as déjà vu. La première analyse peut prendre un peu de temps (récupération de tout ton historique).")

        if st.button("🔍 Analyser"):

            barre = st.progress(0, text="Démarrage...")

            if "historique" not in st.session_state:
                st.session_state["historique"] = obtenir_historique(st.session_state["access_token"], barre)

            resultats, stats_listes, doublons, doublons_detail = analyser_tout(st.session_state["access_token"], st.session_state["historique"], barre)

            st.session_state["resultats"] = resultats
            st.session_state["stats_listes"] = stats_listes
            st.session_state["doublons"] = doublons
            st.session_state["doublons_detail"] = doublons_detail

            barre.empty()
            st.rerun()

    else:

        historique = st.session_state["historique"]
        resultats = st.session_state["resultats"]
        stats_listes = st.session_state["stats_listes"]
        doublons = st.session_state["doublons"]
        doublons_detail = st.session_state["doublons_detail"]

        if st.session_state.get("message_suppression"):
            st.success(st.session_state["message_suppression"])

        col_relance_rapide, col_relance_totale = st.columns(2)

        with col_relance_rapide:
            if st.button("🔄 Relancer l'analyse des listes (rapide)"):
                st.session_state.pop("message_suppression", None)
                del st.session_state["resultats"]
                del st.session_state["stats_listes"]
                del st.session_state["doublons"]
                del st.session_state["doublons_detail"]
                st.rerun()

        with col_relance_totale:
            if st.button("🔃 Tout rafraîchir, historique inclus (plus long)"):
                st.session_state.pop("message_suppression", None)
                del st.session_state["historique"]
                del st.session_state["resultats"]
                del st.session_state["stats_listes"]
                del st.session_state["doublons"]
                del st.session_state["doublons_detail"]
                st.rerun()

        st.divider()
        st.subheader("📊 Statistiques")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Films vus", historique["nb_films"])
        col2.metric("Séries vues", historique["nb_series"])
        col3.metric("Épisodes vus", historique["nb_episodes"])
        col4.metric("Listes personnalisées", len(stats_listes) - 1)

        total_items = sum(s["total"] for s in stats_listes)
        total_deja_vus = sum(s["deja_vus"] for s in stats_listes)
        pourcentage_global = round(total_deja_vus / total_items * 100, 1) if total_items else 0

        col5, col6, col7 = st.columns(3)
        col5.metric("Contenus dans tes listes + watchlist", total_items)
        col6.metric("Déjà vus (tous confondus)", total_deja_vus)
        col7.metric("% potentiellement nettoyable", f"{pourcentage_global}%")

        excel_bytes = generer_excel(pseudo, historique, resultats, stats_listes, doublons)
        st.download_button(
            "📥 Télécharger le rapport Excel complet",
            data=excel_bytes,
            file_name="trakt_smart_lists_rapport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        df_stats_listes = pd.DataFrame(stats_listes)
        df_stats_listes["% nettoyable"] = (
            df_stats_listes["deja_vus"] / df_stats_listes["total"].replace(0, 1) * 100
        ).round(1)

        st.bar_chart(df_stats_listes.set_index("nom")["% nettoyable"])

        st.divider()
        st.subheader("🔁 Doublons entre listes")

        if not doublons_detail:
            st.info("Aucun doublon trouvé entre tes listes.")
        else:
            st.write(f"**{len(doublons)}** contenu(s) présent(s) dans plusieurs listes à la fois. Coche les lignes à retirer d'une liste précise :")

            afficher_tableau_selectionnable(
                doublons_detail,
                colonnes=["type", "titre", "annee", "liste"],
                noms_colonnes=["Type", "Titre", "Année", "Liste"],
                cle="doublons",
                access_token=st.session_state["access_token"],
            )

        st.divider()
        st.subheader("🧹 Contenus déjà vus")

        if not resultats:
            st.info("Aucun contenu déjà vu trouvé. Tout est propre ! 🎉")
        else:
            st.write(f"**{len(resultats)}** contenu(s) déjà vu(s) trouvé(s). Coche ceux à supprimer :")

            afficher_tableau_selectionnable(
                resultats,
                colonnes=["type", "titre", "annee", "vues", "dernier_visionnage", "liste"],
                noms_colonnes=["Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"],
                cle="deja_vus",
                access_token=st.session_state["access_token"],
            )
