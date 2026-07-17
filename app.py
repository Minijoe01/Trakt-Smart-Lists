import streamlit as st
import requests
import time
import qrcode
import io
import pandas as pd
from streamlit_cookies_controller import CookieController

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
                })

    return resultats


def analyser_tout(access_token, historique, barre=None):
    """Compare la watchlist + toutes les listes personnalisées avec l'historique.
    Retourne le détail des contenus déjà vus, et des statistiques par liste."""

    resultats = []
    stats_listes = []

    if barre:
        barre.progress(0.7, text="Analyse de la watchlist...")

    watchlist = obtenir_watchlist(access_token)
    matches = comparer_items_avec_historique(watchlist, historique)
    for m in matches:
        m["liste"] = "Watchlist"
        m["liste_id"] = "watchlist"
    resultats.extend(matches)
    stats_listes.append({"nom": "Watchlist (officielle)", "total": len(watchlist), "deja_vus": len(matches)})

    listes = obtenir_listes(access_token)

    for i, liste in enumerate(listes):

        if barre:
            barre.progress(0.7 + (i + 1) / max(len(listes), 1) * 0.3, text=f"Analyse de la liste : {liste['name']}")

        items = obtenir_contenu_liste(access_token, liste["ids"]["trakt"])
        matches = comparer_items_avec_historique(items, historique)
        for m in matches:
            m["liste"] = liste["name"]
            m["liste_id"] = liste["ids"]["trakt"]
        resultats.extend(matches)

        stats_listes.append({"nom": liste["name"], "total": len(items), "deja_vus": len(matches)})

    return resultats, stats_listes


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

        st.write("Compare ta watchlist et tes listes à ton historique pour repérer ce que tu as déjà vu.")

        if st.button("🔍 Analyser"):

            barre = st.progress(0, text="Démarrage...")

            historique = obtenir_historique(st.session_state["access_token"], barre)
            resultats, stats_listes = analyser_tout(st.session_state["access_token"], historique, barre)

            st.session_state["historique"] = historique
            st.session_state["resultats"] = resultats
            st.session_state["stats_listes"] = stats_listes

            barre.empty()
            st.rerun()

    else:

        historique = st.session_state["historique"]
        resultats = st.session_state["resultats"]
        stats_listes = st.session_state["stats_listes"]

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

        df_stats_listes = pd.DataFrame(stats_listes)
        df_stats_listes["% nettoyable"] = (
            df_stats_listes["deja_vus"] / df_stats_listes["total"].replace(0, 1) * 100
        ).round(1)

        st.bar_chart(df_stats_listes.set_index("nom")["% nettoyable"])

        st.divider()
        st.subheader("🧹 Contenus déjà vus")

        if not resultats:
            st.info("Aucun contenu déjà vu trouvé. Tout est propre ! 🎉")
        else:
            st.write(f"**{len(resultats)}** contenu(s) déjà vu(s) trouvé(s). Coche ceux à supprimer :")

            tableau = pd.DataFrame(resultats)
            tableau_affichage = tableau[["type", "titre", "annee", "vues", "dernier_visionnage", "liste"]].copy()
            tableau_affichage["dernier_visionnage"] = pd.to_datetime(tableau_affichage["dernier_visionnage"]).dt.strftime("%d/%m/%Y")
            tableau_affichage.insert(0, "Sélectionner", False)
            tableau_affichage.columns = ["Sélectionner", "Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"]

            edite = st.data_editor(
                tableau_affichage,
                use_container_width=True,
                hide_index=True,
                disabled=["Type", "Titre", "Année", "Vues", "Dernier visionnage", "Liste"],
                key="editeur_resultats",
            )

            nb_selectionnes = int(edite["Sélectionner"].sum())

            if nb_selectionnes > 0:

                if not st.session_state.get("confirmation_suppression", False):

                    if st.button(f"🗑️ Supprimer les {nb_selectionnes} élément(s) sélectionné(s)"):
                        st.session_state["confirmation_suppression"] = True
                        st.rerun()

                else:

                    st.warning(f"Confirmer la suppression de {nb_selectionnes} élément(s) de tes listes Trakt ? Cette action est irréversible.")

                    col_oui, col_non = st.columns(2)

                    with col_oui:
                        if st.button("✅ Oui, supprimer"):

                            indices = edite[edite["Sélectionner"]].index
                            items_a_supprimer = [resultats[i] for i in indices]

                            with st.spinner("Suppression en cours..."):
                                supprimer_selection(st.session_state["access_token"], items_a_supprimer)

                            st.session_state["confirmation_suppression"] = False
                            del st.session_state["historique"]
                            del st.session_state["resultats"]
                            del st.session_state["stats_listes"]

                            st.success(f"{nb_selectionnes} élément(s) supprimé(s) !")
                            time.sleep(1.5)
                            st.rerun()

                    with col_non:
                        if st.button("❌ Annuler"):
                            st.session_state["confirmation_suppression"] = False
                            st.rerun()

        if st.button("🔄 Relancer l'analyse"):
            del st.session_state["historique"]
            del st.session_state["resultats"]
            del st.session_state["stats_listes"]
            st.rerun()
