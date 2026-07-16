import streamlit as st
import requests
import time
import qrcode
import io
from streamlit_cookies_controller import CookieController

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🎬")

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
# FONCTIONS TRAKT
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


def obtenir_pseudo_trakt(access_token):
    """Récupère le pseudo de l'utilisateur connecté, pour confirmer que ça fonctionne."""

    headers = {
        "Content-Type": "application/json",
        "trakt-api-version": "2",
        "trakt-api-key": CLIENT_ID,
        "Authorization": f"Bearer {access_token}",
    }

    response = requests.get("https://api.trakt.tv/users/settings", headers=headers)
    response.raise_for_status()

    return response.json()["user"]["username"]


def generer_qr_code(url):
    """Génère un QR code en mémoire (aucun fichier écrit sur disque)."""

    image = qrcode.make(url)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")

    return buffer.getvalue()


# ==================================================
# RECONNEXION AUTOMATIQUE (si déjà connecté avant)
# ==================================================

if "access_token" not in st.session_state:

    refresh_token_sauvegarde = cookies.get("trakt_refresh_token")

    if refresh_token_sauvegarde:
        tokens = rafraichir_token(refresh_token_sauvegarde)
        if tokens:
            sauvegarder_connexion(tokens)
            st.rerun()
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

    st.success(f"Connecté à Trakt en tant que **{pseudo}** ✅")

    if st.button("Se déconnecter"):
        cookies.remove("trakt_refresh_token")
        st.session_state.clear()
        st.rerun()
