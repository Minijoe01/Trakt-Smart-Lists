import streamlit as st
import requests
import time

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🎬")

# ==================================================
# CONFIGURATION (lue depuis les secrets Streamlit)
# ==================================================

CLIENT_ID = st.secrets["TRAKT_CLIENT_ID"]
CLIENT_SECRET = st.secrets["TRAKT_CLIENT_SECRET"]

DEVICE_CODE_URL = "https://api.trakt.tv/oauth/device/code"
DEVICE_TOKEN_URL = "https://api.trakt.tv/oauth/device/token"


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


# ==================================================
# INTERFACE
# ==================================================

st.title("🎬 Trakt Smart Lists")

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

        st.markdown(f"**1.** Va sur : "
                    f'<a href="{url}" target="_blank">{url}</a>',
                    unsafe_allow_html=True)
        st.markdown(f"**2.** Entre ce code : `{code}`")
        st.caption("Tu peux faire ça sur ton téléphone si tu préfères. Garde cette page ouverte, elle se met à jour toute seule.")

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
                    st.session_state["access_token"] = tokens["access_token"]
                    st.session_state["refresh_token"] = tokens["refresh_token"]
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
        st.session_state.clear()
        st.rerun()
