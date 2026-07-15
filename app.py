import streamlit as st
import requests
import urllib.parse

st.set_page_config(page_title="Trakt Smart Lists", page_icon="🎬")

# ==================================================
# CONFIGURATION (lue depuis les secrets Streamlit,
# jamais écrite en dur dans le code)
# ==================================================

CLIENT_ID = st.secrets["TRAKT_CLIENT_ID"]
CLIENT_SECRET = st.secrets["TRAKT_CLIENT_SECRET"]
REDIRECT_URI = st.secrets["TRAKT_REDIRECT_URI"]

AUTH_URL = "https://trakt.tv/oauth/authorize"
TOKEN_URL = "https://api.trakt.tv/oauth/token"


# ==================================================
# FONCTIONS TRAKT
# ==================================================

def echanger_code_contre_token(code):
    """Échange le code reçu de Trakt contre un access_token (côté serveur)."""

    payload = {
        "code": code,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
    }

    response = requests.post(TOKEN_URL, json=payload)
    response.raise_for_status()

    return response.json()


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

    # Trakt vient-il de nous rediriger avec un code ?
    code = st.query_params.get("code")

    if code:
        with st.spinner("Connexion à Trakt en cours..."):
            try:
                tokens = echanger_code_contre_token(code)
                st.session_state["access_token"] = tokens["access_token"]
                st.session_state["refresh_token"] = tokens["refresh_token"]
                st.query_params.clear()
                st.rerun()
            except Exception:
                st.error("La connexion à Trakt a échoué. Réessaie ci-dessous.")
                st.query_params.clear()

    if "access_token" not in st.session_state:

        st.write("Connecte ton compte Trakt pour commencer.")

        params = {
            "response_type": "code",
            "client_id": CLIENT_ID,
            "redirect_uri": REDIRECT_URI,
        }

        auth_url = AUTH_URL + "?" + urllib.parse.urlencode(params)

        st.markdown(
            f'<a href="{auth_url}" target="_self" '
            f'style="display:inline-block; background-color:#ED2224; color:white; '
            f'padding:0.6em 1.4em; border-radius:8px; text-decoration:none; '
            f'font-weight:600;">Se connecter à Trakt</a>',
            unsafe_allow_html=True,
        )

else:

    pseudo = obtenir_pseudo_trakt(st.session_state["access_token"])

    st.success(f"Connecté à Trakt en tant que **{pseudo}** ✅")

    if st.button("Se déconnecter"):
        st.session_state.clear()
        st.rerun()
