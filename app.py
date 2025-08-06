import streamlit as st
import requests
import pandas as pd

# ----------------------------
# 🔐 Votre clé API Google Places
# ----------------------------
API_KEY = "AIzaSyDClKdC1ikm9K_GvPYwda6k0_7yfc8NBB4"  # ⚠️ à remplacer

# ----------------------------
# 🔎 Fonction de recherche
# ----------------------------
def search_places(api_key, location, category):
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {'query': f'{category} in {location}', 'key': api_key}
    places = []

    response = requests.get(url, params=params)
    if response.status_code == 200:
        results = response.json()
        places = results.get('results', [])[:20]  # uniquement la 1ère page

        data = []
        for place in places:
            data.append({
                'Nom': place.get('name'),
                'Adresse': place.get('formatted_address'),
                'Note (/5)': place.get('rating', 'Non noté'),
                'Nombre d’avis': place.get('user_ratings_total', 0),
                'Google Maps': f"https://www.google.com/maps/place/?q=place_id:{place.get('place_id')}"
            })

        return pd.DataFrame(data)
    else:
        st.error(f"Erreur {response.status_code} avec l’API Google.")
        return pd.DataFrame()

# ----------------------------
# 🌐 Interface Streamlit
# ----------------------------
st.title("🔎 Lieux populaires par ville")

location = st.text_input("Entrez une ville ou une localisation", value="Lisbonne")
category = st.selectbox("Catégorie", ["restaurants", "museums", "parks", "bars", "cafes"])

if st.button("Afficher les lieux populaires"):
    with st.spinner("Recherche en cours..."):
        df = search_places(API_KEY, location, category)
        if not df.empty:
            st.success(f"Voici les {len(df)} lieux les plus populaires pour '{category}' à {location} :")
            st.dataframe(df)
        else:
            st.warning("Aucun lieu trouvé ou erreur API.")