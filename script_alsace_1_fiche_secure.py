import os
import base64
import asyncio
import pandas as pd
import streamlit as st
from playwright.async_api import async_playwright

# ---------------------------------------------------------
# CONFIGURATION DE LA PAGE
# ---------------------------------------------------------
st.set_page_config(
    page_title="Générateur Fiches Typterres Alsace",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# SYSTEME D'AUTHENTIFICATION SIMPLE (VIA SECRETS)
# ---------------------------------------------------------
def check_password():
    """Vérifie le mot de passe saisi contre les secrets configurés."""
    def password_entered():
        # Vérification si le mot de passe existe dans la section [passwords] des secrets
        user_pwd = st.session_state["password_input"]
        valid_passwords = st.secrets.get("passwords", {}).values()
        
        if user_pwd in valid_passwords:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]  # Ne pas garder le mot de passe en mémoire
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Premier affichage : formulaire de connexion
        st.title("🔒 Accès Protégé")
        st.text_input(
            "Veuillez saisir le mot de passe d'accès :", 
            type="password", 
            on_change=password_entered, 
            key="password_input"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Mot de passe incorrect
        st.title("🔒 Accès Protégé")
        st.text_input(
            "Veuillez saisir le mot de passe d'accès :", 
            type="password", 
            on_change=password_entered, 
            key="password_input"
        )
        st.error("😕 Mot de passe incorrect.")
        return False
    else:
        # Mot de passe correct
        return True

if not check_password():
    st.stop()

# Bouton de déconnexion dans le menu latéral
if st.sidebar.button("Déconnexion"):
    del st.session_state["password_correct"]
    st.rerun()

# ---------------------------------------------------------
# FONCTIONS UTILITAIRES
# ---------------------------------------------------------
def img_to_base64(img_path: str) -> str:
    """Convertit une image locale en chaîne base64 pour l'intégration HTML."""
    if os.path.exists(img_path):
        with open(img_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode()
            ext = os.path.splitext(img_path)[1].replace(".", "").lower()
            mime = "image/png" if ext == "png" else "image/jpeg"
            return f"data:{mime};base64,{encoded}"
    return ""

async def create_pdf_bytes(html_content: str) -> bytes:
    """Génère un buffer PDF en mémoire à partir du HTML via Playwright."""
    async with async_playwright() as p:
        try:
            browser = await p.chromium.launch(headless=True)
        except Exception:
            browser = await p.chromium.launch(
                headless=True,
                executable_path="/usr/bin/chromium"
            )
        page = await browser.new_page()
        await page.set_content(html_content, wait_until="networkidle")
        pdf_bytes = await page.pdf(
            format="A4",
            print_background=True,
            margin={"top": "8mm", "bottom": "8mm", "left": "12mm", "right": "12mm"}
        )
        await browser.close()
        return pdf_bytes

@st.cache_data
def load_data(excel_path: str):
    """Charge le fichier Excel de données sols."""
    return pd.read_excel(excel_path)

def generate_html_template(data: dict, map_b64: str, logos_b64: list) -> str:
    """Génère le rendu HTML/CSS de la fiche technique."""
    logos_html = "".join([f'<img src="{logo}" style="height: 38px; object-fit: contain;"/>' for logo in logos_b64 if logo])
    
    html = f"""
    <!DOCTYPE html>
    <html lang="fr">
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4; margin: 0; }}
            body {{
                font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
                margin: 0;
                padding: 10mm;
                color: #2c3e50;
                box-sizing: border-box;
                height: 100vh;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
            }}
            .header {{
                border-bottom: 3px solid #2e7d32;
                padding-bottom: 8px;
                margin-bottom: 15px;
            }}
            .title {{ font-size: 22px; font-weight: bold; color: #2e7d32; margin: 0; }}
            .subtitle {{ font-size: 14px; color: #555; margin-top: 4px; }}
            .content-grid {{ display: flex; gap: 15px; flex-grow: 1; }}
            .col-left {{ width: 35%; display: flex; flex-direction: column; gap: 10px; }}
            .col-right {{ width: 65%; }}
            .map-box {{
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
                text-align: center;
                background-color: #f9f9f9;
            }}
            .map-box img {{ max-width: 100%; max-height: 220px; object-fit: contain; }}
            .info-card {{
                background: #f4f6f7;
                border-left: 4px solid #2e7d32;
                padding: 10px;
                border-radius: 2px;
                font-size: 12px;
            }}
            .info-card h4 {{ margin: 0 0 5px 0; color: #1b5e20; }}
            .footer-logos {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-top: 1px solid #ddd;
                padding-top: 10px;
                margin-top: auto;
            }}
        </style>
    </head>
    <body>
        <div>
            <div class="header">
                <div class="title">Série de Sol : {data.get('NOM_SERIE', 'N/A')}</div>
                <div class="subtitle">Identifiant Typterres : {data.get('ID_TYPTERRE', 'N/A')}</div>
            </div>

            <div class="content-grid">
                <div class="col-left">
                    <div class="map-box">
                        <strong style="font-size: 11px; display: block; margin-bottom: 4px;">Localisation Alsace</strong>
                        {f'<img src="{map_b64}"/>' if map_b64 else '<em>Carte non disponible</em>'}
                    </div>
                    <div class="info-card">
                        <h4>Secteur / Répartition</h4>
                        <p>{data.get('SECTEUR', 'Donnée non renseignée')}</p>
                    </div>
                </div>
                
                <div class="col-right">
                    <div class="info-card" style="margin-bottom: 10px;">
                        <h4>Description du Sol</h4>
                        <p>{data.get('DESCRIPTION', 'Aucune description disponible pour ce profil.')}</p>
                    </div>
                    <div class="info-card">
                        <h4>Caractéristiques Agronomiques</h4>
                        <ul>
                            <li><strong>Réserve Utile (RU) :</strong> {data.get('RU', 'N/A')} mm</li>
                            <li><strong>Drainage :</strong> {data.get('DRAINAGE', 'N/A')}</li>
                            <li><strong>Profondeur :</strong> {data.get('PROFONDEUR', 'N/A')} cm</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>

        <div class="footer-logos">
            {logos_html}
        </div>
    </body>
    </html>
    """
    return html

# ---------------------------------------------------------
# APPLICATION PRINCIPALE
# ---------------------------------------------------------
st.title("🌾 Générateur de Fiches Techniques Typterres")

EXCEL_FILE = "70_Typterres_Alsace_v04_2018_publipostageREVU.xlsx"

if not os.path.exists(EXCEL_FILE):
    st.error(f"Fichier de données introuvable : `{EXCEL_FILE}`")
    st.stop()

df = load_data(EXCEL_FILE)

map_b64 = img_to_base64("assets/alsace.png")

logo_files = [
    "assets/logo1.png", "assets/logo2.png", "assets/logo3.png",
    "assets/logo4.png", "assets/logo5.png", "assets/logo6.png", "assets/logo7.png"
]
logos_b64 = [img_to_base64(p) for p in logo_files if os.path.exists(p)]

col_left, col_right = st.columns([1, 2])

with col_left:
    st.subheader("Sélection du Sol")
    
    soil_ids = df['ID_TYPTERRE'].unique() if 'ID_TYPTERRE' in df.columns else df.index
    selected_id = st.selectbox("Choisir l'identifiant du sol :", soil_ids)
    
    soil_data = df[df['ID_TYPTERRE'] == selected_id].iloc[0].to_dict() if 'ID_TYPTERRE' in df.columns else df.loc[selected_id].to_dict()
    
    html_payload = generate_html_template(soil_data, map_b64, logos_b64)
    
    st.markdown("---")
    if st.button("📄 Générer le PDF", type="primary"):
        with st.spinner("Rendu PDF via Playwright en cours..."):
            pdf_bytes = asyncio.run(create_pdf_bytes(html_payload))
            
            st.download_button(
                label="⬇️ Télécharger la fiche PDF",
                data=pdf_bytes,
                file_name=f"Fiche_Typterres_{selected_id}.pdf",
                mime="application/pdf"
            )

with col_right:
    st.subheader("Aperçu du rendu")
    if html_payload:
        st.components.v1.html(html_payload, height=680, scrolling=True)