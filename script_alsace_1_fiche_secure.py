import os
import glob
import asyncio
import base64
import pandas as pd
import streamlit as st
from playwright.async_api import async_playwright

st.set_page_config(
    page_title="Générateur Fiches Typterres",
    page_icon="🌱",
    layout="wide"
)



# ---------------------------------------------------------
# SYSTEME D'AUTHENTIFICATION SIMPLE (VIA SECRETS)
# ---------------------------------------------------------
def check_password():
    """Vérifie le mot de passe saisi et récupère l'identifiant utilisateur."""
    def password_entered():
        user_pwd = st.session_state["password_input"]
        passwords_dict = st.secrets.get("passwords", {})
        
        # Recherche du nom d'utilisateur correspondant au mot de passe
        matched_user = None
        for username, password in passwords_dict.items():
            if user_pwd == password:
                matched_user = username
                break
        
        if matched_user:
            st.session_state["password_correct"] = True
            st.session_state["user_name"] = matched_user  # Stocke "admin", "ServiceMTA", etc.
            del st.session_state["password_input"]
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
        return True

if not check_password():
    st.stop()

# ---------------------------------------------------------
# AFFICHAGE DU MESSAGE ET BOUTON DE DECONNEXION (BARRE LATERALE)
# ---------------------------------------------------------
current_user = st.session_state.get("user_name", "Utilisateur")
st.sidebar.markdown(f"**Bienvenue {current_user}**")

if st.sidebar.button("Déconnexion"):
    del st.session_state["password_correct"]
    if "user_name" in st.session_state:
        del st.session_state["user_name"]
    st.rerun()

# ==========================================
# Fichiers d'images et logos
# ==========================================
MAP_IMAGE_PATH = "alsace.png"

LOGO_FILES = [
    "TypTerres_alsace.png",
    "CA_GE.png",
    "GIS_SOL.png",
    "GE.png",
    "Agence_eau_Rhin_Meuse.png",
    "CASDAR.png",
    "UE.png"
]

def get_base64_img_src(file_path):
    """Convertit une image locale en source base64 pour le HTML."""
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
            return f"data:image/png;base64,{encoded}"
    return None

def get_base64_logos_html(image_paths):
    img_tags = []
    for path in image_paths:
        src = get_base64_img_src(path)
        if src:
            img_tags.append(f'<img src="{src}" alt="Logo" />')
        else:
            st.toast(f"⚠️ Logo introuvable : {path}", icon="⚠️")
    
    if img_tags:
        return f'<div class="logos-container">{"".join(img_tags)}</div>'
    return ""

@st.cache_resource
def clear_old_pdfs():
    pdfs = glob.glob("*.pdf")
    count = 0
    for file in pdfs:
        try:
            os.remove(file)
            count += 1
        except Exception:
            pass
    return count

cleaned_count = clear_old_pdfs()

EXCEL_PATH = "70_Typterres_Alsace_v04_2018_publipostageREVU.xlsx"
COL_ID = 'Identifiant Typterres (1 à 70)'

@st.cache_data
def load_dataset():
    if not os.path.exists(EXCEL_PATH):
        st.error(f"⚠️ Fichier Excel introuvable : `{EXCEL_PATH}`")
        return None
    return pd.read_excel(EXCEL_PATH)

df = load_dataset()

DRAINAGE_MAP = {
    1.0: "Excessif", 2.0: "Bon", 3.0: "Modéré", 
    4.0: "Imparfait", 5.0: "Pauvre", 6.0: "Très pauvre"
}

PIERROSITE_MAP = {
    "0": "nulle à très faible (<5%)", "1": "faible (5% à 15%)",
    "2": "moyenne (15% à 30%)", "3": "forte (30% à 50%)"
}

def generate_html(df_data, target_id, logos_html=""):
    sub_df = df_data[df_data[COL_ID] == target_id].sort_values('Numéro couche Typterres')
    if sub_df.empty:
        return None
    
    first = sub_df.iloc[0]
    
    titre_typterre = f"TYPTERRE {int(first[COL_ID])}"
    nom_typterre = str(first['NOM TYPTERRES (70)']) if pd.notna(first['NOM TYPTERRES (70)']) else ""
    ref_pedo = str(first['NOM REFERENTIEL PEDOLOGIQUE']) if pd.notna(first['NOM REFERENTIEL PEDOLOGIQUE']) else ""
    petite_region = str(first['Petite Région Typterres (11)']) if pd.notna(first['Petite Région Typterres (11)']) else ""
    mat_parental = str(first['NOM MATERIAU PARENTAL']) if pd.notna(first['NOM MATERIAU PARENTAL']) else ""
    
    surface = f"{int(first['Surface Totale ha Typterres Simp'])} ha" if pd.notna(first['Surface Totale ha Typterres Simp']) else ""
    correspondance_typt = str(first['Identifiant sous TypSimplifié (70)']) if pd.notna(first['Identifiant sous TypSimplifié (70)']) else ""
    guide_sols = str(first['exemple FICHE GUIDE des sols']) if pd.notna(first['exemple FICHE GUIDE des sols']) else ""
    directive_nitrates = str(first['correspondance GREN Directive Nitrates']) if pd.notna(first['correspondance GREN Directive Nitrates']) else ""
    
    ep_val, ep_min, ep_max = first["Epaisseur Sol"], first["Epaisseur Sol 'min'"], first["Epaisseur Sol 'max'"]
    epaisseur = f"{int(ep_val)} cm (min : {int(ep_min)} cm , max : {int(ep_max)} cm)" if pd.notna(ep_val) else ""
    
    pierrosite_val = str(first['Pierrosité surface']) if pd.notna(first['Pierrosité surface']) else ""
    pierrosite_txt = PIERROSITE_MAP.get(pierrosite_val, pierrosite_val)
        
    ru_val, ru_min, ru_max = first['Estimation RU du Sol (mm)'], first["Estimation RU du Sol 'min' (mm)"], first["Estimation RU du Sol 'max' (mm)"]
    ru_sol = f"{round(ru_val)} mm ( min : {round(ru_min)} mm , max : {round(ru_max)} mm)" if pd.notna(ru_val) else ""
    
    effervescence = str(first['Effervescence en clair']) if pd.notna(first['Effervescence en clair']) else ""
    drainage_txt = DRAINAGE_MAP.get(first['Drainage naturel'], str(first['Drainage naturel']) if pd.notna(first['Drainage naturel']) else "")
    
    # Image de la carte alsace
    map_src = get_base64_img_src(MAP_IMAGE_PATH)
    if map_src:
        map_html = f'<img src="{map_src}" class="map-img" alt="Carte Alsace" />'
    else:
        map_html = '<div style="color:red; font-size:8pt; margin-top:5px;">⚠️ Image alsace.png introuvable</div>'

    horizons = []
    for idx, row in sub_df.iterrows():
        horizons.append({
            'num': f"H{int(row['Numéro couche Typterres'])} ({row['Nom couche Typterres'] if pd.notna(row['Nom couche Typterres']) else ''})",
            'ep': int(row['Epaissseur couche']) if pd.notna(row['Epaissseur couche']) else "",
            'geppa': str(row['TEXTURE GEPPA']) if pd.notna(row['TEXTURE GEPPA']) else "",
            'argile': round(row['TAUX ARGILE'] / 10, 1) if pd.notna(row['TAUX ARGILE']) else "",
            'limon': round(row['TAUX LIMON'] / 10, 1) if pd.notna(row['TAUX LIMON']) else "",
            'sable': round(row['TAUX SABLE'] / 10, 1) if pd.notna(row['TAUX SABLE']) else "",
            'eg': int(row['Abondance volumique en éléments grossiers']) if pd.notna(row['Abondance volumique en éléments grossiers']) else "",
            'mo': round(row['Matière organique'] / 10, 1) if pd.notna(row['Matière organique']) else "",
            'ph': round(row['pH eau'], 1) if pd.notna(row['pH eau']) else "",
            'calc': int(row['Calcaire total']) if pd.notna(row['Calcaire total']) else "",
            'cec': round(row['CEC'], 1) if pd.notna(row['CEC']) else "",
            'da': round(row['Densité apparente'], 2) if pd.notna(row['Densité apparente']) else "",
            'couleur': str(row['Couleur']) if pd.notna(row['Couleur']) else ""
        })

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<style>
@page {{ size: A4 portrait; margin: 8mm 12mm 8mm 12mm; }}
*, *::before, *::after {{ box-sizing: border-box; }}

html, body {{ height: 100%; margin: 0; padding: 0; }}
body {{ 
    font-family: Arial, sans-serif; 
    color: #111; 
    font-size: 8.5pt; 
    line-height: 1.2; 
    display: flex; 
    flex-direction: column; 
}}

.content-wrapper {{ flex: 1 0 auto; }}

.header-top {{ text-align: right; font-weight: bold; font-size: 14pt; margin-bottom: 4px; }}
.title-box {{ background-color: #e0f7fa; border: 1.5px solid #000; padding: 5px 8px; font-size: 12.5pt; font-weight: bold; }}
.subtitle-box {{ background-color: #00a896; border: 1.5px solid #000; border-top: none; padding: 5px 8px; font-size: 10pt; font-weight: bold; color: #fff; margin-bottom: 8px; }}

.info-grid {{ width: 100%; border-collapse: collapse; margin-bottom: 6px; }}
.info-grid td {{ vertical-align: top; padding: 2px 0; }}
.label-cyan {{ color: #00a896; font-weight: bold; }}

/* Style pour la carte Alsace (Agrandie) */
.map-img {{ 
    width: 100%; 
    max-height: 220px; 
    object-fit: contain; 
    margin-top: 6px; 
    display: block; 
}}

.section-title {{ text-align: center; color: #b25900; font-size: 11pt; font-weight: bold; margin: 6px 0 4px 0; }}
.params-table {{ width: 100%; border-collapse: collapse; margin-bottom: 6px; }}
.params-table td {{ width: 50%; vertical-align: top; padding: 2px 4px; }}

.data-table {{ width: 100%; border-collapse: collapse; margin-top: 4px; margin-bottom: 6px; }}
.data-table th, .data-table td {{ border: 1px dashed #a0a0a0; padding: 2.5px 4px; font-size: 7.5pt; text-align: left; }}
.data-table th {{ background-color: #ffffff; font-weight: bold; }}

.footer-note {{ font-size: 6.8pt; font-style: italic; color: #333; margin-top: 4px; }}

/* Zone footer calée automatiquement en bas de page */
.footer-wrapper {{ 
    margin-top: auto; 
    padding-top: 4px; 
    width: 100%; 
}}

.logos-container {{ 
    width: 100%; 
    display: flex; 
    justify-content: space-between; 
    align-items: center; 
    padding-top: 4px; 
    border-top: 1px solid #ccc; 
}}
.logos-container img {{ 
    max-height: 36px; 
    width: auto; 
    object-fit: contain; 
}}
</style>
</head>
<body>

<div class="content-wrapper">
    <div class="header-top">{titre_typterre}</div>
    <div class="title-box">{nom_typterre}</div>
    <div class="subtitle-box">{ref_pedo}</div>

    <table class="info-grid">
        <tr>
            <td style="width: 42%;">
                <span class="label-cyan">Petite Région :</span><br>{petite_region}<br>
                {map_html}
            </td>
            <td style="width: 58%;">
                <span class="label-cyan">Matériau parental :</span> {mat_parental}<br><br>
                <span class="label-cyan">Surface occupée par le {titre_typterre} :</span> {surface}<br><br><br>
                <span class="label-cyan">Correspondances Typterres :</span> {correspondance_typt}<br><br>
                <span class="label-cyan">Guide des sols :</span> {guide_sols}<br><br>
                <span class="label-cyan">Directive Nitrates GREN :</span> {directive_nitrates}
            </td>
        </tr>
    </table>

    <div class="section-title">Caractéristiques physico-chimiques</div>

    <table class="params-table">
        <tr>
            <td>
                <span class="label-cyan">Epaisseur du Sol :</span> {epaisseur}<br><br>
                <span class="label-cyan">Estimation réserve en eau du sol :</span><br>{ru_sol}<br><br>
                <span class="label-cyan">Drainage naturel :</span> {drainage_txt}
            </td>
            <td>
                <span class="label-cyan">Pierrosité en surface :</span> {pierrosite_txt}<br><br>
                <span class="label-cyan">Effervescence en surface :</span> {effervescence}
            </td>
        </tr>
    </table>

    <table class="data-table">
        <thead>
            <tr>
                <th style="width: 35%;">Propriétés</th>
                <th colspan="5">Horizons de sol</th>
            </tr>
            <tr>
                <th>N° Horizon (nom)</th>
                {"".join([f"<td>{h['num']}</td>" for h in horizons])}
                {"".join(["<td>H" + str(i+len(horizons)+1) + " ()</td>" for i in range(5 - len(horizons))])}
            </tr>
        </thead>
        <tbody>
            <tr><td>Epaisseur (cm)</td>{"".join([f"<td>{h['ep']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Texture (classe GEPPA)</td>{"".join([f"<td>{h['geppa']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Argile (%)</td>{"".join([f"<td>{h['argile']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Limons (%)</td>{"".join([f"<td>{h['limon']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Sables (%)</td>{"".join([f"<td>{h['sable']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Eléments grossiers (abondance vol, %)</td>{"".join([f"<td>{h['eg']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>MO (%)</td>{"".join([f"<td>{h['mo']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>pH</td>{"".join([f"<td>{h['ph']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Calcaire total (g/kg)</td>{"".join([f"<td>{h['calc']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>CEC (cmol/kg)</td>{"".join([f"<td>{h['cec']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>Densité apparente</td>{"".join([f"<td>{h['da']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
            <tr><td>couleur</td>{"".join([f"<td>{h['couleur']}</td>" for h in horizons])}{"".join(["<td></td>" for _ in range(5 - len(horizons))])}</tr>
        </tbody>
    </table>
</div>

<div class="footer-wrapper">
    <div class="footer-note">Ces résultats sont calculés à partir des données du Référentiel Régional Pédologique Alsace. Ils sont indicatifs et ne se substituent pas à une analyse de terre.</div>
    {logos_html}
</div>

</body>
</html>"""

async def create_pdf_bytes(html_content):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        await page.set_content(html_content)
        pdf_bytes = await page.pdf(
            format="A4",
            print_background=True,
            margin={"top": "8mm", "bottom": "8mm", "left": "12mm", "right": "12mm"}
        )
        await browser.close()
        return pdf_bytes

# ==========================================
# Interface Streamlit
# ==========================================
st.title("🌾 Générateur de Fiches Techniques Typterres - Alsace")
st.markdown("Sélectionnez un identifiant pour prévisualiser les données et générer le document PDF final.")

logos_html_block = get_base64_logos_html(LOGO_FILES)

if df is not None:
    raw_series = df[COL_ID].dropna()
    valid_series = raw_series[pd.to_numeric(raw_series, errors='coerce').notna()]
    ordered_ids = [int(x) for x in pd.unique(valid_series)]

    st.sidebar.header("🕹️ Contrôles")
    
    selected_id = st.sidebar.selectbox(
        "Choisissez l'identifiant Typterre :",
        options=ordered_ids,
        index=0,
        format_func=lambda x: f"Typterre n°{x}"
    )

    st.sidebar.markdown("---")
    st.sidebar.metric("Fiches disponibles", len(ordered_ids))
    if cleaned_count > 0:
        st.sidebar.caption(f"🧹 Purge automatique : {cleaned_count} ancien(s) PDF supprimé(s).")

    col_left, col_right = st.columns([1, 1.2])

    html_payload = generate_html(df, selected_id, logos_html_block)

    with col_left:
        st.subheader(f"📄 Typterre {selected_id}")
        
        sub = df[df[COL_ID] == selected_id]
        nom_sol = sub['NOM TYPTERRES (70)'].iloc[0] if pd.notna(sub['NOM TYPTERRES (70)'].iloc[0]) else "N/A"
        ref_sol = sub['NOM REFERENTIEL PEDOLOGIQUE'].iloc[0] if pd.notna(sub['NOM REFERENTIEL PEDOLOGIQUE'].iloc[0]) else "N/A"
        nb_horizons = len(sub)

        st.info(f"**Nom :** {nom_sol}\n\n**Référentiel :** {ref_sol}\n\n**Horizons :** {nb_horizons} couche(s)")
        
        if st.button("⚙️ Générer la fiche PDF", type="primary", use_container_width=True):
            with st.spinner("Génération du PDF par Playwright en cours..."):
                pdf_bytes = asyncio.run(create_pdf_bytes(html_payload))
                filename = f"fiche_typterre_{selected_id}.pdf"
                
                with open(filename, "wb") as f:
                    f.write(pdf_bytes)
                
                st.success(f"Fiche générée avec succès : `{filename}`")
                
                st.download_button(
                    label="📥 Télécharger le fichier PDF",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True
                )

    with col_right:
        st.subheader("Aperçu du rendu")
        if html_payload:
            st.components.v1.html(html_payload, height=680, scrolling=True)

            ## commit test_typterre_alsace