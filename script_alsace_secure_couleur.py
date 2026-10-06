import base64
import glob
import io
import os
import re
import pandas as pd
import streamlit as st

try:
  from weasyprint import HTML

  HAS_WEASYPRINT = True
except ImportError:
  from xhtml2pdf import pisa

  HAS_WEASYPRINT = False

st.set_page_config(
    page_title="Générateur Fiches Typterres", page_icon="🌱", layout="wide"
)


# ---------------------------------------------------------
# AUTHENTIFICATION SIMPLE
# ---------------------------------------------------------
def check_password():
  def password_entered():
    user_pwd = st.session_state["password_input"]
    passwords_dict = st.secrets.get("passwords", {})

    matched_user = None
    for username, password in passwords_dict.items():
      if user_pwd == password:
        matched_user = username
        break

    if matched_user:
      st.session_state["password_correct"] = True
      st.session_state["user_name"] = matched_user
      del st.session_state["password_input"]
    else:
      st.session_state["password_correct"] = False

  if "password_correct" not in st.session_state:
    st.title("🔒 Accès Protégé")
    st.text_input(
        "Veuillez saisir le mot de passe d'accès :",
        type="password",
        on_change=password_entered,
        key="password_input",
    )
    return False
  elif not st.session_state["password_correct"]:
    st.title("🔒 Accès Protégé")
    st.text_input(
        "Veuillez saisir le mot de passe d'accès :",
        type="password",
        on_change=password_entered,
        key="password_input",
    )
    st.error("😕 Mot de passe incorrect.")
    return False
  else:
    return True


if not check_password():
  st.stop()

current_user = st.session_state.get("user_name", "Utilisateur")
st.sidebar.markdown(f"**Bienvenue {current_user}**")

if st.sidebar.button("Déconnexion"):
  del st.session_state["password_correct"]
  if "user_name" in st.session_state:
    del st.session_state["user_name"]
  st.rerun()

# ---------------------------------------------------------
# LOGOS ET IMAGES
# ---------------------------------------------------------
MAP_IMAGE_PATH = "alsace.png"
LOGO_FILES = [
    "TypTerres_alsace.png",
    "CA_GE.png",
    "GIS_SOL.png",
    "GE.png",
    "Agence_eau_Rhin_Meuse.png",
    "CASDAR.png",
    "UE.png",
]


def get_base64_img_src(file_path):
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

# ---------------------------------------------------------
# CHARGEMENT DU DATASET ET DÉTECTION DES COLONNES
# ---------------------------------------------------------
EXCEL_PATH = "70_Typterres_Alsace_v04_2018_publipostageREVU.xlsx"


@st.cache_data
def load_dataset():
  if not os.path.exists(EXCEL_PATH):
    st.error(f"⚠️ Fichier Excel introuvable : `{EXCEL_PATH}`")
    return None
  return pd.read_excel(EXCEL_PATH)


df = load_dataset()


def find_col(df_data, keywords, fallback_index=None):
  """Recherche une colonne par mots-clés ou par son index de secours."""
  for kw in keywords:
    for col in df_data.columns:
      if kw.lower() in str(col).strip().lower():
        return col
  if fallback_index is not None and len(df_data.columns) > fallback_index:
    return df_data.columns[fallback_index]
  return None


DRAINAGE_MAP = {
    1.0: "Excessif",
    2.0: "Bon",
    3.0: "Modéré",
    4.0: "Imparfait",
    5.0: "Pauvre",
    6.0: "Très pauvre",
}
PIERROSITE_MAP = {
    "0": "nulle à très faible (<5%)",
    "1": "faible (5% à 15%)",
    "2": "moyenne (15% à 30%)",
    "3": "forte (30% à 50%)",
}

COULEUR_PAR_DEFAUT = "#00a896"


def extract_hex_color(val):
  if pd.isna(val):
    return None
  v = str(val).strip()
  match = re.search(r"#[0-9a-fA-F]{6}", v)
  if match:
    return match.group(0).upper()
  match_nohash = re.search(r"^[0-9a-fA-F]{6}$", v)
  if match_nohash:
    return f"#{match_nohash.group(0)}".upper()
  return None


def hex_to_rgba(hex_color, alpha=0.3):
  if not hex_color or not hex_color.startswith("#") or len(hex_color) != 7:
    return f"rgba(0, 168, 150, {alpha})"
  h = hex_color.lstrip("#")
  r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
  return f"rgba({r}, {g}, {b}, {alpha})"


def text_color_for(bg_hex):
  h = bg_hex.lstrip("#")
  if len(h) != 6:
    return "#ffffff"
  r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
  luminance = 0.299 * r + 0.587 * g + 0.114 * b
  return "#111111" if luminance > 150 else "#ffffff"


@st.cache_data
def read_ar_fills():
  fills = {}
  try:
    from openpyxl import load_workbook

    ws = load_workbook(EXCEL_PATH, data_only=True).active
    col_ar_idx = 44  # Colonne AR (1-based index)
    for r in range(2, ws.max_row + 1):
      f = ws.cell(row=r, column=col_ar_idx).fill
      if f is not None and f.fill_type == "solid":
        rgb = f.fgColor.rgb
        if isinstance(rgb, str) and len(rgb) == 8 and rgb[2:].upper() != "000000":
          fills[r - 2] = "#" + rgb[2:]
  except Exception:
    pass
  return fills


def get_couleur_info(sub_df, col_couleur):
  if col_couleur is None or col_couleur not in sub_df.columns:
    return "AR", None, COULEUR_PAR_DEFAUT, "colonne non trouvée"

  fills = read_ar_fills()
  raw_found = None

  for idx, row in sub_df.iterrows():
    val = row[col_couleur]
    if raw_found is None and pd.notna(val):
      raw_found = val

    c = extract_hex_color(val)
    if c:
      return col_couleur, val, c, "valeur hex de la cellule"

    if idx in fills:
      return col_couleur, val, fills[idx], "remplissage de la cellule"

  return col_couleur, raw_found, COULEUR_PAR_DEFAUT, "couleur par défaut"


def generate_html(
    df_data,
    target_id,
    col_id,
    col_couleur,
    col_layer,
    col_nom,
    col_ref,
    logos_html="",
):
  sub_df = df_data[df_data[col_id] == target_id]
  if col_layer and col_layer in sub_df.columns:
    sub_df = sub_df.sort_values(col_layer)

  if sub_df.empty:
    return None

  first = sub_df.iloc[0]

  titre_typterre = f"TYPTERRE {int(first[col_id])}"
  nom_typterre = (
      str(first[col_nom]) if col_nom and pd.notna(first[col_nom]) else ""
  )
  ref_pedo = str(first[col_ref]) if col_ref and pd.notna(first[col_ref]) else ""

  c_region = find_col(df_data, ["petite région", "région"])
  petite_region = (
      str(first[c_region]) if c_region and pd.notna(first[c_region]) else ""
  )

  c_mat = find_col(df_data, ["matériau parental", "materiau"])
  mat_parental = str(first[c_mat]) if c_mat and pd.notna(first[c_mat]) else ""

  c_surf = find_col(df_data, ["surface totale", "surface"])
  surface = (
      f"{int(first[c_surf])} ha" if c_surf and pd.notna(first[c_surf]) else ""
  )

  c_sub = find_col(df_data, ["sous typsimplifié", "simplifié"])
  correspondance_typt = str(first[c_sub]) if c_sub and pd.notna(first[c_sub]) else ""

  c_guide = find_col(df_data, ["fiche guide", "guide"])
  guide_sols = str(first[c_guide]) if c_guide and pd.notna(first[c_guide]) else ""

  c_gren = find_col(df_data, ["gren", "nitrates"])
  directive_nitrates = str(first[c_gren]) if c_gren and pd.notna(first[c_gren]) else ""

  # Récupération couleur
  _, _, couleur_fond, _ = get_couleur_info(sub_df, col_couleur)
  couleur_transparente = hex_to_rgba(couleur_fond, alpha=0.30)
  couleur_texte_subtitle = text_color_for(couleur_fond)

  # Épaisseur & RU
  c_ep = find_col(df_data, ["epaisseur sol"], fallback_index=None)
  c_ep_min = find_col(df_data, ["epaisseur sol 'min'"])
  c_ep_max = find_col(df_data, ["epaisseur sol 'max'"])

  epaisseur = ""
  if c_ep and pd.notna(first[c_ep]):
    ep_val = int(first[c_ep])
    ep_min = int(first[c_ep_min]) if c_ep_min and pd.notna(first[c_ep_min]) else ""
    ep_max = int(first[c_ep_max]) if c_ep_max and pd.notna(first[c_ep_max]) else ""
    epaisseur = f"{ep_val} cm (min : {ep_min} cm , max : {ep_max} cm)"

  c_pierr = find_col(df_data, ["pierrosité surface", "pierrosité"])
  pierrosite_val = str(first[c_pierr]) if c_pierr and pd.notna(first[c_pierr]) else ""
  pierrosite_txt = PIERROSITE_MAP.get(pierrosite_val, pierrosite_val)

  c_ru = find_col(df_data, ["estimation ru du sol (mm)"])
  c_ru_min = find_col(df_data, ["estimation ru du sol 'min'"])
  c_ru_max = find_col(df_data, ["estimation ru du sol 'max'"])

  ru_sol = ""
  if c_ru and pd.notna(first[c_ru]):
    ru_val = round(first[c_ru])
    ru_min = round(first[c_ru_min]) if c_ru_min and pd.notna(first[c_ru_min]) else ""
    ru_max = round(first[c_ru_max]) if c_ru_max and pd.notna(first[c_ru_max]) else ""
    ru_sol = f"{ru_val} mm ( min : {ru_min} mm , max : {ru_max} mm)"

  c_eff = find_col(df_data, ["effervescence"])
  effervescence = str(first[c_eff]) if c_eff and pd.notna(first[c_eff]) else ""

  c_drain = find_col(df_data, ["drainage naturel", "drainage"])
  drain_val = first[c_drain] if c_drain and pd.notna(first[c_drain]) else ""
  drainage_txt = DRAINAGE_MAP.get(drain_val, str(drain_val))

  map_src = get_base64_img_src(MAP_IMAGE_PATH)
  map_html = (
      f'<img src="{map_src}" class="map-img" alt="Carte Alsace" />'
      if map_src
      else ""
  )

  # Récupération dynamique des colonnes de tableau d'horizons
  c_nom_couche = find_col(df_data, ["nom couche"])
  c_epaiss_couche = find_col(df_data, ["epaissseur couche", "épaisseur couche"])
  c_geppa = find_col(df_data, ["texture geppa", "geppa"])
  c_argile = find_col(df_data, ["taux argile", "argile"])
  c_limon = find_col(df_data, ["taux limon", "limon"])
  c_sable = find_col(df_data, ["taux sable", "sable"])
  c_eg = find_col(
      df_data, ["abondance volumique", "éléments grossiers", "grossiers"]
  )
  c_mo = find_col(df_data, ["matière organique"])
  c_ph = find_col(df_data, ["ph eau", "ph"])
  c_calc = find_col(df_data, ["calcaire total"])
  c_cec = find_col(df_data, ["cec"])
  c_da = find_col(df_data, ["densité apparente"])
  c_couleur_txt = find_col(df_data, ["couleur"], fallback_index=43)

  horizons = []
  for idx, row in sub_df.iterrows():
    num_c = (
        int(row[col_layer]) if col_layer and pd.notna(row[col_layer]) else ""
    )
    nom_c = (
        str(row[c_nom_couche])
        if c_nom_couche and pd.notna(row[c_nom_couche])
        else ""
    )

    horizons.append({
        "num": f"H{num_c} ({nom_c})",
        "ep": (
            int(row[c_epaiss_couche])
            if c_epaiss_couche and pd.notna(row[c_epaiss_couche])
            else ""
        ),
        "geppa": (
            str(row[c_geppa]) if c_geppa and pd.notna(row[c_geppa]) else ""
        ),
        "argile": (
            round(row[c_argile] / 10, 1)
            if c_argile and pd.notna(row[c_argile])
            else ""
        ),
        "limon": (
            round(row[c_limon] / 10, 1)
            if c_limon and pd.notna(row[c_limon])
            else ""
        ),
        "sable": (
            round(row[c_sable] / 10, 1)
            if c_sable and pd.notna(row[c_sable])
            else ""
        ),
        "eg": int(row[c_eg]) if c_eg and pd.notna(row[c_eg]) else "",
        "mo": (
            round(row[c_mo] / 10, 1) if c_mo and pd.notna(row[c_mo]) else ""
        ),
        "ph": round(row[c_ph], 1) if c_ph and pd.notna(row[c_ph]) else "",
        "calc": int(row[c_calc]) if c_calc and pd.notna(row[c_calc]) else "",
        "cec": round(row[c_cec], 1) if c_cec and pd.notna(row[c_cec]) else "",
        "da": round(row[c_da], 2) if c_da and pd.notna(row[c_da]) else "",
        "couleur": (
            str(row[c_couleur_txt])
            if c_couleur_txt and pd.notna(row[c_couleur_txt])
            else ""
        ),
    })

  return f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<style>
@page {{
    size: A4 portrait;
    margin: 8mm 12mm 10mm 12mm;
}}
*, *::before, *::after {{ box-sizing: border-box; }}

html, body {{
    height: 100%;
    margin: 0;
    padding: 0;
    font-family: Arial, sans-serif;
    color: #111;
    font-size: 8.5pt;
    line-height: 1.2;
}}

.header-top {{ text-align: right; font-weight: bold; font-size: 14pt; margin-bottom: 4px; }}

.title-box {{ 
    background-color: {couleur_transparente}; 
    color: #111111; 
    border: 1.5px solid #000; 
    padding: 5px 8px; 
    font-size: 12.5pt; 
    font-weight: bold; 
}}

.subtitle-box {{ 
    background-color: {couleur_fond}; 
    border: 1.5px solid #000; 
    border-top: none; 
    padding: 5px 8px; 
    font-size: 10pt; 
    font-weight: bold; 
    color: {couleur_texte_subtitle}; 
    margin-bottom: 8px; 
}}

.info-grid {{ width: 100%; border-collapse: collapse; margin-bottom: 6px; }}
.info-grid td {{ vertical-align: top; padding: 2px 0; }}
.label-cyan {{ color: #00a896; font-weight: bold; }}

.map-img {{
    width: 100%;
    max-height: 180px;
    object-fit: contain;
    margin-top: 4px;
    display: block;
}}

.section-title {{ text-align: center; color: #b25900; font-size: 11pt; font-weight: bold; margin: 6px 0 4px 0; }}
.params-table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; }}
.params-table td {{ width: 50%; vertical-align: top; padding: 2px 4px; }}

.data-table {{ 
    width: 100%; 
    border-collapse: collapse; 
    margin-top: 4px; 
    margin-bottom: 4px; 
    border: 1px solid #777;
}}
.data-table th, .data-table td {{ 
    border: 1px solid #777;
    padding: 3px 4px; 
    font-size: 7.5pt; 
    text-align: left; 
}}
.data-table th {{ 
    background-color: #ffffff; 
    font-weight: bold; 
}}

.footer-note {{ 
    font-size: 6.8pt; 
    font-style: italic; 
    color: #333; 
    margin-top: 4px;
    margin-bottom: 8px; 
    text-align: left;
}}

.logos-container {{
    width: 100%;
    padding-top: 4px;
    border-top: 1px solid #aaa;
    text-align: center;
}}
.logos-container img {{
    max-height: 32px;
    width: auto;
    display: inline-block;
    vertical-align: middle;
    margin: 0 4px;
}}

@media print {{
    .logos-container {{
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
    }}
}}
</style>
</head>
<body>

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

<div class="footer-note">Ces résultats sont calculés à partir des données du Référentiel Régional Pédologique Alsace. Ils sont indicatifs et ne se substituent pas à une analyse de terre.</div>

{logos_html}

</body>
</html>"""


def create_pdf_bytes(html_content):
  if HAS_WEASYPRINT:
    return HTML(string=html_content).write_pdf()
  else:
    result = io.BytesIO()
    pisa_status = pisa.CreatePDF(html_content, dest=result)
    if pisa_status.err:
      return None
    return result.getvalue()

# ==========================================
# Interface Streamlit
# ==========================================
st.title("🌾 Générateur de Fiches Techniques Typterres - Alsace")
st.markdown(
    "Sélectionnez un identifiant pour prévisualiser les données et générer le"
    " document PDF final."
)

logos_html_block = get_base64_logos_html(LOGO_FILES)

if df is not None:
  raw_series = df[COL_ID].dropna()
  valid_series = raw_series[pd.to_numeric(raw_series, errors="coerce").notna()]
  ordered_ids = [int(x) for x in pd.unique(valid_series)]

  st.sidebar.header("🕹️ Contrôles")

  selected_id = st.sidebar.selectbox(
      "Choisissez l'identifiant Typterre :",
      options=ordered_ids,
      index=0,
      format_func=lambda x: f"Typterre n°{x}",
  )

  st.sidebar.markdown("---")
  st.sidebar.metric("Fiches disponibles", len(ordered_ids))
  if cleaned_count > 0:
    st.sidebar.caption(
        f"🧹 Purge automatique : {cleaned_count} ancien(s) PDF supprimé(s)."
    )

  col_left, col_right = st.columns([1, 1.2])

  html_payload = generate_html(df, selected_id, logos_html_block)

  with col_left:
    st.subheader(f"📄 Typterre {selected_id}")

    sub = df[df[COL_ID] == selected_id]
    nom_sol = (
        sub["NOM TYPTERRES (70)"].iloc[0]
        if pd.notna(sub["NOM TYPTERRES (70)"].iloc[0])
        else "N/A"
    )
    ref_sol = (
        sub["NOM REFERENTIEL PEDOLOGIQUE"].iloc[0]
        if pd.notna(sub["NOM REFERENTIEL PEDOLOGIQUE"].iloc[0])
        else "N/A"
    )
    nb_horizons = len(sub)

    st.info(
        f"**Nom :** {nom_sol}\n\n**Référentiel :** {ref_sol}\n\n**Horizons :**"
        f" {nb_horizons} couche(s)"
    )

    k_name, k_raw, k_hex, k_source = get_couleur_info(sub)
    st.caption(
        f"🎨 Colonne K : « {k_name} » | valeur lue : `{k_raw}` | couleur"
        f" appliquée : `{k_hex}` ({k_source})"
    )

    if st.button(
        "⚙️ Générer la fiche PDF", type="primary", use_container_width=True
    ):
      with st.spinner("Génération du PDF en cours..."):
        pdf_bytes = create_pdf_bytes(html_payload)
        filename = f"fiche_typterre_{selected_id}.pdf"

        if pdf_bytes:
          with open(filename, "wb") as f:
            f.write(pdf_bytes)

          st.success(f"Fiche générée avec succès : `{filename}`")

          st.download_button(
              label="📥 Télécharger le fichier PDF",
              data=pdf_bytes,
              file_name=filename,
              mime="application/pdf",
              use_container_width=True,
          )
        else:
          st.error("Erreur lors de la génération du PDF.")

  with col_right:
    st.subheader("Aperçu du rendu")
    if html_payload:
      st.components.v1.html(html_payload, height=680, scrolling=True)