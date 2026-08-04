import os
import glob
import asyncio
import pandas as pd
from playwright.async_api import async_playwright

# ==========================================
# Nettoyage initial des fichiers PDF existants
# ==========================================
print("Nettoyage des anciens fichiers PDF...")
pdf_files = glob.glob("*.pdf")
for pdf_file in pdf_files:
    try:
        os.remove(pdf_file)
        print(f"Supprimé : {pdf_file}")
    except Exception as e:
        print(f"Impossible de supprimer {pdf_file}: {e}")
print("Nettoyage terminé.\n")


# 1. Chargement des données Excel
excel_path = "70_Typterres_Alsace_v04_2018_publipostageREVU.xlsx"
df = pd.read_excel(excel_path)

# Nom de la colonne Q
COL_ID_TYPTERRE = 'Identifiant Typterres (1 à 70)'

# Dictionnaires de correspondance
drainage_map = {
    1.0: "Excessif",
    2.0: "Bon",
    3.0: "Modéré",
    4.0: "Imparfait",
    5.0: "Pauvre",
    6.0: "Très pauvre"
}

pierrosite_map = {
    "0": "nulle à très faible (<5%)",
    "1": "faible (5% à 15%)",
    "2": "moyenne (15% à 30%)",
    "3": "forte (30% à 50%)"
}

def generate_html_fiche(df_all, typterre_id):
    """Génère la structure HTML complète d'une fiche Typterre pour un ID donné."""
    sub_df = df_all[df_all[COL_ID_TYPTERRE] == typterre_id].sort_values('Numéro couche Typterres')
    if sub_df.empty:
        return None
    
    first = sub_df.iloc[0]
    
    # En-tête et métadonnées
    titre_typterre = f"TYPTERRE {int(first[COL_ID_TYPTERRE])}"
    nom_typterre = str(first['NOM TYPTERRES (70)']) if pd.notna(first['NOM TYPTERRES (70)']) else ""
    ref_pedo = str(first['NOM REFERENTIEL PEDOLOGIQUE']) if pd.notna(first['NOM REFERENTIEL PEDOLOGIQUE']) else ""
    petite_region = str(first['Petite Région Typterres (11)']) if pd.notna(first['Petite Région Typterres (11)']) else ""
    mat_parental = str(first['NOM MATERIAU PARENTAL']) if pd.notna(first['NOM MATERIAU PARENTAL']) else ""
    
    surface = f"{int(first['Surface Totale ha Typterres Simp'])} ha" if pd.notna(first['Surface Totale ha Typterres Simp']) else ""
    correspondance_typt = str(first['Identifiant sous TypSimplifié (70)']) if pd.notna(first['Identifiant sous TypSimplifié (70)']) else ""
    guide_sols = str(first['exemple FICHE GUIDE des sols']) if pd.notna(first['exemple FICHE GUIDE des sols']) else ""
    directive_nitrates = str(first['correspondance GREN Directive Nitrates']) if pd.notna(first['correspondance GREN Directive Nitrates']) else ""
    
    # Caractéristiques globales du sol
    ep_val, ep_min, ep_max = first["Epaisseur Sol"], first["Epaisseur Sol 'min'"], first["Epaisseur Sol 'max'"]
    epaisseur = f"{int(ep_val)} cm (min : {int(ep_min)} cm , max : {int(ep_max)} cm)" if pd.notna(ep_val) else ""
    
    pierrosite_val = str(first['Pierrosité surface']) if pd.notna(first['Pierrosité surface']) else ""
    pierrosite_txt = pierrosite_map.get(pierrosite_val, pierrosite_val)
        
    ru_val, ru_min, ru_max = first['Estimation RU du Sol (mm)'], first["Estimation RU du Sol 'min' (mm)"], first["Estimation RU du Sol 'max' (mm)"]
    ru_sol = f"{round(ru_val)} mm ( min : {round(ru_min)} mm , max : {round(ru_max)} mm)" if pd.notna(ru_val) else ""
    
    effervescence = str(first['Effervescence en clair']) if pd.notna(first['Effervescence en clair']) else ""
    drainage_txt = drainage_map.get(first['Drainage naturel'], str(first['Drainage naturel']) if pd.notna(first['Drainage naturel']) else "")
    
    # Extraction des données par horizon (jusqu'à 5 couches)
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
        
    # Code HTML + CSS
    html_content = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<style>
@page {{
    size: A4 portrait;
    margin: 10mm 15mm 10mm 15mm;
}}
*, *::before, *::after {{ box-sizing: border-box; }}
body {{
    font-family: Arial, sans-serif;
    color: #111;
    margin: 0; padding: 0;
    font-size: 9.5pt;
    line-height: 1.3;
}}
.header-top {{ text-align: right; font-weight: bold; font-size: 15pt; margin-bottom: 6px; }}
.title-box {{ background-color: #e0f7fa; border: 1.5px solid #000; padding: 8px 12px; font-size: 14pt; font-weight: bold; }}
.subtitle-box {{ background-color: #00a896; border: 1.5px solid #000; border-top: none; padding: 7px 12px; font-size: 11.5pt; font-weight: bold; color: #fff; margin-bottom: 12px; }}
.info-grid {{ width: 100%; border-collapse: collapse; margin-bottom: 12px; }}
.info-grid td {{ vertical-align: top; padding: 2px 0; }}
.label-cyan {{ color: #00a896; font-weight: bold; }}
.section-title {{ text-align: center; color: #b25900; font-size: 13pt; font-weight: bold; margin: 12px 0 10px 0; }}
.params-table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
.params-table td {{ width: 50%; vertical-align: top; padding: 3px 5px; }}
.data-table {{ width: 100%; border-collapse: collapse; margin-top: 5px; margin-bottom: 15px; }}
.data-table th, .data-table td {{ border: 1px dashed #a0a0a0; padding: 4px 6px; font-size: 8.5pt; text-align: left; }}
.data-table th {{ background-color: #ffffff; font-weight: bold; }}
.footer-note {{ font-size: 7.5pt; font-style: italic; color: #333; margin-top: 5px; }}
</style>
</head>
<body>

<div class="header-top">{titre_typterre}</div>
<div class="title-box">{nom_typterre}</div>
<div class="subtitle-box">{ref_pedo}</div>

<table class="info-grid">
    <tr>
        <td style="width: 42%;">
            <span class="label-cyan">Petite Région :</span><br>{petite_region}
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
            {"".join([f"<td>Surface<br>{h['num']}</td>" for h in horizons])}
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

<div class="footer-note">
Ces résultats sont calculés à partir des données du Référentiel Régional Pédologique Alsace. Ils sont indicatifs et ne se substituent pas à une analyse de terre.
</div>

</body>
</html>
"""
    return html_content


async def export_all_fiches():
    """Génère uniquement les PDF correspondant aux identifiants renseignés en colonne Q."""
    # Sélection des identifiants non nuls et conversion explicite en entiers
    valid_ids = df[COL_ID_TYPTERRE].dropna()
    valid_ids = valid_ids[pd.to_numeric(valid_ids, errors='coerce').notna()]
    unique_ids = sorted(valid_ids.astype(int).unique())

    print(f"Identifiants uniques détectés ({len(unique_ids)}) : {unique_ids}\n")

    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page()
        
        for t_id in unique_ids:
            try:
                html_out = generate_html_fiche(df, t_id)
                
                if html_out:
                    filename_pdf = f"fiche_typterre_{t_id}.pdf"
                    
                    await page.set_content(html_out)
                    await page.pdf(
                        path=filename_pdf,
                        format="A4",
                        print_background=True,
                        margin={"top": "10mm", "bottom": "10mm", "left": "15mm", "right": "15mm"}
                    )
                    print(f"Fiche générée : {filename_pdf}")
            except Exception as err:
                print(f"❌ Erreur lors de la génération de la fiche ID {t_id} : {err}")
                
        await browser.close()

if __name__ == "__main__":
    asyncio.run(export_all_fiches())