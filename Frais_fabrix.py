import datetime
import io
import os
import re
from PIL import Image
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Frais_fabrix - Gestion des Notes de Frais",
    page_icon="💶",
    layout="wide",
)

st.title("📄 Frais_fabrix - Suivi des Frais (Saisie Rapide & Persistante)")
st.markdown(
    "Application de gestion des notes de frais avec archivage mensuel et"
    " génération de rapport Excel."
)

# Fichiers et dossiers de persistance
FICHIER_SAUVEGARDE = "frais_sauvegarde.csv"
DOSSIER_ARCHIVES = "archives_frais"

# Création du dossier d'archives s'il n'existe pas
if not os.path.exists(DOSSIER_ARCHIVES):
  os.makedirs(DOSSIER_ARCHIVES)

# Initialisation des données (chargement depuis le fichier courant s'il existe)
if "frais_data" not in st.session_state:
  if os.path.exists(FICHIER_SAUVEGARDE):
    try:
      st.session_state["frais_data"] = pd.read_csv(FICHIER_SAUVEGARDE)
      if (
          not st.session_state["frais_data"].empty
          and "ID" in st.session_state["frais_data"].columns
      ):
        st.session_state["id_counter"] = int(
            st.session_state["frais_data"]["ID"].max()
        )
      else:
        st.session_state["id_counter"] = len(st.session_state["frais_data"])
    except Exception:
      st.session_state["frais_data"] = pd.DataFrame(columns=[
          "ID",
          "Période",
          "Nom",
          "Date",
          "Libellé",
          "Chantier",
          "Parking",
          "Entretien matériel",
          "Réception",
          "Hôtel",
          "Restauration",
          "Gasoil",
          "TVA",
          "Total",
          "Justificatif",
      ])
      st.session_state["id_counter"] = 0
  else:
    st.session_state["frais_data"] = pd.DataFrame(columns=[
        "ID",
        "Période",
        "Nom",
        "Date",
        "Libellé",
        "Chantier",
        "Parking",
        "Entretien matériel",
        "Réception",
        "Hôtel",
        "Restauration",
        "Gasoil",
        "TVA",
        "Total",
        "Justificatif",
    ])
    st.session_state["id_counter"] = 0


def sauvegarder_disque():
  """Sauvegarde automatique du DataFrame dans le fichier CSV courant"""
  st.session_state["frais_data"].to_csv(FICHIER_SAUVEGARDE, index=False)


# --- BARRE LATÉRALE : Gestion des Archives ---
st.sidebar.header("📂 Gestion des Mois")
fichiers_archives = [f for f in os.listdir(DOSSIER_ARCHIVES) if f.endswith(".csv")]

if fichiers_archives:
  st.sidebar.subheader("Recharger un ancien mois")
  archive_choisie = st.sidebar.selectbox(
      "Sélectionnez une archive", ["-- Mois en cours --"] + fichiers_archives
  )
  if archive_choisie != "-- Mois en cours --":
    if st.sidebar.button("📂 Charger cette archive"):
      chemin_archive = os.path.join(DOSSIER_ARCHIVES, archive_choisie)
      st.session_state["frais_data"] = pd.read_csv(chemin_archive)
      st.success(f"Archive '{archive_choisie}' chargée avec succès !")
      st.rerun()

# --- ÉTAPE 1 : Le Scan / Justificatif ---
st.subheader("📸 1. Joignez ou scannez votre ticket de caisse")

methode_scan = st.radio(
    "Choisissez le mode d'ajout du justificatif",
    [
        "Télécharger une image depuis l'ordinateur/téléphone",
        "Prendre une photo avec la caméra",
    ],
    horizontal=True,
)

photo_prise = None
if methode_scan == "Prendre une photo avec la caméra":
  photo_prise = st.camera_input("Prenez en photo votre ticket")
else:
  photo_prise = st.file_uploader(
      "Sélectionnez l'image de votre ticket (PNG, JPG)",
      type=["png", "jpg", "jpeg"],
  )

# Variables par défaut
libelle_auto = "Resto"
montant_auto = "0.00"
tva_auto = "0.00"
date_auto_obj = datetime.date.today()
nom_fichier = "Aucun justificatif"

# Si une image est chargée
if photo_prise is not None:
  try:
    image_pil = Image.open(photo_prise)

    # Outil interactif pour redresser/pivoter l'image si elle est de travers
    col_rot1, col_rot2 = st.columns([1, 3])
    with col_rot1:
      angle_rotation = st.selectbox(
          "Pivoter l'image", [0, 90, 180, 270], index=0
      )
    if angle_rotation > 0:
      image_pil = image_pil.rotate(-angle_rotation, expand=True)

    # Affichage de l'aperçu avec curseur de zoom dynamique
    col_img1, col_img2 = st.columns([1, 2])
    with col_img1:
      zoom_taille = st.slider(
          "🔍 Ajuster la taille / Zoom de l'aperçu",
          min_value=150,
          max_value=600,
          value=280,
          step=20,
      )
      st.image(image_pil, caption="Aperçu du ticket", width=zoom_taille)

    nom_fichier = getattr(photo_prise, "name", "ticket_camera.jpg")
    nom_lower = nom_fichier.lower()

    # Détection rapide du libellé selon le nom du fichier
    if (
        "parking" in nom_lower
        or "cofiroute" in nom_lower
        or "indigo" in nom_lower
    ):
      libelle_auto = "Parking / Péage"
    elif (
        "resto" in nom_lower
        or "restaurant" in nom_lower
        or "cafeteria" in nom_lower
    ):
      libelle_auto = "Restauration"
    elif "hotel" in nom_lower:
      libelle_auto = "Hôtel"
    elif "gasoil" in nom_lower or "essence" in nom_lower:
      libelle_auto = "Gasoil"

  except Exception as e:
    st.error(f"Erreur lors du traitement de l'image : {e}")

st.markdown("---")

# --- ÉTAPE 2 : Formulaire de validation de la dépense ---
st.subheader("✍️ 2. Saisissez ou vérifiez les informations")

with st.form("form_frais", clear_on_submit=True):
  col1, col2, col3 = st.columns(3)

  with col1:
    periode = st.text_input("Période (ex: août-26)", value="août-26")
    nom = st.text_input("Nom du salarié", value="Alain Autier")
    date_frais = st.date_input("Date de la dépense", value=date_auto_obj)

  with col2:
    libelle = st.text_input(
        "Libellé / Type de frais (ex: Parking, resto)", value=libelle_auto
    )
    chantier = st.text_input("Nom du chantier (ex: AEU, Inspire)", value="")

    cat_defaut_idx = 0
    if "Parking" in libelle or "Péage" in libelle:
      cat_defaut_idx = 0
    elif "Restauration" in libelle or "Resto" in libelle or "Cafeteria" in libelle:
      cat_defaut_idx = 1

    categorie = st.selectbox(
        "Catégorie de dépense",
        [
            "Parking / Péage",
            "Restauration",
            "Gasoil",
            "Hôtel",
            "Entretien matériel",
            "Réception",
        ],
        index=cat_defaut_idx,
    )

  with col3:
    montant_str = st.text_input("Montant (€)", value=str(montant_auto))
    tva_str = st.text_input("TVA (€)", value=str(tva_auto))

  submitted = st.form_submit_button(
      "💾 Enregistrer définitivement la dépense"
  )

  if submitted:
    try:
      montant = float(str(montant_str).replace(",", "."))
      tva = float(str(tva_str).replace(",", "."))
    except ValueError:
      montant = 0.0
      tva = 0.0

    if montant > 0:
      st.session_state["id_counter"] += 1
      ligne_id = st.session_state["id_counter"]

      parking = montant if categorie == "Parking / Péage" else 0.0
      entretien = montant if categorie == "Entretien matériel" else 0.0
      reception = montant if categorie == "Réception" else 0.0
      hotel = montant if categorie == "Hôtel" else 0.0
      restauration = montant if categorie == "Restauration" else 0.0
      gasoil = montant if categorie == "Gasoil" else 0.0

      nouvelle_ligne = {
          "ID": ligne_id,
          "Période": periode,
          "Nom": nom,
          "Date": date_frais.strftime("%d-%b"),
          "Libellé": libelle,
          "Chantier": chantier,
          "Parking": parking if parking > 0 else 0.0,
          "Entretien matériel": entretien if entretien > 0 else 0.0,
          "Réception": reception if reception > 0 else 0.0,
          "Hôtel": hotel if hotel > 0 else 0.0,
          "Restauration": restauration if restauration > 0 else 0.0,
          "Gasoil": gasoil if gasoil > 0 else 0.0,
          "TVA": tva if tva > 0 else 0.0,
          "Total": montant,
          "Justificatif": nom_fichier,
      }

      st.session_state["frais_data"] = pd.concat(
          [
              st.session_state["frais_data"],
              pd.DataFrame([nouvelle_ligne]),
          ],
          ignore_index=True,
      )

      sauvegarder_disque()

      st.success(
          "Dépense enregistrée et sauvegardée avec succès ! Vous pouvez la"
          " modifier ci-dessous si besoin."
      )
    else:
      st.error(
          "Veuillez vérifier le montant (doit être supérieur à 0) avant"
          " d'enregistrer."
      )

# --- Affichage et Édition interactive du Tableau Récapitulatif ---
st.markdown("---")
st.subheader("📊 Tableau récapitulatif des frais (Modifiable en direct)")

if not st.session_state["frais_data"].empty:
  st.markdown(
      "💡 *Astuce : Vous pouvez double-cliquer directement dans n'importe quelle"
      " case du tableau ci-dessous pour modifier une information après"
      " enregistrement ! Les modifications sont enregistrées automatiquement.*"
  )

  df_affiche = st.session_state["frais_data"].drop(columns=["ID"])
  df_modifie = st.data_editor(
      df_affiche, use_container_width=True, key="tableau_edition"
  )

  st.session_state["frais_data"].update(df_modifie)
  sauvegarder_disque()

  total_general = st.session_state["frais_data"]["Total"].sum()
  st.metric(
      label="Total Général des Dépenses", value=f"{total_general:.2f} €"
  )

  # Suppression d'une dépense
  st.markdown("### 🗑️ Supprimer une dépense")
  options_suppr = {}
  for idx, row in st.session_state["frais_data"].iterrows():
    libelle_court = (
        f"Ligne {row.get('ID', idx)} - {row['Date']} : {row['Libellé']} ("
        f"{row['Total']}€)"
    )
    options_suppr[libelle_court] = idx

  col_suppr1, col_suppr2 = st.columns([2, 1])
  with col_suppr1:
    ligne_a_supprimer = st.selectbox(
        "Sélectionnez la dépense à retirer", options=list(options_suppr.keys())
    )

  with col_suppr2:
    st.write("")
    st.write("")
    if st.button("Supprimer cette ligne"):
      index_a_retirer = options_suppr[ligne_a_supprimer]
      st.session_state["frais_data"] = st.session_state["frais_data"].drop(
          index_a_retirer
      )
      st.session_state["frais_data"] = st.session_state["frais_data"].reset_index(
          drop=True
      )
      sauvegarder_disque()
      st.success("Ligne supprimée et enregistrée !")
      st.rerun()

  # --- SECTION : Déclaration et Export Excel Professionnel ---
  st.markdown("---")
  st.subheader("📁 Déclaration officielle de la note de frais")

  col_dec1, col_dec2 = st.columns(2)

  with col_dec1:
    if st.button(
        "📄 Déclarer note de frais (Générer le fichier Excel)", type="primary"
    ):
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.title = "Note de Frais"

      font_titre = Font(name="Arial", size=14, bold=True, color="1F4E78")
      font_en_tete = Font(name="Arial", size=10, bold=True, color="FFFFFF")
      fill_en_tete = PatternFill(
          start_color="1F4E78", end_color="1F4E78", fill_type="solid"
      )
      font_donnees = Font(name="Arial", size=10)
      font_total = Font(name="Arial", size=10, bold=True)
      align_centre = Alignment(horizontal="center", vertical="center")
      align_droite = Alignment(horizontal="right", vertical="center")
      align_gauche = Alignment(horizontal="left", vertical="center")
      bordure_fine = Border(
          left=Side(style="thin", color="D9D9D9"),
          right=Side(style="thin", color="D9D9D9"),
          top=Side(style="thin", color="D9D9D9"),
          bottom=Side(style="thin", color="D9D9D9"),
      )
      bordure_total = Border(
          top=Side(style="thin", color="000000"),
          bottom=Side(style="double", color="000000"),
      )

      ws["A1"] = "NOTE DE FRAIS - RAPPORT MENSUEL"
      ws["A1"].font = font_titre
      ws.append([])

      colonnes = list(df_modifie.columns)
      ws.append(colonnes)

      for col_num in range(1, len(colonnes) + 1):
        cell = ws.cell(row=3, column=col_num)
        cell.font = font_en_tete
        cell.fill = fill_en_tete
        cell.alignment = align_centre
        cell.border = bordure_fine

      for r_idx, row in df_modifie.iterrows():
        ligne_valeurs = []
        for c_name in colonnes:
          val = row[c_name]
          if c_name in [
              "Parking",
              "Entretien matériel",
              "Réception",
              "Hôtel",
              "Restauration",
              "Gasoil",
              "TVA",
              "Total",
          ]:
            try:
              val = float(val) if pd.notna(val) and val != "" else 0.0
            except ValueError:
              val = 0.0
          ligne_valeurs.append(val)
        ws.append(ligne_valeurs)

        current_row = ws.max_row
        for col_num in range(1, len(colonnes) + 1):
          cell = ws.cell(row=current_row, column=col_num)
          cell.font = font_donnees
          cell.border = bordure_fine
          c_name = colonnes[col_num - 1]
          if c_name in [
              "Parking",
              "Entretien matériel",
              "Réception",
              "Hôtel",
              "Restauration",
              "Gasoil",
              "TVA",
              "Total",
          ]:
            cell.number_format = "#,##0.00 €"
            cell.alignment = align_droite
          else:
            cell.alignment = align_gauche

      ligne_totaux = []
      for c_name in colonnes:
        if c_name in [
            "Parking",
            "Entretien matériel",
            "Réception",
            "Hôtel",
            "Restauration",
            "Gasoil",
            "TVA",
            "Total",
        ]:
          total_col = (
              df_modifie[c_name].apply(pd.to_numeric, errors="coerce").sum()
          )
          ligne_totaux.append(total_col)
        elif c_name == "Libellé":
          ligne_totaux.append("TOTAL GÉNÉRAL")
        else:
          ligne_totaux.append("")

      ws.append(ligne_totaux)
      tot_row_idx = ws.max_row
      for col_num in range(1, len(colonnes) + 1):
        cell = ws.cell(row=tot_row_idx, column=col_num)
        cell.font = font_total
        cell.border = bordure_total
        c_name = colonnes[col_num - 1]
        if c_name in [
            "Parking",
            "Entretien matériel",
            "Réception",
            "Hôtel",
            "Restauration",
            "Gasoil",
            "TVA",
            "Total",
        ]:
          cell.number_format = "#,##0.00 €"
          cell.alignment = align_droite
        else:
          cell.alignment = align_gauche

      for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
          if cell.value is not None:
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

      output = io.BytesIO()
      wb.save(output)
      output.seek(0)

      nom_salarie = (
          str(df_modifie["Nom"].iloc[0])
          if "Nom" in df_modifie.columns and not df_modifie.empty
          else "Salarie"
      )
      periode_rapport = (
          str(df_modifie["Période"].iloc[0])
          if "Période" in df_modifie.columns and not df_modifie.empty
          else "Mois"
      )
      nom_fichier_excel = (
          f"NDF-{nom_salarie.replace(' ', '-')}-{periode_rapport}.xlsx"
      )

      st.success("🎉 Fichier Excel généré avec succès !")
      st.download_button(
          label="📥 Télécharger le fichier Excel de la note de frais",
          data=output,
          file_name=nom_fichier_excel,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

  with col_dec2:
    # Bouton d'archivage et de réinitialisation pour le mois suivant
    if st.button("🧹 Archiver et vider pour un nouveau mois"):
      if not df_modifie.empty:
        # Nom de l'archive basé sur la période et le salarié
        periode_actuelle = (
            str(df_modifie["Période"].iloc[0])
            if "Période" in df_modifie.columns
            else "archive"
        )
        nom_archive = f"frais_{periode_actuelle}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        chemin_archive = os.path.join(DOSSIER_ARCHIVES, nom_archive)

        # Sauvegarde dans le dossier d'archives
        st.session_state["frais_data"].to_csv(chemin_archive, index=False)

        # Remise à zéro du DataFrame courant et suppression du fichier de sauvegarde
        st.session_state["frais_data"] = pd.DataFrame(columns=[
            "ID",
            "Période",
            "Nom",
            "Date",
            "Libellé",
            "Chantier",
            "Parking",
            "Entretien matériel",
            "Réception",
            "Hôtel",
            "Restauration",
            "Gasoil",
            "TVA",
            "Total",
            "Justificatif",
        ])
        st.session_state["id_counter"] = 0
        if os.path.exists(FICHIER_SAUVEGARDE):
          os.remove(FICHIER_SAUVEGARDE)

        st.success(
            "📁 Le mois a été archivé avec succès ! Le tableau est prêt pour"
            " le mois prochain."
        )
        st.rerun()

else:
  st.info(
      "Aucune dépense enregistrée pour le moment. Ajoutez un ticket ci-dessus."
  )
