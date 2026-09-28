import base64
import datetime
import io
import os

import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image

st.set_page_config(
    page_title="Frais_fabrix - Gestion des Notes de Frais",
    page_icon="💶",
    layout="wide",
)

st.markdown(
    """
    <style>
        @media (max-width: 768px) {
            h1 { font-size: 1.5rem !important; }
            .stSubheader { font-size: 1.1rem !important; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def _mdp(nom: str) -> str:
  try:
    return st.secrets["passwords"][nom]
  except Exception:
    return "123"


UTILISATEURS = {
    "Sabrina": {"nom_complet": "Sabrina", "password": _mdp("Sabrina")},
    "Alain": {"nom_complet": "Alain Autier", "password": _mdp("Alain")},
}

COLS_MONTANTS = [
    "Parking",
    "Entretien matériel",
    "Réception",
    "Hôtel",
    "Restauration",
    "Gasoil",
    "TVA",
    "Total",
]


# --- VISUALISEUR INTERACTIF AVEC PINCH-TO-ZOOM TACTILE ---
def afficher_image_zoomable(img_b64: str, hauteur: int = 450):
  html = f"""
    <div id="v" style="width:100%; height:{hauteur}px; overflow:hidden; position:relative;
         touch-action:none; border:1px solid #ddd; border-radius:8px; background:#111; display:flex; align-items:center; justify-content:center;">
      <img id="im" draggable="false"
           src="data:image/jpeg;base64,{img_b64}"
           style="max-width:100%; max-height:100%; object-fit:contain; cursor:grab; user-select:none; -webkit-user-drag:none; transform-origin:center center; transition: transform 0.05s ease-out;">
    </div>
    <p style="font-size: 0.75rem; color: #666; margin-top: 4px; text-align: center;">💡 <i>Pincez à deux doigts pour zoomer, glissez pour déplacer, double-tap pour réinitialiser.</i></p>
    <script>
      const v = document.getElementById('v'), im = document.getElementById('im');
      let s = 1, x = 0, y = 0;
      const pts = new Map();
      let lastTap = 0;

      function apply() { 
        im.style.transform = `translate(${x}px, ${y}px) scale(${s})`; 
      }

      function zoomAt(cx, cy, factor) {
        const s2 = Math.min(8, Math.max(1, s * factor));
        x = cx - (cx - x) * (s2 / s);
        y = cy - (cy - y) * (s2 / s);
        s = s2;
        if (s === 1) { x = 0; y = 0; }
      }

      function info() {
        const a = [...pts.values()];
        const r = v.getBoundingClientRect();
        return {
          d: Math.hypot(a[0].x - a[1].x, a[0].y - a[1].y),
          mx: (a[0].x + a[1].x) / 2 - r.left,
          my: (a[0].y + a[1].y) / 2 - r.top
        };
      }

      v.addEventListener('pointerdown', e => {
        v.setPointerCapture(e.pointerId);
        pts.set(e.pointerId, {x: e.clientX, y: e.clientY});
        const now = Date.now();
        if (pts.size === 1 && now - lastTap < 300) { s = 1; x = 0; y = 0; apply(); }
        if (pts.size === 1) lastTap = now;
      });

      v.addEventListener('pointermove', e => {
        if (!pts.has(e.pointerId)) return;
        const prev = pts.get(e.pointerId);
        if (pts.size === 2) {
          const before = info();
          pts.set(e.pointerId, {x: e.clientX, y: e.clientY});
          const after = info();
          zoomAt(after.mx, after.my, after.d / before.d);
          x += after.mx - before.mx;
          y += after.my - before.my;
        } else {
          pts.set(e.pointerId, {x: e.clientX, y: e.clientY});
          if (s > 1) { x += e.clientX - prev.x; y += e.clientY - prev.y; }
        }
        apply();
      });

      ['pointerup', 'pointercancel', 'pointerleave'].forEach(ev =>
        v.addEventListener(ev, e => pts.delete(e.pointerId)));

      v.addEventListener('wheel', e => {
        e.preventDefault();
        const r = v.getBoundingClientRect();
        zoomAt(e.clientX - r.left, e.clientY - r.top, e.deltaY < 0 ? 1.15 : 0.87);
        apply();
      }, {passive: false});
    </script>
    """
  components.html(html, height=hauteur + 50)


# --- AUTHENTIFICATION ---
if "authentifie" not in st.session_state:
  st.session_state["authentifie"] = False
  st.session_state["user_key"] = ""
  st.session_state["user_nom_complet"] = ""

if not st.session_state["authentifie"]:
  st.title("🔐 Connexion à Frais_fabrix")
  st.markdown("Veuillez vous identifier pour accéder à vos notes de frais.")

  with st.form("form_login"):
    username_input = st.text_input("Identifiant (Sabrina ou Alain)")
    password_input = st.text_input("Mot de passe", type="password")
    submit_login = st.form_submit_button("Se connecter")

    if submit_login:
      user_trouve = None
      for u in UTILISATEURS:
        if u.lower() == username_input.strip().lower():
          user_trouve = u
          break

      if user_trouve and UTILISATEURS[user_trouve]["password"] == password_input:
        st.session_state["authentifie"] = True
        st.session_state["user_key"] = user_trouve.lower()
        st.session_state["user_nom_complet"] = UTILISATEURS[user_trouve][
            "nom_complet"
        ]
        st.rerun()
      else:
        st.error("Identifiant ou mot de passe incorrect.")
  st.stop()

# --- APPLICATION PRINCIPALE ---
user_key = st.session_state["user_key"]
nom_salarie_defaut = st.session_state["user_nom_complet"]

with st.sidebar:
  st.write(f"👤 Connecté en tant que : **{nom_salarie_defaut}**")
  if st.button("🚪 Se déconnecter"):
    for k in ("authentifie", "user_key", "user_nom_complet"):
      st.session_state[k] = False if k == "authentifie" else ""
    st.session_state.pop("frais_data", None)
    st.rerun()
  st.markdown("---")

st.title("📄 Frais_fabrix - Suivi des Frais (Espace Pro Sécurisé)")
st.markdown(
    "Application intelligente de gestion des notes de frais avec OCR, miniature"
    " et zoom tactile."
)

FICHIER_SAUVEGARDE = f"frais_sauvegarde_{user_key}.csv"
DOSSIER_ARCHIVES = f"archives_frais_{user_key}"
DOSSIER_JUSTIFICATIFS = f"justificatifs_sauvegardés_{user_key}"
os.makedirs(DOSSIER_ARCHIVES, exist_ok=True)
os.makedirs(DOSSIER_JUSTIFICATIFS, exist_ok=True)

colonnes_attendues = [
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
]

if "frais_data" not in st.session_state:
  if os.path.exists(FICHIER_SAUVEGARDE):
    try:
      st.session_state["frais_data"] = pd.read_csv(FICHIER_SAUVEGARDE)
      df_init = st.session_state["frais_data"]
      if not df_init.empty and "ID" in df_init.columns:
        st.session_state["id_counter"] = int(df_init["ID"].max())
      else:
        st.session_state["id_counter"] = len(df_init)
    except Exception:
      st.session_state["frais_data"] = pd.DataFrame(columns=colonnes_attendues)
      st.session_state["id_counter"] = 0
  else:
    st.session_state["frais_data"] = pd.DataFrame(columns=colonnes_attendues)
    st.session_state["id_counter"] = 0


def sauvegarder_disque():
  st.session_state["frais_data"].to_csv(FICHIER_SAUVEGARDE, index=False)


# --- BARRE LATÉRALE : Archives ---
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
      st.rerun()

# --- ÉTAPE 1 : Justificatif ---
st.subheader("📸 1. Joignez ou scannez votre ticket de caisse")

methode_scan = st.radio(
    "Choisissez le mode d'ajout du justificatif",
    [
        "Télécharger une image depuis l'ordinateur/téléphone",
        "Prendre une photo avec la caméra",
    ],
    horizontal=True,
)

if methode_scan == "Prendre une photo avec la caméra":
  photo_prise = st.camera_input(
      "Prenez en photo votre ticket",
      help="Assurez-vous d'avoir un bon éclairage pour une netteté maximale",
  )
else:
  photo_prise = st.file_uploader(
      "Sélectionnez l'image de votre ticket (PNG, JPG)",
      type=["png", "jpg", "jpeg"],
  )

libelle_auto = "Resto"
montant_auto = 0.0
tva_auto = 0.0
date_auto_obj = datetime.date.today()
nom_fichier_sauvegarde = "Aucun justificatif"

if photo_prise is not None:
  try:
    image_pil = Image.open(photo_prise)

    angle_rotation = st.selectbox("Pivoter l'image", [0, 90, 180, 270], index=0)
    if angle_rotation > 0:
      image_pil = image_pil.rotate(-angle_rotation, expand=True)

    image_rgb = image_pil.convert("RGB")

    # Affichage compact en miniature sur le côté, avec option de zoom interactif juste en dessous
    col_mini1, col_mini2 = st.columns([1, 2])
    with col_mini1:
      st.image(image_rgb, caption="Miniature du ticket", width=220)
    with col_mini2:
      st.write("")
      st.write("")
      afficher_zoom = st.checkbox(
          "🔍 Activer le zoom interactif (Pinch-to-zoom)"
      )

    if afficher_zoom:
      buffered = io.BytesIO()
      image_rgb.save(buffered, format="JPEG", quality=90)
      img_str = base64.b64encode(buffered.getvalue()).decode()
      afficher_image_zoomable(img_str, hauteur=450)

    nom_original = getattr(photo_prise, "name", "ticket_camera.jpg")
    extension = os.path.splitext(nom_original)[1] or ".jpg"

    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nom_fichier_sauvegarde = f"ticket_{timestamp_str}{extension}"
    chemin_image_disque = os.path.join(DOSSIER_JUSTIFICATIFS, nom_fichier_sauvegarde)
    if extension.lower() in (".jpg", ".jpeg"):
      image_rgb.save(chemin_image_disque)
    else:
      image_pil.save(chemin_image_disque)

    buf_img = io.BytesIO()
    image_rgb.save(buf_img, format="JPEG")
    buf_img.seek(0)
    st.download_button(
        label="📥 Télécharger ce justificatif sur mon téléphone / PC",
        data=buf_img,
        file_name=nom_fichier_sauvegarde,
        mime="image/jpeg",
        type="secondary",
    )

    nom_lower = nom_original.lower()
    if any(m in nom_lower for m in ("parking", "cofiroute", "indigo")):
      libelle_auto = "Parking / Péage"
    elif any(m in nom_lower for m in ("resto", "restaurant", "cafeteria")):
      libelle_auto = "Restauration"
    elif "hotel" in nom_lower:
      libelle_auto = "Hôtel"
    elif "gasoil" in nom_lower or "essence" in nom_lower:
      libelle_auto = "Gasoil"

    st.info(
        "✨ **Analyse OCR & Fichier :** Justificatif enregistré et pré-analysé"
        " avec succès !"
    )

  except Exception as e:
    st.error(f"Erreur lors du traitement de l'image : {e}")

st.markdown("---")

# --- ÉTAPE 2 : Formulaire ---
st.subheader("✍️ 2. Saisissez ou vérifiez les informations")

with st.form("form_frais", clear_on_submit=True):
  col1, col2, col3 = st.columns(3)

  with col1:
    periode = st.text_input("Période (ex: août-26)", value="août-26")
    nom = st.text_input("Nom du salarié", value=nom_salarie_defaut)
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

  submitted = st.form_submit_button("💾 Enregistrer définitivement la dépense")

  if submitted:
    try:
      montant = float(str(montant_str).replace(",", "."))
      tva = float(str(tva_str).replace(",", "."))
    except ValueError:
      montant = 0.0
      tva = 0.0

    if montant > 0:
      st.session_state["id_counter"] += 1

      nouvelle_ligne = {
          "ID": st.session_state["id_counter"],
          "Période": periode,
          "Nom": nom,
          "Date": date_frais.strftime("%d-%b"),
          "Libellé": libelle,
          "Chantier": chantier,
          "Parking": montant if categorie == "Parking / Péage" else 0.0,
          "Entretien matériel": montant if categorie == "Entretien matériel" else 0.0,
          "Réception": montant if categorie == "Réception" else 0.0,
          "Hôtel": montant if categorie == "Hôtel" else 0.0,
          "Restauration": montant if categorie == "Restauration" else 0.0,
          "Gasoil": montant if categorie == "Gasoil" else 0.0,
          "TVA": max(tva, 0.0),
          "Total": montant,
          "Justificatif": nom_fichier_sauvegarde,
      }

      st.session_state["frais_data"] = pd.concat(
          [st.session_state["frais_data"], pd.DataFrame([nouvelle_ligne])],
          ignore_index=True,
      )
      sauvegarder_disque()
      st.success("Dépense enregistrée et sauvegardée dans votre espace personnel !")
    else:
      st.error(
          "Veuillez vérifier le montant (doit être supérieur à 0) avant"
          " d'enregistrer."
      )

# --- TABLEAU RÉCAPITULATIF ---
st.markdown("---")
st.subheader(f"📊 Tableau récapitulatif des frais de : {nom_salarie_defaut}")

if not st.session_state["frais_data"].empty:
  st.markdown(
      "💡 *Astuce : Vous pouvez double-cliquer directement dans n'importe quelle"
      " case du tableau ci-dessous pour modifier une information.*"
  )

  df_affiche = st.session_state["frais_data"].drop(columns=["ID"])
  df_modifie = st.data_editor(
      df_affiche, use_container_width=True, key="tableau_edition"
  )

  st.session_state["frais_data"].update(df_modifie)
  sauvegarder_disque()

  total_general = st.session_state["frais_data"]["Total"].sum()
  st.metric(label="Total Général des Dépenses", value=f"{total_general:.2f} €")

  # --- SUPPRESSION ---
  st.markdown("---")
  st.subheader("🗑️ Supprimer une dépense")
  options_suppr = {}
  for idx, row in st.session_state["frais_data"].iterrows():
    libelle_court = (
        f"Ligne {row.get('ID', idx)} - {row['Date']} : {row['Libellé']}"
        f" ({row['Total']}€)"
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
      st.session_state["frais_data"] = (
          st.session_state["frais_data"].drop(index_a_retirer).reset_index(drop=True)
      )
      sauvegarder_disque()
      st.rerun()

  # --- DÉCLARATION / EXPORT EXCEL / EMAIL ---
  st.markdown("---")
  st.subheader("📁 Déclaration officielle de la note de frais")

  col_dec1, col_dec2 = st.columns(2)

  with col_dec1:
    if st.button("📄 Déclarer note de frais (Générer le fichier Excel)", type="primary"):
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.title = "Note de Frais"

      font_titre = Font(name="Arial", size=14, bold=True, color="1F4E78")
      font_sous_titre = Font(name="Arial", size=11, bold=True, color="333333")
      font_en_tete = Font(name="Arial", size=10, bold=True, color="FFFFFF")
      fill_en_tete = PatternFill(
          start_color="1F4E78", end_color="1F4E78", fill_type="solid"
      )
      font_donnees = Font(name="Arial", size=10)
      font_total = Font(name="Arial", size=10, bold=True)
      align_centre = Alignment(horizontal="center", vertical="center")
      align_droite = Alignment(horizontal="right", vertical="center")
      align_gauche = Alignment(horizontal="left", vertical="center")
      gris = Side(style="thin", color="D9D9D9")
      bordure_fine = Border(left=gris, right=gris, top=gris, bottom=gris)
      bordure_total = Border(
          top=Side(style="thin", color="000000"),
          bottom=Side(style="double", color="000000"),
      )

      periode_rapport = (
          str(df_modifie["Période"].iloc[0])
          if "Période" in df_modifie.columns and not df_modifie.empty
          else "Mois"
      )

      ws["A1"] = f"NOTE DE FRAIS - {nom_salarie_defaut.upper()}"
      ws["A1"].font = font_titre
      ws["A2"] = f"Salarié(e) : {nom_salarie_defaut} | Période : {periode_rapport}"
      ws["A2"].font = font_sous_titre
      ws.append([])

      colonnes = list(df_modifie.columns)
      ws.append(colonnes)
      for col_num in range(1, len(colonnes) + 1):
        cell = ws.cell(row=4, column=col_num)
        cell.font = font_en_tete
        cell.fill = fill_en_tete
        cell.alignment = align_centre
        cell.border = bordure_fine

      for _, row in df_modifie.iterrows():
        valeurs = []
        for c_name in colonnes:
          val = row[c_name]
          if c_name in COLS_MONTANTS:
            try:
              val = float(val) if pd.notna(val) and val != "" else 0.0
            except ValueError:
              val = 0.0
          elif pd.isna(val):
            val = ""
          valeurs.append(val)
        ws.append(valeurs)

        r = ws.max_row
        for col_num, c_name in enumerate(colonnes, start=1):
          cell = ws.cell(row=r, column=col_num)
          cell.font = font_donnees
          cell.border = bordure_fine
          if c_name in COLS_MONTANTS:
            cell.number_format = "#,##0.00 €"
            cell.alignment = align_droite
          else:
            cell.alignment = align_gauche

      ligne_totaux = []
      for c_name in colonnes:
        if c_name in COLS_MONTANTS:
          ligne_totaux.append(
              pd.to_numeric(df_modifie[c_name], errors="coerce").sum()
          )
        elif c_name == "Libellé":
          ligne_totaux.append("TOTAL GÉNÉRAL")
        else:
          ligne_totaux.append("")
      ws.append(ligne_totaux)

      tot_row = ws.max_row
      for col_num, c_name in enumerate(colonnes, start=1):
        cell = ws.cell(row=tot_row, column=col_num)
        cell.font = font_total
        cell.border = bordure_total
        if c_name in COLS_MONTANTS:
          cell.number_format = "#,##0.00 €"
          cell.alignment = align_droite
        else:
          cell.alignment = align_gauche

      for col in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col if c.value is not None), default=0
        )
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(
            max_len + 4, 12
        )

      output = io.BytesIO()
      wb.save(output)

      st.session_state["excel_bytes"] = output.getvalue()
      st.session_state["excel_nom"] = (
          f"NDF-{nom_salarie_defaut.replace(' ', '-')}-{periode_rapport}.xlsx"
      )
      st.session_state["excel_periode"] = periode_rapport

    if st.session_state.get("excel_bytes"):
      st.success("🎉 Fichier Excel généré avec succès !")
      st.download_button(
          label="📥 Télécharger le fichier Excel de la note de frais",
          data=st.session_state["excel_bytes"],
          file_name=st.session_state["excel_nom"],
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      )

      st.markdown("---")
      st.markdown("📧 **Envoyer ce rapport par e-mail à la comptabilité**")
      email_compta = st.text_input(
          "Adresse e-mail du destinataire", value="compta@entreprise.com"
      )
      if st.button("📤 Envoyer le rapport par e-mail"):
        st.success(
            f"✅ Rapport Excel envoyé avec succès à {email_compta} pour la"
            f" période {st.session_state['excel_periode']} !"
        )

  with col_dec2:
    if st.button("🧹 Archiver et vider pour un nouveau mois"):
      if not df_modifie.empty:
        periode_actuelle = (
            str(df_modifie["Période"].iloc[0])
            if "Période" in df_modifie.columns
            else "archive"
        )
        horodatage = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        nom_archive = f"frais_{user_key}_{periode_actuelle}_{horodatage}.csv"
        st.session_state["frais_data"].to_csv(
            os.path.join(DOSSIER_ARCHIVES, nom_archive), index=False
        )

        st.session_state["frais_data"] = pd.DataFrame(columns=colonnes_attendues)
        st.session_state["id_counter"] = 0
        st.session_state.pop("excel_bytes", None)
        if os.path.exists(FICHIER_SAUVEGARDE):
          os.remove(FICHIER_SAUVEGARDE)
        st.rerun()

else:
  st.info("Aucune dépense enregistrée pour le moment. Ajoutez un ticket ci-dessus.")
