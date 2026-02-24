import os
import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# --- CONFIGURATION DES LISTES ---
FICHIERS = {"CARACT ENTRANT": "Suivi CARACT_ENTRANT.xlsx", "CARACT SORTANT": "CARACT_SORTANT.xlsx"}
LISTE_EQUIPES = ["MATIN", "APRES-MIDI"]
LISTE_LIEUX = ["SILO", "CABINE", "COMPACTEUR", "COLLECTE"]
LISTE_CLIENTS_ENTRANT = ["GARE MONTPARNASSE", "LE PETIT PLUS", "LA COURNEUVE", "PSM", "PSM - ADP BEAUVAIS", "LBM", "AEROVILLE", "CDG", "HAUSSMAN", "ORLY", "CEMEX", "LE BOURGET & LBM", "ROISSY", "DISNEYLAND", "VALLEE VILLAGE & 4 TEMPS", "GARE DE LYON", "SNCF", "VALODEA", "VALOR'AISME", "SITRU", "SYCTOM", "AUTRES"]
LISTE_FLUX_SORTANT = ["JRM", "EMR", "CARTON", "GM", "PETQ9", "FLUX DEV", "PETB", "PEPP", "ELA", "FILM", "ACIER", "ALU", "PETIT ALU", "REFUS"]

COLORS = {"BLUE": "#0070c0", "LIGHT_BLUE": "#00b0f0", "GREEN": "#00b050", "YELLOW": "#ffc000", "ORANGE": "#ed7d31", "RED": "#c00000"}

GROUPES_ENTRANT = {
    "FIBREUX": {"color": COLORS["BLUE"], "items": ["CARTON", "CARTONNETTE"]},
    "FIBREUX 2": {"color": COLORS["LIGHT_BLUE"], "items": ["ECRIT COULEUR", "JRM", "GM"]},
    "ELA": {"color": COLORS["GREEN"], "items": ["ELA"]},
    "PLASTIQUES": {"color": COLORS["YELLOW"], "items": ["PEPP", "PET Q9", "PET Q5", "PET B", "FILM"]},
    "METAUX": {"color": COLORS["ORANGE"], "items": ["ACIER", "ALU", "PETIT ALU"]},
    "REFUS / AUTRES": {"color": COLORS["RED"], "items": ["PB REFUSES", "FILM REFUSES", "EMBALLAGES NOIRS", "AUTRES EMBALLAGES NON RECYCLABLES", "BOIS", "VERRE", "DDS", "D3E", "IMBRIQUES", "PLASTIQUES NON EMBALLAGES", "FERRAILLES", "GRAVATS", "TEXTILE", "EMBALLAGES NON VIDES / SOUILLÉS", "REFUS", "PAPIER MOUILLE", "FINES"]}
}

GROUPES_SORTANT = {
    "FIBREUX": {"color": COLORS["BLUE"], "items": ["JRM", "EMR", "CARTON", "GM"]},
    "PLASTIQUES": {"color": COLORS["YELLOW"], "items": ["PETQ9", "FLUX DEV", "PET B", "PEPP", "ELA", "FILM"]},
    "METAUX": {"color": COLORS["ORANGE"], "items": ["ACIER", "ALU", "PETIT ALU"]},
    "AUTRES": {"color": COLORS["RED"], "items": ["REFUS"]}
}

# --- LOGIQUE D'ENREGISTREMENT EXCEL ---
def enregistrer_donnees(mode, header, dict_poids):
    nom_f = FICHIERS[mode]
    if not os.path.exists(nom_f):
        st.error(f"Fichier {nom_f} introuvable.")
        return False
    try:
        wb = load_workbook(nom_f)
        ws = wb["SAISIE"]
        row = 2
        while ws.cell(row=row, column=2).value: row += 1
        
        groupes = GROUPES_ENTRANT if mode == "CARACT ENTRANT" else GROUPES_SORTANT
        liste_ordonnee = []
        for g in groupes.values(): liste_ordonnee.extend(g["items"])
        
        poids_finaux = [float(str(dict_poids.get(m, 0)).replace(',', '.')) for m in liste_ordonnee]
        
        if mode == "CARACT SORTANT":
            ligne = [header['flux'], header['date'], header['equipe'], header['lieu']] + poids_finaux
        else:
            ligne = [header['flux'], header['date']] + poids_finaux
            
        for i, v in enumerate(ligne):
            cell = ws.cell(row=row, column=i+2, value=v)
            cell.alignment = Alignment(horizontal="center")
            
        wb.save(nom_f)
        return True
    except PermissionError:
        st.error("❌ Erreur : Le fichier Excel est ouvert. Fermez-le avant d'enregistrer.")
        return False
    except Exception as e:
        st.error(f"❌ Erreur Excel : {e}")
        return False

# --- INTERFACE ---
st.set_page_config(page_title="PAPREC - Caractérisation", layout="wide")

if 'mode' not in st.session_state:
    st.session_state.mode = None
if 'session_time' not in st.session_state:
    st.session_state.session_time = datetime.now().strftime("%Hh%M")

# --- ECRAN D'ACCUEIL ---
if st.session_state.mode is None:
    st.title("🏭 Système de Caractérisation")
    c1, c2 = st.columns(2)
    if c1.button("📥 CARACT ENTRANT", use_container_width=True):
        st.session_state.mode = "CARACT ENTRANT"
        st.rerun()
    if c2.button("📤 CARACT SORTANT", use_container_width=True):
        st.session_state.mode = "CARACT SORTANT"
        st.rerun()

# --- ECRAN DE SAISIE ---
else:
    st.button("⬅ Retour", on_click=lambda: setattr(st.session_state, 'mode', None))
    st.header(f"Saisie : {st.session_state.mode}")

    with st.container(border=True):
        col1, col2, col3, col4 = st.columns(4)
        # DATE AU FORMAT JJ-MM-AAAA
        date_saisie = col1.text_input("Date:", datetime.now().strftime("%d-%m-%Y"))
        
        lbl = "Client:" if st.session_state.mode == "CARACT ENTRANT" else "Flux:"
        lst = LISTE_CLIENTS_ENTRANT if st.session_state.mode == "CARACT ENTRANT" else LISTE_FLUX_SORTANT
        flux_sel = col2.selectbox(lbl, lst)
        equipe_sel = col3.selectbox("Équipe:", LISTE_EQUIPES)
        lieu_sel = col4.selectbox("Lieu:", LISTE_LIEUX)

    # Dossier de session pour les photos
    date_folder = datetime.now().strftime("%d-%m-%Y")
    session_id = f"{st.session_state.session_time}_{flux_sel.replace(' ', '_')}"
    path_session = f"PHOTOS_CARACT/{date_folder}/{session_id}"

    groupes = GROUPES_ENTRANT if st.session_state.mode == "CARACT ENTRANT" else GROUPES_SORTANT
    dict_entrees = {}
    
    for g_name, info in groupes.items():
        st.markdown(f"<div style='background-color:{info['color']}; color:white; padding:10px; border-radius:5px; margin-top:20px; margin-bottom:10px;'><b>{g_name}</b></div>", unsafe_allow_html=True)
        
        # Affichage en colonnes pour plus de clarté
        items = info["items"]
        for i in range(0, len(items), 2):
            cols = st.columns(2)
            for j in range(2):
                if i + j < len(items):
                    matiere = items[i+j]
                    with cols[j].container(border=True):
                        c_poids, c_photo = st.columns([1, 1])
                        dict_entrees[matiere] = c_poids.number_input(f"{matiere} (kg)", min_value=0.0, step=0.1, key=f"p_{matiere}")
                        
                        with c_photo.popover("📸 Photo"):
                            img = st.camera_input(f"Caméra {matiere}", key=f"cam_{matiere}")
                            if img:
                                if st.button(f"Sauvegarder {matiere}", key=f"btn_save_{matiere}"):
                                    os.makedirs(path_session, exist_ok=True)
                                    nom_img = f"{matiere}_{datetime.now().strftime('%H%M%S')}.jpg"
                                    with open(os.path.join(path_session, nom_img), "wb") as f:
                                        f.write(img.getbuffer())
                                    st.toast(f"Photo {matiere} enregistrée !", icon="✅")

    st.markdown("---")
    if st.button("💾 ENREGISTRER DANS EXCEL", type="primary", use_container_width=True):
        h = {'date': date_saisie, 'flux': flux_sel, 'equipe': equipe_sel, 'lieu': lieu_sel}
        if enregistrer_donnees(st.session_state.mode, h, dict_entrees):
            st.success(f"Données enregistrées ! Dossier : {session_id}")
            st.balloons()
            st.session_state.session_time = datetime.now().strftime("%Hh%M")
