#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ──────────────────────────────────────────────────────────────────────────────
# Streamlit App: Vues Feedback (Moyennes + Vues filtrées <3/5)
# ──────────────────────────────────────────────────────────────────────────────
# - Upload un Excel (une ou plusieurs feuilles).
# - Détecte automatiquement les paires (Sur une échelle de 0 à 5 … -> Commentaire)
#   où le commentaire est la colonne immédiatement suivante.
# - Calcule la vue "Moyennes" (moyenne sur 5) par catégorie.
# - Construit 5 vues (<3/5) : Coaching, Fiches de cours, Professeurs,
#   Plateforme, Organisation générale.
# - Permet de télécharger un Excel qui contient Moyennes + ces 5 vues.
# ──────────────────────────────────────────────────────────────────────────────

import io
import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Vues Feedback – Diploma Santé", layout="wide")

# Catégories cibles (clés normalisées → libellés affichés)
TARGET_VIEWS = [
    ("coaching", "Coaching"),
    ("fichesdecours", "Fiches de cours"),
    ("fiches cours", "Fiches de cours"),
    ("professeurs", "Professeurs"),
    ("plateforme", "Plateforme"),
    ("organisationgenerale", "Organisation générale"),
    ("organisation generale", "Organisation générale"),
]

REQUIRED_SHEETS = [
    "Moyennes",
    "Coaching",
    "Fiches de cours",
    "Professeurs",
    "Plateforme",
    "Organisation générale",
]

SCALE_PREFIX = "sur une echelle de 0 a 5"  # version normalisée (sans accents)

# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s

def parse_note(val):
    """
    Convertit la note en float sur 5.
    Accepte :
      - '2/5', '4 / 5'
      - '2,5', '3', '4.0'
      - '4 - Plutôt satisfait', '5 – Très bien', etc. (extrait le 1er nombre)
    """
    if pd.isna(val):
        return None
    s = str(val).strip().replace(",", ".")
    # Forme fractionnaire n/den
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)\s*$", s)
    if m:
        num = float(m.group(1))
        den = float(m.group(2))
        return (num / den) * 5.0 if den else None
    # Nombre simple
    try:
        return float(s)
    except ValueError:
        # Extraire le premier nombre rencontré
        m2 = re.search(r"(\d+(?:\.\d+)?)", s)
        if m2:
            try:
                return float(m2.group(1))
            except Exception:
                return None
        return None

def _try_read_excel_all_sheets(bio: io.BytesIO):
    """
    Essaie plusieurs moteurs de lecture pour éviter l'ImportError d'openpyxl.
    Ordre d'essai : openpyxl (xlsx), xlrd (xls legacy), calamine (xlsx/xls), auto.
    """
    engines = ["openpyxl", "xlrd", "calamine", None]
    last_err = None
    for eng in engines:
        try:
            bio.seek(0)
            if eng is None:
                return pd.read_excel(bio, sheet_name=None)
            return pd.read_excel(bio, sheet_name=None, engine=eng)
        except Exception as e:
            last_err = e
            continue
    # Si aucun moteur ne fonctionne, on remonte l'erreur
    raise last_err if last_err else RuntimeError("Lecture Excel impossible.")

def read_all_sheets(uploaded_file) -> pd.DataFrame:
    """
    Lit toutes les feuilles d’un Excel uploadé et les concatène.
    Robuste aux environnements sans openpyxl : essaie plusieurs moteurs.
    """
    bytes_data = uploaded_file.read()
    if not bytes_data:
        return pd.DataFrame()
    bio = io.BytesIO(bytes_data)
    try:
        sheets = _try_read_excel_all_sheets(bio)
    except Exception as e:
        st.error(
            "Impossible de lire le fichier Excel. "
            "Installez au moins l’un des moteurs suivants dans votre environnement : "
            "`openpyxl` (xlsx), `pandas-calamine` (xlsx/xls) ou `xlrd<=1.2.0` (xls). "
            f"\n\nDétails techniques : {e}"
        )
        return pd.DataFrame()

    frames = []
    for sheet_name, df in sheets.items():
        if df is not None and not df.empty:
            df["_source_sheet"] = str(sheet_name)
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

def find_identity_columns(df: pd.DataFrame):
    cols = {normalize(c): c for c in df.columns}
    first_name = next((cols[k] for k in cols if any(w in k for w in ["prenom", "prénom", "first name", "given name"])), None)
    last_name  = next((cols[k] for k in cols if any(w in k for w in ["nom", "last name", "surname", "family name"]) and "prenom" not in k and "prénom" not in k), None)
    email      = next((cols[k] for k in cols if any(w in k for w in ["email", "e-mail", "mail", "adresse email", "adresse e mail"])), None)
    if first_name is None:
        first_name = next((cols[k] for k in cols if "prenom" in k or "prénom" in k), None)
    if last_name is None:
        last_name = next((cols[k] for k in cols if k.startswith("nom")), None)
    if email is None:
        email = next((cols[k] for k in cols if "adresse" in k and "mail" in k), None)
    return first_name, last_name, email

def _extract_category_from_scale_header(norm_header: str) -> str:
    """
    Reçoit un en-tête déjà normalisé.
    Si l’en-tête commence par 'sur une echelle de 0 a 5', on renvoie la partie
    après ce préfixe (nettoyée).
    """
    if not norm_header.startswith(SCALE_PREFIX):
        return ""
    tail = norm_header[len(SCALE_PREFIX):].strip(" :–—-")
    # Supprimer ponctuation résiduelle, espaces, etc.
    tail = re.sub(r"\s+", " ", tail)
    return tail

def build_pairs(df: pd.DataFrame):
    """
    Détecte les colonnes où une question commence par 'Sur une échelle de 0 à 5 ...'
    et où la colonne immédiatement suivante est le commentaire.
    Retourne: dict { display_name: (scale_col, comment_col) }
    """
    columns = list(df.columns)
    norm = [normalize(c) for c in columns]
    target_keys = {k for k, _ in TARGET_VIEWS}
    display_map = {k: disp for k, disp in TARGET_VIEWS}
    pairs = {}

    for i, ncol in enumerate(norm):
        # 1) Identifier les colonnes qui débutent par "Sur une échelle de 0 à 5"
        if not ncol.startswith(SCALE_PREFIX):
            continue
        # 2) La colonne suivante est le commentaire
        if i + 1 >= len(columns):
            continue
        next_norm = norm[i + 1]
        if not any(x in next_norm for x in ["comment", "commentaire", "remarque", "avis"]):
            # Si la colonne suivante n'est pas un commentaire, ignorer cette paire
            continue

        # 3) Extraire une clé catégorie depuis la "queue" de l'en-tête
        cat_tail = _extract_category_from_scale_header(ncol)
        cat_key_simple = re.sub(r"[^a-z0-9 ]", "", cat_tail)
        cat_key_simple = re.sub(r"\s+", "", cat_key_simple)

        # 4) Faire correspondre aux vues cibles
        match_key = None
        for tk in target_keys:
            if tk in cat_key_simple or cat_key_simple in tk:
                match_key = tk
                break
        if match_key is None:
            # Correspondance plus permissive : partage de mots
            for tk in target_keys:
                words = re.findall(r"[a-z]+", tk)
                if any(w in cat_key_simple for w in words):
                    match_key = tk
                    break

        if match_key:
            disp = display_map[match_key]
            pairs[disp] = (columns[i], columns[i + 1])

    return pairs

def compute_averages(df: pd.DataFrame, pairs: dict) -> pd.DataFrame:
    rows = []
    for view, (scale_col, _) in pairs.items():
        series = df[scale_col].map(parse_note)
        series = series.dropna()
        if not series.empty:
            rows.append({"Catégorie": view, "Moyenne (/5)": round(float(series.mean()), 2)})
    return pd.DataFrame(rows).sort_values("Catégorie").reset_index(drop=True) if rows else pd.DataFrame(columns=["Catégorie", "Moyenne (/5)"])

def build_views(df: pd.DataFrame, prenom_col: str, nom_col: str, email_col: str, pairs: dict):
    sheets = {}
    for display, (scale_col, comm_col) in pairs.items():
        cols = [c for c in [prenom_col, nom_col, email_col, scale_col, comm_col] if c and c in df.columns]
        if not cols:
            continue
        temp = df[cols].copy()
        rename_map = {}
        if prenom_col in temp.columns: rename_map[prenom_col] = "Prénom"
        if nom_col in temp.columns:    rename_map[nom_col]    = "Nom"
        if email_col in temp.columns:  rename_map[email_col]  = "Email"
        rename_map[scale_col] = "Note"
        rename_map[comm_col]  = "Commentaire"
        temp.rename(columns=rename_map, inplace=True)
        if "Note" not in temp.columns:
            continue
        temp["__note_num"] = temp["Note"].map(parse_note)
        temp = temp[temp["__note_num"] < 3.0].drop(columns="__note_num")
        ordered = [c for c in ["Prénom", "Nom", "Email", "Note", "Commentaire"] if c in temp.columns]
        sheets[display] = temp[ordered]
    return sheets

def generate_excel(df_avg: pd.DataFrame, sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_avg.to_excel(writer, sheet_name="Moyennes", index=False)
        for view in ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]:
            df_view = sheets.get(view, pd.DataFrame(columns=["Prénom", "Nom", "Email", "Note", "Commentaire"]))
            df_view.to_excel(writer, sheet_name=view[:31], index=False)
    output.seek(0)
    return output.getvalue()

# ──────────────────────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────────────────────

st.title("📊 Vues Feedback – Générateur d’Excel")
st.write(
    "Dépose ton export Excel, calcule les **moyennes** par catégorie et récupère **5 vues filtrées (< 3/5)**.\n\n"
    "ℹ️ Hypothèse : chaque **commentaire** est **immédiatement après** la colonne dont l’en-tête **commence par** "
    "“Sur une échelle de 0 à 5 …”."
)

uploaded = st.file_uploader(
    "Dépose ton fichier Excel (.xlsx ou .xls)",
    type=["xlsx", "xls"],
    accept_multiple_files=False
)

if not uploaded:
    st.info("🔺 Dépose un fichier pour commencer.")
    st.stop()

df = read_all_sheets(uploaded)
if df.empty:
    st.error("Impossible de lire des données depuis ce fichier. Vérifie le format et les dépendances (openpyxl, pandas-calamine ou xlrd<=1.2.0).")
    st.stop()

# Détection identité & paires
prenom_col, nom_col, email_col = find_identity_columns(df)
pairs = build_pairs(df)

with st.expander("🔎 Colonnes détectées", expanded=False):
    st.write("**Prénom** :", prenom_col or "non détecté")
    st.write("**Nom** :", nom_col or "non détecté")
    st.write("**Email** :", email_col or "non détecté")
    st.write("**Paires (Échelle 0–5 → Commentaire)** :")
    if pairs:
        st.json({k: {"Echelle 0–5": v[0], "Commentaire": v[1]} for k, v in pairs.items()})
    else:
        st.warning(
            "Aucune paire détectée. Vérifie que l’en-tête de la **colonne de note** "
            "commence bien par **“Sur une échelle de 0 à 5”** et que la **colonne suivante** est le **commentaire**."
        )

if not pairs:
    st.stop()

# Moyennes
df_avg = compute_averages(df, pairs)
st.subheader("📈 Moyennes par catégorie (/5)")
st.dataframe(df_avg, use_container_width=True)
if not df_avg.empty:
    chart_df = df_avg.set_index("Catégorie")["Moyenne (/5)"]
    st.bar_chart(chart_df)

# Vues filtrées < 3/5
sheets = build_views(df, prenom_col, nom_col, email_col, pairs)

tabs = st.tabs(REQUIRED_SHEETS)
with tabs[0]:
    st.markdown("**Moyennes** par catégorie (sur 5).")
    st.dataframe(df_avg, use_container_width=True)

view_names = ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]
for i, view in enumerate(view_names, start=1):
    with tabs[i]:
        st.markdown(f"**{view}** – élèves avec **Note < 3/5**")
        df_view = sheets.get(view, pd.DataFrame(columns=["Prénom", "Nom", "Email", "Note", "Commentaire"]))
        st.dataframe(df_view, use_container_width=True)

# Export Excel
xls_bytes = generate_excel(df_avg, sheets)
st.download_button(
    "📥 Télécharger l’Excel (Moyennes + Vues)",
    data=xls_bytes,
    file_name="vues_feedback.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.caption("ℹ️ Le **commentaire** est **toujours** la colonne immédiatement **après** la colonne dont l’en-tête commence par “Sur une échelle de 0 à 5 …”.")
