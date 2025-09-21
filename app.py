
import io
import re
import unicodedata
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="Vues Feedback – Diploma Santé", layout="wide")

st.title("📊 Vues Feedback – Générateur d’Excel")
st.write("Importe ton export (Excel) et obtiens un fichier avec des vues filtrées + un tableau des moyennes.")

TARGET_VIEWS = [
    ("coaching", "Coaching"),
    ("fichesdecours", "Fiches de cours"),
    ("fiches cours", "Fiches de cours"),
    ("professeurs", "Professeurs"),
    ("plateforme", "Plateforme"),
    ("organisationgenerale", "Organisation générale"),
    ("organisation generale", "Organisation générale"),
]
REQUIRED_SHEETS = ["Moyennes", "Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]

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
    """Retourne une note numérique ramenée sur 5 (float) à partir de:
       - '2/5', '4 / 5'
       - '2,5' (virgule)
       - '3' (déjà sur 5)
    """
    if pd.isna(val):
        return np.nan
    s = str(val).strip().replace(",", ".")
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)\s*$", s)
    if m:
        num, den = float(m.group(1)), float(m.group(2))
        if den != 0:
            return (num / den) * 5.0
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan

def read_all_sheets(uploaded_file) -> pd.DataFrame:
    try:
        xls = pd.ExcelFile(uploaded_file)
        frames = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            if not df.empty:
                df["_source_sheet"] = sheet
                frames.append(df)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    except Exception as e:
        st.error(f"Erreur de lecture du fichier : {e}")
        return pd.DataFrame()

def find_identity_columns(df):
    cols = {normalize(c): c for c in df.columns}
    first_name = next((cols[k] for k in cols if any(w in k for w in ["prenom", "prénom", "first name", "given name"])), None)
    last_name  = next((cols[k] for k in cols if any(w in k for w in ["nom", "last name", "surname", "family name"]) and "prenom" not in k and "prénom" not in k), None)
    email      = next((cols[k] for k in cols if any(w in k for w in ["email", "e-mail", "mail", "adresse email", "adresse e mail"])), None)
    # Fallbacks
    if first_name is None:
        first_name = next((cols[k] for k in cols if "prenom" in k or "prénom" in k), None)
    if last_name is None:
        last_name = next((cols[k] for k in cols if k.startswith("nom")), None)
    if email is None:
        email = next((cols[k] for k in cols if "adresse" in k and "mail" in k), None)
    return first_name, last_name, email

def build_pairs(df):
    """Repère chaque colonne 'Note ...' et associe la colonne suivante si c'est un commentaire.
       Retour: dict { display_name: (note_col, comment_col) }
    """
    columns = list(df.columns)
    norm = [normalize(c) for c in columns]
    target_keys = {k for k, _ in TARGET_VIEWS}
    display_map = {k: disp for k, disp in TARGET_VIEWS}
    pairs = {}
    for i, ncol in enumerate(norm):
        if "note" in ncol and i + 1 < len(columns):
            next_norm = norm[i + 1]
            if any(x in next_norm for x in ["comment", "commentaire", "remarque", "avis"]):
                cat_key = normalize(re.sub(r"\bnote\b|:|-|–|—", " ", ncol)).replace("note", "").strip()
                cat_key_simple = re.sub(r"[^a-z0-9 ]", "", cat_key)
                cat_key_simple = re.sub(r"\s+", "", cat_key_simple)

                match_key = None
                for tk in target_keys:
                    if tk in cat_key_simple or cat_key_simple in tk:
                        match_key = tk
                        break
                if match_key is None:
                    for tk in target_keys:
                        words = re.findall(r"[a-z]+", tk)
                        if any(w in cat_key_simple for w in words):
                            match_key = tk
                            break
                if match_key:
                    disp = display_map[match_key]
                    pairs[disp] = (columns[i], columns[i+1])
    return pairs

def compute_averages(df, pairs):
    """Calcule la moyenne (sur 5) pour chaque catégorie disponible."""
    data = []
    for view, (note_col, _) in pairs.items():
        series = df[note_col].map(parse_note)
        if series.notna().any():
            mean_val = series.mean()
            data.append({"Catégorie": view, "Moyenne (/5)": round(float(mean_val), 2)})
    if not data:
        return pd.DataFrame(columns=["Catégorie", "Moyenne (/5)"])
    df_avg = pd.DataFrame(data).sort_values("Catégorie").reset_index(drop=True)
    return df_avg

def build_views(df, prenom_col, nom_col, email_col, pairs):
    sheets = {}
    for display, (note_col, comm_col) in pairs.items():
        cols = [c for c in [prenom_col, nom_col, email_col, note_col, comm_col] if c is not None]
        temp = df[cols].copy()
        rename_map = {}
        if prenom_col: rename_map[prenom_col] = "Prénom"
        if nom_col:    rename_map[nom_col]    = "Nom"
        if email_col:  rename_map[email_col]  = "Email"
        rename_map[note_col] = "Note"
        rename_map[comm_col] = "Commentaire"
        temp.rename(columns=rename_map, inplace=True)
        temp["__note_num"] = temp["Note"].map(parse_note)
        temp = temp[temp["__note_num"] < 3.0].drop(columns="__note_num")
        ordered = [c for c in ["Prénom", "Nom", "Email", "Note", "Commentaire"] if c in temp.columns]
        sheets[display] = temp[ordered]
    return sheets

def generate_excel(df_avg, sheets):
    """Retourne un bytes buffer d'un Excel contenant Moyennes + vues filtrées."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Moyennes en premier
        df_avg.to_excel(writer, sheet_name="Moyennes", index=False)
        # Autres vues
        for view in ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]:
            df_view = sheets.get(view, pd.DataFrame(columns=["Prénom", "Nom", "Email", "Note", "Commentaire"]))
            df_view.to_excel(writer, sheet_name=view[:31], index=False)
    output.seek(0)
    return output

uploaded = st.file_uploader("Dépose ton fichier Excel (export)", type=["xlsx", "xls"])

if uploaded is not None:
    df = read_all_sheets(uploaded)
    if df.empty:
        st.warning("Aucune donnée lisible n’a été trouvée.")
        st.stop()

    prenom_col, nom_col, email_col = find_identity_columns(df)
    if not any([prenom_col, nom_col, email_col]):
        st.warning("Impossible de détecter les colonnes d'identité (Prénom/Nom/Email). Le fichier sera traité quand même si les colonnes de notes existent.")
    pairs = build_pairs(df)
    if not pairs:
        st.error("Aucune paire (Note → Commentaire) n’a été détectée. Vérifie les en-têtes.")
        st.stop()

    # Moyennes d'abord
    df_avg = compute_averages(df, pairs)
    st.subheader("📈 Moyennes par catégorie (/5)")
    st.dataframe(df_avg, use_container_width=True)

    # Vues par catégorie (<3/5)
    sheets = build_views(df, prenom_col, nom_col, email_col, pairs)

    tabs = st.tabs(["Moyennes", "Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"])
    with tabs[0]:
        st.markdown("Tableau récapitulatif des moyennes.")
        st.dataframe(df_avg, use_container_width=True)
        # Graphique simple
        if not df_avg.empty:
            chart_df = df_avg.set_index("Catégorie")["Moyenne (/5)"]
            st.bar_chart(chart_df)

    view_names = ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]
    for i, view in enumerate(view_names, start=1):
        with tabs[i]:
            st.markdown(f"**{view}** – élèves avec **Note < 3/5**")
            df_view = sheets.get(view, pd.DataFrame(columns=["Prénom", "Nom", "Email", "Note", "Commentaire"]))
            st.dataframe(df_view, use_container_width=True)

    # Génération Excel téléchargeable
    xls_bytes = generate_excel(df_avg, sheets)
    st.download_button(
        label="📥 Télécharger l’Excel (Moyennes + Vues)",
        data=xls_bytes,
        file_name="vues_feedback.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

else:
    st.info("🔺 Dépose un fichier pour commencer.")

st.caption("ℹ️ Le script associe **chaque Note** à **la colonne Commentaire immédiatement suivante**.")
