# ──────────────────────────────────────────────────────────────────────────────
# Streamlit App: Vues Feedback (Moyennes + Vues filtrées <3/5)
# ──────────────────────────────────────────────────────────────────────────────
# - Upload un Excel (une ou plusieurs feuilles).
# - Détecte automatiquement les paires (Note -> Commentaire) où le commentaire
#   est la colonne immédiatement suivante.
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
    Accepte : '2/5', '4 / 5', '2,5', '3'
    """
    if pd.isna(val):
        return None
    s = str(val).strip().replace(",", ".")
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)\s*$", s)
    if m:
        num = float(m.group(1))
        den = float(m.group(2))
        return (num / den) * 5.0 if den else None
    try:
        return float(s)
    except ValueError:
        return None

def read_all_sheets(uploaded_file) -> pd.DataFrame:
    """
    Lit toutes les feuilles d’un Excel uploadé et les concatène.
    Utilise openpyxl si possible (xlsx), sinon fallback.
    """
    bytes_data = uploaded_file.read()
    if not bytes_data:
        return pd.DataFrame()
    bio = io.BytesIO(bytes_data)
    try:
        # .xlsx recommandé : moteur openpyxl
        sheets = pd.read_excel(bio, sheet_name=None, engine="openpyxl")
    except Exception:
        bio.seek(0)
        sheets = pd.read_excel(bio, sheet_name=None)  # fallback générique
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

def build_pairs(df: pd.DataFrame):
    """
    Détecte les colonnes où une colonne Note est immédiatement suivie d’une colonne Commentaire.
    Retourne: dict { display_name: (note_col, comment_col) }
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
                # extraire la clé de catégorie depuis l'intitulé de la note
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
                        if any(w in cat_key_simple for w in re.findall(r"[a-z]+", tk)):
                            match_key = tk
                            break
                if match_key:
                    disp = display_map[match_key]
                    pairs[disp] = (columns[i], columns[i+1])
    return pairs

def compute_averages(df: pd.DataFrame, pairs: dict) -> pd.DataFrame:
    rows = []
    for view, (note_col, _) in pairs.items():
        series = df[note_col].map(parse_note)
        series = series.dropna()
        if not series.empty:
            rows.append({"Catégorie": view, "Moyenne (/5)": round(float(series.mean()), 2)})
    return pd.DataFrame(rows).sort_values("Catégorie").reset_index(drop=True) if rows else pd.DataFrame(columns=["Catégorie", "Moyenne (/5)"])

def build_views(df: pd.DataFrame, prenom_col: str, nom_col: str, email_col: str, pairs: dict):
    sheets = {}
    for display, (note_col, comm_col) in pairs.items():
        cols = [c for c in [prenom_col, nom_col, email_col, note_col, comm_col] if c and c in df.columns]
        if not cols:
            continue
        temp = df[cols].copy()
        rename_map = {}
        if prenom_col in temp.columns: rename_map[prenom_col] = "Prénom"
        if nom_col in temp.columns:    rename_map[nom_col]    = "Nom"
        if email_col in temp.columns:  rename_map[email_col]  = "Email"
        rename_map[note_col] = "Note"
        rename_map[comm_col] = "Commentaire"
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
st.write("Dépose ton export Excel, calcule les **moyennes** par catégorie et récupère **5 vues filtrées (< 3/5)**.")

uploaded = st.file_uploader(
    "Dépose ton fichier Excel (.xlsx de préférence ; .xls accepté)",
    type=["xlsx", "xls"],
    accept_multiple_files=False
)

if not uploaded:
    st.info("🔺 Dépose un fichier pour commencer.")
    st.stop()

df = read_all_sheets(uploaded)
if df.empty:
    st.error("Impossible de lire des données depuis ce fichier. Vérifie le format (idéalement .xlsx).")
    st.stop()

# Détection identité & paires
prenom_col, nom_col, email_col = find_identity_columns(df)
pairs = build_pairs(df)
with st.expander("🔎 Colonnes détectées", expanded=False):
    st.write("**Prénom** :", prenom_col or "non détecté")
    st.write("**Nom** :", nom_col or "non détecté")
    st.write("**Email** :", email_col or "non détecté")
    st.write("**Paires Note → Commentaire** :")
    if pairs:
        st.json({k: {"Note": v[0], "Commentaire": v[1]} for k, v in pairs.items()})
    else:
        st.warning("Aucune paire détectée. Vérifie que la **colonne Commentaire** est **juste après** la **colonne Note**.")

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

st.caption("ℹ️ Le commentaire est **toujours** la colonne immédiatement **après** la Note.")
