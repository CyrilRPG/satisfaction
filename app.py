import streamlit as st
import pandas as pd
from pathlib import Path
from cli_export import (
    read_all_sheets,
    find_identity_columns,
    find_pseudo_column,
    build_pairs,
    compute_averages_by_fac,
    build_views,
    build_commentaires_view,
    build_recommandations_view,
    write_output,
)

st.set_page_config(page_title="Générateur Feedback", layout="centered")

st.title("📊 Générateur de Feedback – Diploma Santé")

uploaded = st.file_uploader("Upload du fichier Excel exporté", type=["xlsx", "xls"])

if uploaded:
    df = pd.ExcelFile(uploaded)
    st.success("✔️ Fichier chargé")

    # Lecture réelle
    df_all = read_all_sheets(uploaded)

    # Détection colonnes
    prenom_col, nom_col, email_col = find_identity_columns(df_all, None, None, None)
    pseudo_col = find_pseudo_column(df_all, None)

    # Détection paires
    pairs = build_pairs(df_all)

    # Construction des vues
    df_avg = compute_averages_by_fac(df_all, pairs, pseudo_col, prenom_col, nom_col, email_col)
    standard_views = build_views(df_all, prenom_col, nom_col, email_col, pairs)
    commentaires_df = build_commentaires_view(df_all, prenom_col, nom_col, email_col, pseudo_col)
    reco_df = build_recommandations_view(df_all, prenom_col, nom_col, email_col, pseudo_col)

    # Génération fichier
    out_path = Path("resultat.xlsx")
    write_output(out_path, df_avg, standard_views, commentaires_df, reco_df)

    with open(out_path, "rb") as f:
        st.download_button(
            label="📥 Télécharger le fichier généré",
            data=f,
            file_name="vues_feedback.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.success("✔️ Fichier généré !")
