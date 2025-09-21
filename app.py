#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Génère un Excel avec:
- Onglet "Moyennes" (moyenne sur 5 par catégorie)
- 5 vues filtrées (<3/5): Coaching, Fiches de cours, Professeurs, Plateforme, Organisation générale

Règle de couplage:
- Pour chaque colonne "Note …" OU "Sur une échelle de 0 à 5 …",
  on rattache le **Commentaire** trouvé dans l'une des 2 colonnes suivantes.

Usage:
    python vues_feedback_cli.py -i "export.xlsx" -o "vues_feedback.xlsx"

Options (facultatives):
    --prenom "Prénom" --nom "Nom" --email "Email"
pour forcer les noms de colonnes d’identité si la détection automatique échoue.
"""

import argparse
import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# Clés de catégories (normalisées) → libellés d’affichage
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


def parse_note(val) -> Optional[float]:
    """Convertit une note en float sur 5.
    Accepte:
      - '2/5', '4 / 5'
      - '2,5', '3' (déjà sur 5)
      - '4 - Satisfait', '5: Très bien' (prend le nombre initial)
    """
    if pd.isna(val):
        return None
    s = str(val).strip()

    # 1) Forme fraction
    m = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)\s*$", s)
    if m:
        num = float(m.group(1).replace(",", "."))
        den = float(m.group(2).replace(",", "."))
        return (num / den) * 5.0 if den else None

    # 2) Nombre simple
    try:
        return float(s.replace(",", "."))
    except ValueError:
        pass

    # 3) Nombre en tête de chaîne (ex: "4 - Plutôt satisfait")
    m2 = re.match(r"^\s*(\d+(?:[.,]\d+)?)", s)
    if m2:
        return float(m2.group(1).replace(",", "."))

    return None


def read_all_sheets(path: Path) -> pd.DataFrame:
    """Lit toutes les feuilles d’un Excel et concatène."""
    try:
        # Essaye openpyxl pour .xlsx ; fallback sinon
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
        except Exception:
            xls = pd.ExcelFile(path)
        frames: List[pd.DataFrame] = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            if not df.empty:
                df["_source_sheet"] = sheet
                frames.append(df)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    except Exception as e:
        raise RuntimeError(f"Erreur lecture Excel: {e}")


def find_identity_columns(df: pd.DataFrame,
                          forced_prenom: Optional[str],
                          forced_nom: Optional[str],
                          forced_email: Optional[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    if forced_prenom or forced_nom or forced_email:
        return forced_prenom, forced_nom, forced_email

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


def _is_comment_col(n: str) -> bool:
    return any(x in n for x in ["comment", "commentaire", "remarque", "avis"])


def build_pairs(df: pd.DataFrame) -> Dict[str, Tuple[str, str]]:
    """Détecte les paires (ColonneNoteOuEchelle, ColonneCommentaireSuivante) par catégorie.
    Retourne dict { 'Coaching': ('<col note/echelle>', '<col commentaire>'), ... }
    """
    columns = list(df.columns)
    norm = [normalize(c) for c in columns]
    target_keys = {k for k, _ in TARGET_VIEWS}
    display_map = {k: disp for k, disp in TARGET_VIEWS}
    pairs: Dict[str, Tuple[str, str]] = {}

    for i, ncol in enumerate(norm):
        # Candidat "Note …"
        is_note = "note" in ncol
        # Candidat "Sur une échelle de 0 à 5 …" (robuste aux accents)
        is_scale = (
            ncol.startswith("sur une echelle de 0 a 5") or
            ncol.startswith("sur une echelle 0 a 5") or
            "echelle de 0 a 5" in ncol
        )

        if not (is_note or is_scale):
            continue

        # Cherche la colonne de commentaire dans les 2 colonnes suivantes
        comment_col = None
        for j in (i + 1, i + 2):
            if j < len(columns) and _is_comment_col(norm[j]):
                comment_col = columns[j]
                break
        if comment_col is None:
            continue

        # Extraire un "cat_key" depuis l'en-tête pour mapper à nos 5 vues
        if is_note:
            base = re.sub(r"\bnote\b|:|-|–|—", " ", ncol)
        else:
            # Retire le préfixe "sur une echelle de 0 a 5"
            base = re.sub(r"^sur une echelle( de)? 0 a 5", " ", ncol)
        cat_key = normalize(base)
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
        if not match_key:
            # Si on ne parvient pas à mapper la catégorie, on ignore ce couple
            continue

        display = display_map[match_key]
        pairs[display] = (columns[i], comment_col)

    return pairs


def compute_averages(df: pd.DataFrame, pairs: Dict[str, Tuple[str, str]]) -> pd.DataFrame:
    rows = []
    for view, (note_col, _) in pairs.items():
        series = df[note_col].map(parse_note)
        series = series.dropna()
        if not series.empty:
            rows.append({"Catégorie": view, "Moyenne (/5)": round(float(series.mean()), 2)})
    if not rows:
        return pd.DataFrame(columns=["Catégorie", "Moyenne (/5)"])
    return pd.DataFrame(rows).sort_values("Catégorie").reset_index(drop=True)


def build_views(df: pd.DataFrame,
                prenom_col: Optional[str],
                nom_col: Optional[str],
                email_col: Optional[str],
                pairs: Dict[str, Tuple[str, str]]) -> Dict[str, pd.DataFrame]:
    sheets: Dict[str, pd.DataFrame] = {}
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


def write_output(output_path: Path, df_avg: pd.DataFrame, sheets: Dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_avg.to_excel(writer, sheet_name="Moyennes", index=False)
        for view in ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]:
            df_view = sheets.get(view, pd.DataFrame(columns=["Prénom", "Nom", "Email", "Note", "Commentaire"]))
            df_view.to_excel(writer, sheet_name=view[:31], index=False)


def main():
    parser = argparse.ArgumentParser(description="Génère un Excel 'Moyennes + Vues (<3/5)' à partir d’un export.")
    parser.add_argument("-i", "--input", required=True, help="Chemin du fichier Excel source (xlsx/xls).")
    parser.add_argument("-o", "--output", default="vues_feedback.xlsx", help="Chemin du fichier Excel de sortie.")
    parser.add_argument("--prenom", help="Nom exact de la colonne Prénom (optionnel).")
    parser.add_argument("--nom", help="Nom exact de la colonne Nom (optionnel).")
    parser.add_argument("--email", help="Nom exact de la colonne Email (optionnel).")
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        print(f"Erreur: fichier introuvable: {in_path}", file=sys.stderr)
        sys.exit(1)

    # Lecture
    df = read_all_sheets(in_path)
    if df.empty:
        print("Erreur: aucune donnée lisible dans le fichier source.", file=sys.stderr)
        sys.exit(1)

    # Colonnes identité + paires
    prenom_col, nom_col, email_col = find_identity_columns(df, args.prenom, args.nom, args.email)
    pairs = build_pairs(df)

    if not pairs:
        print("Erreur: aucune paire détectée. "
              "Astuce: pour chaque catégorie, mets un commentaire dans l'une des 2 colonnes suivant "
              "la colonne 'Note …' ou 'Sur une échelle de 0 à 5 …'.",
              file=sys.stderr)
        # Affiche un aperçu des en-têtes pour debug rapide
        for c in df.columns:
            print(f"- {c}", file=sys.stderr)
        sys.exit(1)

    # Moyennes + vues filtrées
    df_avg = compute_averages(df, pairs)
    sheets = build_views(df, prenom_col, nom_col, email_col, pairs)

    # Écriture
    write_output(out_path, df_avg, sheets)

    # Résumé console
    print(f"✅ Fichier généré: {out_path}")
    print("Feuilles écrites:", ", ".join(REQUIRED_SHEETS))
    if prenom_col or nom_col or email_col:
        print("Colonnes identité détectées/forcées:",
              f"Prénom={prenom_col or '-'} | Nom={nom_col or '-'} | Email={email_col or '-'}")
    print("Paires (colonne de note/échelle → commentaire):")
    for k, (ncol, ccol) in pairs.items():
        print(f"  - {k}: Note/Echelle='{ncol}'  |  Commentaire='{ccol}'")


if __name__ == "__main__":
    main()
