#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# ========= CONFIG ========= #

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
    "Commentaires",
    "Recommandations",
]

FAC_ORDER = ["UPC", "UPEC", "UPS", "UVSQ", "SU", "USPN"]
FAC_DISPLAY = {
    "UPC": "UPC",
    "UPEC": "UPEC L1",
    "UPS": "UPS",
    "UVSQ": "UVSQ",
    "SU": "SU",
    "USPN": "USPN",
}

RECO_COL_EXACT = (
    "Si vous avez des besoins, des demandes ou des améliorations à proposer avant le concours, écrivez-les ici !"
)


# ========= UTILS ========= #

def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s.lower().strip())


def parse_note(val) -> Optional[float]:
    if pd.isna(val):
        return None
    s = str(val).strip()

    m = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)\s*$", s)
    if m:
        num = float(m.group(1).replace(",", "."))
        den = float(m.group(2).replace(",", "."))
        return (num / den) * 5.0 if den else None

    try:
        return float(s.replace(",", "."))
    except ValueError:
        pass

    m2 = re.match(r"^\s*(\d+(?:[.,]\d+)?)", s)
    if m2:
        return float(m2.group(1).replace(",", "."))

    return None


def read_all_sheets(path: Path) -> pd.DataFrame:
    try:
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
        except Exception:
            xls = pd.ExcelFile(path)

        frames = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            if not df.empty:
                df["_source_sheet"] = sheet
                frames.append(df)
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    except Exception as e:
        raise RuntimeError(f"Erreur lecture Excel: {e}")


def find_identity_columns(df, forced_prenom, forced_nom, forced_email):
    if forced_prenom or forced_nom or forced_email:
        return forced_prenom, forced_nom, forced_email

    cols = {normalize(c): c for c in df.columns}

    first = next((cols[k] for k in cols if "prenom" in k or "prénom" in k or "first" in k), None)
    last = next((cols[k] for k in cols if k.startswith("nom") or "last" in k), None)
    mail = next((cols[k] for k in cols if "mail" in k or "email" in k), None)

    return first, last, mail


def find_pseudo_column(df, forced):
    if forced and forced in df.columns:
        return forced
    for col in df.columns:
        if any(x in normalize(col) for x in ["pseudo", "identifiant", "username", "login"]):
            return col
    return None


def infer_faculty_from_value(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = normalize(v).upper()
    for f in FAC_ORDER:
        if f in s:
            return f
    return None


def infer_faculty_for_row(row, pseudo_col, prenom_col, nom_col, email_col):
    if pseudo_col and row.get(pseudo_col):
        f = infer_faculty_from_value(row[pseudo_col])
        if f:
            return f
    if email_col and row.get(email_col):
        f = infer_faculty_from_value(row[email_col])
        if f:
            return f

    concat = []
    if prenom_col and row.get(prenom_col): concat.append(str(row[prenom_col]))
    if nom_col and row.get(nom_col): concat.append(str(row[nom_col]))
    if concat:
        return infer_faculty_from_value(" ".join(concat))

    return None


def _is_comment_col(n):
    return any(x in n for x in ["comment", "commentaire", "remarque", "avis"])


# ========= CORE LOGIC ========= #

def build_pairs(df):
    columns = list(df.columns)
    norm = [normalize(c) for c in columns]

    target_keys = {k for k, _ in TARGET_VIEWS}
    display_map = dict(TARGET_VIEWS)
    pairs = {}

    for i, ncol in enumerate(norm):
        is_note = "note" in ncol
        is_scale = ncol.startswith("sur une echelle") or "echelle de 0 a 5" in ncol

        if not (is_note or is_scale):
            continue

        # commentaire dans les 2 prochaines colonnes
        comm = None
        for j in (i + 1, i + 2):
            if j < len(columns) and _is_comment_col(norm[j]):
                comm = columns[j]
                break
        if not comm:
            continue

        base = re.sub(r"\bnote\b|:|-", " ", ncol) if is_note else re.sub(r"^sur une echelle( de)? 0 a 5", " ", ncol)
        cat_key = re.sub(r"\s+", "", re.sub(r"[^a-z0-9 ]", "", normalize(base)))

        match = None
        for tk in target_keys:
            if tk in cat_key or cat_key in tk:
                match = tk
                break
        if not match:
            continue

        display = display_map[match]
        pairs[display] = (columns[i], comm)

    return pairs


def compute_averages_by_fac(df, pairs, pseudo_col, prenom_col, nom_col, email_col):
    rows = []
    for view, (note_col, _) in pairs.items():
        notes = df[note_col].map(parse_note)
        facs = df.apply(lambda r: infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col), axis=1)

        tmp = pd.DataFrame({"fac": facs, "note": notes}).dropna()
        if tmp.empty:
            rows.append({"Catégorie": view, **{FAC_DISPLAY[f]: None for f in FAC_ORDER}})
            continue

        means = tmp.groupby("fac")["note"].mean().to_dict()
        row = {"Catégorie": view}
        for f in FAC_ORDER:
            row[FAC_DISPLAY[f]] = round(float(means.get(f, float("nan"))), 2) if f in means else None

        rows.append(row)

    df_avg = pd.DataFrame(rows)
    ordered_cols = ["Catégorie"] + [FAC_DISPLAY[f] for f in FAC_ORDER]
    return df_avg.reindex(columns=ordered_cols).sort_values("Catégorie").reset_index(drop=True)


def build_views(df, prenom_col, nom_col, email_col, pairs):
    sheets = {}
    for display, (note_col, comm_col) in pairs.items():
        needed = [c for c in [prenom_col, nom_col, email_col, note_col, comm_col] if c in df.columns]

        temp = df[needed].copy()
        rename = {
            prenom_col: "Prénom",
            nom_col: "Nom",
            email_col: "Email",
            note_col: "Note",
            comm_col: "Commentaire",
        }
        temp.rename(columns={k: v for k, v in rename.items() if k}, inplace=True)

        temp["__note"] = temp["Note"].map(parse_note)
        temp = temp[temp["__note"] < 3].drop(columns="__note")

        final_cols = [c for c in ["Prénom", "Nom", "Email", "Note", "Commentaire"] if c in temp.columns]
        sheets[display] = temp[final_cols]

    return sheets


# ========= NEW FEATURES ========= #

def build_commentaires_view(df, prenom_col, nom_col, email_col, pseudo_col):
    comment_cols = [c for c in df.columns if _is_comment_col(normalize(c))]

    rows = []
    for _, r in df.iterrows():
        comments = []
        for col in comment_cols:
            v = r.get(col)
            if isinstance(v, str) and v.strip():
                comments.append(f"{col}: {v.strip()}")

        if not comments:
            continue

        fac = infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col)

        rows.append({
            "Prénom": r.get(prenom_col, ""),
            "Nom": r.get(nom_col, ""),
            "Email": r.get(email_col, ""),
            "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
            "Commentaires": "\n".join(comments),
        })

    return pd.DataFrame(rows)


def build_recommandations_view(df, prenom_col, nom_col, email_col, pseudo_col):
    if RECO_COL_EXACT not in df.columns:
        return pd.DataFrame(columns=["Prénom", "Nom", "Email", "Fac", "Recommandation"])

    rows = []
    for _, r in df.iterrows():
        rec = r.get(RECO_COL_EXACT)
        if isinstance(rec, str) and rec.strip():
            fac = infer_faculty_for_row(r, pseudo_col, prenom_col, nom_col, email_col)
            rows.append({
                "Prénom": r.get(prenom_col, ""),
                "Nom": r.get(nom_col, ""),
                "Email": r.get(email_col, ""),
                "Fac": FAC_DISPLAY.get(fac, fac) if fac else "",
                "Recommandation": rec.strip(),
            })

    return pd.DataFrame(rows)


def write_output(path, df_avg, standard_views, commentaires_df, reco_df):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df_avg.to_excel(writer, sheet_name="Moyennes", index=False)

        for view in ["Coaching", "Fiches de cours", "Professeurs", "Plateforme", "Organisation générale"]:
            df_v = standard_views.get(view, pd.DataFrame(columns=["Prénom", "Nom", "Email", "Note", "Commentaire"]))
            df_v.to_excel(writer, sheet_name=view[:31], index=False)

        commentaires_df.to_excel(writer, sheet_name="Commentaires", index=False)
        reco_df.to_excel(writer, sheet_name="Recommandations", index=False)


# ========= MAIN ========= #

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", required=True)
    parser.add_argument("-o", "--output", default="vues_feedback.xlsx")
    parser.add_argument("--prenom")
    parser.add_argument("--nom")
    parser.add_argument("--email")
    parser.add_argument("--pseudo")
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        print(f"Erreur: fichier introuvable → {in_path}", file=sys.stderr)
        sys.exit(1)

    df = read_all_sheets(in_path)
    if df.empty:
        print("Erreur: fichier vide", file=sys.stderr)
        sys.exit(1)

    prenom_col, nom_col, email_col = find_identity_columns(df, args.prenom, args.nom, args.email)
    pseudo_col = find_pseudo_column(df, args.pseudo)

    pairs = build_pairs(df)
    df_avg = compute_averages_by_fac(df, pairs, pseudo_col, prenom_col, nom_col, email_col)
    standard_views = build_views(df, prenom_col, nom_col, email_col, pairs)

    commentaires_df = build_commentaires_view(df, prenom_col, nom_col, email_col, pseudo_col)
    reco_df = build_recommandations_view(df, prenom_col, nom_col, email_col, pseudo_col)

    write_output(out_path, df_avg, standard_views, commentaires_df, reco_df)

    print(f"✅ Fichier généré : {out_path}")
    print(f"📄 Feuilles : {', '.join(REQUIRED_SHEETS)}")


if __name__ == "__main__":
    main()
