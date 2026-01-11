# -*- coding: utf-8 -*-
"""
pks_analysis_2.py (nur T01 Zeitreihe Bund)

Fixes:
- Zeitraum: ab 2012 bis 2024
- Analyse-Deliktgruppen (ohne Untergruppen):
    * diebstahl_insgesamt  -> Schlüssel '****00' (exakt)
    * wohnungseinbruch     -> Schlüssel '435*00' (exakt)
    * kfz_diebstahl        -> Schlüssel '***100' (exakt)
- Aufklärungsquote (AQ): wird DIREKT aus Excel ("in % (AQ)") übernommen.
  Fix für Dezimalpunkt: 28.0 bleibt 28.0 (nicht 280).
- Outputs:
    * time_series_cases_clean.csv
    * time_series_clearance_clean.csv  (AQ in %, gerundet auf 1 Stelle)
    * tree_all_diebstahl.txt
    * tree_relevant_diebstahl.txt
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


# -----------------------------
# CONFIG
# -----------------------------
FILE_T01 = Path("T01-ZR-Bund-Fälle_xls.xlsx")
SHEET_T01 = "T01 Zeitreihe"

YEAR_MIN = 2014
YEAR_MAX = 2024

DROP_2024_Z22 = True

ANALYSIS_KEYS = {
    "diebstahl_insgesamt": "****00",
    "wohnungseinbruch": "435*00",
    "kfz_diebstahl": "***100",
}

TREE_RELEVANT_NODES = {
    "root": "****00",
    "kfz_total": "***100",
    "kfz_simple": "3**100",
    "kfz_aggravated": "4**100",
    "wohnung_total": "**35*00",
    "wed": "435*00",
}


# -----------------------------
# Helpers
# -----------------------------
def flatten_col(col_tuple) -> str:
    parts = [str(x) for x in col_tuple if x is not None and str(x) != "nan"]
    parts = [p for p in parts if not p.startswith("Unnamed")]
    return "_".join(parts).replace("\n", " ").strip()


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def find_col(cols: List[str], patterns: List[str]) -> Optional[str]:
    for pat in patterns:
        rx = re.compile(pat, re.IGNORECASE)
        for c in cols:
            if rx.search(c):
                return c
    return None


def to_num_de_counts(series: pd.Series) -> pd.Series:
    """
    Für Fallzahlen etc. (Tausenderpunkt möglich, Dezimalkomma möglich)
    Beispiel: '2.379.725' -> 2379725
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    s = series.astype(str).str.strip()
    s = s.replace({"-": None, "—": None, "": None})

    s = s.str.replace(" ", "", regex=False)
    s = s.str.replace(".", "", regex=False)      # Tausenderpunkt weg
    s = s.str.replace(",", ".", regex=False)     # Komma -> Punkt
    return pd.to_numeric(s, errors="coerce")


def to_percent_aq(series: pd.Series) -> pd.Series:
    """
    Für AQ-Spalte (in % (AQ)):
    - Excel liefert oft echte Zahlen (float) -> direkt übernehmen.
    - Wenn Text: Komma-Variante behandeln (29,2 -> 29.2)
    - Dezimalpunkt darf NICHT entfernt werden (sonst 28.0 -> 280).
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    def parse_one(x):
        if pd.isna(x):
            return pd.NA
        s = str(x).strip()
        if s in {"-", "—", ""}:
            return pd.NA

        s = s.replace(" ", "")

        # Wenn Komma drin ist: deutsches Dezimalformat -> Punkt, und Tausenderpunkte entfernen
        if "," in s:
            s = s.replace(".", "")      # Tausenderpunkt weg
            s = s.replace(",", ".")     # Dezimalcomma -> dot
        else:
            # Kein Komma: Punkt ist sehr wahrscheinlich Dezimalpunkt (AQ <= 100)
            # Falls doch Tausenderformat '1.234' (unwahrscheinlich bei AQ), könnte man das abfangen:
            if re.match(r"^\d{1,3}\.\d{3}$", s):
                s = s.replace(".", "")

        try:
            return float(s)
        except ValueError:
            return pd.NA

    return series.apply(parse_one)


def matches_pattern(key: str, pattern: str) -> bool:
    key = str(key).strip()
    pattern = str(pattern).strip()
    if len(key) != len(pattern):
        return False
    for k, p in zip(key, pattern):
        if p != "*" and k != p:
            return False
    return True


# -----------------------------
# Load + clean T01
# -----------------------------
def load_t01_time_series(file_path: Path) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name=SHEET_T01, header=[10, 12, 13])
    df.columns = [flatten_col(c) for c in df.columns]
    cols = df.columns.tolist()

    col_key = find_col(cols, [r"^Schl[uü]ssel$"])
    col_off = find_col(cols, [r"^Straftat$"])
    col_year = find_col(cols, [r"^Jahr$"])
    col_cases = find_col(cols, [r"Anzahl erfasste F[aä]lle$"])
    col_aq = find_col(cols, [r"in\s*%\s*\(AQ\)", r"Aufkl[aä]rung.*in\s*%"])

    if not all([col_key, col_off, col_year, col_cases, col_aq]):
        raise ValueError(
            "Spalten nicht gefunden. "
            f"key={col_key}, off={col_off}, year={col_year}, cases={col_cases}, aq={col_aq}"
        )

    df = df[df[col_key].notna() & df[col_year].notna()].copy()

    df["year_str"] = df[col_year].astype(str).map(normalize_text)
    df["year"] = pd.to_numeric(df["year_str"].str.extract(r"(\d{4})")[0], errors="coerce")
    df = df.dropna(subset=["year"]).copy()
    df["year"] = df["year"].astype(int)

    df = df[(df["year"] >= YEAR_MIN) & (df["year"] <= YEAR_MAX)].copy()

    if DROP_2024_Z22:
        df = df[~((df["year"] == 2024) & df["year_str"].str.contains("Z22", na=False))].copy()

    df["key"] = df[col_key].astype(str).map(normalize_text)
    df["offence"] = df[col_off].astype(str).map(normalize_text)

    df["cases"] = to_num_de_counts(df[col_cases])
    df["aq_percent"] = to_percent_aq(df[col_aq]).round(1)  # <<< FIX

    return df[["year", "year_str", "key", "offence", "cases", "aq_percent"]]


# -----------------------------
# Build outputs (exact keys only)
# -----------------------------
def build_exact_series(df: pd.DataFrame, key_exact: str) -> pd.DataFrame:
    sub = df[df["key"] == key_exact].copy()
    out = (
        sub.sort_values("year")
        .groupby("year", as_index=False)
        .agg({"cases": "first", "aq_percent": "first"})
    )
    return out


def write_clean_csvs(df: pd.DataFrame) -> None:
    series = {name: build_exact_series(df, k) for name, k in ANALYSIS_KEYS.items()}

    base_years = series["diebstahl_insgesamt"]["year"]

    # Cases
    cases_out = pd.DataFrame({"year": base_years})
    for name, s in series.items():
        cases_out[name] = s.set_index("year").reindex(base_years)["cases"].values
    cases_out.to_csv("time_series_cases_clean_2014_2024.csv", index=False)

    # AQ (in Prozent aus Excel)
    aq_out = pd.DataFrame({"year": base_years})
    for name, s in series.items():
        aq_out[name] = s.set_index("year").reindex(base_years)["aq_percent"].values
    aq_out.to_csv("time_series_clearance_clean_2014_2024.csv", index=False)


# -----------------------------
# Trees
# -----------------------------
def build_parent_child_edges(keys: List[str]) -> List[Tuple[str, str]]:
    keyset = set(keys)
    edges = []
    for k in keys:
        for i, ch in enumerate(k):
            if ch != "*":
                parent = k[:i] + "*" + k[i + 1 :]
                if parent in keyset:
                    edges.append((parent, k))
    return list(dict.fromkeys(edges))


def pick_label(df: pd.DataFrame, key_or_pattern: str) -> str:
    if key_or_pattern in set(df["key"]):
        sub = df[df["key"] == key_or_pattern].sort_values("year")
        return sub["offence"].iloc[0]
    return key_or_pattern


def ascii_tree(root: str, children_map: Dict[str, List[str]], label_map: Dict[str, str]) -> str:
    lines = []

    def rec(node: str, pref: str):
        kids = children_map.get(node, [])
        for idx, child in enumerate(kids):
            last = idx == len(kids) - 1
            branch = "└─ " if last else "├─ "
            lines.append(f"{pref}{branch}{label_map.get(child, child)} ({child})")
            rec(child, pref + ("   " if last else "│  "))

    lines.append(f"{label_map.get(root, root)} ({root})")
    rec(root, "")
    return "\n".join(lines) + "\n"


def write_trees(df: pd.DataFrame) -> None:
    theft_df = df[df["offence"].str.lower().str.contains("diebstahl", na=False)].copy()
    theft_keys = sorted(theft_df["key"].unique().tolist())

    edges = build_parent_child_edges(theft_keys)
    children_map: Dict[str, List[str]] = {}
    for p, c in edges:
        children_map.setdefault(p, []).append(c)

    # sort children for stable output
    for p in list(children_map.keys()):
        children_map[p] = sorted(children_map[p])

    root = "****00" if "****00" in theft_keys else theft_keys[0]
    label_map = {k: pick_label(df, k) for k in theft_keys}
    Path("tree_all_diebstahl.txt").write_text(ascii_tree(root, children_map, label_map), encoding="utf-8")

    # Relevant tree EXACT like screenshot structure
    rel_children = {
        TREE_RELEVANT_NODES["root"]: [TREE_RELEVANT_NODES["kfz_total"], TREE_RELEVANT_NODES["wohnung_total"]],
        TREE_RELEVANT_NODES["kfz_total"]: [TREE_RELEVANT_NODES["kfz_simple"], TREE_RELEVANT_NODES["kfz_aggravated"]],
        TREE_RELEVANT_NODES["wohnung_total"]: [TREE_RELEVANT_NODES["wed"]],
    }
    rel_label_map = {k: pick_label(df, k) for k in TREE_RELEVANT_NODES.values()}
    Path("tree_relevant_diebstahl.txt").write_text(
        ascii_tree(TREE_RELEVANT_NODES["root"], rel_children, rel_label_map),
        encoding="utf-8"
    )


# -----------------------------
# MAIN
# -----------------------------
def main():
    df = load_t01_time_series(FILE_T01)
    write_clean_csvs(df)
    write_trees(df)
    print("✅ Fertig: time_series_cases_clean_2014_2024.csv, time_series_clearance_clean_2014_2024.csv, tree_all_diebstahl.txt, tree_relevant_diebstahl.txt")


if __name__ == "__main__":
    main()
