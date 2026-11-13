"""Rumus capaian CPMK -> IK -> CPL (metode jumlah MK pembawa IK).

Rumus ini sama dengan file "Peta CPL-IK-CPMK D3 TE 2025/2026":

- Bobot IK = jumlah MK yang membawa IK / total jumlah MK semua IK.
- Bobot IK dalam CPL = bobot IK / jumlah bobot IK pada CPL tersebut.
- Bobot CPMK di MK = jumlah bobot IK yang diukur CPMK / jumlah bobot IK semua CPMK di MK itu x 100.

Capaian mahasiswa:

- Nilai CPMK = nilai CPMK mahasiswa (rata-rata berbobot komponen bila ada beberapa baris).
- Nilai IK per MK = rata-rata nilai CPMK di MK itu yang mengukur IK tersebut.
- Nilai IK = rata-rata nilai IK per MK di semua MK pembawa IK, sehingga setiap MK pembawa
  dihitung sekali (sama dengan "jumlah MK pembawa IK" pada bobot IK).
- Nilai CPL = jumlah (nilai IK x bobot IK dalam CPL), dibagi jumlah bobot IK yang punya nilai.

CPMK yang mengukur lebih dari satu IK (Magang Industri, Proyek Akhir) boleh ditulis sebagai
"IK7.1, IK8.2" dalam satu sel atau sebagai beberapa baris; nilainya dihitung untuk setiap IK.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional

import pandas as pd

_IK_PATTERN = re.compile(r"IK\s*[-_]?\s*0*(\d+)\s*[.,_-]\s*0*(\d+)", re.IGNORECASE)
_IK_SPLIT_PATTERN = re.compile(r"[;,/\n&]|\bdan\b", re.IGNORECASE)
_CPL_PATTERN = re.compile(r"CPL\s*[-_]?\s*0*(\d+)", re.IGNORECASE)
_BLANK = {"", "NAN", "NONE", "<NA>", "-"}


def _text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    return "" if text.upper() in _BLANK else text


def canonical_ik(value: object) -> str:
    """'IK01.02', 'ik 1.2', 'IK1,2' -> 'IK1.2'. Kode tak dikenal dikembalikan apa adanya (huruf besar)."""
    text = _text(value)
    if not text:
        return ""
    match = _IK_PATTERN.search(text)
    if not match:
        return text.upper()
    return f"IK{int(match.group(1))}.{int(match.group(2))}"


def canonical_cpl(value: object) -> str:
    text = _text(value)
    match = _CPL_PATTERN.search(text)
    if not match:
        return text.upper()
    return f"CPL{int(match.group(1)):02d}"


def split_ik_codes(value: object) -> List[str]:
    """'IK7.1, IK8.2' -> ['IK7.1', 'IK8.2']. Satu kode tetap jadi daftar berisi satu kode."""
    text = _text(value)
    if not text:
        return []
    found = [f"IK{int(a)}.{int(b)}" for a, b in _IK_PATTERN.findall(text)]
    if found:
        return list(dict.fromkeys(found))
    parts = [part.strip().upper() for part in _IK_SPLIT_PATTERN.split(text)]
    return [part for part in dict.fromkeys(parts) if part]


def cpl_of_ik(ik: str, master_lookup: Optional[Dict[str, str]] = None) -> str:
    if master_lookup and ik in master_lookup and master_lookup[ik]:
        return master_lookup[ik]
    match = _IK_PATTERN.search(ik or "")
    if match:
        return f"CPL{int(match.group(1)):02d}"
    return ""


def master_ik_lookup(master_ik: Optional[pd.DataFrame]) -> Dict[str, str]:
    if master_ik is None or master_ik.empty or "Kode IK" not in master_ik.columns:
        return {}
    master = master_ik.copy()
    if "Kode CPL" in master.columns:
        cpl = master["Kode CPL"].map(_text).replace("", pd.NA).ffill()
    else:
        cpl = pd.Series(pd.NA, index=master.index)
    lookup: Dict[str, str] = {}
    for ik_value, cpl_value in zip(master["Kode IK"], cpl):
        ik = canonical_ik(ik_value)
        if ik and ik not in lookup:
            lookup[ik] = canonical_cpl(cpl_value) if _text(cpl_value) else cpl_of_ik(ik)
    return lookup


def _code(series: pd.Series) -> pd.Series:
    return series.map(_text).str.upper()


def explode_mapping_ik(mapping: pd.DataFrame, master_ik: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Satu baris per (Kode MK, Kode CPMK, Kode IK) dengan Kode CPL dari IK."""
    columns = ["Kode MK", "Kode CPMK", "Kode IK", "Kode CPL"]
    if mapping is None or mapping.empty or "Kode CPMK" not in mapping.columns:
        return pd.DataFrame(columns=columns)
    lookup = master_ik_lookup(master_ik)
    rows = []
    mk_series = _code(mapping["Kode MK"]) if "Kode MK" in mapping.columns else pd.Series("", index=mapping.index)
    ik_series = mapping["Kode IK"] if "Kode IK" in mapping.columns else pd.Series("", index=mapping.index)
    cpl_series = mapping["Kode CPL"] if "Kode CPL" in mapping.columns else pd.Series("", index=mapping.index)
    for mk, cpmk, ik_cell, cpl_cell in zip(mk_series, _code(mapping["Kode CPMK"]), ik_series, cpl_series):
        if not cpmk:
            continue
        iks = split_ik_codes(ik_cell)
        for ik in iks:
            cpl = cpl_of_ik(ik, lookup)
            if not cpl and len(iks) == 1:
                cpl = canonical_cpl(cpl_cell)
            rows.append({"Kode MK": mk, "Kode CPMK": cpmk, "Kode IK": ik, "Kode CPL": cpl})
    result = pd.DataFrame(rows, columns=columns)
    return result.drop_duplicates().reset_index(drop=True)


def ik_weights(mapping: pd.DataFrame, master_ik: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Bobot IK dan bobot IK dalam CPL dari jumlah MK pembawa IK (sheet BOBOT IK)."""
    columns = ["Kode CPL", "Kode IK", "Jumlah MK Pembawa", "Bobot IK (%)", "Bobot dalam CPL (%)"]
    links = explode_mapping_ik(mapping, master_ik)
    links = links[(links["Kode IK"] != "") & (links["Kode CPL"] != "")]
    if links.empty:
        return pd.DataFrame(columns=columns)
    counts = (
        links.groupby(["Kode CPL", "Kode IK"], as_index=False)
        .agg(**{"Jumlah MK Pembawa": ("Kode MK", "nunique")})
    )
    total = counts["Jumlah MK Pembawa"].sum()
    counts["Bobot IK (%)"] = counts["Jumlah MK Pembawa"] / total * 100 if total else 0.0
    cpl_total = counts.groupby("Kode CPL")["Jumlah MK Pembawa"].transform("sum")
    counts["Bobot dalam CPL (%)"] = (counts["Jumlah MK Pembawa"] / cpl_total * 100).fillna(0)
    return counts[columns]


def cpmk_weights(mapping: pd.DataFrame, master_ik: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """Bobot CPMK di MK (%) = jumlah bobot IK CPMK / jumlah bobot IK semua CPMK di MK x 100."""
    columns = ["Kode MK", "Kode CPMK", "Bobot CPMK Rumus (%)"]
    links = explode_mapping_ik(mapping, master_ik)
    weights = ik_weights(mapping, master_ik)
    if links.empty or weights.empty:
        return pd.DataFrame(columns=columns)
    merged = links.merge(weights[["Kode IK", "Bobot IK (%)"]], on="Kode IK", how="left")
    merged["Bobot IK (%)"] = merged["Bobot IK (%)"].fillna(0)
    per_cpmk = merged.groupby(["Kode MK", "Kode CPMK"], as_index=False)["Bobot IK (%)"].sum()
    mk_total = per_cpmk.groupby("Kode MK")["Bobot IK (%)"].transform("sum")
    per_cpmk["Bobot CPMK Rumus (%)"] = (per_cpmk["Bobot IK (%)"] / mk_total * 100).fillna(0)
    return per_cpmk[columns]


def _weighted_score(group: pd.DataFrame) -> float:
    values = pd.to_numeric(group["_nilai"], errors="coerce").clip(0, 100)
    weights = pd.to_numeric(group["_bobot"], errors="coerce")
    valid = values.notna()
    values, weights = values[valid], weights[valid]
    if values.empty:
        return float("nan")
    weights = weights.where(weights > 0).fillna(1)
    return float((values * weights).sum() / weights.sum())


def _identity_columns(nilai: pd.DataFrame) -> List[str]:
    return [column for column in ["Periode", "Kelas", "NIM"] if column in nilai.columns]


def student_cpmk_ik(
    nilai: pd.DataFrame, mapping: pd.DataFrame, master_ik: Optional[pd.DataFrame] = None
) -> pd.DataFrame:
    """Nilai CPMK tiap mahasiswa, satu baris per IK yang diukur CPMK itu."""
    if nilai is None or nilai.empty or "NIM" not in nilai.columns or "Kode CPMK" not in nilai.columns:
        return pd.DataFrame(columns=["NIM", "Nama Mahasiswa", "Kode MK", "Kode CPMK", "Kode IK", "Kode CPL", "Nilai CPMK"])
    work = nilai.copy()
    value_column = "Nilai_Bersih" if "Nilai_Bersih" in work.columns else "Nilai"
    work["_nilai"] = pd.to_numeric(work[value_column], errors="coerce").clip(0, 100)
    weight_column = next(
        (column for column in ["Bobot_Komponen_Bersih", "Bobot Komponen"] if column in work.columns),
        None,
    )
    work["_bobot"] = pd.to_numeric(work[weight_column], errors="coerce") if weight_column else 1.0
    work["NIM"] = work["NIM"].map(_text)
    work = work[(work["NIM"] != "") & work["_nilai"].notna()]
    work["Kode MK"] = _code(work["Kode MK"]) if "Kode MK" in work.columns else ""
    work["Kode CPMK"] = _code(work["Kode CPMK"])
    work = work[work["Kode CPMK"] != ""]
    if "Nama Mahasiswa" not in work.columns:
        work["Nama Mahasiswa"] = ""
    if "Kode IK" not in work.columns:
        work["Kode IK"] = ""
    if "Kode CPL" not in work.columns:
        work["Kode CPL"] = ""

    keys = _identity_columns(work) + ["Kode MK", "Kode CPMK"]
    rows = []
    for key, group in work.groupby(keys, dropna=False, sort=False):
        key = key if isinstance(key, tuple) else (key,)
        row = dict(zip(keys, key))
        row["Nama Mahasiswa"] = next((name for name in group["Nama Mahasiswa"].map(_text) if name), "")
        row["Nilai CPMK"] = _weighted_score(group)
        row["_ik_nilai"] = ",".join(dict.fromkeys(i for cell in group["Kode IK"] for i in split_ik_codes(cell)))
        row["_cpl_nilai"] = next((c for c in group["Kode CPL"].map(_text) if c), "")
        rows.append(row)
    scores = pd.DataFrame(rows)
    if scores.empty:
        return pd.DataFrame(columns=keys + ["Nama Mahasiswa", "Kode IK", "Kode CPL", "Nilai CPMK"])

    # IK diambil dari Mapping_CPMK (kurikulum). Bila kode MK + CPMK tidak ada di mapping,
    # coba kode CPMK saja (bila unik), lalu IK yang tertulis di Nilai_CPMK.
    links = explode_mapping_ik(mapping, master_ik)
    by_mk_cpmk: Dict[tuple, List[tuple]] = {}
    for mk, cpmk, ik, cpl in links.itertuples(index=False):
        by_mk_cpmk.setdefault((mk, cpmk), []).append((ik, cpl))
    by_cpmk: Dict[str, List[tuple]] = {}
    for (mk, cpmk), pairs in by_mk_cpmk.items():
        by_cpmk.setdefault(cpmk, []).append((mk, pairs))
    lookup = master_ik_lookup(master_ik)

    out = []
    for row in scores.to_dict("records"):
        pairs = by_mk_cpmk.get((row["Kode MK"], row["Kode CPMK"]))
        if pairs is None and len(by_cpmk.get(row["Kode CPMK"], [])) == 1:
            pairs = by_cpmk[row["Kode CPMK"]][0][1]
        if pairs is None:
            iks = split_ik_codes(row["_ik_nilai"])
            pairs = [
                (ik, cpl_of_ik(ik, lookup) or (canonical_cpl(row["_cpl_nilai"]) if len(iks) == 1 else ""))
                for ik in iks
            ]
        for ik, cpl in pairs:
            if not ik or not cpl:
                continue
            item = {k: v for k, v in row.items() if not k.startswith("_")}
            item["Kode IK"] = ik
            item["Kode CPL"] = cpl
            out.append(item)
    columns = keys + ["Nama Mahasiswa", "Kode IK", "Kode CPL", "Nilai CPMK"]
    return pd.DataFrame(out, columns=columns)


def student_ik(
    nilai: pd.DataFrame, mapping: pd.DataFrame, master_ik: Optional[pd.DataFrame] = None
) -> pd.DataFrame:
    """Nilai IK tiap mahasiswa: rata-rata per MK dulu, lalu rata-rata antar MK pembawa."""
    base = student_cpmk_ik(nilai, mapping, master_ik)
    identity = [column for column in ["Periode", "Kelas", "NIM"] if column in base.columns]
    columns = identity + ["Nama Mahasiswa", "Kode CPL", "Kode IK", "Nilai IK", "Jumlah MK Dinilai"]
    if base.empty:
        return pd.DataFrame(columns=columns)
    per_mk = (
        base.groupby(identity + ["Kode CPL", "Kode IK", "Kode MK"], as_index=False, dropna=False)
        .agg(**{"Nilai IK MK": ("Nilai CPMK", "mean"), "Nama Mahasiswa": ("Nama Mahasiswa", "first")})
    )
    result = (
        per_mk.groupby(identity + ["Kode CPL", "Kode IK"], as_index=False, dropna=False)
        .agg(
            **{
                "Nama Mahasiswa": ("Nama Mahasiswa", "first"),
                "Nilai IK": ("Nilai IK MK", "mean"),
                "Jumlah MK Dinilai": ("Kode MK", "nunique"),
            }
        )
    )
    result["Nilai IK"] = result["Nilai IK"].clip(0, 100)
    return result[columns]


def student_cpl(
    nilai: pd.DataFrame, mapping: pd.DataFrame, master_ik: Optional[pd.DataFrame] = None
) -> pd.DataFrame:
    """Nilai CPL tiap mahasiswa = rata-rata nilai IK berbobot 'bobot IK dalam CPL'."""
    ik_scores = student_ik(nilai, mapping, master_ik)
    identity = [column for column in ["Periode", "Kelas", "NIM"] if column in ik_scores.columns]
    columns = identity + ["Nama Mahasiswa", "Kode CPL", "Nilai CPL", "Jumlah IK Dinilai"]
    if ik_scores.empty:
        return pd.DataFrame(columns=columns)
    weights = ik_weights(mapping, master_ik)[["Kode IK", "Bobot dalam CPL (%)"]]
    work = ik_scores.merge(weights, on="Kode IK", how="left")
    # IK yang tidak ada di Mapping_CPMK (hanya di nilai) diberi bobot 1 MK.
    work["_w"] = pd.to_numeric(work["Bobot dalam CPL (%)"], errors="coerce").fillna(0)
    work.loc[work["_w"] <= 0, "_w"] = work["_w"][work["_w"] > 0].min() if (work["_w"] > 0).any() else 1.0
    work["_wx"] = work["_w"] * work["Nilai IK"]
    result = (
        work.groupby(identity + ["Kode CPL"], as_index=False, dropna=False)
        .agg(
            **{
                "Nama Mahasiswa": ("Nama Mahasiswa", "first"),
                "_wx": ("_wx", "sum"),
                "_w": ("_w", "sum"),
                "Jumlah IK Dinilai": ("Kode IK", "nunique"),
            }
        )
    )
    result["Nilai CPL"] = (result["_wx"] / result["_w"]).clip(0, 100)
    return result[columns]
