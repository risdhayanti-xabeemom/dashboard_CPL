from __future__ import annotations

import re
from html import escape
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


APP_TITLE = "Dashboard Asesmen CPL dan CQI D3 Teknik Elektronika"

REQUIRED_SHEETS = {
    "Master_CPL": ["Kode CPL", "Rumusan CPL", "GA IABEE"],
    "Master_IK": ["Kode CPL", "Kode IK", "Rumusan IK", "Notasi IK"],
    "Mapping_CPMK": [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Rumusan CPMK",
        "Kode CPL",
        "Kode IK",
        "Bobot CPMK",
        "Komponen Asesmen",
        "Bobot Komponen",
    ],
    "Nilai_CPMK": [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Kode CPL",
        "Kode IK",
        "NIM",
        "Nama Mahasiswa",
        "Nilai",
    ],
    "Target": ["Parameter", "Nilai"],
}

BASE_REQUIRED_COLUMNS = {
    "Master_CPL": ["Kode CPL", "Rumusan CPL", "GA IABEE"],
    "Master_IK": ["Kode CPL", "Kode IK", "Rumusan IK"],
    "Mapping_CPMK": [
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Rumusan CPMK",
        "Kode CPL",
        "Kode IK",
        "Bobot CPMK",
    ],
    "Nilai_CPMK": ["NIM", "Nama Mahasiswa", "Kode MK", "Mata Kuliah", "Kode CPMK", "Nilai"],
    "Target": ["Parameter", "Nilai"],
}

OPTIONAL_SHEETS = {
    "Nilai_Asesmen_Detail": [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Kode CPL",
        "Kode IK",
        "Notasi IK",
        "Komponen",
        "NIM",
        "Nama Mahasiswa",
        "Nilai",
        "Bobot Komponen",
        "Label Jadwal Asesmen",
    ]
}

ACCEPTED_COMPONENTS = [
    "Tugas",
    "Quiz",
    "CM",
    "PBL/Produk",
    "UTS",
    "UAS",
    "PBL/Praktik Kerja",
    "Ujian Magang",
    "Ujian Akhir",
]

RECOMMENDATIONS = {
    "CPL1": "Perkuat latihan penerapan matematika, sains, dan konsep elektro dalam tugas terapan.",
    "CPL2": "Tambahkan studi kasus troubleshooting dan analisis masalah well-defined.",
    "CPL3": "Tingkatkan praktik penggunaan instrumen dan software modern.",
    "CPL4": "Perkuat praktikum pengukuran, pengujian, kalibrasi, dan interpretasi data.",
    "CPL5": "Tambahkan proyek implementasi sistem berbasis standar IEC, DIN, ISO, dan K3LH.",
    "CPL6": "Tambahkan studi kasus etika profesi dan tanggung jawab sosial.",
    "CPL7": "Tambahkan peer assessment dan rubrik kerja tim.",
    "CPL8": "Perkuat laporan teknis, presentasi, dan dokumentasi engineering.",
    "CPL9": "Tambahkan evaluasi tugas berbasis standar, waktu, biaya, dan sumber daya.",
    "CPL10": "Tambahkan aktivitas literasi teknologi dan refleksi pembelajaran mandiri.",
}

STATUS_COLORS = {
    "Tercapai": "#1f9d55",
    "Perlu Perhatian": "#d89b00",
    "Belum Tercapai": "#d64545",
}

MASTER_WORKBOOK_CANDIDATES = [
    Path("Master_CPMK_44_MK_D3_PSTE_draft.xlsx"),
    Path("C:/Users/Lenovo/Downloads/Master_CPMK_44_MK_D3_PSTE_draft.xlsx"),
]

LOGO_CANDIDATES = {
    "polinema": [
        Path("logo_polinema.png"),
        Path("logo_polinema.jpg"),
        Path("logo_polinema.jpeg"),
        Path("C:/Users/Lenovo/Downloads/OIP.jpg"),
    ],
    "jte": [
        Path("logo_jte.png"),
        Path("logo_jte.jpg"),
        Path("logo_jte.jpeg"),
        Path("C:/Users/Lenovo/Downloads/OIP (1).jpg"),
    ],
}

COLUMN_ALIASES = {
    "kode cpl": "Kode CPL",
    "rumusan cpl": "Rumusan CPL",
    "ga iabee": "GA IABEE",
    "kode ik": "Kode IK",
    "rumusan ik": "Rumusan IK",
    "notasi ik": "Notasi IK",
    "tahun akademik": "Tahun Akademik",
    "semester": "Semester",
    "kode mk": "Kode MK",
    "mata kuliah": "Mata Kuliah",
    "nama mata kuliah": "Mata Kuliah",
    "kode cpmk": "Kode CPMK",
    "rumusan cpmk": "Rumusan CPMK",
    "bobot cpmk": "Bobot CPMK",
    "komponen asesmen": "Komponen Asesmen",
    "komponen": "Komponen",
    "bobot komponen": "Bobot Komponen",
    "nim": "NIM",
    "nama mahasiswa": "Nama Mahasiswa",
    "nilai": "Nilai",
    "parameter": "Parameter",
    "label jadwal asesmen": "Label Jadwal Asesmen",
}

COMPONENT_ALIASES = {
    "kuis": "Quiz",
    "quiz": "Quiz",
    "tugas": "Tugas",
    "cm": "CM",
    "pbl": "PBL/Produk",
    "pbl/produk": "PBL/Produk",
    "pbl produk": "PBL/Produk",
    "uts": "UTS",
    "uas": "UAS",
    "pbl/praktik kerja": "PBL/Praktik Kerja",
    "pbl praktik kerja": "PBL/Praktik Kerja",
    "ujian magang": "Ujian Magang",
    "ujian akhir": "Ujian Akhir",
}


def load_custom_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --primary: #0f4c81;
            --primary-dark: #12324f;
            --surface: #ffffff;
            --muted: #64748b;
            --line: #e2e8f0;
            --green: #1f9d55;
            --orange: #d89b00;
            --red: #d64545;
        }
        .block-container {
            padding-top: 1.25rem;
            padding-bottom: 2.5rem;
        }
        .app-header {
            padding: 1.25rem 1.5rem;
            border: 1px solid var(--line);
            border-left: 6px solid var(--primary);
            border-radius: 8px;
            background: linear-gradient(135deg, #ffffff 0%, #eef6fb 100%);
            margin-bottom: 1rem;
        }
        .app-header h1 {
            margin: 0;
            font-size: 2rem;
            line-height: 1.2;
            color: var(--primary-dark);
            letter-spacing: 0;
        }
        .app-header .subtitle {
            margin-top: .45rem;
            color: #334155;
            font-size: 1.02rem;
            line-height: 1.55;
        }
        .section-caption {
            margin: .75rem 0 .35rem;
            color: var(--primary-dark);
            font-weight: 700;
            font-size: .92rem;
        }
        .sidebar-logo-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: .65rem;
            margin-bottom: .75rem;
        }
        .logo-placeholder {
            min-height: 76px;
            border: 1px dashed #94a3b8;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
            color: #475569;
            background: #f8fafc;
            font-size: .84rem;
            font-weight: 600;
        }
        .kpi-card {
            min-height: 118px;
            padding: 1rem 1rem .95rem;
            border: 1px solid var(--line);
            border-radius: 8px;
            background: var(--surface);
            box-shadow: 0 10px 24px rgba(15, 23, 42, .07);
            position: relative;
            overflow: hidden;
        }
        .kpi-card::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 6px;
            background: var(--accent);
        }
        .kpi-label {
            color: var(--muted);
            font-size: .88rem;
            font-weight: 700;
            margin-bottom: .4rem;
        }
        .kpi-value {
            color: #0f172a;
            font-size: 2rem;
            line-height: 1;
            font-weight: 800;
        }
        .kpi-note {
            color: var(--muted);
            font-size: .82rem;
            margin-top: .45rem;
        }
        .narrative-box {
            border: 1px solid #bfdbfe;
            background: #eff6ff;
            color: #1e3a5f;
            border-radius: 8px;
            padding: .95rem 1rem;
            margin: 1rem 0;
            line-height: 1.5;
        }
        div[data-testid="stSidebar"] h2, div[data-testid="stSidebar"] h3 {
            color: var(--primary-dark);
        }
        div[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
            margin-bottom: .35rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_main_header() -> None:
    st.markdown(
        """
        <div class="app-header">
            <h1>Dashboard Asesmen CPL dan CQI</h1>
            <div class="subtitle">
                Program Studi D3 Teknik Elektronika<br>
                Jurusan Teknik Elektro<br>
                Politeknik Negeri Malang
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def first_existing_path(paths: Iterable[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None


def render_sidebar_logos() -> None:
    logo_polinema = first_existing_path(LOGO_CANDIDATES["polinema"])
    logo_jte = first_existing_path(LOGO_CANDIDATES["jte"])
    col1, col2 = st.sidebar.columns(2)
    with col1:
        if logo_polinema:
            st.image(str(logo_polinema), width=120)
        else:
            st.markdown('<div class="logo-placeholder">Logo Polinema</div>', unsafe_allow_html=True)
    with col2:
        if logo_jte:
            st.image(str(logo_jte), width=120)
        else:
            st.markdown('<div class="logo-placeholder">Logo JTE</div>', unsafe_allow_html=True)


def sidebar_section(title: str) -> None:
    st.markdown(f'<div class="section-caption">{escape(title)}</div>', unsafe_allow_html=True)


def render_sidebar_system_info(mode: str) -> None:
    sidebar_section("Informasi Sistem")
    st.caption(f"Mode aktif: {mode}")
    st.caption("Format input Excel lama tetap digunakan.")
    st.caption("Perhitungan berbasis CPMK, IK, CPL, dan CQI OBE/IABEE.")
    if first_existing_path(MASTER_WORKBOOK_CANDIDATES):
        st.caption("Master CPL/IK/CPMK terdeteksi untuk template input.")


def render_kpi_card(label: str, value: str, note: str, color: str) -> None:
    st.markdown(
        f"""
        <div class="kpi-card" style="--accent:{escape(color)};">
            <div class="kpi-label">{escape(label)}</div>
            <div class="kpi-value">{escape(value)}</div>
            <div class="kpi-note">{escape(note)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_dashboard_narrative(
    average_cpl: float, achieved: int, attention: int, not_achieved: int
) -> None:
    st.markdown(
        f"""
        <div class="narrative-box">
            Pada periode ini, rata-rata capaian CPL adalah <strong>{average_cpl:.2f}%</strong>.
            Sebanyak <strong>{achieved}</strong> CPL tercapai,
            <strong>{attention}</strong> CPL perlu perhatian, dan
            <strong>{not_achieved}</strong> CPL belum tercapai.
        </div>
        """,
        unsafe_allow_html=True,
    )


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def normalize_label(value: object) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower())


def canonical_column_name(column: object) -> str:
    key = normalize_label(column)
    return COLUMN_ALIASES.get(key, str(column).strip())


def normalize_component(value: object) -> str:
    if pd.isna(value):
        return ""
    text = re.sub(r"\s+", " ", str(value).strip())
    return COMPONENT_ALIASES.get(text.lower(), text)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [canonical_column_name(col) for col in cleaned.columns]
    return cleaned.dropna(how="all")


def read_workbook(file_obj) -> Dict[str, pd.DataFrame]:
    workbook = pd.read_excel(file_obj, sheet_name=None, engine="openpyxl")
    cleaned = {str(sheet).strip(): clean_dataframe(df) for sheet, df in workbook.items()}
    return normalize_input_workbook(cleaned)


def normalize_input_workbook(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    normalized = {sheet: df.copy() for sheet, df in data.items()}

    if "Master_CPMK" in normalized and "Mapping_CPMK" not in normalized:
        master_cpmk = normalized["Master_CPMK"].rename(
            columns={
                "Nama Mata Kuliah": "Mata Kuliah",
            }
        )
        if "Bobot CPMK" not in master_cpmk.columns:
            master_cpmk["Bobot CPMK"] = 1
        mapping = pd.DataFrame(columns=REQUIRED_SHEETS["Mapping_CPMK"])
        for column in mapping.columns:
            if column in master_cpmk.columns:
                mapping[column] = master_cpmk[column]
        normalized["Mapping_CPMK"] = mapping

    if "Master_CPMK" in normalized and "Nilai_CPMK" not in normalized:
        normalized["Nilai_CPMK"] = pd.DataFrame(columns=REQUIRED_SHEETS["Nilai_CPMK"])

    if "Master_CPMK" in normalized and "Target" not in normalized:
        normalized["Target"] = pd.DataFrame(
            [
                {"Parameter": "Batas Nilai Minimum IK", "Nilai": 70},
                {"Parameter": "Target Pemenuhan CPL (%)", "Nilai": 70},
            ]
        )

    if "Master_IK" in normalized and "Notasi IK" not in normalized["Master_IK"].columns:
        normalized["Master_IK"]["Notasi IK"] = normalized["Master_IK"].get("Kode IK", "")

    if "Mapping_CPMK" in normalized:
        mapping = normalized["Mapping_CPMK"]
        if "Komponen" in mapping.columns and "Komponen Asesmen" not in mapping.columns:
            mapping = mapping.rename(columns={"Komponen": "Komponen Asesmen"})
        for column in REQUIRED_SHEETS["Mapping_CPMK"]:
            if column not in mapping.columns:
                mapping[column] = "" if column != "Bobot Komponen" else pd.NA
        mapping["Komponen Asesmen"] = mapping["Komponen Asesmen"].map(normalize_component)
        normalized["Mapping_CPMK"] = mapping

    if "Nilai_CPMK" in normalized:
        nilai = normalized["Nilai_CPMK"]
        for column in REQUIRED_SHEETS["Nilai_CPMK"]:
            if column not in nilai.columns:
                nilai[column] = ""
        if "Mapping_CPMK" in normalized:
            enrich_columns = ["Kode CPMK", "Kode MK", "Kode CPL", "Kode IK", "Tahun Akademik", "Semester"]
            mapping_enrich = (
                normalized["Mapping_CPMK"][enrich_columns]
                .drop_duplicates(subset=["Kode CPMK", "Kode MK"])
                .copy()
            )
            merge_keys = [key for key in ["Kode CPMK", "Kode MK"] if key in nilai.columns]
            if merge_keys:
                nilai = nilai.merge(mapping_enrich, on=merge_keys, how="left", suffixes=("", "_Map"))
                for column in ["Kode CPL", "Kode IK", "Tahun Akademik", "Semester"]:
                    mapped = f"{column}_Map"
                    if mapped in nilai.columns:
                        nilai[column] = nilai[column].where(nilai[column].astype(str).str.strip() != "", nilai[mapped])
                        nilai = nilai.drop(columns=[mapped])
        normalized["Nilai_CPMK"] = nilai

    if "Nilai_Asesmen_Detail" in normalized:
        detail = normalized["Nilai_Asesmen_Detail"]
        for column in OPTIONAL_SHEETS["Nilai_Asesmen_Detail"]:
            if column not in detail.columns:
                detail[column] = ""
        detail["Komponen"] = detail["Komponen"].map(normalize_component)
        normalized["Nilai_Asesmen_Detail"] = detail

    return normalized


def validate_workbook(data: Dict[str, pd.DataFrame], label: str = "file") -> List[str]:
    errors: List[str] = []

    missing_sheets = [sheet for sheet in REQUIRED_SHEETS if sheet not in data]
    for sheet in missing_sheets:
        errors.append(f"{label}: sheet wajib '{sheet}' belum tersedia.")

    for sheet, required_columns in BASE_REQUIRED_COLUMNS.items():
        if sheet not in data:
            continue
        missing_columns = [col for col in required_columns if col not in data[sheet].columns]
        if missing_columns:
            errors.append(
                f"{label}: sheet '{sheet}' belum memiliki kolom wajib: {', '.join(missing_columns)}."
            )

    if errors:
        return errors

    master_cpl = data["Master_CPL"].copy()
    master_ik = data["Master_IK"].copy()
    mapping = data["Mapping_CPMK"].copy()
    nilai = data["Nilai_CPMK"].copy()

    for df, columns in [
        (master_cpl, ["Kode CPL"]),
        (master_ik, ["Kode IK", "Kode CPL"]),
        (mapping, ["Kode CPMK", "Kode IK", "Kode CPL"]),
        (nilai, ["Kode CPMK"]),
    ]:
        for column in columns:
            df[column] = df[column].map(normalize_code)

    missing_cpmk = sorted(
        code
        for code in set(nilai["Kode CPMK"].dropna()) - set(mapping["Kode CPMK"].dropna())
        if code
    )
    if missing_cpmk:
        errors.append(
            f"{label}: Kode CPMK pada 'Nilai_CPMK' belum ada di 'Mapping_CPMK': "
            + ", ".join(missing_cpmk)
            + "."
        )

    missing_ik = sorted(
        code
        for code in set(mapping["Kode IK"].dropna()) - set(master_ik["Kode IK"].dropna())
        if code
    )
    if missing_ik:
        errors.append(
            f"{label}: Kode IK pada 'Mapping_CPMK' belum ada di 'Master_IK': "
            + ", ".join(missing_ik)
            + "."
        )

    missing_cpl = sorted(
        code
        for code in set(master_ik["Kode CPL"].dropna()) - set(master_cpl["Kode CPL"].dropna())
        if code
    )
    if missing_cpl:
        errors.append(
            f"{label}: Kode CPL pada 'Master_IK' belum ada di 'Master_CPL': "
            + ", ".join(missing_cpl)
            + "."
        )

    if pd.to_numeric(nilai["Nilai"], errors="coerce").dropna().empty:
        errors.append(f"{label}: sheet 'Nilai_CPMK' belum memiliki nilai mahasiswa.")

    return errors


def get_workbook_warnings(data: Dict[str, pd.DataFrame], label: str = "file") -> List[str]:
    warnings: List[str] = []

    for sheet, columns in REQUIRED_SHEETS.items():
        if sheet not in data:
            continue
        missing_optional = [
            column
            for column in columns
            if column not in BASE_REQUIRED_COLUMNS.get(sheet, []) and column not in data[sheet].columns
        ]
        if missing_optional:
            warnings.append(
                f"{label}: sheet '{sheet}' belum memiliki kolom opsional template baru: "
                + ", ".join(missing_optional)
                + "."
            )

    if "Nilai_Asesmen_Detail" in data:
        missing_detail = [
            column for column in OPTIONAL_SHEETS["Nilai_Asesmen_Detail"] if column not in data["Nilai_Asesmen_Detail"].columns
        ]
        if missing_detail:
            warnings.append(
                f"{label}: sheet opsional 'Nilai_Asesmen_Detail' belum memiliki kolom: "
                + ", ".join(missing_detail)
                + "."
            )

    if "Mapping_CPMK" in data:
        mapping = data["Mapping_CPMK"].copy()
        if "Komponen Asesmen" in mapping.columns:
            unknown = sorted(
                component
                for component in mapping["Komponen Asesmen"].dropna().map(normalize_component).unique()
                if component and component not in ACCEPTED_COMPONENTS
            )
            if unknown:
                warnings.append(f"{label}: komponen asesmen belum dikenali: {', '.join(unknown)}.")
        warnings.extend(validate_component_rules(mapping, label))
        warnings.extend(validate_component_weights(mapping, label))

    if "Target" in data and "Parameter" in data["Target"].columns:
        target_keys = {
            normalize_label(value).replace("%", "")
            for value in data["Target"]["Parameter"].dropna().astype(str)
        }
        if not {"batas nilai minimum ik", "batas nilai minimum"} & target_keys:
            warnings.append(
                f"{label}: Target belum memuat parameter 'Batas Nilai Minimum IK' atau 'Batas Nilai Minimum'."
            )
        if not {"target pemenuhan cpl ()", "target pemenuhan cpl", "target ketercapaian cpl"} & target_keys:
            warnings.append(
                f"{label}: Target belum memuat parameter 'Target Pemenuhan CPL (%)' atau 'Target Ketercapaian CPL'."
            )

    return warnings


def validate_component_rules(mapping: pd.DataFrame, label: str) -> List[str]:
    warnings: List[str] = []
    if "Mata Kuliah" not in mapping.columns or "Komponen Asesmen" not in mapping.columns:
        return warnings

    magang_allowed = {"PBL/Praktik Kerja", "Ujian Magang", ""}
    proyek_akhir_allowed = {"PBL/Produk", "Ujian Akhir", ""}
    for _, row in mapping.iterrows():
        mk = str(row.get("Mata Kuliah", "")).lower()
        component = normalize_component(row.get("Komponen Asesmen", ""))
        if "magang" in mk and component not in magang_allowed:
            warnings.append(
                f"{label}: komponen '{component}' pada Magang Industri tidak termasuk komponen khusus magang."
            )
        if "proyek akhir" in mk and component not in proyek_akhir_allowed:
            warnings.append(
                f"{label}: komponen '{component}' pada Proyek Akhir tidak termasuk komponen khusus proyek akhir."
            )
    return sorted(set(warnings))


def validate_component_weights(mapping: pd.DataFrame, label: str) -> List[str]:
    required = [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Kode CPMK",
        "Kode CPL",
        "Kode IK",
        "Bobot CPMK",
        "Bobot Komponen",
    ]
    if any(column not in mapping.columns for column in required):
        return []

    check = mapping.copy()
    check["Bobot CPMK"] = pd.to_numeric(check["Bobot CPMK"], errors="coerce")
    check["Bobot Komponen"] = pd.to_numeric(check["Bobot Komponen"], errors="coerce")
    check = check.dropna(subset=["Bobot CPMK", "Bobot Komponen"])
    if check.empty:
        return []

    key_columns = ["Tahun Akademik", "Semester", "Kode MK", "Kode CPMK", "Kode CPL", "Kode IK"]
    grouped = check.groupby(key_columns, dropna=False).agg(
        **{"Bobot CPMK": ("Bobot CPMK", "first"), "Total Bobot Komponen": ("Bobot Komponen", "sum")}
    )
    mismatch = grouped[(grouped["Bobot CPMK"] - grouped["Total Bobot Komponen"]).abs() > 0.01]
    warnings = []
    for key, row in mismatch.head(10).iterrows():
        kode_cpmk = key[3] if isinstance(key, tuple) else key
        warnings.append(
            f"{label}: total Bobot Komponen untuk {kode_cpmk} ({row['Total Bobot Komponen']:.2f}) "
            f"belum sama dengan Bobot CPMK ({row['Bobot CPMK']:.2f})."
        )
    return warnings


def prepare_data(data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    prepared = {sheet: df.copy() for sheet, df in data.items()}
    code_columns = {
        "Master_CPL": ["Kode CPL"],
        "Master_IK": ["Kode IK", "Kode CPL"],
        "Mapping_CPMK": ["Kode MK", "Kode CPMK", "Kode IK", "Kode CPL"],
        "Nilai_CPMK": ["NIM", "Kode MK", "Kode CPMK", "Kode IK", "Kode CPL"],
        "Nilai_Asesmen_Detail": ["NIM", "Kode MK", "Kode CPMK", "Kode IK", "Kode CPL"],
    }
    for sheet, columns in code_columns.items():
        if sheet not in prepared:
            continue
        for column in columns:
            if column in prepared[sheet].columns:
                prepared[sheet][column] = prepared[sheet][column].map(normalize_code)

    prepared["Mapping_CPMK"]["Bobot CPMK"] = pd.to_numeric(
        prepared["Mapping_CPMK"]["Bobot CPMK"], errors="coerce"
    )
    prepared["Mapping_CPMK"]["Bobot Komponen"] = pd.to_numeric(
        prepared["Mapping_CPMK"]["Bobot Komponen"], errors="coerce"
    )
    prepared["Nilai_CPMK"]["Nilai"] = pd.to_numeric(prepared["Nilai_CPMK"]["Nilai"], errors="coerce")
    if "Nilai_Asesmen_Detail" in prepared:
        prepared["Nilai_Asesmen_Detail"]["Nilai"] = pd.to_numeric(
            prepared["Nilai_Asesmen_Detail"]["Nilai"], errors="coerce"
        )
        prepared["Nilai_Asesmen_Detail"]["Bobot Komponen"] = pd.to_numeric(
            prepared["Nilai_Asesmen_Detail"]["Bobot Komponen"], errors="coerce"
        )
    prepared["Target"]["Nilai"] = pd.to_numeric(prepared["Target"]["Nilai"], errors="coerce")
    return prepared


def get_target_value(
    target_df: pd.DataFrame, parameter: str, default: float, aliases: Optional[List[str]] = None
) -> float:
    if target_df.empty:
        return default
    accepted = [parameter, *(aliases or [])]
    accepted_keys = {normalize_label(item).replace("%", "") for item in accepted}
    target_keys = target_df["Parameter"].astype(str).map(lambda value: normalize_label(value).replace("%", ""))
    mask = target_keys.isin(accepted_keys)
    values = target_df.loc[mask, "Nilai"].dropna()
    if values.empty:
        return default
    return float(values.iloc[0])


def status_by_target(capaian: float, target: float) -> str:
    if capaian >= target:
        return "Tercapai"
    if capaian >= target - 10:
        return "Perlu Perhatian"
    return "Belum Tercapai"


def sort_cpl_key(code: str) -> Tuple[int, str]:
    text = normalize_code(code)
    number = "".join(char for char in text if char.isdigit())
    return (int(number) if number else 999, text)


def parse_period_label(label: str) -> Tuple[str, str]:
    text = str(label).strip() or "Periode Tanpa Nama"
    year_match = re.search(r"(\d{4}(?:\s*/\s*\d{4})?)", text)
    semester_match = re.search(r"(ganjil|genap|gasal|semester\s+\d+)", text, re.IGNORECASE)
    tahun = year_match.group(1).replace(" ", "") if year_match else text
    semester = semester_match.group(1).title() if semester_match else text
    if semester.lower() == "gasal":
        semester = "Ganjil"
    return tahun, semester


def add_period_columns(df: pd.DataFrame, periode: str, tahun_akademik: str, semester: str) -> pd.DataFrame:
    output = df.copy()
    if "Semester" in output.columns:
        output = output.rename(columns={"Semester": "Semester Data"})
    if "Tahun Akademik" in output.columns:
        output = output.rename(columns={"Tahun Akademik": "Tahun Akademik Data"})
    if "Periode" in output.columns:
        output = output.rename(columns={"Periode": "Periode Data"})
    for column, value in [
        ("Semester", semester),
        ("Tahun Akademik", tahun_akademik),
        ("Periode", periode),
    ]:
        output.insert(0, column, value)
    return output


def calculate_rekap_cpmk(
    mapping: pd.DataFrame, nilai: pd.DataFrame, batas_nilai_minimum: float
) -> pd.DataFrame:
    nilai_valid = nilai.dropna(subset=["Kode CPMK", "Nilai"]).copy()
    nilai_valid["Tercapai"] = nilai_valid["Nilai"] >= batas_nilai_minimum
    group_keys = [
        column
        for column in ["Tahun Akademik", "Semester", "Kode MK", "Mata Kuliah", "Kode CPMK", "Kode CPL", "Kode IK"]
        if column in nilai_valid.columns
    ]
    if not group_keys:
        group_keys = ["Kode CPMK"]
    grouped = (
        nilai_valid.groupby(group_keys, as_index=False, dropna=False)
        .agg(
            **{
                "Jumlah Mahasiswa": ("NIM", "nunique"),
                "Jumlah Mahasiswa Tercapai": ("Tercapai", "sum"),
                "Rata-rata Nilai": ("Nilai", "mean"),
            }
        )
    )
    grouped["Persentase CPMK"] = (
        grouped["Jumlah Mahasiswa Tercapai"] / grouped["Jumlah Mahasiswa"] * 100
    ).fillna(0)

    mapping_columns = [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Rumusan CPMK",
        "Kode CPL",
        "Kode IK",
        "Bobot CPMK",
    ]
    mapping_unique_key = [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Kode CPMK",
        "Kode CPL",
        "Kode IK",
    ]
    mapping_unique = mapping[mapping_columns].drop_duplicates(subset=mapping_unique_key)

    merge_keys = [key for key in mapping_unique_key if key in grouped.columns]
    if not merge_keys:
        merge_keys = ["Kode CPMK"]
    rekap = mapping_unique.merge(grouped, on=merge_keys, how="left", suffixes=("", "_Nilai"))
    for column in ["Mata Kuliah"]:
        nilai_column = f"{column}_Nilai"
        if nilai_column in rekap.columns:
            rekap[column] = rekap[column].where(rekap[column].astype(str).str.strip() != "", rekap[nilai_column])
            rekap = rekap.drop(columns=[nilai_column])
    fill_zero_columns = [
        "Jumlah Mahasiswa",
        "Jumlah Mahasiswa Tercapai",
        "Rata-rata Nilai",
        "Persentase CPMK",
    ]
    rekap[fill_zero_columns] = rekap[fill_zero_columns].fillna(0)
    return rekap.sort_values(["Kode CPL", "Kode IK", "Kode MK", "Kode CPMK"]).reset_index(drop=True)


def weighted_average(group: pd.DataFrame) -> float:
    weights = pd.to_numeric(group["Bobot CPMK"], errors="coerce")
    values = pd.to_numeric(group["Persentase CPMK"], errors="coerce").fillna(0)
    if weights.notna().any() and weights.fillna(0).sum() > 0:
        return float((values * weights.fillna(0)).sum() / weights.fillna(0).sum())
    return float(values.mean()) if len(values) else 0.0


def calculate_rekap_ik(
    rekap_cpmk: pd.DataFrame, master_ik: pd.DataFrame, target_ketercapaian: float
) -> pd.DataFrame:
    ik_values = pd.DataFrame(
        [
            {"Kode IK": kode_ik, "Capaian IK": weighted_average(group)}
            for kode_ik, group in rekap_cpmk.groupby("Kode IK")
        ],
        columns=["Kode IK", "Capaian IK"],
    )
    counts = (
        rekap_cpmk.groupby("Kode IK", as_index=False)
        .agg(**{"Jumlah CPMK Pendukung": ("Kode CPMK", "nunique")})
    )
    rekap = master_ik.merge(ik_values, on="Kode IK", how="left").merge(counts, on="Kode IK", how="left")
    rekap["Capaian IK"] = rekap["Capaian IK"].fillna(0)
    rekap["Jumlah CPMK Pendukung"] = rekap["Jumlah CPMK Pendukung"].fillna(0).astype(int)
    rekap["Target"] = target_ketercapaian
    rekap["Status"] = rekap["Capaian IK"].apply(lambda value: status_by_target(value, target_ketercapaian))
    return rekap.sort_values(["Kode CPL", "Kode IK"]).reset_index(drop=True)


def calculate_rekap_cpl(
    rekap_ik: pd.DataFrame, master_cpl: pd.DataFrame, target_ketercapaian: float
) -> pd.DataFrame:
    cpl_values = (
        rekap_ik.groupby("Kode CPL", as_index=False)
        .agg(**{"Capaian CPL": ("Capaian IK", "mean"), "Jumlah IK Pendukung": ("Kode IK", "nunique")})
    )
    rekap = master_cpl.merge(cpl_values, on="Kode CPL", how="left")
    rekap["Capaian CPL"] = rekap["Capaian CPL"].fillna(0)
    rekap["Jumlah IK Pendukung"] = rekap["Jumlah IK Pendukung"].fillna(0).astype(int)
    rekap["Target"] = target_ketercapaian
    rekap["Status"] = rekap["Capaian CPL"].apply(lambda value: status_by_target(value, target_ketercapaian))
    return rekap.sort_values("Kode CPL", key=lambda col: col.map(sort_cpl_key)).reset_index(drop=True)


def make_root_cause(status: str) -> str:
    if status == "Belum Tercapai":
        return (
            "Capaian indikator pendukung masih rendah; perlu evaluasi keselarasan CPMK, rubrik, "
            "aktivitas pembelajaran, dan tingkat kesulitan asesmen."
        )
    return (
        "Capaian mendekati target tetapi belum stabil; perlu penguatan aktivitas pembelajaran "
        "dan konsistensi rubrik asesmen."
    )


def generate_cqi(rekap_cpl: pd.DataFrame, rekap_ik: pd.DataFrame) -> pd.DataFrame:
    needs_attention = rekap_cpl[rekap_cpl["Status"].isin(["Perlu Perhatian", "Belum Tercapai"])].copy()
    rows = []
    for _, cpl in needs_attention.iterrows():
        kode_cpl = cpl["Kode CPL"]
        weak_ik = rekap_ik[rekap_ik["Kode CPL"] == kode_cpl].sort_values("Capaian IK").head(3)
        weak_labels = [
            f"{row['Kode IK']} ({row['Capaian IK']:.1f}%)" for _, row in weak_ik.iterrows()
        ]
        recommendation = RECOMMENDATIONS.get(
            kode_cpl,
            "Lakukan telaah rubrik, metode pembelajaran, dan instrumen asesmen pada CPL terkait.",
        )
        rows.append(
            {
                "Kode CPL": kode_cpl,
                "Rumusan CPL": cpl["Rumusan CPL"],
                "Capaian Aktual": round(float(cpl["Capaian CPL"]), 2),
                "Target": round(float(cpl["Target"]), 2),
                "Status": cpl["Status"],
                "Status Awal": cpl["Status"],
                "Indikator Lemah": ", ".join(weak_labels) if weak_labels else "-",
                "Dugaan Akar Penyebab": make_root_cause(cpl["Status"]),
                "Rekomendasi CQI": recommendation,
                "Rekomendasi tindak lanjut": recommendation,
                "PIC": "Koordinator Prodi / Tim Kurikulum / Dosen Pengampu",
                "Target Waktu": "Semester berikutnya",
                "Status Tindak Lanjut": "Belum dimulai",
            }
        )
    return pd.DataFrame(
        rows,
        columns=[
            "Kode CPL",
            "Rumusan CPL",
            "Capaian Aktual",
            "Target",
            "Status",
            "Status Awal",
            "Indikator Lemah",
            "Dugaan Akar Penyebab",
            "Rekomendasi CQI",
            "Rekomendasi tindak lanjut",
            "PIC",
            "Target Waktu",
            "Status Tindak Lanjut",
        ],
    )


def calculate_all(
    data: Dict[str, pd.DataFrame],
    batas_nilai: Optional[float] = None,
    target_cpl: Optional[float] = None,
) -> Dict[str, pd.DataFrame | float]:
    prepared = prepare_data(data)
    batas = (
        float(batas_nilai)
        if batas_nilai is not None
        else get_target_value(
            prepared["Target"], "Batas Nilai Minimum IK", 70, ["Batas Nilai Minimum"]
        )
    )
    target = (
        float(target_cpl)
        if target_cpl is not None
        else get_target_value(
            prepared["Target"], "Target Pemenuhan CPL (%)", 70, ["Target Ketercapaian CPL"]
        )
    )
    rekap_cpmk = calculate_rekap_cpmk(prepared["Mapping_CPMK"], prepared["Nilai_CPMK"], batas)
    rekap_ik = calculate_rekap_ik(rekap_cpmk, prepared["Master_IK"], target)
    rekap_cpl = calculate_rekap_cpl(rekap_ik, prepared["Master_CPL"], target)
    cqi = generate_cqi(rekap_cpl, rekap_ik)
    return {
        "prepared": prepared,
        "batas_nilai": batas,
        "target_cpl": target,
        "rekap_cpmk": rekap_cpmk,
        "rekap_ik": rekap_ik,
        "rekap_cpl": rekap_cpl,
        "cqi": cqi,
    }


def calculate_period_result(period_label: str, data: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    tahun_akademik, semester = parse_period_label(period_label)
    result = calculate_all(data)
    output: Dict[str, pd.DataFrame] = {}
    for key in ["rekap_cpmk", "rekap_ik", "rekap_cpl", "cqi"]:
        output[key] = add_period_columns(
            result[key], period_label, tahun_akademik, semester  # type: ignore[arg-type]
        )
    prepared = result["prepared"]  # type: ignore[assignment]
    if isinstance(prepared, dict) and "Nilai_Asesmen_Detail" in prepared:
        output["detail_asesmen"] = add_period_columns(
            prepared["Nilai_Asesmen_Detail"], period_label, tahun_akademik, semester
        )
    output["target_cpl"] = result["target_cpl"]  # type: ignore[assignment]
    return output


def make_trend_cpl(rekap_cpl_all: pd.DataFrame) -> pd.DataFrame:
    if rekap_cpl_all.empty:
        return pd.DataFrame()
    return (
        rekap_cpl_all[
            ["Periode", "Tahun Akademik", "Semester", "Kode CPL", "Rumusan CPL", "Capaian CPL", "Target", "Status"]
        ]
        .rename(columns={"Capaian CPL": "Capaian Aktual"})
        .reset_index(drop=True)
    )


def make_trend_ik(rekap_ik_all: pd.DataFrame) -> pd.DataFrame:
    if rekap_ik_all.empty:
        return pd.DataFrame()
    return rekap_ik_all[
        [
            "Periode",
            "Tahun Akademik",
            "Semester",
            "Kode CPL",
            "Kode IK",
            "Rumusan IK",
            "Capaian IK",
            "Target",
            "Status",
        ]
    ].reset_index(drop=True)


def effectiveness_label(change: Optional[float]) -> str:
    if change is None or pd.isna(change):
        return "Belum Ada Periode Berikutnya"
    if change >= 5:
        return "Efektif"
    if change >= 0:
        return "Perlu Penguatan"
    return "Belum Efektif"


def recommendation_follow_up(effectiveness: str, kode_cpl: str) -> str:
    if effectiveness == "Efektif":
        return "Pertahankan tindak lanjut dan standarkan praktik baik pada RPS, rubrik, dan asesmen."
    if effectiveness == "Perlu Penguatan":
        return "Perkuat implementasi rekomendasi CQI dan pantau indikator lemah pada asesmen berikutnya."
    if effectiveness == "Belum Efektif":
        return RECOMMENDATIONS.get(
            kode_cpl,
            "Lakukan analisis ulang akar penyebab dan revisi strategi pembelajaran serta asesmen.",
        )
    return "Evaluasi efektivitas setelah data periode berikutnya tersedia."


def get_next_period(period_order: List[str], current: str) -> Optional[str]:
    try:
        index = period_order.index(current)
    except ValueError:
        return None
    if index + 1 >= len(period_order):
        return None
    return period_order[index + 1]


def add_cqi_tracking(cqi_all: pd.DataFrame, rekap_cpl_all: pd.DataFrame, period_order: List[str]) -> pd.DataFrame:
    if cqi_all.empty:
        return cqi_all.copy()
    rows = []
    for _, row in cqi_all.iterrows():
        current_period = row["Periode"]
        next_period = get_next_period(period_order, current_period)
        next_value = None
        change = None
        if next_period:
            match = rekap_cpl_all[
                (rekap_cpl_all["Periode"] == next_period)
                & (rekap_cpl_all["Kode CPL"] == row["Kode CPL"])
            ]
            if not match.empty:
                next_value = float(match.iloc[0]["Capaian CPL"])
                change = next_value - float(row["Capaian Aktual"])
        effectiveness = effectiveness_label(change)
        enriched = row.to_dict()
        enriched["Capaian Semester Berikutnya"] = round(next_value, 2) if next_value is not None else None
        enriched["Efektivitas Tindak Lanjut"] = effectiveness
        rows.append(enriched)
    return pd.DataFrame(rows)


def make_cqi_evaluation(
    cqi_all: pd.DataFrame, rekap_cpl_all: pd.DataFrame, period_order: List[str]
) -> pd.DataFrame:
    rows = []
    if cqi_all.empty:
        return pd.DataFrame(
            columns=[
                "Periode Awal",
                "Periode Berikut",
                "Kode CPL",
                "Indikator Lemah",
                "Capaian Awal",
                "Capaian Berikut",
                "Perubahan",
                "Efektivitas",
                "Rekomendasi Lanjutan",
            ]
        )

    for _, row in cqi_all.iterrows():
        next_period = get_next_period(period_order, row["Periode"])
        if not next_period:
            continue
        match = rekap_cpl_all[
            (rekap_cpl_all["Periode"] == next_period) & (rekap_cpl_all["Kode CPL"] == row["Kode CPL"])
        ]
        if match.empty:
            continue
        capaian_awal = float(row["Capaian Aktual"])
        capaian_berikut = float(match.iloc[0]["Capaian CPL"])
        change = capaian_berikut - capaian_awal
        effectiveness = effectiveness_label(change)
        rows.append(
            {
                "Periode Awal": row["Periode"],
                "Periode Berikut": next_period,
                "Kode CPL": row["Kode CPL"],
                "Indikator Lemah": row["Indikator Lemah"],
                "Capaian Awal": round(capaian_awal, 2),
                "Capaian Berikut": round(capaian_berikut, 2),
                "Perubahan": round(change, 2),
                "Efektivitas": effectiveness,
                "Rekomendasi Lanjutan": recommendation_follow_up(effectiveness, row["Kode CPL"]),
            }
        )
    return pd.DataFrame(rows)


def dataframe_to_excel(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


def make_template_data() -> Dict[str, pd.DataFrame]:
    template = {sheet: pd.DataFrame(columns=columns) for sheet, columns in REQUIRED_SHEETS.items()}
    master_path = first_existing_path(MASTER_WORKBOOK_CANDIDATES)
    if not master_path:
        return template

    try:
        master_data = read_workbook(master_path)
    except Exception:
        return template

    for sheet in ["Master_CPL", "Master_IK", "Mapping_CPMK"]:
        if sheet in master_data:
            required_columns = REQUIRED_SHEETS[sheet]
            frame = master_data[sheet].copy()
            for column in required_columns:
                if column not in frame.columns:
                    frame[column] = ""
            template[sheet] = frame[required_columns]

    template["Target"] = pd.DataFrame(
        [
            {"Parameter": "Batas Nilai Minimum", "Nilai": 70},
            {"Parameter": "Target Ketercapaian CPL", "Nilai": 70},
        ]
    )
    return template


def make_sample_data() -> Dict[str, pd.DataFrame]:
    master_cpl = pd.DataFrame(
        [
            {
                "Kode CPL": "CPL1",
                "Rumusan CPL": "Mampu menerapkan matematika, sains, dan konsep teknik elektro terapan.",
                "GA IABEE": "GA1",
            },
            {
                "Kode CPL": "CPL2",
                "Rumusan CPL": "Mampu menganalisis dan menyelesaikan masalah teknik elektro well-defined.",
                "GA IABEE": "GA2",
            },
            {
                "Kode CPL": "CPL3",
                "Rumusan CPL": "Mampu menggunakan instrumen, perangkat lunak, dan teknologi modern.",
                "GA IABEE": "GA5",
            },
        ]
    )
    master_ik = pd.DataFrame(
        [
            {
                "Kode IK": "IK1.1",
                "Rumusan IK": "Menggunakan konsep matematika untuk rangkaian listrik dasar.",
                "Kode CPL": "CPL1",
            },
            {
                "Kode IK": "IK1.2",
                "Rumusan IK": "Menerapkan konsep sains dan elektro dalam tugas terapan.",
                "Kode CPL": "CPL1",
            },
            {
                "Kode IK": "IK2.1",
                "Rumusan IK": "Mengidentifikasi gejala gangguan pada sistem elektronika.",
                "Kode CPL": "CPL2",
            },
            {
                "Kode IK": "IK2.2",
                "Rumusan IK": "Menentukan alternatif solusi troubleshooting.",
                "Kode CPL": "CPL2",
            },
            {
                "Kode IK": "IK3.1",
                "Rumusan IK": "Mengoperasikan instrumen pengukuran elektronika.",
                "Kode CPL": "CPL3",
            },
            {
                "Kode IK": "IK3.2",
                "Rumusan IK": "Menggunakan software simulasi dan analisis rangkaian.",
                "Kode CPL": "CPL3",
            },
        ]
    )
    mapping = pd.DataFrame(
        [
            {
                "Tahun Akademik": "2025/2026",
                "Semester": "Ganjil",
                "Kode MK": "TE101",
                "Mata Kuliah": "Rangkaian Listrik",
                "Kode CPMK": "CPMK1",
                "Rumusan CPMK": "Mahasiswa mampu menghitung parameter rangkaian DC.",
                "Kode CPL": "CPL1",
                "Kode IK": "IK1.1",
                "Bobot CPMK": 2,
                "Komponen Asesmen": "Tugas",
                "Bobot Komponen": 2,
            },
            {
                "Tahun Akademik": "2025/2026",
                "Semester": "Ganjil",
                "Kode MK": "TE102",
                "Mata Kuliah": "Fisika Terapan",
                "Kode CPMK": "CPMK2",
                "Rumusan CPMK": "Mahasiswa mampu menerapkan konsep fisika pada sistem elektro.",
                "Kode CPL": "CPL1",
                "Kode IK": "IK1.2",
                "Bobot CPMK": 1,
                "Komponen Asesmen": "Quiz",
                "Bobot Komponen": 1,
            },
            {
                "Tahun Akademik": "2025/2026",
                "Semester": "Genap",
                "Kode MK": "TE201",
                "Mata Kuliah": "Dasar Elektronika",
                "Kode CPMK": "CPMK3",
                "Rumusan CPMK": "Mahasiswa mampu mengidentifikasi gangguan komponen elektronika.",
                "Kode CPL": "CPL2",
                "Kode IK": "IK2.1",
                "Bobot CPMK": 1,
                "Komponen Asesmen": "CM",
                "Bobot Komponen": 1,
            },
            {
                "Tahun Akademik": "2025/2026",
                "Semester": "Genap",
                "Kode MK": "TE202",
                "Mata Kuliah": "Troubleshooting Elektronika",
                "Kode CPMK": "CPMK4",
                "Rumusan CPMK": "Mahasiswa mampu memilih solusi troubleshooting.",
                "Kode CPL": "CPL2",
                "Kode IK": "IK2.2",
                "Bobot CPMK": 2,
                "Komponen Asesmen": "PBL/Produk",
                "Bobot Komponen": 2,
            },
            {
                "Tahun Akademik": "2025/2026",
                "Semester": "Genap",
                "Kode MK": "TE301",
                "Mata Kuliah": "Pengukuran Elektronika",
                "Kode CPMK": "CPMK5",
                "Rumusan CPMK": "Mahasiswa mampu menggunakan osiloskop dan multimeter.",
                "Kode CPL": "CPL3",
                "Kode IK": "IK3.1",
                "Bobot CPMK": 1,
                "Komponen Asesmen": "UTS",
                "Bobot Komponen": 1,
            },
            {
                "Tahun Akademik": "2025/2026",
                "Semester": "Genap",
                "Kode MK": "TE302",
                "Mata Kuliah": "Simulasi Rangkaian",
                "Kode CPMK": "CPMK6",
                "Rumusan CPMK": "Mahasiswa mampu menggunakan software simulasi rangkaian.",
                "Kode CPL": "CPL3",
                "Kode IK": "IK3.2",
                "Bobot CPMK": 1,
                "Komponen Asesmen": "UAS",
                "Bobot Komponen": 1,
            },
        ],
        columns=REQUIRED_SHEETS["Mapping_CPMK"],
    )
    students = [
        ("2303001", "Andi Saputra"),
        ("2303002", "Budi Santoso"),
        ("2303003", "Citra Lestari"),
        ("2303004", "Dewi Anggraini"),
        ("2303005", "Eko Prasetyo"),
        ("2303006", "Fitri Nabila"),
        ("2303007", "Gilang Ramadhan"),
        ("2303008", "Hana Maharani"),
        ("2303009", "Indra Wijaya"),
        ("2303010", "Joko Permana"),
    ]
    scores = {
        "CPMK1": [82, 76, 69, 88, 73, 65, 91, 70, 67, 80],
        "CPMK2": [75, 72, 68, 81, 66, 62, 85, 74, 69, 78],
        "CPMK3": [71, 64, 60, 78, 69, 55, 82, 67, 63, 74],
        "CPMK4": [68, 66, 58, 75, 62, 57, 79, 65, 61, 72],
        "CPMK5": [88, 84, 76, 90, 79, 72, 92, 81, 77, 86],
        "CPMK6": [80, 74, 70, 86, 71, 68, 89, 75, 73, 82],
    }
    mapping_lookup = mapping.set_index("Kode CPMK")[["Kode MK", "Mata Kuliah"]].to_dict("index")
    nilai_rows = []
    for cpmk, values in scores.items():
        for (nim, nama), nilai in zip(students, values):
            nilai_rows.append(
                {
                    "NIM": nim,
                    "Nama Mahasiswa": nama,
                    "Kode MK": mapping_lookup[cpmk]["Kode MK"],
                    "Mata Kuliah": mapping_lookup[cpmk]["Mata Kuliah"],
                    "Kode CPMK": cpmk,
                    "Nilai": nilai,
                }
            )
    target = pd.DataFrame(
        [
            {"Parameter": "Batas Nilai Minimum", "Nilai": 70},
            {"Parameter": "Target Ketercapaian CPL", "Nilai": 70},
        ]
    )
    return {
        "Master_CPL": master_cpl,
        "Master_IK": master_ik,
        "Mapping_CPMK": mapping,
        "Nilai_CPMK": pd.DataFrame(nilai_rows),
        "Target": target,
    }


def ensure_example_workbooks() -> None:
    examples = {
        "template_input.xlsx": make_template_data(),
        "sample_data.xlsx": make_sample_data(),
    }
    for filename, sheets in examples.items():
        path = Path(filename)
        if not path.exists():
            try:
                path.write_bytes(dataframe_to_excel(sheets))
            except OSError:
                pass


def format_display(df: pd.DataFrame) -> pd.DataFrame:
    display = df.copy()
    numeric_columns = [
        "Persentase CPMK",
        "Capaian IK",
        "Capaian CPL",
        "Capaian Aktual",
        "Capaian Awal",
        "Capaian Berikut",
        "Capaian Semester Berikutnya",
        "Target",
        "Perubahan",
        "Rata-rata Nilai",
    ]
    for column in numeric_columns:
        if column in display.columns:
            display[column] = pd.to_numeric(display[column], errors="coerce").round(2)
    return display


def unique_filter_options(data: Dict[str, pd.DataFrame], column: str) -> List[str]:
    values: List[str] = []
    for sheet in ["Mapping_CPMK", "Nilai_CPMK", "Nilai_Asesmen_Detail"]:
        if sheet in data and column in data[sheet].columns:
            values.extend(
                str(value).strip()
                for value in data[sheet][column].dropna().unique()
                if str(value).strip()
            )
    return sorted(set(values), key=lambda value: (sort_cpl_key(value) if column == "Kode CPL" else (999, value)))


def render_filter_controls(data: Dict[str, pd.DataFrame], key_prefix: str) -> Dict[str, List[str]]:
    sidebar_section("Filter Dashboard")
    filters = {
        "Tahun Akademik": st.multiselect(
            "Tahun Akademik", unique_filter_options(data, "Tahun Akademik"), key=f"{key_prefix}_filter_tahun"
        ),
        "Semester": st.multiselect(
            "Semester", unique_filter_options(data, "Semester"), key=f"{key_prefix}_filter_semester"
        ),
        "Mata Kuliah": st.multiselect(
            "Mata Kuliah", unique_filter_options(data, "Mata Kuliah"), key=f"{key_prefix}_filter_mk"
        ),
        "Kode CPL": st.multiselect(
            "CPL", unique_filter_options(data, "Kode CPL"), key=f"{key_prefix}_filter_cpl"
        ),
        "Kode IK": st.multiselect("IK", unique_filter_options(data, "Kode IK"), key=f"{key_prefix}_filter_ik"),
    }
    component_options = sorted(
        set(unique_filter_options(data, "Komponen Asesmen") + unique_filter_options(data, "Komponen"))
    )
    filters["Komponen"] = st.multiselect(
        "Komponen Asesmen", component_options, key=f"{key_prefix}_filter_component"
    )
    return filters


def filter_frame(df: pd.DataFrame, filters: Dict[str, List[str]], component_column: Optional[str] = None) -> pd.DataFrame:
    filtered = df.copy()
    for column in ["Tahun Akademik", "Semester", "Mata Kuliah", "Kode CPL", "Kode IK"]:
        selected = filters.get(column, [])
        if selected and column in filtered.columns:
            filtered = filtered[filtered[column].astype(str).isin(selected)]
    selected_components = filters.get("Komponen", [])
    if selected_components and component_column and component_column in filtered.columns:
        filtered = filtered[filtered[component_column].astype(str).isin(selected_components)]
    return filtered


def apply_workbook_filters(data: Dict[str, pd.DataFrame], filters: Dict[str, List[str]]) -> Dict[str, pd.DataFrame]:
    if not any(filters.values()):
        return data
    filtered = {sheet: df.copy() for sheet, df in data.items()}
    if "Mapping_CPMK" in filtered:
        filtered["Mapping_CPMK"] = filter_frame(filtered["Mapping_CPMK"], filters, "Komponen Asesmen")
    if "Nilai_CPMK" in filtered:
        filtered["Nilai_CPMK"] = filter_frame(filtered["Nilai_CPMK"], filters)
        if "Mapping_CPMK" in filtered and filtered["Mapping_CPMK"].empty:
            filtered["Nilai_CPMK"] = filtered["Nilai_CPMK"].iloc[0:0]
        elif "Mapping_CPMK" in filtered:
            keys = ["Tahun Akademik", "Semester", "Kode MK", "Kode CPMK", "Kode CPL", "Kode IK"]
            available_keys = [key for key in keys if key in filtered["Nilai_CPMK"].columns and key in filtered["Mapping_CPMK"].columns]
            allowed = filtered["Mapping_CPMK"][available_keys].drop_duplicates()
            filtered["Nilai_CPMK"] = filtered["Nilai_CPMK"].merge(allowed, on=available_keys, how="inner")
    if "Nilai_Asesmen_Detail" in filtered:
        filtered["Nilai_Asesmen_Detail"] = filter_frame(filtered["Nilai_Asesmen_Detail"], filters, "Komponen")
    return filtered


def load_default_sample() -> Optional[Dict[str, pd.DataFrame]]:
    try:
        return read_workbook("sample_data.xlsx")
    except FileNotFoundError:
        return make_sample_data()


def render_download_buttons_single(
    rekap_cpmk: pd.DataFrame, rekap_ik: pd.DataFrame, rekap_cpl: pd.DataFrame, cqi: pd.DataFrame
) -> None:
    st.subheader("Download Laporan")
    col1, col2, col3, col4, col5 = st.columns(5)
    downloads = [
        (col1, "Rekap CPMK", "Rekap_CPMK.xlsx", {"Rekap_CPMK": rekap_cpmk}),
        (col2, "Rekap IK", "Rekap_IK.xlsx", {"Rekap_IK": rekap_ik}),
        (col3, "Rekap CPL", "Rekap_CPL.xlsx", {"Rekap_CPL": rekap_cpl}),
        (col4, "Laporan CQI", "Laporan_CQI.xlsx", {"Laporan_CQI": cqi}),
        (
            col5,
            "Semua Rekap",
            "Semua_Rekap_Asesmen_CPL.xlsx",
            {
                "Rekap_CPMK": rekap_cpmk,
                "Rekap_IK": rekap_ik,
                "Rekap_CPL": rekap_cpl,
                "Laporan_CQI": cqi,
            },
        ),
    ]
    for col, label, filename, sheets in downloads:
        with col:
            st.download_button(
                label,
                data=dataframe_to_excel(sheets),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


def render_download_buttons_multi(
    rekap_cpmk_all: pd.DataFrame,
    rekap_ik_all: pd.DataFrame,
    rekap_cpl_all: pd.DataFrame,
    trend_cpl: pd.DataFrame,
    trend_ik: pd.DataFrame,
    cqi_all: pd.DataFrame,
    evaluasi_cqi: pd.DataFrame,
) -> None:
    st.subheader("Download Multi Semester")
    col1, col2, col3, col4 = st.columns(4)
    downloads = [
        (col1, "Tren CPL", "Tren_CPL.xlsx", {"Tren_CPL": trend_cpl}),
        (col2, "Tren IK", "Tren_IK.xlsx", {"Tren_IK": trend_ik}),
        (col3, "Evaluasi CQI", "Evaluasi_CQI.xlsx", {"Evaluasi_CQI": evaluasi_cqi}),
        (
            col4,
            "Semua Rekap Multi",
            "Semua_Rekap_Multi_Semester.xlsx",
            {
                "Rekap_CPMK_All": rekap_cpmk_all,
                "Rekap_IK_All": rekap_ik_all,
                "Rekap_CPL_All": rekap_cpl_all,
                "Tren_CPL": trend_cpl,
                "Tren_IK": trend_ik,
                "CQI_All": cqi_all,
                "Evaluasi_CQI": evaluasi_cqi,
            },
        ),
    ]
    for col, label, filename, sheets in downloads:
        with col:
            st.download_button(
                label,
                data=dataframe_to_excel(sheets),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


def render_dashboard(rekap_cpl: pd.DataFrame, rekap_ik: pd.DataFrame) -> None:
    status_counts = rekap_cpl["Status"].value_counts()
    achieved = int(status_counts.get("Tercapai", 0))
    attention = int(status_counts.get("Perlu Perhatian", 0))
    not_achieved = int(status_counts.get("Belum Tercapai", 0))
    average_cpl = float(pd.to_numeric(rekap_cpl["Capaian CPL"], errors="coerce").mean())
    if pd.isna(average_cpl):
        average_cpl = 0.0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_kpi_card("Jumlah CPL Tercapai", str(achieved), "Status memenuhi target", STATUS_COLORS["Tercapai"])
    with col2:
        render_kpi_card(
            "Jumlah CPL Perlu Perhatian",
            str(attention),
            "Mendekati target",
            STATUS_COLORS["Perlu Perhatian"],
        )
    with col3:
        render_kpi_card(
            "Jumlah CPL Belum Tercapai",
            str(not_achieved),
            "Prioritas perbaikan",
            STATUS_COLORS["Belum Tercapai"],
        )
    with col4:
        render_kpi_card("Rata-rata Capaian CPL", f"{average_cpl:.2f}%", "Rerata seluruh CPL", "#0f4c81")

    render_dashboard_narrative(average_cpl, achieved, attention, not_achieved)

    bar = px.bar(
        rekap_cpl,
        x="Kode CPL",
        y="Capaian CPL",
        color="Status",
        text=rekap_cpl["Capaian CPL"].round(1),
        color_discrete_map=STATUS_COLORS,
        range_y=[0, 100],
        title="Capaian CPL",
    )
    bar.update_traces(texttemplate="%{text}%", textposition="outside")
    bar.update_layout(yaxis_title="Capaian (%)", xaxis_title="CPL", legend_title_text="Status")
    st.plotly_chart(bar, use_container_width=True)

    radar_df = rekap_cpl.sort_values("Kode CPL", key=lambda col: col.map(sort_cpl_key))
    theta = radar_df["Kode CPL"].tolist()
    values = radar_df["Capaian CPL"].round(2).tolist()
    if values:
        radar = go.Figure()
        radar.add_trace(
            go.Scatterpolar(
                r=values + [values[0]],
                theta=theta + [theta[0]],
                fill="toself",
                name="Capaian CPL",
                line_color="#2563eb",
            )
        )
        radar.update_layout(
            title="Radar Capaian CPL",
            polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
            showlegend=False,
        )
        st.plotly_chart(radar, use_container_width=True)

    heatmap_source = rekap_ik.pivot_table(
        index="Kode IK", columns="Kode CPL", values="Capaian IK", aggfunc="mean"
    ).fillna(0)
    heatmap = px.imshow(
        heatmap_source,
        labels=dict(x="CPL", y="IK", color="Capaian (%)"),
        text_auto=".1f",
        aspect="auto",
        color_continuous_scale="RdYlGn",
        zmin=0,
        zmax=100,
        title="Heatmap Capaian IK per CPL",
    )
    st.plotly_chart(heatmap, use_container_width=True)


def render_trend_cpl(trend_cpl: pd.DataFrame) -> None:
    st.subheader("Tren Capaian CPL per Periode")
    if trend_cpl.empty:
        st.info("Data tren CPL belum tersedia.")
        return
    cpl_options = sorted(trend_cpl["Kode CPL"].unique(), key=sort_cpl_key)
    selected = st.multiselect("Filter CPL", cpl_options, default=cpl_options)
    filtered = trend_cpl[trend_cpl["Kode CPL"].isin(selected)]
    chart = px.line(
        filtered,
        x="Periode",
        y="Capaian Aktual",
        color="Kode CPL",
        markers=True,
        range_y=[0, 100],
        title="Line Chart Tren Capaian CPL",
    )
    chart.update_layout(yaxis_title="Capaian (%)", xaxis_title="Periode")
    st.plotly_chart(chart, use_container_width=True)
    st.dataframe(
        format_display(filtered[["Periode", "Kode CPL", "Rumusan CPL", "Capaian Aktual", "Target", "Status"]]),
        use_container_width=True,
        hide_index=True,
    )


def render_trend_ik(trend_ik: pd.DataFrame) -> None:
    st.subheader("Tren Capaian IK per CPL")
    if trend_ik.empty:
        st.info("Data tren IK belum tersedia.")
        return
    cpl_options = sorted(trend_ik["Kode CPL"].unique(), key=sort_cpl_key)
    selected_cpl = st.selectbox("Filter Kode CPL", cpl_options)
    filtered = trend_ik[trend_ik["Kode CPL"] == selected_cpl]
    chart = px.line(
        filtered,
        x="Periode",
        y="Capaian IK",
        color="Kode IK",
        markers=True,
        range_y=[0, 100],
        title=f"Line Chart Tren IK dalam {selected_cpl}",
    )
    chart.update_layout(yaxis_title="Capaian IK (%)", xaxis_title="Periode")
    st.plotly_chart(chart, use_container_width=True)
    st.dataframe(
        format_display(filtered[["Periode", "Kode CPL", "Kode IK", "Rumusan IK", "Capaian IK", "Status"]]),
        use_container_width=True,
        hide_index=True,
    )


def render_cqi(cqi: pd.DataFrame) -> None:
    st.subheader("Laporan CQI")
    if cqi.empty:
        st.success("Semua CPL sudah tercapai. Tidak ada item CQI prioritas.")
    else:
        st.dataframe(format_display(cqi), use_container_width=True, hide_index=True)


def render_evaluasi_cqi(evaluasi_cqi: pd.DataFrame) -> None:
    st.subheader("Evaluasi Efektivitas CQI")
    if evaluasi_cqi.empty:
        st.info("Evaluasi CQI membutuhkan minimal dua periode dan CPL bermasalah pada periode awal.")
        return
    st.dataframe(format_display(evaluasi_cqi), use_container_width=True, hide_index=True)


def render_detail_asesmen(detail: pd.DataFrame) -> None:
    st.subheader("Detail Asesmen per Komponen")
    if detail.empty:
        st.info("Sheet `Nilai_Asesmen_Detail` tersedia, tetapi tidak ada data setelah filter.")
        return

    detail_display = detail.copy()
    if "Label Jadwal Asesmen" in detail_display.columns:
        detail_display["Label Jadwal Asesmen"] = detail_display["Label Jadwal Asesmen"].fillna("")
    summary = (
        detail_display.groupby(["Mata Kuliah", "Kode CPMK", "Kode IK", "Komponen"], dropna=False)
        .agg(
            **{
                "Jumlah Mahasiswa": ("NIM", "nunique"),
                "Rata-rata Nilai": ("Nilai", "mean"),
                "Bobot Komponen": ("Bobot Komponen", "first"),
            }
        )
        .reset_index()
    )
    chart = px.bar(
        summary,
        x="Komponen",
        y="Rata-rata Nilai",
        color="Kode IK",
        barmode="group",
        range_y=[0, 100],
        title="Rata-rata Nilai Komponen Asesmen",
        hover_data=["Mata Kuliah", "Kode CPMK", "Bobot Komponen"],
    )
    st.plotly_chart(chart, use_container_width=True)
    st.dataframe(format_display(summary), use_container_width=True, hide_index=True)

    table_columns = [
        "Tahun Akademik",
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Kode CPL",
        "Kode IK",
        "Notasi IK",
        "Komponen",
        "Label Jadwal Asesmen",
        "NIM",
        "Nama Mahasiswa",
        "Nilai",
        "Bobot Komponen",
    ]
    existing_columns = [column for column in table_columns if column in detail_display.columns]
    st.dataframe(format_display(detail_display[existing_columns]), use_container_width=True, hide_index=True)


def render_guide() -> None:
    st.subheader("Panduan Input Data")
    st.markdown(
        """
        Format Excel lama tetap digunakan tanpa perubahan. File harus memiliki sheet `Master_CPL`,
        `Master_IK`, `Mapping_CPMK`, `Nilai_CPMK`, dan `Target`.

        Pada mode `Single Semester`, unggah satu file asesmen seperti biasa. Pada mode
        `Multi Semester`, isi nama periode seperti `2025 Ganjil`, unggah file dengan format yang
        sama, lalu tekan tombol tambah periode. Ulangi untuk periode berikutnya.

        `Bobot CPMK` boleh dikosongkan. Jika bobot tersedia, capaian IK dihitung sebagai rata-rata
        berbobot. Jika bobot tidak tersedia, capaian IK dihitung sebagai rata-rata sederhana.
        """
    )
    st.info("Pastikan kode CPMK, IK, dan CPL konsisten antar sheet dan antar periode.")


def init_period_state() -> None:
    if "period_uploads" not in st.session_state:
        st.session_state.period_uploads = []


def render_single_mode() -> None:
    with st.sidebar:
        sidebar_section("Upload Data")
        uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"], key="single_upload")

    raw_data = read_workbook(uploaded_file) if uploaded_file is not None else load_default_sample()
    if raw_data is None:
        st.warning("Upload file Excel asesmen atau letakkan `sample_data.xlsx` di folder aplikasi.")
        render_guide()
        return

    validation_errors = validate_workbook(raw_data)
    if validation_errors:
        st.error("Data Excel belum sesuai.")
        for error in validation_errors:
            st.write(f"- {error}")
        render_guide()
        return
    validation_warnings = get_workbook_warnings(raw_data)
    if validation_warnings:
        with st.expander("Catatan validasi template", expanded=False):
            for warning in validation_warnings:
                st.warning(warning)

    prepared = prepare_data(raw_data)
    default_batas = get_target_value(
        prepared["Target"], "Batas Nilai Minimum IK", 70, ["Batas Nilai Minimum"]
    )
    default_target = get_target_value(
        prepared["Target"], "Target Pemenuhan CPL (%)", 70, ["Target Ketercapaian CPL"]
    )

    with st.sidebar:
        sidebar_section("Target Asesmen")
        st.caption("Target dibaca dari sheet `Target` dan dapat disesuaikan di sini.")
        batas_nilai = st.number_input(
            "Batas Nilai Minimum IK",
            min_value=0.0,
            max_value=100.0,
            value=float(default_batas),
            step=1.0,
        )
        target_cpl = st.number_input(
            "Target Pemenuhan CPL (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(default_target),
            step=1.0,
        )
        st.divider()
        filters = render_filter_controls(raw_data, "single")
        st.divider()
        render_sidebar_system_info("Single Semester")

    filtered_data = apply_workbook_filters(raw_data, filters)
    result = calculate_all(filtered_data, batas_nilai, target_cpl)
    rekap_cpmk = result["rekap_cpmk"]
    rekap_ik = result["rekap_ik"]
    rekap_cpl = result["rekap_cpl"]
    cqi = result["cqi"]

    render_download_buttons_single(rekap_cpmk, rekap_ik, rekap_cpl, cqi)

    rekap_cpl_period = add_period_columns(rekap_cpl, "Single Semester", "-", "-")
    rekap_ik_period = add_period_columns(rekap_ik, "Single Semester", "-", "-")
    cqi_period = add_period_columns(cqi, "Single Semester", "-", "-")
    trend_cpl = make_trend_cpl(rekap_cpl_period)
    trend_ik = make_trend_ik(rekap_ik_period)
    evaluasi_cqi = make_cqi_evaluation(cqi_period, rekap_cpl_period, ["Single Semester"])

    tab_names = [
        "Dashboard CPL",
        "Rekap CPMK",
        "Rekap IK",
        "Rekap CPL",
        "Tren CPL",
        "Tren IK",
    ]
    if "Nilai_Asesmen_Detail" in filtered_data:
        tab_names.append("Detail Asesmen")
    tab_names.extend(["CQI", "Evaluasi CQI", "Panduan Input Data"])
    tabs = st.tabs(tab_names)
    with tabs[0]:
        render_dashboard(rekap_cpl, rekap_ik)
    with tabs[1]:
        st.dataframe(format_display(rekap_cpmk), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(format_display(rekap_ik), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(format_display(rekap_cpl), use_container_width=True, hide_index=True)
    with tabs[4]:
        render_trend_cpl(trend_cpl)
    with tabs[5]:
        render_trend_ik(trend_ik)
    next_index = 6
    if "Nilai_Asesmen_Detail" in filtered_data:
        with tabs[next_index]:
            render_detail_asesmen(prepare_data(filtered_data)["Nilai_Asesmen_Detail"])
        next_index += 1
    with tabs[next_index]:
        render_cqi(cqi)
    with tabs[next_index + 1]:
        render_evaluasi_cqi(evaluasi_cqi)
    with tabs[next_index + 2]:
        render_guide()


def render_period_manager() -> None:
    init_period_state()
    with st.sidebar:
        sidebar_section("Upload Data")
        period_label = st.text_input("Nama Periode/Semester", placeholder="Contoh: 2025 Ganjil")
        uploaded_file = st.file_uploader("Upload file Excel periode ini", type=["xlsx"], key="multi_upload")
        if st.button("Tambah Periode", use_container_width=True):
            if not period_label.strip():
                st.error("Isi nama periode terlebih dahulu.")
            elif uploaded_file is None:
                st.error("Upload file Excel terlebih dahulu.")
            else:
                st.session_state.period_uploads.append(
                    {
                        "periode": period_label.strip(),
                        "filename": uploaded_file.name,
                        "content": uploaded_file.getvalue(),
                    }
                )
                st.success(f"Periode '{period_label.strip()}' ditambahkan.")
        st.divider()
        sidebar_section("Target Asesmen")
        st.caption("Batas nilai dan target CPL dibaca dari sheet `Target` pada masing-masing file periode.")
        st.divider()
        sidebar_section("Upload Data Tersimpan")
        if not st.session_state.period_uploads:
            st.caption("Belum ada periode.")
        for index, item in enumerate(st.session_state.period_uploads):
            cols = st.columns([3, 1])
            cols[0].caption(f"{index + 1}. {item['periode']} - {item['filename']}")
            if cols[1].button("Hapus", key=f"remove_period_{index}"):
                st.session_state.period_uploads.pop(index)
                st.rerun()
        if st.session_state.period_uploads and st.button("Hapus Semua Periode", use_container_width=True):
            st.session_state.period_uploads = []
            st.rerun()
        st.divider()
        render_sidebar_system_info("Multi Semester")


def render_multi_mode() -> None:
    render_period_manager()
    uploads = st.session_state.get("period_uploads", [])
    if not uploads:
        st.warning("Tambahkan minimal satu periode di sidebar.")
        render_guide()
        return

    period_results = []
    validation_errors: List[str] = []
    validation_warnings: List[str] = []
    for item in uploads:
        try:
            data = read_workbook(BytesIO(item["content"]))
        except Exception as exc:  # pragma: no cover - shown to Streamlit users
            validation_errors.append(f"{item['periode']}: file tidak dapat dibaca. Detail: {exc}")
            continue
        errors = validate_workbook(data, item["periode"])
        if errors:
            validation_errors.extend(errors)
            continue
        validation_warnings.extend(get_workbook_warnings(data, item["periode"]))
        period_results.append(calculate_period_result(item["periode"], data))

    if validation_errors:
        st.error("Ada data periode yang belum sesuai.")
        for error in validation_errors:
            st.write(f"- {error}")
        return
    if validation_warnings:
        with st.expander("Catatan validasi template multi-semester", expanded=False):
            for warning in validation_warnings:
                st.warning(warning)

    rekap_cpmk_all = pd.concat([result["rekap_cpmk"] for result in period_results], ignore_index=True)
    rekap_ik_all = pd.concat([result["rekap_ik"] for result in period_results], ignore_index=True)
    rekap_cpl_all = pd.concat([result["rekap_cpl"] for result in period_results], ignore_index=True)
    cqi_all_raw = pd.concat([result["cqi"] for result in period_results], ignore_index=True)
    detail_frames = [result["detail_asesmen"] for result in period_results if "detail_asesmen" in result]
    detail_asesmen_all = pd.concat(detail_frames, ignore_index=True) if detail_frames else pd.DataFrame()
    period_order = [item["periode"] for item in uploads]

    filter_source = {"Mapping_CPMK": rekap_cpmk_all}
    if not detail_asesmen_all.empty:
        filter_source["Nilai_Asesmen_Detail"] = detail_asesmen_all
    with st.sidebar:
        st.divider()
        multi_filters = render_filter_controls(filter_source, "multi")
    rekap_cpmk_all = filter_frame(rekap_cpmk_all, multi_filters)
    rekap_ik_all = filter_frame(rekap_ik_all, multi_filters)
    rekap_cpl_all = filter_frame(rekap_cpl_all, multi_filters)
    cqi_all_raw = filter_frame(cqi_all_raw, multi_filters)
    if not detail_asesmen_all.empty:
        detail_asesmen_all = filter_frame(detail_asesmen_all, multi_filters, "Komponen")

    trend_cpl = make_trend_cpl(rekap_cpl_all)
    trend_ik = make_trend_ik(rekap_ik_all)
    cqi_all = add_cqi_tracking(cqi_all_raw, rekap_cpl_all, period_order)
    evaluasi_cqi = make_cqi_evaluation(cqi_all_raw, rekap_cpl_all, period_order)

    render_download_buttons_multi(
        rekap_cpmk_all,
        rekap_ik_all,
        rekap_cpl_all,
        trend_cpl,
        trend_ik,
        cqi_all,
        evaluasi_cqi,
    )

    tab_names = [
        "Dashboard CPL",
        "Rekap CPMK",
        "Rekap IK",
        "Rekap CPL",
        "Tren CPL",
        "Tren IK",
    ]
    if not detail_asesmen_all.empty:
        tab_names.append("Detail Asesmen")
    tab_names.extend(["CQI", "Evaluasi CQI", "Panduan Input Data"])
    tabs = st.tabs(tab_names)
    with tabs[0]:
        selected_period = st.selectbox("Pilih periode dashboard", period_order, index=len(period_order) - 1)
        render_dashboard(
            rekap_cpl_all[rekap_cpl_all["Periode"] == selected_period],
            rekap_ik_all[rekap_ik_all["Periode"] == selected_period],
        )
    with tabs[1]:
        st.dataframe(format_display(rekap_cpmk_all), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(format_display(rekap_ik_all), use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(format_display(rekap_cpl_all), use_container_width=True, hide_index=True)
    with tabs[4]:
        render_trend_cpl(trend_cpl)
    with tabs[5]:
        render_trend_ik(trend_ik)
    next_index = 6
    if not detail_asesmen_all.empty:
        with tabs[next_index]:
            render_detail_asesmen(detail_asesmen_all)
        next_index += 1
    with tabs[next_index]:
        render_cqi(cqi_all)
    with tabs[next_index + 1]:
        render_evaluasi_cqi(evaluasi_cqi)
    with tabs[next_index + 2]:
        render_guide()


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    load_custom_css()
    ensure_example_workbooks()
    render_main_header()

    with st.sidebar:
        render_sidebar_logos()
        st.header("Panel Kontrol")
        sidebar_section("Mode Analisis")
        mode = st.radio("Mode", ["Single Semester", "Multi Semester"], horizontal=False)
        st.divider()

    if mode == "Single Semester":
        render_single_mode()
    else:
        render_multi_mode()


if __name__ == "__main__":
    main()
