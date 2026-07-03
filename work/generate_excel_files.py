from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]

REQUIRED_SHEETS = {
    "Master_CPL": ["Kode CPL", "Rumusan CPL", "GA IABEE"],
    "Master_IK": ["Kode IK", "Rumusan IK", "Kode CPL"],
    "Mapping_CPMK": [
        "Semester",
        "Kode MK",
        "Mata Kuliah",
        "Kode CPMK",
        "Rumusan CPMK",
        "Kode IK",
        "Kode CPL",
        "Bobot CPMK",
    ],
    "Nilai_CPMK": ["NIM", "Nama Mahasiswa", "Kode MK", "Mata Kuliah", "Kode CPMK", "Nilai"],
    "Target": ["Parameter", "Nilai"],
}


def col_name(index: int) -> str:
    name = ""
    number = index + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def sheet_xml(rows: list[list[object]]) -> str:
    xml_rows = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row):
            ref = f"{col_name(col_index)}{row_index}"
            if value in (None, ""):
                cells.append(f'<c r="{ref}"/>')
            elif isinstance(value, (int, float)):
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
            else:
                cells.append(
                    f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'
                )
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheetData>{"".join(xml_rows)}</sheetData></worksheet>'
    )


def rows_from_objects(columns: list[str], objects: list[dict[str, object]]) -> list[list[object]]:
    return [columns, *[[item.get(column, "") for column in columns] for item in objects]]


def write_workbook(path: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    sheet_overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i, _ in enumerate(sheets, start=1)
    )
    workbook_sheets = "".join(
        f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, (name, _) in enumerate(sheets, start=1)
    )
    workbook_rels = "".join(
        f'<Relationship Id="rId{i}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{i}.xml"/>'
        for i, _ in enumerate(sheets, start=1)
    )
    workbook_rels += (
        f'<Relationship Id="rId{len(sheets) + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            f"{sheet_overrides}"
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>'
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{workbook_sheets}</sheets></workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{workbook_rels}</Relationships>"
        ),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            "</styleSheet>"
        ),
    }
    for index, (_, rows) in enumerate(sheets, start=1):
        files[f"xl/worksheets/sheet{index}.xml"] = sheet_xml(rows)

    with ZipFile(path, "w", ZIP_DEFLATED) as workbook:
        for filename, content in files.items():
            workbook.writestr(filename, content)


def sample_sheets() -> list[tuple[str, list[list[object]]]]:
    master_cpl = [
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
    master_ik = [
        {"Kode IK": "IK1.1", "Rumusan IK": "Menggunakan konsep matematika untuk rangkaian listrik dasar.", "Kode CPL": "CPL1"},
        {"Kode IK": "IK1.2", "Rumusan IK": "Menerapkan konsep sains dan elektro dalam tugas terapan.", "Kode CPL": "CPL1"},
        {"Kode IK": "IK2.1", "Rumusan IK": "Mengidentifikasi gejala gangguan pada sistem elektronika.", "Kode CPL": "CPL2"},
        {"Kode IK": "IK2.2", "Rumusan IK": "Menentukan alternatif solusi troubleshooting.", "Kode CPL": "CPL2"},
        {"Kode IK": "IK3.1", "Rumusan IK": "Mengoperasikan instrumen pengukuran elektronika.", "Kode CPL": "CPL3"},
        {"Kode IK": "IK3.2", "Rumusan IK": "Menggunakan software simulasi dan analisis rangkaian.", "Kode CPL": "CPL3"},
    ]
    mapping = [
        {"Semester": 1, "Kode MK": "TE101", "Mata Kuliah": "Rangkaian Listrik", "Kode CPMK": "CPMK1", "Rumusan CPMK": "Mahasiswa mampu menghitung parameter rangkaian DC.", "Kode IK": "IK1.1", "Kode CPL": "CPL1", "Bobot CPMK": 2},
        {"Semester": 1, "Kode MK": "TE102", "Mata Kuliah": "Fisika Terapan", "Kode CPMK": "CPMK2", "Rumusan CPMK": "Mahasiswa mampu menerapkan konsep fisika pada sistem elektro.", "Kode IK": "IK1.2", "Kode CPL": "CPL1", "Bobot CPMK": 1},
        {"Semester": 2, "Kode MK": "TE201", "Mata Kuliah": "Dasar Elektronika", "Kode CPMK": "CPMK3", "Rumusan CPMK": "Mahasiswa mampu mengidentifikasi gangguan komponen elektronika.", "Kode IK": "IK2.1", "Kode CPL": "CPL2", "Bobot CPMK": 1},
        {"Semester": 2, "Kode MK": "TE202", "Mata Kuliah": "Troubleshooting Elektronika", "Kode CPMK": "CPMK4", "Rumusan CPMK": "Mahasiswa mampu memilih solusi troubleshooting.", "Kode IK": "IK2.2", "Kode CPL": "CPL2", "Bobot CPMK": 2},
        {"Semester": 3, "Kode MK": "TE301", "Mata Kuliah": "Pengukuran Elektronika", "Kode CPMK": "CPMK5", "Rumusan CPMK": "Mahasiswa mampu menggunakan osiloskop dan multimeter.", "Kode IK": "IK3.1", "Kode CPL": "CPL3", "Bobot CPMK": 1},
        {"Semester": 3, "Kode MK": "TE302", "Mata Kuliah": "Simulasi Rangkaian", "Kode CPMK": "CPMK6", "Rumusan CPMK": "Mahasiswa mampu menggunakan software simulasi rangkaian.", "Kode IK": "IK3.2", "Kode CPL": "CPL3", "Bobot CPMK": 1},
    ]
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
    mapping_by_cpmk = {row["Kode CPMK"]: row for row in mapping}
    nilai = []
    for cpmk, values in scores.items():
        for (nim, nama), score in zip(students, values):
            nilai.append(
                {
                    "NIM": nim,
                    "Nama Mahasiswa": nama,
                    "Kode MK": mapping_by_cpmk[cpmk]["Kode MK"],
                    "Mata Kuliah": mapping_by_cpmk[cpmk]["Mata Kuliah"],
                    "Kode CPMK": cpmk,
                    "Nilai": score,
                }
            )
    target = [
        {"Parameter": "Batas Nilai Minimum", "Nilai": 70},
        {"Parameter": "Target Ketercapaian CPL", "Nilai": 70},
    ]
    objects = {
        "Master_CPL": master_cpl,
        "Master_IK": master_ik,
        "Mapping_CPMK": mapping,
        "Nilai_CPMK": nilai,
        "Target": target,
    }
    return [(sheet, rows_from_objects(columns, objects[sheet])) for sheet, columns in REQUIRED_SHEETS.items()]


def main() -> None:
    template = [(sheet, [columns]) for sheet, columns in REQUIRED_SHEETS.items()]
    write_workbook(ROOT / "template_input.xlsx", template)
    write_workbook(ROOT / "sample_data.xlsx", sample_sheets())


if __name__ == "__main__":
    main()
