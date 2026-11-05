import os
import sys
import tempfile
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app  # noqa: E402


class KodeValidTest(unittest.TestCase):
    def test_kode_cpmk_berawalan_mk_tetap_valid(self):
        for code in ["CPMK1", "CPMK1.1", "REC231001-CPMK1", "RTE261002_CPMK 2", "cpmk08.02"]:
            self.assertTrue(app.is_valid_cpmk_code(code), code)
        for code in ["", None, "IK1.1", "-"]:
            self.assertFalse(app.is_valid_cpmk_code(code), code)

    def test_filter_valid_codes_tidak_membuang_cpmk_berawalan_mk(self):
        df = pd.DataFrame(
            {
                "Kode CPL": ["CPL01", "CPL1"],
                "Kode IK": ["IK1.1", "IK01.02"],
                "Kode CPMK": ["REC231001-CPMK1", "CPMK2"],
            }
        )
        result = app.filter_valid_codes(df, require_cpl=True, require_ik=True, require_cpmk=True)
        self.assertEqual(len(result), 2)


class SelGabunganTest(unittest.TestCase):
    def test_master_ik_dengan_kode_cpl_merge_tetap_lengkap(self):
        master_ik = pd.DataFrame(
            {
                "Kode IK": ["IK1.1", "IK1.2", "IK2.1", "IK2.2"],
                "Rumusan IK": ["a", "b", "c", "d"],
                "Kode CPL": ["CPL01", None, "CPL02", None],
            }
        )
        data = {
            "Master_CPL": pd.DataFrame({"Kode CPL": ["CPL01", "CPL02"], "Rumusan CPL": ["x", "y"], "GA IABEE": ["", ""]}),
            "Master_IK": master_ik,
            "Mapping_CPMK": pd.DataFrame(
                columns=["Kode MK", "Mata Kuliah", "Kode CPMK", "Kode CPL", "Kode IK", "Bobot CPMK", "Bobot Komponen"]
            ),
            "Nilai_CPMK": pd.DataFrame(
                columns=["Kode MK", "Mata Kuliah", "Kode CPMK", "Kode CPL", "Kode IK", "NIM", "Nama Mahasiswa", "Nilai"]
            ),
            "Target": pd.DataFrame({"Parameter": ["Batas Nilai Minimum"], "Nilai": [70]}),
        }
        prepared = app.prepare_data(data)
        self.assertEqual(prepared["Master_IK"]["Kode CPL"].tolist(), ["CPL01", "CPL01", "CPL02", "CPL02"])


class SampleDataTest(unittest.TestCase):
    def test_rekap_cpmk_ik_dan_tren_terisi(self):
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        sheets = pd.read_excel(os.path.join(root, "sample_data.xlsx"), sheet_name=None)
        mapping = sheets["Mapping_CPMK"]
        nilai = sheets["Nilai_CPMK"]
        if "Kode IK" not in nilai.columns:
            nilai = nilai.merge(
                mapping[["Kode CPMK", "Kode IK", "Kode CPL"]].drop_duplicates("Kode CPMK"),
                on="Kode CPMK",
                how="left",
            )
            sheets["Nilai_CPMK"] = nilai
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "input.xlsx")
            with pd.ExcelWriter(path) as writer:
                for name, frame in sheets.items():
                    frame.to_excel(writer, sheet_name=name, index=False)
            data = app.read_workbook(path)
        result = app.calculate_all(data, 70, 70)
        rekap_cpmk = result["rekap_cpmk"]
        self.assertFalse(rekap_cpmk.empty)
        self.assertGreater(int((rekap_cpmk["Jumlah Mahasiswa"] > 0).sum()), 0)
        self.assertFalse(app.prepare_trend_cpmk(rekap_cpmk).empty)
        # Semua IK pada Master_IK tetap muncul walau Kode CPL ditulis sebagai sel gabungan.
        self.assertEqual(len(result["rekap_ik"]), len(sheets["Master_IK"]))


if __name__ == "__main__":
    unittest.main()
