import os
import sys
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app  # noqa: E402
import rumus_ik_cpl as rumus  # noqa: E402


def contoh_mapping():
    # MK A: 2 CPMK (IK1.1, IK1.2). MK B: 1 CPMK (IK1.1).
    # MK C: 1 CPMK yang mengukur 2 IK sekaligus (gaya Magang/PA): "IK1.1, IK2.1".
    return pd.DataFrame(
        {
            "Kode MK": ["MKA", "MKA", "MKB", "MKC"],
            "Mata Kuliah": ["A", "A", "B", "C"],
            "Kode CPMK": ["MKA-CPMK1", "MKA-CPMK2", "MKB-CPMK1", "MKC-CPMK1"],
            "Kode CPL": ["CPL01", "CPL01", "CPL01", "CPL01, CPL02"],
            "Kode IK": ["IK1.1", "IK1.2", "IK1.1", "IK1.1, IK2.1"],
            "Bobot CPMK": [None, None, None, None],
        }
    )


def contoh_nilai():
    return pd.DataFrame(
        {
            "NIM": ["1"] * 4,
            "Nama Mahasiswa": ["Ani"] * 4,
            "Kode MK": ["MKA", "MKA", "MKB", "MKC"],
            "Kode CPMK": ["MKA-CPMK1", "MKA-CPMK2", "MKB-CPMK1", "MKC-CPMK1"],
            "Nilai": [80, 60, 70, 90],
        }
    )


class KodeTest(unittest.TestCase):
    def test_kode_ik_dinormalkan_dan_dipisah(self):
        self.assertEqual(rumus.canonical_ik("IK01.02"), "IK1.2")
        self.assertEqual(rumus.split_ik_codes("IK7.1, IK8.2; ik 10.1"), ["IK7.1", "IK8.2", "IK10.1"])
        self.assertEqual(rumus.cpl_of_ik("IK10.1"), "CPL10")


class BobotTest(unittest.TestCase):
    def test_bobot_ik_dari_jumlah_mk_pembawa(self):
        bobot = rumus.ik_weights(contoh_mapping()).set_index("Kode IK")
        # IK1.1 dibawa 3 MK, IK1.2 1 MK, IK2.1 1 MK: total 5.
        self.assertEqual(bobot.loc["IK1.1", "Jumlah MK Pembawa"], 3)
        self.assertAlmostEqual(bobot.loc["IK1.1", "Bobot IK (%)"], 60.0)
        self.assertAlmostEqual(bobot.loc["IK1.1", "Bobot dalam CPL (%)"], 75.0)
        self.assertAlmostEqual(bobot.loc["IK2.1", "Bobot dalam CPL (%)"], 100.0)
        self.assertEqual(bobot.loc["IK2.1", "Kode CPL"], "CPL02")

    def test_bobot_cpmk_di_mk(self):
        bobot = rumus.cpmk_weights(contoh_mapping()).set_index("Kode CPMK")["Bobot CPMK Rumus (%)"]
        # MK A: IK1.1 (60) dan IK1.2 (20) -> 75% dan 25%.
        self.assertAlmostEqual(bobot["MKA-CPMK1"], 75.0)
        self.assertAlmostEqual(bobot["MKA-CPMK2"], 25.0)
        self.assertAlmostEqual(bobot["MKC-CPMK1"], 100.0)


class NilaiMahasiswaTest(unittest.TestCase):
    def test_nilai_ik_rata_rata_per_mk_pembawa(self):
        ik = rumus.student_ik(contoh_nilai(), contoh_mapping()).set_index("Kode IK")
        # IK1.1: MK A 80, MK B 70, MK C 90 -> 80. IK2.1 hanya dari CPMK multi-IK MK C.
        self.assertAlmostEqual(ik.loc["IK1.1", "Nilai IK"], 80.0)
        self.assertEqual(ik.loc["IK1.1", "Jumlah MK Dinilai"], 3)
        self.assertAlmostEqual(ik.loc["IK1.2", "Nilai IK"], 60.0)
        self.assertAlmostEqual(ik.loc["IK2.1", "Nilai IK"], 90.0)

    def test_nilai_cpl_berbobot_ik(self):
        cpl = rumus.student_cpl(contoh_nilai(), contoh_mapping()).set_index("Kode CPL")
        # CPL01 = (80 x 3 + 60 x 1) / 4 = 75.
        self.assertAlmostEqual(cpl.loc["CPL01", "Nilai CPL"], 75.0)
        self.assertAlmostEqual(cpl.loc["CPL02", "Nilai CPL"], 90.0)

    def test_mapping_satu_baris_per_ik_sama_hasilnya(self):
        mapping = contoh_mapping()
        baris = mapping.iloc[[3]].copy()
        mapping = pd.concat([mapping.iloc[:3], baris.assign(**{"Kode IK": "IK1.1"}), baris.assign(**{"Kode IK": "IK2.1"})])
        cpl = rumus.student_cpl(contoh_nilai(), mapping).set_index("Kode CPL")
        self.assertAlmostEqual(cpl.loc["CPL01", "Nilai CPL"], 75.0)
        self.assertAlmostEqual(cpl.loc["CPL02", "Nilai CPL"], 90.0)

    def test_kode_cpmk_lama_dipakai_banyak_mk_tidak_tercampur(self):
        mapping = pd.DataFrame(
            {
                "Kode MK": ["MKA", "MKB"],
                "Kode CPMK": ["CPMK1.1", "CPMK1.1"],
                "Kode CPL": ["CPL01", "CPL01"],
                "Kode IK": ["IK1.1", "IK1.1"],
            }
        )
        nilai = pd.DataFrame(
            {"NIM": ["1", "1"], "Kode MK": ["MKA", "MKB"], "Kode CPMK": ["CPMK1.1", "CPMK1.1"], "Nilai": [90, 50]}
        )
        ik = rumus.student_ik(nilai, mapping)
        self.assertAlmostEqual(ik["Nilai IK"].iloc[0], 70.0)
        self.assertEqual(ik["Jumlah MK Dinilai"].iloc[0], 2)


class CalculateAllTest(unittest.TestCase):
    def test_rekap_memakai_rumus_baru(self):
        mapping = contoh_mapping()
        mapping["Komponen Asesmen"] = ""
        mapping["Bobot Komponen"] = None
        nilai = contoh_nilai()
        nilai["Mata Kuliah"] = ""
        nilai["Kode CPL"] = mapping["Kode CPL"]
        nilai["Kode IK"] = mapping["Kode IK"]
        data = {
            "Master_CPL": pd.DataFrame({"Kode CPL": ["CPL01", "CPL02"], "Rumusan CPL": ["x", "y"], "GA IABEE": ["", ""]}),
            "Master_IK": pd.DataFrame(
                {"Kode CPL": ["CPL01", "CPL01", "CPL02"], "Kode IK": ["IK1.1", "IK1.2", "IK2.1"], "Rumusan IK": ["a", "b", "c"]}
            ),
            "Mapping_CPMK": mapping,
            "Nilai_CPMK": nilai,
            "Target": pd.DataFrame({"Parameter": ["Batas Nilai Minimum"], "Nilai": [70]}),
        }
        result = app.calculate_all(data, 70, 70)
        ik = result["rekap_ik"].set_index("Kode IK")
        self.assertAlmostEqual(ik.loc["IK1.1", "Capaian IK"], 80.0)
        self.assertEqual(ik.loc["IK1.1", "Jumlah MK Pembawa"], 3)
        self.assertAlmostEqual(ik.loc["IK1.1", "Bobot dalam CPL (%)"], 75.0)
        cpl = pd.DataFrame(result["rekap_cpl"]).set_index("Kode CPL")
        self.assertAlmostEqual(cpl.loc["CPL01", "Rata-rata Nilai CPL"], 75.0)
        self.assertAlmostEqual(cpl.loc["CPL02", "Rata-rata Nilai CPL"], 90.0)
        cpmk = result["rekap_cpmk"].set_index("Kode CPMK")
        self.assertAlmostEqual(cpmk.loc["MKA-CPMK1", "Bobot CPMK Rumus (%)"], 75.0)
        radar = app.calculate_student_cpl_table(data).set_index("Kode CPL Radar")
        self.assertAlmostEqual(radar.loc["CPL01", "Nilai CPL"], 75.0)


if __name__ == "__main__":
    unittest.main()
