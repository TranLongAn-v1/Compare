import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from modules.module_01_bh_02_bh import logic_01_bh_02_bh as logic


PDF_SAMPLE = Path(__file__).resolve().parents[1] / "file_xlsx/rptTH_DENGHITHANHTOAN_MS02_TT122026_VTU.pdf"
EXPECTED = {
    "so_luot": Decimal("7713"),
    "tong_chi_bv": Decimal("4611581257"),
    "tong_chi_bh": Decimal("4611581257"),
    "bao_hiem_tt": Decimal("4384141284"),
    "benh_nhan_cct": Decimal("227439973"),
}


class Mau02PdfTests(unittest.TestCase):
    def test_sample_totals(self):
        self.assertEqual(logic._doc_tong_mau02(str(PDF_SAMPLE)), EXPECTED)

    def test_shifted_layout(self):
        items = [(y + 20, x * 1.4 + 70, text) for y, x, text in logic._trich_text_pdf(PDF_SAMPLE)]
        with patch.object(logic, "_trich_text_pdf", return_value=items):
            self.assertEqual(logic._doc_tong_mau02_pdf("shifted.pdf"), EXPECTED)

    def test_missing_amount_raises_instead_of_returning_zero(self):
        items = [item for item in logic._trich_text_pdf(PDF_SAMPLE)
                 if not (340 < item[0] < 350 and item[2] == "227.439.973")]
        with patch.object(logic, "_trich_text_pdf", return_value=items):
            with self.assertRaises(ValueError):
                logic._doc_tong_mau02_pdf("incomplete.pdf")

    def test_explicit_zero_is_valid(self):
        items = [(y, x, "0" if 340 < y < 350 and text == "227.439.973" else text)
                 for y, x, text in logic._trich_text_pdf(PDF_SAMPLE)]
        with patch.object(logic, "_trich_text_pdf", return_value=items):
            self.assertEqual(logic._doc_tong_mau02_pdf("zero.pdf"), {**EXPECTED, "benh_nhan_cct": Decimal("0")})

    def test_missing_header_raises(self):
        items = [item for item in logic._trich_text_pdf(PDF_SAMPLE) if item[2] != "(5)"]
        with patch.object(logic, "_trich_text_pdf", return_value=items):
            with self.assertRaises(ValueError):
                logic._doc_tong_mau02_pdf("incomplete.pdf")

    def test_pdf_byte_escapes(self):
        cmap = {8: "b", 12: ")", 92: "y", 40: "(", 41: ")"}
        self.assertEqual(logic._decode_pdf_text(b"\x00\\b\x00\\f\x00\\\\\x00\\(\x00\\)", cmap), "b)y()")
        self.assertEqual(logic._decode_pdf_text(b"\x00\\010\x00\\014", cmap), "b)")

    def test_compare_and_export_with_sample(self):
        # A real Excel input exercises the complete comparison path. All costs
        # are placed on the first visit; remaining visits have zero cost.
        df = pd.DataFrame(0, index=range(7713), columns=["Tong chi BV", "Tong chi BH", "Bao hiem TT", "Benh nhan CCT"])
        df.iloc[0] = [4611581257, 4611581257, 4384141284, 227439973]
        with tempfile.TemporaryDirectory() as directory:
            mau01 = str(Path(directory) / "Mau01_test.xlsx")
            df.to_excel(mau01, index=False)
            success, result = logic.doi_chieu_01_bh_02_bh([mau01], str(PDF_SAMPLE))
            self.assertTrue(success, result)
            self.assertEqual(result["Trạng thái"].tolist(), ["Khớp"] * 5)
            self.assertEqual(result["Tổng Mẫu 02"].tolist(), ["7.713", "4.611.581.257", "4.611.581.257", "4.384.141.284", "227.439.973"])
            output = logic.xuat_excel_01_bh_02_bh(result, str(Path(directory) / "comparison.xlsx"))
            exported = pd.read_excel(output)
            self.assertEqual(exported["Trạng thái"].tolist(), ["Khớp"] * 5)


if __name__ == "__main__":
    unittest.main()
