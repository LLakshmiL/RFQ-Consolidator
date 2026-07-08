import unittest
import zipfile
from io import BytesIO
from pathlib import Path

from streamlit.testing.v1 import AppTest

REPO_ROOT = Path(__file__).resolve().parent.parent
APP_PATH = REPO_ROOT / "webapp" / "streamlit_app.py"
SAMPLE1 = REPO_ROOT / "Sample1.xlsx"
SAMPLE2 = REPO_ROOT / "Sample2.xlsx"


@unittest.skipUnless(SAMPLE1.exists() and SAMPLE2.exists(), "sample workbooks not present")
class TestStreamlitApp(unittest.TestCase):
    def test_upload_consolidate_and_download_flow(self):
        at = AppTest.from_file(str(APP_PATH))
        at.run(timeout=30)
        self.assertFalse(at.exception)

        uploader = at.file_uploader[0]
        uploader.upload(SAMPLE1.name, SAMPLE1.read_bytes())
        uploader.upload(SAMPLE2.name, SAMPLE2.read_bytes())
        uploader.run(timeout=30)
        self.assertFalse(at.exception)

        at.button[0].click().run(timeout=60)
        self.assertFalse(at.exception)

        self.assertEqual(len(at.success), 1)
        self.assertIn("24/25", at.success[0].value)

        summary_df = at.dataframe[0].value
        self.assertEqual(len(summary_df), 25)
        self.assertEqual((summary_df["status"] == "processed").sum(), 24)
        self.assertEqual((summary_df["status"] == "skipped_hidden").sum(), 1)

        self.assertEqual(len(at.download_button), 1)
        self.assertFalse(at.download_button[0].disabled)

        # The download button's data comes straight from session_state, built
        # by the same zip-writing code exercised in test_consolidate.py — spot
        # check it here rather than reaching into Streamlit's internal proto.
        with zipfile.ZipFile(BytesIO(at.session_state["result"]["zip_bytes"])) as zf:
            names = set(zf.namelist())
        self.assertEqual(
            names, {"consolidated_lcl_ratecard.xlsx", "consolidated_bid_matrix.xlsx", "audit_report.xlsx"}
        )

    def test_consolidate_button_disabled_without_upload(self):
        at = AppTest.from_file(str(APP_PATH))
        at.run(timeout=30)
        self.assertTrue(at.button[0].disabled)


if __name__ == "__main__":
    unittest.main()
