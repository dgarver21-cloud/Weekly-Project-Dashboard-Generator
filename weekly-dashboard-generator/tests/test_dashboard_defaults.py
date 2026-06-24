import unittest
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from dashboard_defaults import current_report_date


class DashboardDefaultTests(unittest.TestCase):
    def test_report_date_defaults_to_today(self):
        try:
            expected = datetime.now(ZoneInfo("America/New_York")).date()
        except ZoneInfoNotFoundError:
            expected = date.today()

        self.assertEqual(current_report_date(), expected)

    def test_streamlit_uses_today_helper_and_keeps_date_input_editable(self):
        app_source = Path("app.py").read_text(encoding="utf-8")

        self.assertIn("default_report_date = current_report_date()", app_source)
        self.assertIn('st.date_input("Report Date", value=default_report_date)', app_source)
        self.assertNotIn("October 10, 2025", app_source)
        self.assertNotIn("2025-10-10", app_source)

    def test_project_information_widgets_default_blank(self):
        app_source = Path("app.py").read_text(encoding="utf-8")

        self.assertIn('st.text_input("Project Name", value="")', app_source)
        self.assertIn('st.text_input("Project Manager", value="")', app_source)
        self.assertIn('st.text_input("Project Sponsor", value="")', app_source)
        self.assertIn('st.date_input("Go-Live Date", value=None)', app_source)
        self.assertIn('st.session_state["project_status_narrative"] = ""', app_source)
        for demo_value in ("Test Project", "Test Test", "Dylan Garver"):
            self.assertNotIn(demo_value, app_source)

        self.assertNotIn("% Complete, Per WIP Hours", app_source)
        self.assertNotIn("percent_complete", app_source)

    def test_streamlit_has_optional_client_logo_uploader(self):
        app_source = Path("app.py").read_text(encoding="utf-8")

        self.assertIn('st.subheader("Client Logo")', app_source)
        self.assertIn('"Upload Client Logo"', app_source)
        self.assertIn('type=["png", "jpg", "jpeg"]', app_source)
        self.assertIn("PNG with a transparent background is recommended", app_source)
        self.assertIn("client_logo_bytes=", app_source)

    def test_streamlit_download_uses_fully_saved_powerpoint_bytes(self):
        app_source = Path("app.py").read_text(encoding="utf-8")
        generator_source = Path("generate_dashboard.py").read_text(encoding="utf-8")

        self.assertIn("pptx_bytes = build_dashboard_bytes(", app_source)
        self.assertIn("tempfile.TemporaryDirectory", generator_source)
        self.assertIn("_save_presentation_safely(presentation, output)", generator_source)
        self.assertIn("return output.read_bytes()", generator_source)


if __name__ == "__main__":
    unittest.main()
