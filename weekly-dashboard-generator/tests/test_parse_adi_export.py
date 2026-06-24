from datetime import date
import unittest

from parse_adi_export import (
    AdiExportError,
    build_description,
    extract_latest_update,
    parse_row,
    read_csv_text,
    summarize_current_status,
)


class ExtractLatestUpdateTests(unittest.TestCase):
    def test_multiple_dated_updates_uses_first_only(self):
        notes = (
            "2/6 - Cogsdale is coordinating with Danvers on open questions.\n"
            "1/22 - The initial design discussion was completed."
        )
        self.assertEqual(
            extract_latest_update(notes),
            "Cogsdale is coordinating with Danvers on open questions.",
        )

    def test_date_prefix_is_removed(self):
        self.assertEqual(
            extract_latest_update("01/22 - Danvers is reviewing the mapping."),
            "Danvers is reviewing the mapping.",
        )

    def test_notes_without_date_prefix_are_cleaned(self):
        self.assertEqual(
            extract_latest_update("  Danvers is updating   the templates.  "),
            "Danvers is updating the templates.",
        )

    def test_multiple_paragraphs_are_collapsed(self):
        notes = "2/6 - Danvers is reviewing the design.\n\nApproval is expected Friday."
        self.assertEqual(
            extract_latest_update(notes),
            "Danvers is reviewing the design. Approval is expected Friday.",
        )


class SummarizeCurrentStatusTests(unittest.TestCase):
    def test_blank_notes_return_default(self):
        self.assertEqual(
            summarize_current_status("Chart of Accounts", ""),
            "No recent update provided.",
        )

    def test_we_discussed_chart_of_accounts_is_rewritten(self):
        self.assertEqual(
            summarize_current_status(
                "Chart of Accounts",
                "We discussed the Chart of Accounts on our 1/21 call. "
                "Angelica confirmed that they want to keep the GL numbers within the COA.",
            ),
            "Danvers is updating the chart of accounts structure based on the latest project feedback.",
        )

    def test_cogsdale_requested_munis_meeting_is_rewritten(self):
        self.assertEqual(
            summarize_current_status(
                "Munis Interface Design",
                "Cogsdale requested a meeting to discuss a number of open questions "
                "regarding the Munis Integration.",
            ),
            "Cogsdale is coordinating with Danvers to resolve open Munis integration design questions.",
        )

    def test_we_confirmed_fixed_assets_is_neutralized(self):
        self.assertEqual(
            summarize_current_status(
                "Fixed Assets",
                "We confirmed Fixed Assets as a dimension.",
            ),
            "The team confirmed Fixed Assets as a dimension.",
        )

    def test_danvers_targeting_finance_template_is_rewritten(self):
        self.assertEqual(
            summarize_current_status(
                "Data Preparation Templates - Finance",
                "Danvers is targeting to deliver the Finance template by 2/7.",
            ),
            "Danvers is targeting completion of the finance data preparation template.",
        )

    def test_generic_discussion_uses_complete_neutral_sentence(self):
        summary = summarize_current_status(
            "Security Roles",
            "We discussed the security-role approach on our call.",
        )
        self.assertEqual(
            summary,
            "The team is reviewing Security Roles based on the latest project discussion.",
        )

    def test_very_long_notes_are_truncated_cleanly(self):
        notes = "2/6 - " + ("Danvers is reviewing project feedback and next steps " * 8)
        summary = summarize_current_status("Finance Templates", notes)
        self.assertLessEqual(len(summary), 141)
        self.assertTrue(summary.endswith("."))
        self.assertFalse(summary.endswith(" ."))

    def test_description_has_status_label_and_safety_limit(self):
        description = build_description(
            "Data Preparation Templates - Finance",
            "Danvers is targeting completion of the finance data preparation templates.",
        )
        self.assertIn("\nStatus: ", description)
        self.assertLessEqual(len(description), 220)


class AdiCsvColumnTests(unittest.TestCase):
    REQUIRED_HEADERS = (
        "Action Items,Status,Owner,Priority,Notes,Modified,Due Date,"
        "TestRail/JIRA Link"
    )
    SAMPLE_VALUES = (
        "Chart of Accounts,In Progress,Angelica,2. High,"
        "Danvers is reviewing the structure.,06/09/2026,06/15/2026,"
        "https://example.test/item"
    )

    def test_csv_with_attachments_still_works(self):
        rows = read_csv_text(
            self.REQUIRED_HEADERS
            + ",Attachments\n"
            + self.SAMPLE_VALUES
            + ",2\n"
        )
        parsed = parse_row(rows[0], row_number=2, today=date(2026, 6, 9))
        self.assertEqual(parsed.raid_item["attachments"], "2")

    def test_csv_without_attachments_works(self):
        rows = read_csv_text(
            self.REQUIRED_HEADERS + "\n" + self.SAMPLE_VALUES + "\n"
        )
        parsed = parse_row(rows[0], row_number=2, today=date(2026, 6, 9))
        self.assertEqual(parsed.raid_item["attachments"], "")

    def test_missing_required_column_has_clear_error(self):
        headers = self.REQUIRED_HEADERS.replace(",Owner", "")
        values = self.SAMPLE_VALUES.replace(",Angelica", "")
        with self.assertRaisesRegex(AdiExportError, "Owner"):
            read_csv_text(headers + "\n" + values + "\n")


if __name__ == "__main__":
    unittest.main()
