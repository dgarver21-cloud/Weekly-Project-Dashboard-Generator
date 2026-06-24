import unittest
from io import BytesIO
from pathlib import Path

from docx import Document

from parse_meeting_transcript import (
    clean_transcript_text,
    clean_transcript_artifacts,
    combine_transcript_results,
    is_high_quality_project_bullet,
    is_irrelevant_transcript_content,
    parse_meeting_transcript,
    parse_meeting_transcripts,
    parse_transcript_text,
    rewrite_as_project_bullet,
    score_project_relevance,
)


SAMPLE_TRANSCRIPT = """
Angelica: We completed the finance data validation and reviewed configuration.
PM: Next week, we plan to prepare the team for UAT activities.
Alex: The integration design is blocked while we are waiting on client feedback.
PM: We need confirmation on the final interface approach.
Angelica: Action item: Alex is responsible for the fixed assets follow up by Friday.
"""


class MeetingTranscriptParserTests(unittest.TestCase):
    def test_speaker_names_and_timestamps_are_removed(self):
        rough = (
            "Dylan Garver 5:28. Yeah, I will follow up on the BC configuration. "
            "Daniel Pearce 7:21. Henry Virkler 7:22. Right, confirm the UAT plan. "
            "Ramesh Kumar 1:00:12. Okay, review the data validation results."
        )
        cleaned = clean_transcript_artifacts(rough)

        for unwanted in (
            "Dylan Garver",
            "Daniel Pearce",
            "Henry Virkler",
            "Ramesh Kumar",
            "5:28",
            "7:21",
            "7:22",
            "1:00:12",
        ):
            self.assertNotIn(unwanted, cleaned)
        self.assertFalse(cleaned.casefold().startswith(("yeah", "right", "okay")))

    def test_rough_fragments_are_rewritten_as_pm_bullets(self):
        cases = (
            (
                "How is that? What's the plan to get that into test? The master data as well.",
                "upcoming_focus_areas",
                "Confirm the approach for moving master data and balances from production into the test environment.",
            ),
            (
                "Yeah, I'll ask him around that. I should say, what's the plan to get that then from prod to test?",
                "upcoming_focus_areas",
                "Confirm the approach for moving master data and balances from production into the test environment.",
            ),
            (
                "My takeaway yesterday is that we need a shell of BC that holds the old chart of accounts.",
                "decisions_needed",
                "Determine whether a Business Central shell company is needed to retain the legacy chart of accounts.",
            ),
            (
                "Obviously, the interview data payroll Word document, we'll get that updated for you.",
                "possible_raid_items_for_review",
                "Update the payroll integration document and confirm any remaining open questions.",
            ),
            (
                "See if it's making sense or not, or you need any custom reports in the future.",
                "decisions_needed",
                "Review reporting needs and determine whether any custom reports are required.",
            ),
        )
        for rough, category, expected in cases:
            with self.subTest(rough=rough):
                self.assertEqual(
                    rewrite_as_project_bullet(rough, category), expected
                )

    def test_final_quality_gate_rejects_vague_conversational_bullets(self):
        rejected = (
            "Confirm the plan to get that into test.",
            "Address yep, can you go back to Power BI, sorry, Business Central again.",
            "Address master data with shelf numbers.",
            "Address I was thinking that we could schedule a call once it is set up.",
            "Follow up on you said it is in the production environment.",
            "Follow up on fixed assets is kind of up in the air.",
        )
        for bullet in rejected:
            with self.subTest(bullet=bullet):
                self.assertFalse(is_high_quality_project_bullet(bullet))

    def test_topic_grouping_rewrites_requested_project_topics(self):
        transcript = """
        What's the plan to get master data and balances from prod into test?
        Yep, can you go back to Power BI, sorry, Business Central again once test is updated?
        We need a BC shell company that retains the old chart of accounts.
        Fixed assets is kind of up in the air and the remaining data needs are unclear.
        """
        result = parse_transcript_text(transcript)
        all_items = [
            item
            for category in (
                "accomplishments",
                "upcoming_focus_areas",
                "blockers_or_concerns",
                "decisions_needed",
                "possible_raid_items_for_review",
            )
            for item in result[category]
        ]
        bullets = {item["topic"]: item["text"] for item in all_items}

        self.assertEqual(
            bullets["data_environment"],
            "Confirm the approach for moving master data and balances from production into the test environment.",
        )
        self.assertEqual(
            bullets["power_bi_business_central"],
            "Review Power BI and Business Central configuration once the test environment is updated.",
        )
        self.assertEqual(
            bullets["shell_company_chart_of_accounts"],
            "Determine whether a Business Central shell company is needed to retain the legacy chart of accounts.",
        )
        self.assertEqual(
            bullets["fixed_assets"],
            "Confirm the fixed assets configuration approach and remaining data needs before testing begins.",
        )

    def test_general_power_bi_review_is_not_a_blocker(self):
        result = parse_transcript_text(
            "Review Power BI and Business Central configuration once the test "
            "environment is updated."
        )

        self.assertTrue(result["upcoming_focus_areas"])
        self.assertFalse(result["blockers_or_concerns"])

    def test_same_topic_is_not_repeated_across_categories(self):
        first = parse_meeting_transcript(
            b"Plan to move master data and balances from production into test.",
            "planning.txt",
        )
        second = parse_meeting_transcript(
            b"We need to confirm the decision for moving master data into test.",
            "decision.txt",
        )
        result = combine_transcript_results([first, second])
        topic_locations = [
            category
            for category in (
                "accomplishments",
                "upcoming_focus_areas",
                "blockers_or_concerns",
                "decisions_needed",
                "possible_raid_items_for_review",
            )
            for item in result[category]
            if item.get("topic") == "data_environment"
        ]

        self.assertEqual(topic_locations, ["decisions_needed"])

    def test_low_quality_fragments_are_dropped(self):
        result = parse_transcript_text(
            "Yep, can you go back to that? You said it is there. "
            "I was thinking we could schedule a call."
        )

        self.assertFalse(
            any(
                result[category]
                for category in (
                    "accomplishments",
                    "upcoming_focus_areas",
                    "blockers_or_concerns",
                    "decisions_needed",
                    "possible_raid_items_for_review",
                )
            )
        )
        self.assertEqual(result["suggested_project_status_narrative"], "")

    def test_deduplication_happens_after_rewriting_related_data_items(self):
        first = (
            "What's the plan to move master data into the test environment? "
            "The team will continue Business Central validation."
        )
        second = (
            "Confirm the plan to move balances from production to test. "
            "The team will continue Business Central validation."
        )
        result = parse_meeting_transcripts(
            [
                (first.encode("utf-8"), "first.txt"),
                (second.encode("utf-8"), "second.txt"),
            ]
        )
        data_to_test = [
            item
            for item in result["upcoming_focus_areas"]
            if "test environment" in item["text"].casefold()
        ]

        self.assertEqual(len(data_to_test), 1)
        self.assertEqual(data_to_test[0]["source_files"], ["first.txt", "second.txt"])

    def test_main_bullet_text_excludes_sources_and_extraction_metadata(self):
        result = parse_meeting_transcript(
            SAMPLE_TRANSCRIPT.encode("utf-8"), "weekly-call.txt"
        )
        item = result["accomplishments"][0]

        self.assertNotIn("weekly-call.txt", item["text"])
        self.assertNotIn("Confidence", item["text"])
        self.assertNotIn(item["reason"], item["text"])
        self.assertIn("raw_text", item)

    def test_app_hides_raw_details_in_optional_expander(self):
        app_source = Path("app.py").read_text(encoding="utf-8")
        main_display = app_source.split("def show_transcript_items", 1)[1].split(
            "def show_transcript_details", 1
        )[0]

        self.assertIn("Show transcript extraction details", app_source)
        self.assertNotIn("confidence_score", main_display)
        self.assertNotIn("source_file", main_display)
        self.assertNotIn("raw_text", main_display)

    def test_suggested_narrative_uses_rewritten_bullets(self):
        transcript = (
            "Dylan Garver 5:28. We completed Business Central configuration validation. "
            "What's the plan to move master data into the test environment?"
        )
        result = parse_transcript_text(transcript)
        narrative = result["suggested_project_status_narrative"]

        self.assertIn("Key follow-ups include", narrative)
        self.assertIn("confirming the approach for moving master data", narrative.casefold())
        self.assertNotIn("Dylan Garver", narrative)
        self.assertNotIn("5:28", narrative)

    def test_single_transcript_batch_still_works(self):
        result = parse_meeting_transcripts(
            [(SAMPLE_TRANSCRIPT.encode("utf-8"), "weekly-call.txt")]
        )

        self.assertTrue(result["suggested_project_status_narrative"])
        self.assertEqual(
            result["accomplishments"][0]["source_file"], "weekly-call.txt"
        )

    def test_multiple_transcripts_are_combined(self):
        first = (
            "We completed the fixed assets validation. "
            "The team plans to prepare the UAT test cases."
        )
        second = (
            "The Munis integration is blocked while waiting on client feedback. "
            "Danvers needs to confirm the interface decision."
        )
        result = parse_meeting_transcripts(
            [
                (first.encode("utf-8"), "configuration-call.txt"),
                (second.encode("utf-8"), "integration-call.txt"),
            ]
        )

        self.assertTrue(result["accomplishments"])
        self.assertTrue(result["upcoming_focus_areas"])
        self.assertTrue(result["decisions_needed"])

    def test_one_bad_transcript_does_not_block_valid_files(self):
        result = parse_meeting_transcripts(
            [
                (SAMPLE_TRANSCRIPT.encode("utf-8"), "weekly-call.txt"),
                (b"not a supported transcript", "broken.pdf"),
            ]
        )

        self.assertTrue(result["accomplishments"])
        self.assertTrue(
            any("broken.pdf" in warning for warning in result["transcript_warnings"])
        )

    def test_duplicate_items_are_merged_and_sources_are_preserved(self):
        transcript = (
            "The team finalized the chart of accounts mapping. "
            "Next steps focus on preparing the UAT test cases."
        )
        result = parse_meeting_transcripts(
            [
                (transcript.encode("utf-8"), "call-one.txt"),
                (transcript.encode("utf-8"), "call-two.txt"),
            ]
        )

        matching_items = [
            item
            for item in result["accomplishments"]
            if "chart of accounts" in item["text"].casefold()
        ]
        self.assertEqual(len(matching_items), 1)
        self.assertEqual(
            matching_items[0]["source_files"], ["call-one.txt", "call-two.txt"]
        )
        self.assertIn("call-one.txt", matching_items[0]["source_file"])
        self.assertIn("call-two.txt", matching_items[0]["source_file"])

    def test_combined_narrative_stays_within_two_to_three_sentences(self):
        first = parse_meeting_transcript(
            (
                "We completed Business Central configuration validation. "
                "We plan to prepare UAT test cases."
            ).encode("utf-8"),
            "first.txt",
        )
        second = parse_meeting_transcript(
            (
                "The Munis integration is blocked waiting on client confirmation. "
                "The team needs to confirm the interface decision."
            ).encode("utf-8"),
            "second.txt",
        )
        result = combine_transcript_results([first, second])
        narrative = result["suggested_project_status_narrative"]
        sentence_count = sum(narrative.count(mark) for mark in ".!?")

        self.assertGreaterEqual(sentence_count, 2)
        self.assertLessEqual(sentence_count, 3)

    def test_parse_meeting_transcript_accepts_no_project_context(self):
        result = parse_meeting_transcript(
            SAMPLE_TRANSCRIPT.encode("utf-8"), "weekly-meeting.txt"
        )

        self.assertTrue(result["suggested_project_status_narrative"])

    def test_parse_meeting_transcript_accepts_app_project_context_keyword(self):
        result = parse_meeting_transcript(
            SAMPLE_TRANSCRIPT.encode("utf-8"),
            "weekly-meeting.txt",
            project_context={
                "project_name": "Danvers Business Central",
                "project_status_narrative": "Configuration is underway.",
                "adi_titles": ["Fixed Assets Configuration"],
                "clarizen_deliverable_names": ["UAT Preparation"],
            },
        )

        self.assertTrue(result["accomplishments"])
        self.assertTrue(result["upcoming_focus_areas"])

    def test_none_project_context_is_handled_safely(self):
        result = parse_meeting_transcript(
            SAMPLE_TRANSCRIPT.encode("utf-8"),
            "weekly-meeting.txt",
            project_context=None,
        )

        self.assertTrue(result["suggested_project_status_narrative"])

    def test_txt_transcript_extracts_all_review_categories(self):
        result = parse_meeting_transcript(
            SAMPLE_TRANSCRIPT.encode("utf-8"), "weekly-meeting.txt"
        )

        self.assertTrue(result["accomplishments"])
        self.assertTrue(result["upcoming_focus_areas"])
        self.assertTrue(result["decisions_needed"])
        self.assertTrue(result["possible_raid_items_for_review"])
        self.assertFalse(result["transcript_warnings"])
        for category in (
            "accomplishments",
            "upcoming_focus_areas",
            "blockers_or_concerns",
            "decisions_needed",
            "possible_raid_items_for_review",
        ):
            for item in result[category]:
                self.assertIn("text", item)
                self.assertEqual(item["category"], category)
                self.assertIn("confidence_score", item)
                self.assertIn("reason", item)

    def test_vtt_timestamps_cue_numbers_and_speaker_labels_are_removed(self):
        vtt = """WEBVTT

1
00:00:01.000 --> 00:00:05.000
<v Angelica>We completed data validation.</v>

2
00:00:06.000 --> 00:00:10.000
PM: Next week, we plan to prepare for UAT.
"""
        cleaned = clean_transcript_text(vtt, is_vtt=True)
        result = parse_meeting_transcript(vtt.encode("utf-8"), "meeting.vtt")

        self.assertNotIn("00:00", cleaned)
        self.assertNotIn("WEBVTT", cleaned)
        self.assertNotIn("Angelica", cleaned)
        self.assertTrue(result["accomplishments"])
        self.assertTrue(result["upcoming_focus_areas"])

    def test_docx_transcript_reads_paragraph_text(self):
        document = Document()
        for line in SAMPLE_TRANSCRIPT.strip().splitlines():
            document.add_paragraph(line)
        output = BytesIO()
        document.save(output)

        result = parse_meeting_transcript(output.getvalue(), "meeting.docx")

        self.assertTrue(result["accomplishments"])
        self.assertTrue(result["upcoming_focus_areas"])
        self.assertTrue(result["suggested_project_status_narrative"])

    def test_narrative_contains_two_to_three_sentences(self):
        result = parse_transcript_text(SAMPLE_TRANSCRIPT)
        narrative = result["suggested_project_status_narrative"]
        sentence_count = sum(narrative.count(mark) for mark in ".!?")

        self.assertGreaterEqual(sentence_count, 2)
        self.assertLessEqual(sentence_count, 3)
        self.assertNotIn("Angelica:", narrative)

    def test_categories_are_limited_to_five_items(self):
        transcript = "\n".join(
            f"We completed project task {index}." for index in range(10)
        ) + "\nNext week, we plan to prepare for UAT."
        result = parse_transcript_text(transcript)

        for category in (
            "accomplishments",
            "upcoming_focus_areas",
            "blockers_or_concerns",
            "decisions_needed",
            "possible_raid_items_for_review",
        ):
            self.assertLessEqual(len(result[category]), 5)

    def test_blank_transcript_returns_warning_and_no_narrative(self):
        result = parse_transcript_text("Okay. Thank you. Yep.")

        self.assertEqual(result["suggested_project_status_narrative"], "")
        self.assertTrue(result["transcript_warnings"])

    def test_low_quality_transcript_returns_warning(self):
        result = parse_transcript_text(
            "We reviewed the agenda. General conversation continued afterward."
        )

        self.assertEqual(result["suggested_project_status_narrative"], "")
        self.assertIn("high-confidence project status", result["transcript_warnings"][0])

    def test_pto_travel_and_doctor_appointment_comments_are_excluded(self):
        transcript = """
        Henry's out next week starting Tuesday through the 6th of July.
        John has a doctor's appointment throughout the middle of the day Tuesday.
        I won't see you for a few weeks because I'll be traveling.
        """
        result = parse_transcript_text(transcript)

        self.assertFalse(any(result[name] for name in result if name in (
            "accomplishments",
            "upcoming_focus_areas",
            "blockers_or_concerns",
            "decisions_needed",
            "possible_raid_items_for_review",
        )))
        self.assertEqual(result["suggested_project_status_narrative"], "")

    def test_general_meeting_scheduling_chatter_is_excluded(self):
        result = parse_transcript_text(
            "Can we move next week's meeting to Thursday? "
            "I will send a new calendar invite."
        )

        self.assertFalse(result["upcoming_focus_areas"])
        self.assertEqual(result["suggested_project_status_narrative"], "")

    def test_unrelated_employee_problem_is_excluded(self):
        result = parse_transcript_text(
            "I was invited to a meeting tomorrow about one of my other employees "
            "and a problem."
        )

        self.assertFalse(result["blockers_or_concerns"])
        self.assertFalse(result["possible_raid_items_for_review"])

    def test_project_relevant_fixed_assets_content_is_included(self):
        result = parse_transcript_text(
            "The team finalized the fixed assets configuration. "
            "Henry is out next week, which may delay completion of the fixed assets "
            "configuration."
        )

        self.assertTrue(result["blockers_or_concerns"])
        self.assertFalse(result["accomplishments"])
        self.assertFalse(
            is_irrelevant_transcript_content(
                "Henry is out next week, which may delay completion of the fixed "
                "assets configuration."
            )
        )

    def test_project_relevant_chart_of_accounts_content_is_included(self):
        result = parse_transcript_text(
            "Danvers needs to finalize the remaining chart of accounts mapping. "
            "The team needs to confirm the final COA structure."
        )

        self.assertTrue(result["decisions_needed"])

    def test_project_relevant_integration_content_is_included(self):
        result = parse_transcript_text(
            "Open Munis integration questions may impact UAT testing. "
            "The team plans to finalize the interface approach."
        )

        self.assertTrue(result["blockers_or_concerns"])
        self.assertTrue(result["upcoming_focus_areas"])

    def test_project_context_increases_relevance_score(self):
        text = "The team needs to finalize the remaining mapping."
        without_context = score_project_relevance(text)
        with_context = score_project_relevance(
            text,
            {"adi_item_titles": ["Remaining Finance Mapping"]},
        )

        self.assertGreater(with_context, without_context)

    def test_only_actionable_project_items_are_possible_raid_suggestions(self):
        transcript = """
        Follow up with John on the fixed assets list.
        Action item: create a Business Central shell company for the old COA.
        Follow up about lunch next week.
        """
        result = parse_transcript_text(transcript)
        raid_text = " ".join(
            item["text"] for item in result["possible_raid_items_for_review"]
        )

        self.assertIn("fixed assets", raid_text)
        self.assertNotIn("lunch", raid_text)


if __name__ == "__main__":
    unittest.main()
