import unittest

from report_requests import (
    is_download_followup, localize_source_links, parse_report_tool_result, report_type_for,
)


class ReportRequestTests(unittest.TestCase):
    def test_sdk_structured_mirror_does_not_break_json_parsing(self):
        self.assertEqual(
            parse_report_tool_result(['{"status":"success"}', '{"result":"mirror"}']),
            {"status": "success"},
        )

    def test_invalid_report_tool_results_raise(self):
        for texts in ([], ["null"], ["not JSON"]):
            with self.subTest(texts=texts), self.assertRaises(ValueError):
                parse_report_tool_result(texts)

    def test_weekly_report(self):
        self.assertEqual(
            report_type_for("Generate the weekly vendor and commitment status report "
                            "across all active vendor contracts."),
            "weekly_vendor_status",
        )

    def test_readiness_report(self):
        self.assertEqual(report_type_for("Generate the Joint Commission readiness report"),
                         "jc_readiness")

    def test_non_report_question(self):
        self.assertIsNone(report_type_for("What are the vendor risks?"))

    def test_download_followup(self):
        self.assertTrue(is_download_followup("yes please give the download link"))
        self.assertTrue(is_download_followup("Can you share the download link again?"))

    def test_unrelated_links_are_not_report_followups(self):
        self.assertFalse(is_download_followup("Show the policy document link"))
        self.assertFalse(is_download_followup("Download the medication spreadsheet"))

    def test_simulated_source_url_becomes_local_preview(self):
        url = "https://northbridge-health.sharepoint.com/checklist.xlsx"
        index = {"FILE-004": ("file", {"url": url})}
        self.assertEqual(
            localize_source_links(f"[Checklist]({url})", index),
            "[Checklist](/citation/file/FILE-004)",
        )

    def test_download_and_unknown_urls_are_preserved(self):
        answer = "[Report](/download/report/report.xlsx) https://example.com"
        self.assertEqual(localize_source_links(answer, {}), answer)


if __name__ == "__main__":
    unittest.main()
