import json
import os
import tempfile
import threading
import unittest
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import app
from openpyxl import load_workbook
from http.server import ThreadingHTTPServer


class CnamisSeedTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database = Path(self.temp_dir.name) / "cnamis-test.db"
        app.initialize_database(
            self.database,
            (app.SURVEY_FILE, app.ROSTER_FILE, app.ANALYSIS_FILE),
        )

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_seeded_roster_and_response_reconciliation(self):
        with app.database_session(self.database) as connection:
            agency_id = app.agency_id_for(connection)
            dashboard = app.dashboard_data(connection, agency_id)
            self.assertEqual(dashboard["roster_count"], 77)
            self.assertEqual(dashboard["cna_responses"], 39)
            self.assertEqual(dashboard["matched_participants"], 32)
            self.assertEqual(dashboard["verify_name"], 2)
            self.assertEqual(dashboard["not_in_cna"], 43)
            self.assertEqual(len(app.possible_matches(connection, agency_id)), 2)

    def test_diagnostic_gap_uses_ten_point_cna_scale(self):
        source = load_workbook(app.SURVEY_FILE, data_only=True)
        survey = source.active
        scores = [
            survey.cell(row, 15).value
            for row in range(2, survey.max_row + 1)
            if survey.cell(row, 3).value and isinstance(survey.cell(row, 15).value, (int, float))
        ]
        source.close()
        expected_mean = sum(scores) / len(scores)
        with app.database_session(self.database) as connection:
            agency_id = app.agency_id_for(connection)
            rows = app.item_analytics(connection, agency_id)
            self.assertEqual(len(rows), 20)
            self.assertAlmostEqual(rows[0]["mean_score"], expected_mean)
            self.assertAlmostEqual(rows[0]["gap"], 10 - rows[0]["mean_score"])
            self.assertEqual(rows[0]["band"], "Average")

    def test_five_year_programme_is_seeded_without_invented_costs(self):
        with app.database_session(self.database) as connection:
            agency_id = app.agency_id_for(connection)
            programme = app.programme_data(connection, agency_id)
            self.assertEqual(len(programme), 35)
            self.assertEqual({row["year"] for row in programme}, {2027, 2028, 2029, 2030, 2031})
            self.assertTrue(all(row["estimated_budget"] is None for row in programme))
            self.assertTrue(all(row["actual_expenditure"] is None for row in programme))

    def test_workspace_answers_with_source_and_qualification(self):
        with app.database_session(self.database) as connection:
            agency_id = app.agency_id_for(connection)
            result = app.answer_workspace(connection, agency_id, "How many roster officers did not match?")
            self.assertIn("43 have no confirmed matching CNA name", result["answer"])
            self.assertTrue(result["sources"])

            unsupported = app.answer_workspace(connection, agency_id, "What should the budget be?")
            self.assertIn("unpriced", unsupported["answer"])

    def test_report_states_evidence_limited_agency_compliance(self):
        with app.database_session(self.database) as connection:
            agency_id = app.agency_id_for(connection)
            report = app.capability_report_data(connection, agency_id)
            self.assertIn("has complied with the agency-level CNA submission", report["compliance_statement"])
            self.assertIn("does not certify complete individual participation", report["participation_caveat"])
            self.assertEqual(report["dashboard"]["matched_participants"], 32)
            self.assertEqual(report["dashboard"]["verify_name"], 2)
            self.assertEqual(report["dashboard"]["not_in_cna"], 43)
            self.assertEqual(len(report["diagnostic"]), 20)
            self.assertEqual(len(report["programme"]), 35)


class CnamisHttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.database = Path(cls.temp_dir.name) / "cnamis-http-test.db"
        cls.previous_database = os.environ.get("CNAMIS_DB_PATH")
        os.environ["CNAMIS_DB_PATH"] = str(cls.database)
        app.initialize_database(cls.database, (app.SURVEY_FILE, app.ROSTER_FILE, app.ANALYSIS_FILE))
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), app.CnamisHandler)
        cls.base_url = f"http://127.0.0.1:{cls.server.server_address[1]}"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=5)
        if cls.previous_database is None:
            os.environ.pop("CNAMIS_DB_PATH", None)
        else:
            os.environ["CNAMIS_DB_PATH"] = cls.previous_database
        cls.temp_dir.cleanup()

    def get_json(self, path):
        with urlopen(self.base_url + path, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))

    def post_json(self, path, body):
        request = Request(
            self.base_url + path,
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urlopen(request, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))

    def patch_json(self, path, body):
        request = Request(
            self.base_url + path,
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="PATCH",
        )
        with urlopen(request, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))

    def test_programme_and_monitoring_edits_persist(self):
        activity = self.get_json("/api/programme?agency=OPP")[0]
        activity_id = activity["activity_id"]
        response = self.patch_json(
            f"/api/programme/{activity_id}",
            {"agency": "OPP", "estimated_budget": 1250.5, "approval_status": "Submitted for review"},
        )
        self.assertTrue(response["ok"])
        updated = next(row for row in self.get_json("/api/programme?agency=OPP") if row["activity_id"] == activity_id)
        self.assertEqual(updated["estimated_budget"], 1250.5)
        self.assertEqual(updated["approval_status"], "Submitted for review")

        self.post_json("/api/monitoring", {
            "agency": "OPP", "activity_id": activity_id, "planned_participants": 8,
            "participants_completed": 5, "actual_expenditure": 900.25, "indicator": "Completion",
        })
        monitored = next(row for row in self.get_json("/api/programme?agency=OPP") if row["activity_id"] == activity_id)
        self.assertEqual(monitored["participants_completed"], 5)
        self.assertEqual(monitored["actual_expenditure"], 900.25)

    def test_ambiguous_match_verification_is_limited_to_candidates(self):
        possible = self.get_json("/api/matches?agency=OPP")[0]
        candidate = possible["candidates"][0]
        result = self.post_json("/api/matches/verify", {
            "agency": "OPP", "response_id": possible["response_id"], "staff_id": candidate["id"],
        })
        self.assertTrue(result["ok"])
        dashboard = self.get_json("/api/dashboard?agency=OPP")
        self.assertEqual(dashboard["matched_participants"], 33)
        self.assertEqual(dashboard["verify_name"], 1)

    def test_excel_report_download_is_a_valid_workbook(self):
        with urlopen(self.base_url + "/api/export.xlsx?agency=OPP", timeout=10) as response:
            self.assertEqual(response.status, 200)
            workbook_data = response.read()
        from io import BytesIO
        report = load_workbook(BytesIO(workbook_data), read_only=True, data_only=True)
        self.assertIn("Five-Year Plan", report.sheetnames)
        self.assertEqual(report["Five-Year Plan"].max_row, 36)
        report.close()

    def test_cna_report_download_returns_pdf(self):
        with urlopen(self.base_url + "/api/report.cna.pdf?agency=OPP", timeout=30) as response:
            self.assertEqual(response.status, 200)
            self.assertEqual(response.headers.get_content_type(), "application/pdf")
            self.assertIn("Capability-Needs-Analysis-Report.pdf", response.headers["Content-Disposition"])
            report = response.read()
        self.assertTrue(report.startswith(b"%PDF-"))
        self.assertGreater(len(report), 10000)

    def test_brand_logo_is_served_from_supplied_image(self):
        with urlopen(self.base_url + "/image/opp-logo.webp", timeout=10) as response:
            self.assertEqual(response.status, 200)
            self.assertEqual(response.headers.get_content_type(), "image/webp")
            logo_header = response.read(12)
        self.assertEqual(logo_header[:4], b"RIFF")
        self.assertEqual(logo_header[8:12], b"WEBP")

    def test_negative_budget_is_rejected(self):
        activity = self.get_json("/api/programme?agency=OPP")[0]
        request = Request(
            f"{self.base_url}/api/programme/{activity['activity_id']}",
            data=json.dumps({"agency": "OPP", "estimated_budget": -1}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="PATCH",
        )
        try:
            urlopen(request, timeout=10)
        except HTTPError as error:
            with error:
                self.assertEqual(error.code, 400)
        else:
            self.fail("Negative budget update should be rejected.")


if __name__ == "__main__":
    unittest.main()
