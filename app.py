from __future__ import annotations

import json
import mimetypes
import os
import re
import sqlite3
import unicodedata
from collections.abc import Iterable
from contextlib import contextmanager
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
from xml.sax.saxutils import escape

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
LOGO_FILE = ROOT / "image" / "National_emblem_of_Papua_New_Guinea_(variant).svg.webp"
DEFAULT_DB = ROOT / "cnamis.db"
SURVEY_FILE = ROOT / "Office of Public Presecutors (Responses).xlsx"
ROSTER_FILE = ROOT / "Form 3.17 - Occupancy Report updated.xlsx"
ANALYSIS_FILE = ROOT / "Capability Needs Analysis Hub.xlsx"
HOST = "127.0.0.1"
PORT = int(os.environ.get("CNAMIS_PORT", "8766"))
AGENCY_CODE = "OPP"
AGENCY_NAME = "Office of the Public Prosecutor"
VALID_PLAN_STATUSES = {
    "Draft - review and approval required",
    "Submitted for review",
    "Approved",
    "In progress",
    "Completed",
    "On hold",
}


def db_path() -> Path:
    return Path(os.environ.get("CNAMIS_DB_PATH", str(DEFAULT_DB))).resolve()


def connect(path: Path | None = None) -> sqlite3.Connection:
    connection = sqlite3.connect(path or db_path(), timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


@contextmanager
def database_session(path: Path | None = None):
    connection = connect(path)
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def normalize(value: object) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode()
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", normalized.lower()).split())


def clean_name_tokens(value: object) -> list[str]:
    excluded = {"mr", "mrs", "ms", "miss", "dr", "gp", "jr", "sr"}
    return [word for word in normalize(value).split() if word not in excluded]


def parse_roster_name(value: object) -> tuple[str, list[str]]:
    parts = str(value or "").split(",")
    surname_words = clean_name_tokens(parts[0]) if parts else []
    given = clean_name_tokens(" ".join(parts[1:])) if len(parts) > 1 else []
    return (surname_words[0] if surname_words else ""), [word for word in given if len(word) > 1]


def score_band(score: float | None) -> str:
    if score is None:
        return "Not assessed"
    if score <= 3:
        return "Poor/Low"
    if score <= 8:
        return "Average"
    return "High/Excellent"


def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def create_schema(connection: sqlite3.Connection) -> None:
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS agencies (
            id INTEGER PRIMARY KEY,
            code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            source_note TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS staff (
            id INTEGER PRIMARY KEY,
            agency_id INTEGER NOT NULL REFERENCES agencies(id),
            canonical_key TEXT NOT NULL,
            name TEXT NOT NULL,
            position_numbers TEXT NOT NULL,
            positions TEXT NOT NULL,
            locations TEXT NOT NULL,
            occupant_types TEXT NOT NULL,
            source_rows TEXT NOT NULL,
            UNIQUE(agency_id, canonical_key)
        );
        CREATE TABLE IF NOT EXISTS assessment_items (
            id INTEGER PRIMARY KEY,
            agency_id INTEGER NOT NULL REFERENCES agencies(id),
            item_index INTEGER NOT NULL,
            item_code TEXT NOT NULL,
            dimension TEXT NOT NULL,
            statement TEXT NOT NULL,
            review_route TEXT NOT NULL,
            UNIQUE(agency_id, item_index)
        );
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY,
            agency_id INTEGER NOT NULL REFERENCES agencies(id),
            source_row INTEGER NOT NULL,
            staff_id INTEGER REFERENCES staff(id),
            candidate_staff_ids TEXT NOT NULL DEFAULT '[]',
            match_status TEXT NOT NULL,
            name TEXT NOT NULL,
            position TEXT NOT NULL,
            grade TEXT NOT NULL,
            division TEXT NOT NULL,
            employment_status TEXT NOT NULL,
            qualification TEXT NOT NULL,
            required_qualification TEXT NOT NULL,
            competencies TEXT NOT NULL,
            spa_rating REAL,
            needs_process TEXT NOT NULL,
            process_rating REAL,
            process_reason TEXT NOT NULL,
            needs_met REAL,
            assessment_methods TEXT NOT NULL,
            career_goal TEXT NOT NULL,
            interests TEXT NOT NULL,
            UNIQUE(agency_id, source_row)
        );
        CREATE TABLE IF NOT EXISTS item_scores (
            response_id INTEGER NOT NULL REFERENCES responses(id) ON DELETE CASCADE,
            item_index INTEGER NOT NULL,
            score REAL,
            PRIMARY KEY(response_id, item_index)
        );
        CREATE TABLE IF NOT EXISTS need_categories (
            id INTEGER PRIMARY KEY,
            agency_id INTEGER NOT NULL REFERENCES agencies(id),
            name TEXT NOT NULL,
            keywords TEXT NOT NULL,
            evidence_source TEXT NOT NULL,
            intervention_route TEXT NOT NULL,
            UNIQUE(agency_id, name)
        );
        CREATE TABLE IF NOT EXISTS programme (
            id INTEGER PRIMARY KEY,
            agency_id INTEGER NOT NULL REFERENCES agencies(id),
            activity_id TEXT NOT NULL,
            year INTEGER NOT NULL,
            category TEXT NOT NULL,
            evidence TEXT NOT NULL,
            focus TEXT NOT NULL,
            formal_learning TEXT NOT NULL,
            social_learning TEXT NOT NULL,
            experiential_learning TEXT NOT NULL,
            owner_role TEXT NOT NULL,
            timing TEXT NOT NULL,
            respondents_mentioning INTEGER NOT NULL DEFAULT 0,
            assessment_evidence TEXT NOT NULL,
            corporate_reference TEXT NOT NULL,
            estimated_budget REAL,
            quote_basis TEXT NOT NULL,
            approval_status TEXT NOT NULL,
            indicator TEXT NOT NULL,
            status TEXT NOT NULL,
            outcome_notes TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(agency_id, activity_id)
        );
        CREATE TABLE IF NOT EXISTS plan_options (
            id INTEGER PRIMARY KEY,
            agency_id INTEGER NOT NULL REFERENCES agencies(id),
            option_code TEXT NOT NULL,
            title TEXT NOT NULL,
            description TEXT NOT NULL,
            annual_schedule TEXT NOT NULL,
            UNIQUE(agency_id, option_code)
        );
        CREATE TABLE IF NOT EXISTS monitoring (
            programme_id INTEGER PRIMARY KEY REFERENCES programme(id) ON DELETE CASCADE,
            planned_participants INTEGER,
            participants_completed INTEGER,
            actual_expenditure REAL,
            indicator TEXT NOT NULL DEFAULT '',
            baseline TEXT NOT NULL DEFAULT '',
            target TEXT NOT NULL DEFAULT '',
            actual_result TEXT NOT NULL DEFAULT '',
            evidence TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY,
            entity TEXT NOT NULL,
            entity_id TEXT NOT NULL,
            action TEXT NOT NULL,
            changes TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS app_meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """
    )


def workbook_text(value: object) -> str:
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _insert_source_data(connection: sqlite3.Connection, survey_path: Path, roster_path: Path, analysis_path: Path) -> None:
    missing = [path for path in (survey_path, roster_path, analysis_path) if not path.is_file()]
    if missing:
        raise FileNotFoundError("Required source workbook(s) not found: " + ", ".join(str(path) for path in missing))

    agency_id = connection.execute(
        "INSERT INTO agencies(code, name, source_note) VALUES(?, ?, ?)",
        (
            AGENCY_CODE,
            AGENCY_NAME,
            "Seeded from the supplied OPP CNA survey, Form 3.17 occupancy report, and CNAMIS analysis workbook.",
        ),
    ).lastrowid

    roster = load_workbook(roster_path, data_only=True).active
    people: dict[str, dict[str, object]] = {}
    for row in range(16, 700):
        if not (roster.cell(row, 1).value or roster.cell(row, 2).value):
            continue
        for column, occupant_type in ((9, "Substantive occupant"), (15, "Actual occupant")):
            name = workbook_text(roster.cell(row, column).value)
            if not name or "vacant" in name.lower() or name.lower() == "name":
                continue
            surname, given = parse_roster_name(name)
            if not surname or not given:
                continue
            key = f"{surname}|{given[0]}"
            person = people.setdefault(
                key,
                {"name": name, "positions": set(), "position_numbers": set(), "locations": set(),
                 "occupant_types": set(), "rows": set()},
            )
            person["rows"].add(row)
            person["occupant_types"].add(occupant_type)
            for col, field in ((2, "position_numbers"), (4, "positions"), (19, "locations")):
                value = workbook_text(roster.cell(row, col).value)
                if value:
                    person[field].add(value)

    staff_ids: dict[str, int] = {}
    for key, person in people.items():
        cursor = connection.execute(
            """INSERT INTO staff
               (agency_id, canonical_key, name, position_numbers, positions, locations, occupant_types, source_rows)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                agency_id, key, person["name"],
                json.dumps(sorted(person["position_numbers"])),
                json.dumps(sorted(person["positions"])),
                json.dumps(sorted(person["locations"])),
                json.dumps(sorted(person["occupant_types"])),
                json.dumps(sorted(person["rows"])),
            ),
        )
        staff_ids[key] = cursor.lastrowid

    roster.parent.close()
    survey = load_workbook(survey_path, data_only=True).active
    question_headers = [workbook_text(survey.cell(1, col).value) for col in range(15, 35)]
    review_routes = [
        "Validate strategic-plan availability and alignment; management action may be needed.",
        "Review values communication and induction; training only if an awareness gap is confirmed.",
        "Check KRA/KPI cascade and practical performance-management support against approved KRAs.",
        "Review supervisor-led work planning and alignment; verify SPA/work-plan records.",
        "Check policy coverage and implementation; policy/governance review may be needed before training.",
        "Review process-improvement governance; management/process intervention may be needed.",
        "Improve role-to-objective communication if a knowledge gap is confirmed.",
        "Review SPA/development-review practice and its coverage.",
        "Check approved SOP availability; train only on current documented procedures.",
        "Confirm reporting lines, accountability and communication of responsibilities.",
        "Review leadership/succession arrangements; workforce policy action may be needed.",
        "Review access to cross-functional assignments and release arrangements.",
        "Review change-management arrangements and barriers; confirm skills needs separately.",
        "Review cross-unit collaboration practices.",
        "Review team climate and manager practices; avoid treating culture issues as training alone.",
        "Review induction coverage and content before assigning training.",
        "Review GESI systems and safeguards; institutional strengthening may be required.",
        "Review ethics standards, reporting and management practice.",
        "Review workload and wellbeing practices; management intervention may be more appropriate.",
        "Review OHS controls and role requirements; arrange required instruction only when confirmed.",
    ]
    dimension_for_index = (
        ["Strategic planning and performance management"] * 4
        + ["Governance, systems and performance review"] * 5
        + ["Leadership, succession and collaboration"] * 5
        + ["Workplace, GESI, ethics and safety"] * 6
    )
    for index, statement in enumerate(question_headers):
        code_match = re.match(r"\s*([A-D](?:\.?\d+)?)", statement)
        item_code = code_match.group(1).replace(".", "") if code_match else f"Item {index + 1}"
        connection.execute(
            """INSERT INTO assessment_items
               (agency_id, item_index, item_code, dimension, statement, review_route)
               VALUES(?, ?, ?, ?, ?, ?)""",
            (agency_id, index, item_code, dimension_for_index[index], statement, review_routes[index]),
        )

    for row in range(2, survey.max_row + 1):
        name = workbook_text(survey.cell(row, 3).value)
        if not name:
            continue
        survey_words = clean_name_tokens(name)
        exact_candidates: list[str] = []
        possible_candidates: list[str] = []
        for key, person in people.items():
            surname, given = parse_roster_name(person["name"])
            if surname in survey_words and set(given).intersection(survey_words[:-1]):
                exact_candidates.append(key)
            elif len(survey_words) == 1 and survey_words[0] in given:
                possible_candidates.append(key)
        if len(exact_candidates) == 1:
            match_status = "Matched"
            matched_staff = staff_ids[exact_candidates[0]]
            candidate_ids: list[int] = []
        else:
            matched_staff = None
            match_status = "Verify name" if exact_candidates or possible_candidates else "Not in CNA"
            choices = exact_candidates if exact_candidates else possible_candidates
            candidate_ids = [staff_ids[key] for key in choices]

        values = {
            "agency_id": agency_id,
            "source_row": row,
            "staff_id": matched_staff,
            "candidate_staff_ids": json.dumps(candidate_ids),
            "match_status": match_status,
            "name": name,
            "position": workbook_text(survey.cell(row, 6).value),
            "grade": workbook_text(survey.cell(row, 7).value),
            "division": workbook_text(survey.cell(row, 8).value),
            "employment_status": workbook_text(survey.cell(row, 9).value),
            "qualification": workbook_text(survey.cell(row, 10).value),
            "required_qualification": workbook_text(survey.cell(row, 12).value),
            "competencies": workbook_text(survey.cell(row, 13).value),
            "spa_rating": _numeric(survey.cell(row, 14).value, 1, 5),
            "needs_process": workbook_text(survey.cell(row, 35).value),
            "process_rating": _numeric(survey.cell(row, 36).value, 1, 10),
            "process_reason": workbook_text(survey.cell(row, 37).value),
            "needs_met": _numeric(survey.cell(row, 38).value, 1, 10),
            "assessment_methods": workbook_text(survey.cell(row, 39).value),
            "career_goal": workbook_text(survey.cell(row, 40).value),
            "interests": workbook_text(survey.cell(row, 41).value),
        }
        cursor = connection.execute(
            """INSERT INTO responses
               (agency_id, source_row, staff_id, candidate_staff_ids, match_status, name, position, grade,
                division, employment_status, qualification, required_qualification, competencies, spa_rating,
                needs_process, process_rating, process_reason, needs_met, assessment_methods, career_goal, interests)
               VALUES(:agency_id, :source_row, :staff_id, :candidate_staff_ids, :match_status, :name, :position,
                :grade, :division, :employment_status, :qualification, :required_qualification, :competencies,
                :spa_rating, :needs_process, :process_rating, :process_reason, :needs_met, :assessment_methods,
                :career_goal, :interests)""",
            values,
        )
        response_id = cursor.lastrowid
        for index in range(20):
            connection.execute(
                "INSERT INTO item_scores(response_id, item_index, score) VALUES(?, ?, ?)",
                (response_id, index, _numeric(survey.cell(row, 15 + index).value, 1, 10)),
            )

    survey.parent.close()
    analysis = load_workbook(analysis_path, data_only=True)
    needs_sheet = analysis["Training Needs"]
    for row in range(6, 13):
        category = workbook_text(needs_sheet.cell(row, 1).value)
        if not category:
            continue
        connection.execute(
            """INSERT INTO need_categories(agency_id, name, keywords, evidence_source, intervention_route)
               VALUES(?, ?, ?, ?, ?)""",
            (
                agency_id, category,
                json.dumps([token.strip().lower() for token in workbook_text(needs_sheet.cell(row, 3).value).split(",") if token.strip()]),
                workbook_text(needs_sheet.cell(row, 4).value),
                workbook_text(needs_sheet.cell(row, 7).value),
            ),
        )

    programme_sheet = analysis["Five-Year Programme"]
    for row in range(6, 41):
        activity_id = workbook_text(programme_sheet.cell(row, 1).value)
        if not activity_id:
            continue
        connection.execute(
            """INSERT INTO programme
               (agency_id, activity_id, year, category, evidence, focus, formal_learning, social_learning,
                experiential_learning, owner_role, timing, respondents_mentioning, assessment_evidence,
                corporate_reference, estimated_budget, quote_basis, approval_status, indicator, status,
                outcome_notes, updated_at)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                agency_id, activity_id, int(programme_sheet.cell(row, 2).value),
                workbook_text(programme_sheet.cell(row, 3).value), workbook_text(programme_sheet.cell(row, 4).value),
                workbook_text(programme_sheet.cell(row, 5).value), workbook_text(programme_sheet.cell(row, 6).value),
                workbook_text(programme_sheet.cell(row, 7).value), workbook_text(programme_sheet.cell(row, 8).value),
                workbook_text(programme_sheet.cell(row, 9).value), workbook_text(programme_sheet.cell(row, 10).value),
                int(programme_sheet.cell(row, 11).value or 0), workbook_text(programme_sheet.cell(row, 12).value),
                workbook_text(programme_sheet.cell(row, 13).value), _numeric(programme_sheet.cell(row, 14).value, 0, None),
                workbook_text(programme_sheet.cell(row, 15).value), workbook_text(programme_sheet.cell(row, 16).value),
                workbook_text(programme_sheet.cell(row, 17).value), workbook_text(programme_sheet.cell(row, 18).value),
                workbook_text(programme_sheet.cell(row, 19).value), now_utc(),
            ),
        )
        programme_id = connection.execute(
            "SELECT id FROM programme WHERE agency_id=? AND activity_id=?",
            (agency_id, activity_id),
        ).fetchone()["id"]
        connection.execute(
            "INSERT INTO monitoring(programme_id, updated_at) VALUES(?, ?)",
            (programme_id, now_utc()),
        )

    for option_code, sheet_name in (
        ("A", "Plan A - Foundation"),
        ("B", "Plan B - Prosecution"),
        ("C", "Plan C - Transformation"),
    ):
        option_sheet = analysis[ sheet_name ]
        schedule = []
        for row in range(7, 12):
            schedule.append({
                "year": workbook_text(option_sheet.cell(row, 1).value),
                "focus": workbook_text(option_sheet.cell(row, 2).value),
                "target_cohort": workbook_text(option_sheet.cell(row, 3).value),
                "practice": workbook_text(option_sheet.cell(row, 4).value),
                "peer": workbook_text(option_sheet.cell(row, 5).value),
                "formal": workbook_text(option_sheet.cell(row, 6).value),
                "cost_guardrails": workbook_text(option_sheet.cell(row, 7).value),
                "success_evidence": workbook_text(option_sheet.cell(row, 8).value),
                "corporate_link": workbook_text(option_sheet.cell(row, 9).value),
            })
        connection.execute(
            "INSERT INTO plan_options(agency_id, option_code, title, description, annual_schedule) VALUES(?, ?, ?, ?, ?)",
            (agency_id, option_code, workbook_text(option_sheet.cell(1, 1).value),
             workbook_text(option_sheet.cell(2, 1).value), json.dumps(schedule)),
        )
    analysis.close()

    connection.execute(
        "INSERT OR REPLACE INTO app_meta(key, value) VALUES('source_files', ?)",
        (json.dumps([survey_path.name, roster_path.name, analysis_path.name]),),
    )
    connection.execute(
        "INSERT OR REPLACE INTO app_meta(key, value) VALUES('seeded_at', ?)",
        (now_utc(),),
    )


def _numeric(value: object, minimum: float | None, maximum: float | None) -> float | None:
    if value is None or value == "":
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if minimum is not None and number < minimum:
        return None
    if maximum is not None and number > maximum:
        return None
    return number


def initialize_database(path: Path | None = None, source_files: tuple[Path, Path, Path] | None = None) -> None:
    target = path or db_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    with database_session(target) as connection:
        create_schema(connection)
        exists = connection.execute("SELECT 1 FROM agencies WHERE code=?", (AGENCY_CODE,)).fetchone()
        if exists is None:
            _insert_source_data(connection, *(source_files or (SURVEY_FILE, ROSTER_FILE, ANALYSIS_FILE)))


def agency_id_for(connection: sqlite3.Connection, code: str | None = None) -> int:
    agency = connection.execute(
        "SELECT id FROM agencies WHERE code=?",
        (code or AGENCY_CODE,),
    ).fetchone()
    if agency is None:
        raise ValueError("Selected agency is not available in the local database.")
    return int(agency["id"])


def as_dict(row: sqlite3.Row | None) -> dict[str, object] | None:
    return dict(row) if row is not None else None


def list_agencies(connection: sqlite3.Connection) -> list[dict[str, object]]:
    return [dict(row) for row in connection.execute("SELECT id, code, name FROM agencies ORDER BY name")]


def response_records(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    rows = connection.execute(
        "SELECT * FROM responses WHERE agency_id=? ORDER BY source_row",
        (agency_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def item_analytics(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    rows = connection.execute(
        """SELECT ai.item_index, ai.item_code, ai.dimension, ai.statement, ai.review_route,
                  COUNT(CASE WHEN r.id IS NOT NULL AND s.score IS NOT NULL THEN 1 END) AS valid_scores,
                  AVG(CASE WHEN r.id IS NOT NULL THEN s.score END) AS mean_score,
                  SUM(CASE WHEN r.id IS NOT NULL AND s.score BETWEEN 1 AND 3 THEN 1 ELSE 0 END) AS low_count,
                  SUM(CASE WHEN r.id IS NOT NULL AND s.score BETWEEN 4 AND 8 THEN 1 ELSE 0 END) AS average_count,
                  SUM(CASE WHEN r.id IS NOT NULL AND s.score BETWEEN 9 AND 10 THEN 1 ELSE 0 END) AS high_count
           FROM assessment_items ai
           LEFT JOIN item_scores s ON s.item_index=ai.item_index
           LEFT JOIN responses r ON r.id=s.response_id AND r.agency_id=ai.agency_id
           WHERE ai.agency_id=?
           GROUP BY ai.item_index, ai.item_code, ai.dimension, ai.statement, ai.review_route
           ORDER BY ai.item_index""",
        (agency_id,),
    ).fetchall()
    results = []
    for row in rows:
        item = dict(row)
        mean = item["mean_score"]
        item["gap"] = 10 - mean if mean is not None else None
        item["band"] = score_band(mean)
        results.append(item)
    return results


def category_analytics(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    rows = connection.execute(
        "SELECT * FROM need_categories WHERE agency_id=? ORDER BY id",
        (agency_id,),
    ).fetchall()
    responses = response_records(connection, agency_id)
    results = []
    for row in rows:
        keywords = json.loads(row["keywords"])
        count = sum(
            any(keyword in (response["interests"] or "").lower() for keyword in keywords)
            for response in responses
        )
        results.append({
            "id": row["id"],
            "name": row["name"],
            "respondents_mentioning": count,
            "keywords": keywords,
            "evidence_source": row["evidence_source"],
            "intervention_route": row["intervention_route"],
            "priority_note": "Draft evidence only; confirm strategic fit, operational need, affected population, urgency and eligibility.",
        })
    return results


def dashboard_data(connection: sqlite3.Connection, agency_id: int) -> dict[str, object]:
    agency = connection.execute("SELECT * FROM agencies WHERE id=?", (agency_id,)).fetchone()
    totals = connection.execute(
        "SELECT COUNT(*) AS count FROM staff WHERE agency_id=?",
        (agency_id,),
    ).fetchone()["count"]
    responses = connection.execute(
        "SELECT * FROM responses WHERE agency_id=?",
        (agency_id,),
    ).fetchall()
    covered = connection.execute(
        """SELECT COUNT(DISTINCT staff_id) AS count FROM responses
           WHERE agency_id=? AND staff_id IS NOT NULL AND match_status IN ('Matched','Verified')""",
        (agency_id,),
    ).fetchone()["count"]
    verify = sum(row["match_status"] == "Verify name" for row in responses)
    not_matched = totals - covered
    review_candidates: set[int] = set()
    for row in connection.execute(
        "SELECT candidate_staff_ids FROM responses WHERE agency_id=? AND match_status='Verify name'",
        (agency_id,),
    ):
        review_candidates.update(json.loads(row["candidate_staff_ids"]))
    spa = [row["spa_rating"] for row in responses if row["spa_rating"] is not None]
    high_spa = sum(value in (4, 5) for value in spa)
    process_yes = sum((row["needs_process"] or "").strip().lower() == "yes" for row in responses)
    process_no = sum((row["needs_process"] or "").strip().lower() == "no" for row in responses)
    programmes = connection.execute(
        """SELECT COUNT(*) AS activity_count,
                  SUM(CASE WHEN status='Completed' THEN 1 ELSE 0 END) AS completed_count,
                  SUM(estimated_budget) AS estimated_budget
           FROM programme WHERE agency_id=?""",
        (agency_id,),
    ).fetchone()
    actual_spend = connection.execute(
        """SELECT SUM(m.actual_expenditure) AS total FROM monitoring m
           JOIN programme p ON p.id=m.programme_id WHERE p.agency_id=?""",
        (agency_id,),
    ).fetchone()["total"]
    return {
        "agency": {"code": agency["code"], "name": agency["name"]},
        "roster_count": totals,
        "cna_responses": len(responses),
        "matched_participants": covered,
        "not_in_cna": not_matched - len(review_candidates),
        "verify_name": verify,
        "coverage_percent": round(covered / totals * 100, 1) if totals else 0,
        "spa_4_5_count": high_spa,
        "spa_rated_count": len(spa),
        "needs_process_yes": process_yes,
        "needs_process_no": process_no,
        "plan_activities": programmes["activity_count"] or 0,
        "completed_activities": programmes["completed_count"] or 0,
        "estimated_budget_pgk": programmes["estimated_budget"],
        "actual_spend_pgk": actual_spend,
        "budget_priced": programmes["estimated_budget"] is not None,
    }


def skills_data(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    rows = connection.execute(
        "SELECT * FROM responses WHERE agency_id=? ORDER BY name COLLATE NOCASE",
        (agency_id,),
    ).fetchall()
    result = []
    for row in rows:
        status = "Manual review against approved job description" if row["qualification"] and row["required_qualification"] else "Required or held qualification not recorded"
        result.append({
            "response_id": row["id"],
            "source_row": row["source_row"],
            "name": row["name"],
            "position": row["position"],
            "grade": row["grade"],
            "division": row["division"],
            "employment_status": row["employment_status"] or "Not recorded",
            "qualification": row["qualification"] or "Not recorded",
            "required_qualification": row["required_qualification"] or "Not recorded",
            "qualification_review": status,
            "competencies": row["competencies"] or "Not recorded",
            "competency_standard": "Not supplied; link approved job competency profile",
            "spa_rating": row["spa_rating"],
            "spa_scale": "1–5",
        })
    return result


def staff_data(connection: sqlite3.Connection, agency_id: int, search: str = "") -> list[dict[str, object]]:
    rows = connection.execute(
        "SELECT * FROM staff WHERE agency_id=? ORDER BY name COLLATE NOCASE",
        (agency_id,),
    ).fetchall()
    matched = {
        row["staff_id"]: row
        for row in connection.execute(
            """SELECT * FROM responses WHERE agency_id=? AND staff_id IS NOT NULL
               AND match_status IN ('Matched','Verified')""",
            (agency_id,),
        )
    }
    candidates: set[int] = set()
    for candidate_row in connection.execute(
        "SELECT candidate_staff_ids FROM responses WHERE agency_id=? AND match_status='Verify name'",
        (agency_id,),
    ):
        candidates.update(json.loads(candidate_row["candidate_staff_ids"]))
    results = []
    for row in rows:
        response = matched.get(row["id"])
        item = {
            "id": row["id"],
            "name": row["name"],
            "position_numbers": json.loads(row["position_numbers"]),
            "positions": json.loads(row["positions"]),
            "locations": json.loads(row["locations"]),
            "occupant_types": json.loads(row["occupant_types"]),
            "source_rows": json.loads(row["source_rows"]),
            "cna_status": "Participated" if response else "Verify name" if row["id"] in candidates else "Not in CNA",
            "response_id": response["id"] if response else None,
        }
        if not search or search.casefold() in (
            row["name"] + " " + " ".join(json.loads(row["positions"])) + " "
            + " ".join(json.loads(row["locations"]))
        ).casefold():
            results.append(item)
    return results


def possible_matches(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    pending = connection.execute(
        """SELECT id, name, source_row, candidate_staff_ids FROM responses
           WHERE agency_id=? AND match_status='Verify name' ORDER BY name COLLATE NOCASE""",
        (agency_id,),
    ).fetchall()
    results = []
    for response in pending:
        candidate_ids = json.loads(response["candidate_staff_ids"])
        candidates = []
        if candidate_ids:
            placeholders = ",".join("?" for _ in candidate_ids)
            rows = connection.execute(
                f"SELECT id, name, positions, locations FROM staff WHERE agency_id=? AND id IN ({placeholders})",
                (agency_id, *candidate_ids),
            ).fetchall()
            candidates = [
                {"id": row["id"], "name": row["name"], "positions": json.loads(row["positions"]),
                 "locations": json.loads(row["locations"])}
                for row in rows
            ]
        results.append({
            "response_id": response["id"],
            "respondent": response["name"],
            "source_row": response["source_row"],
            "candidates": candidates,
        })
    return results


def response_card(connection: sqlite3.Connection, response_id: int, agency_id: int) -> dict[str, object] | None:
    response = connection.execute(
        "SELECT * FROM responses WHERE id=? AND agency_id=?",
        (response_id, agency_id),
    ).fetchone()
    if response is None:
        return None
    record = dict(response)
    record["scores"] = [
        dict(row)
        for row in connection.execute(
            """SELECT ai.item_code, ai.dimension, ai.statement, s.score
               FROM item_scores s JOIN assessment_items ai ON ai.item_index=s.item_index
               WHERE s.response_id=? ORDER BY ai.item_index""",
            (response_id,),
        )
    ]
    record["training_eligibility"] = "Not determined: approved agency eligibility rules were not supplied."
    record["qualification_comparison"] = (
        "Manual review required; no qualification equivalency rules supplied."
        if record["qualification"] and record["required_qualification"]
        else "Qualification or position requirement is missing from this response."
    )
    return record


def programme_data(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    rows = connection.execute(
        """SELECT p.*, m.planned_participants, m.participants_completed, m.actual_expenditure,
                  m.indicator AS monitor_indicator, m.baseline, m.target, m.actual_result, m.evidence
           FROM programme p LEFT JOIN monitoring m ON m.programme_id=p.id
           WHERE p.agency_id=? ORDER BY p.year, p.activity_id""",
        (agency_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def options_data(connection: sqlite3.Connection, agency_id: int) -> list[dict[str, object]]:
    options = connection.execute(
        "SELECT option_code, title, description, annual_schedule FROM plan_options WHERE agency_id=? ORDER BY option_code",
        (agency_id,),
    ).fetchall()
    return [{
        "code": row["option_code"],
        "title": row["title"],
        "description": row["description"],
        "annual_schedule": json.loads(row["annual_schedule"]),
    } for row in options]


def methodology_data(connection: sqlite3.Connection, agency_id: int) -> dict[str, object]:
    return {
        "sources": [
            "Office of Public Presecutors (Responses).xlsx: 39 CNA submissions.",
            "Form 3.17 - Occupancy Report updated.xlsx: named substantive and actual occupants; duplicate people collapsed by family/given-name key.",
            "Capability Needs Analysis Hub.xlsx: existing three scenario options and 35 editable 2027–2031 programme activities.",
        ],
        "cna_scale": "1–3 Poor/Low; 4–8 Average; 9–10 High/Excellent.",
        "gap_formula": "Capability gap = 10 − assessment score. Institutional item means are descriptive summaries of respondent scores.",
        "spa_scale": "SPA ratings remain on 1–5 and are never converted to CNA's 1–10 scale.",
        "qualification_method": "Compare self-reported highest qualification with verbatim position requirement; manual review only. No equivalency was inferred.",
        "competency_method": "Survey-reported competencies are shown against an explicit missing approved job competency standard; no pass/fail comparison is invented.",
        "priority_method": "Use assessed gap severity, strategic importance, operational requirements, staff affected and urgency. No weighted ranking is calculated because validated weights, targets and impact data are unavailable.",
        "learning_model": "10% formal, 20% social and 70% experiential learning are treated as learning approaches, not budget shares.",
        "training_eligibility": "Formal learning is conditional on verified role need, approved SPA, employment eligibility, fair access and approval. The authoritative eligibility rules are not in the supplied files.",
        "budget_method": "Estimated and actual budget fields remain blank until approved provider quotes and financial records are entered.",
        "corporate_plan_limit": "The provided corporate plan covers 2025–2030 and is image-only. Confirm the approved plan/KRA for 2027–2030 and the successor plan for 2031.",
        "workspace_mode": "The workspace answers from local database records using deterministic evidence queries; it does not call an external AI service.",
        "local_only": "The server binds to 127.0.0.1. This is a single-user local application, not an authenticated or multi-agency deployment.",
        "need_categories": category_analytics(connection, agency_id),
    }


def answer_workspace(connection: sqlite3.Connection, agency_id: int, question: str) -> dict[str, object]:
    query = normalize(question)
    if not query:
        raise ValueError("Enter a question to search local CNA evidence.")
    dash = dashboard_data(connection, agency_id)
    diagnostic = item_analytics(connection, agency_id)
    categories = category_analytics(connection, agency_id)
    responses = response_records(connection, agency_id)

    for response in responses:
        if normalize(response["name"]) in query or (len(query.split()) > 1 and query in normalize(response["name"])):
            return {
                "answer": (
                    f"{response['name']} is listed as {response['position'] or 'a role not recorded'} in "
                    f"{response['division'] or 'a division not recorded'}. Reported interests: "
                    f"{response['interests'] or 'not recorded'}. Career goal: {response['career_goal'] or 'not recorded'}. "
                    f"SPA rating: {response['spa_rating'] if response['spa_rating'] is not None else 'not recorded'} on the separate 1–5 scale. "
                    "This does not establish training eligibility or a performance finding."
                ),
                "sources": [f"CNA response row {response['source_row']}, fields E5–E7 and SPA"],
            }

    if any(word in query for word in ("participat", "nonrespond", "response rate", "cna response", "roster")):
        return {
            "answer": (
                f"The roster has {dash['roster_count']} distinct named people. "
                f"{dash['matched_participants']} have matched CNA submissions; {dash['verify_name']} possible one-name matches "
                f"need verification; {dash['not_in_cna']} have no confirmed matching CNA name. "
                f"There are {dash['cna_responses']} submitted survey responses."
            ),
            "sources": ["Roster CNA Check", "CNA Survey Response Data"],
        }

    if any(word in query for word in ("lowest", "weak", "gap", "diagnostic", "score")):
        measured = [item for item in diagnostic if item["mean_score"] is not None]
        measured.sort(key=lambda item: item["mean_score"])
        lowest = measured[:5]
        summary = "; ".join(
            f"{item['item_code']} mean {item['mean_score']:.2f}/10, gap {item['gap']:.2f} ({item['band']})"
            for item in lowest
        )
        return {
            "answer": (
                "Lowest reported institutional assessment means are: " + summary
                + ". These are perception signals; they do not prove causes or that training is the remedy."
            ),
            "sources": ["CNA Diagnostic: item means and gap = 10 − mean; 39 source responses"],
        }

    if any(word in query for word in ("budget", "cost", "expenditure", "spend")):
        return {
            "answer": (
                "Training budgets are unpriced because no provider quotations or approved expenditure records were supplied. "
                "Enter approved PGK estimates and actual expenditure in the programme and monitoring pages; the app will then report totals."
            ),
            "sources": ["Five-Year Programme budget fields", "M&E Tracker expenditure fields"],
        }

    matched_categories = [
        category for category in categories
        if category["name"].lower() in query
        or any(len(word) > 4 and word in query for word in category["keywords"])
    ]
    if matched_categories:
        return {
            "answer": " | ".join(
                f"{item['name']}: {item['respondents_mentioning']} of {dash['cna_responses']} respondents mentioned one or more mapped phrases. "
                "This is an overlapping free-text interest count, not a validated priority or staff-affected count."
                for item in matched_categories
            ),
            "sources": ["Training Needs: phrase rules from the workbook and CNA E7 responses"],
        }

    return {
        "answer": (
            f"Local records contain {dash['cna_responses']} CNA responses, {dash['roster_count']} roster people and "
            f"{dash['plan_activities']} draft plan activities for 2027–2031. Ask about roster participation, a named respondent, "
            "CNA score gaps, a training-interest category, or budget status. I cannot infer missing targets, causes or eligibility rules."
        ),
        "sources": ["Local CNAMIS database seeded from the supplied OPP workbooks"],
    }


def export_workbook(connection: sqlite3.Connection, agency_id: int) -> bytes:
    agency = connection.execute("SELECT name FROM agencies WHERE id=?", (agency_id,)).fetchone()
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    dash = dashboard_data(connection, agency_id)
    summary.append(["CNAMIS local report", agency["name"]])
    summary.append(["CNA responses", dash["cna_responses"]])
    summary.append(["Roster people", dash["roster_count"]])
    summary.append(["Matched", dash["matched_participants"]])
    summary.append(["Verify name", dash["verify_name"]])
    summary.append(["No confirmed CNA match", dash["not_in_cna"]])
    summary.append(["SPA 4–5 (survey screen only)", dash["spa_4_5_count"]])
    summary.append(["Budget", "Not priced" if not dash["budget_priced"] else dash["estimated_budget_pgk"]])

    diag = wb.create_sheet("CNA Diagnostic")
    diag.append(["Item", "Dimension", "Assessment statement", "Valid scores", "Mean / 10", "Gap", "Band", "Review route"])
    for item in item_analytics(connection, agency_id):
        diag.append([item["item_code"], item["dimension"], item["statement"], item["valid_scores"],
                     item["mean_score"], item["gap"], item["band"], item["review_route"]])

    skills = wb.create_sheet("Skills Review")
    skills.append(["Respondent", "Position", "Qualification", "Required qualification", "Competencies reported", "SPA 1–5", "Employment status"])
    for item in skills_data(connection, agency_id):
        skills.append([item["name"], item["position"], item["qualification"], item["required_qualification"],
                       item["competencies"], item["spa_rating"], item["employment_status"]])

    plan = wb.create_sheet("Five-Year Plan")
    plan.append(["Activity", "Year", "Category", "Evidence", "Focus", "Formal 10%", "Social 20%",
                 "Experiential 70%", "Owner", "Timing", "Estimate PGK", "Approval", "Status"])
    for item in programme_data(connection, agency_id):
        plan.append([item["activity_id"], item["year"], item["category"], item["evidence"], item["focus"],
                     item["formal_learning"], item["social_learning"], item["experiential_learning"],
                     item["owner_role"], item["timing"], item["estimated_budget"], item["approval_status"], item["status"]])

    monitor = wb.create_sheet("Monitoring")
    monitor.append(["Activity", "Year", "Category", "Planned participants", "Completed", "Actual PGK",
                    "Indicator", "Baseline", "Target", "Actual result", "Evidence"])
    for item in programme_data(connection, agency_id):
        monitor.append([item["activity_id"], item["year"], item["category"], item["planned_participants"],
                       item["participants_completed"], item["actual_expenditure"], item["monitor_indicator"],
                       item["baseline"], item["target"], item["actual_result"], item["evidence"]])

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1A365D")
            cell.font = Font(color="FFFFFF", bold=True)
        for column_cells in ws.columns:
            max_len = min(60, max(len(str(cell.value or "")) for cell in column_cells) + 2)
            ws.column_dimensions[column_cells[0].column_letter].width = max(14, max_len)
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()


def capability_report_data(connection: sqlite3.Connection, agency_id: int) -> dict[str, object]:
    agency = connection.execute(
        "SELECT code, name FROM agencies WHERE id=?",
        (agency_id,),
    ).fetchone()
    dashboard = dashboard_data(connection, agency_id)
    diagnostic = item_analytics(connection, agency_id)
    categories = category_analytics(connection, agency_id)
    skills = skills_data(connection, agency_id)
    programme = programme_data(connection, agency_id)
    options = options_data(connection, agency_id)
    score_summary = connection.execute(
        """SELECT COUNT(s.score) AS score_count, AVG(s.score) AS mean_score,
                  SUM(CASE WHEN s.score BETWEEN 1 AND 3 THEN 1 ELSE 0 END) AS low_count,
                  SUM(CASE WHEN s.score BETWEEN 4 AND 8 THEN 1 ELSE 0 END) AS average_count,
                  SUM(CASE WHEN s.score BETWEEN 9 AND 10 THEN 1 ELSE 0 END) AS high_count
           FROM item_scores s JOIN responses r ON r.id=s.response_id
           WHERE r.agency_id=?""",
        (agency_id,),
    ).fetchone()
    dimensions: dict[str, dict[str, object]] = {}
    for item in diagnostic:
        if item["mean_score"] is None:
            continue
        summary = dimensions.setdefault(item["dimension"], {"weighted_total": 0.0, "score_count": 0})
        summary["weighted_total"] += item["mean_score"] * item["valid_scores"]
        summary["score_count"] += item["valid_scores"]
    dimension_results = [
        {
            "name": name,
            "score_count": item["score_count"],
            "mean_score": item["weighted_total"] / item["score_count"] if item["score_count"] else None,
        }
        for name, item in dimensions.items()
    ]
    response_rows = connection.execute(
        "SELECT employment_status, division, grade FROM responses WHERE agency_id=?",
        (agency_id,),
    ).fetchall()
    employment = _frequency_summary(row["employment_status"] for row in response_rows)
    divisions = _frequency_summary(row["division"] for row in response_rows)
    grades = _frequency_summary(row["grade"] for row in response_rows)
    qualification_records = sum(bool(item["qualification"] and item["required_qualification"]) for item in skills)
    competency_records = sum(bool(item["competencies"] and item["competencies"] != "Not recorded") for item in skills)
    monitored_records = sum(
        item["participants_completed"] is not None
        or item["actual_expenditure"] is not None
        or bool(item["actual_result"])
        or bool(item["evidence"])
        for item in programme
    )
    activity_status: dict[str, int] = {}
    annual_summary: list[dict[str, object]] = []
    for year in range(2027, 2032):
        activities = [item for item in programme if item["year"] == year]
        for item in activities:
            status = item["status"] or "Not started"
            activity_status[status] = activity_status.get(status, 0) + 1
        priced = [item["estimated_budget"] for item in activities if item["estimated_budget"] is not None]
        annual_summary.append({
            "year": year,
            "activity_count": len(activities),
            "category_count": len({item["category"] for item in activities}),
            "budget": sum(priced) if priced else None,
        })
    response_count = dashboard["cna_responses"]
    roster_count = dashboard["roster_count"]
    compliance_statement = (
        f"The {agency['name']} has complied with the agency-level CNA submission and assessment "
        "requirements evidenced by the completed CNA survey returns and the Form 3.17 occupancy "
        "records provided for this analysis."
    )
    participation_caveat = (
        f"This agency-level compliance finding does not certify complete individual participation. "
        f"Of {roster_count} distinct named roster people, {dashboard['matched_participants']} have a "
        f"confirmed CNA name match, {dashboard['verify_name']} possible match(es) require manual "
        f"verification, and {dashboard['not_in_cna']} have no confirmed match. The source contains "
        f"{response_count} survey submissions. Follow-up is required before claiming that every "
        "rostered officer participated."
    )
    return {
        "agency": dict(agency),
        "generated_at": datetime.now().astimezone().strftime("%d %B %Y"),
        "dashboard": dashboard,
        "response_rate": round(response_count / roster_count * 100, 1) if roster_count else 0,
        "confirmed_coverage": dashboard["coverage_percent"],
        "compliance_statement": compliance_statement,
        "participation_caveat": participation_caveat,
        "score_summary": dict(score_summary),
        "overall_band": score_band(score_summary["mean_score"]),
        "diagnostic": diagnostic,
        "dimensions": dimension_results,
        "lowest_items": sorted(
            (item for item in diagnostic if item["mean_score"] is not None),
            key=lambda item: item["mean_score"],
        )[:5],
        "highest_items": sorted(
            (item for item in diagnostic if item["mean_score"] is not None),
            key=lambda item: item["mean_score"],
            reverse=True,
        )[:5],
        "categories": categories,
        "skills": skills,
        "qualification_records": qualification_records,
        "competency_records": competency_records,
        "employment_summary": employment,
        "division_summary": divisions,
        "grade_summary": grades,
        "programme": programme,
        "options": options,
        "annual_summary": annual_summary,
        "activity_status": activity_status,
        "monitored_records": monitored_records,
        "budget_priced": dashboard["budget_priced"],
        "methodology": methodology_data(connection, agency_id),
    }


def _frequency_summary(values: Iterable[object]) -> list[tuple[str, int]]:
    counts: dict[str, int] = {}
    for value in values:
        label = str(value or "").strip() or "Not recorded"
        counts[label] = counts.get(label, 0) + 1
    return sorted(counts.items(), key=lambda item: (-item[1], item[0].casefold()))


def export_capability_report(connection: sqlite3.Connection, agency_id: int) -> bytes:
    report = capability_report_data(connection, agency_id)
    stream = BytesIO()
    regular_font = "Helvetica"
    bold_font = "Helvetica-Bold"
    font_dir = Path(os.environ.get("WINDIR", "C:\\Windows")) / "Fonts"
    regular_path = font_dir / "arial.ttf"
    bold_path = font_dir / "arialbd.ttf"
    if regular_path.is_file() and bold_path.is_file():
        if "CNAMISArial" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("CNAMISArial", str(regular_path)))
            pdfmetrics.registerFont(TTFont("CNAMISArial-Bold", str(bold_path)))
            pdfmetrics.registerFontFamily(
                "CNAMISArial",
                normal="CNAMISArial",
                bold="CNAMISArial-Bold",
                italic="CNAMISArial",
                boldItalic="CNAMISArial-Bold",
            )
        regular_font = "CNAMISArial"
        bold_font = "CNAMISArial-Bold"

    navy = colors.HexColor("#1A365D")
    dark_navy = colors.HexColor("#1A1A40")
    pale_green = colors.HexColor("#EFF8F1")
    ink = colors.HexColor("#203047")
    muted = colors.HexColor("#64748B")
    line = colors.HexColor("#D9E2EC")
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ReportTitle", parent=styles["Title"], fontName=bold_font, fontSize=26,
        leading=32, textColor=dark_navy, alignment=TA_LEFT, spaceAfter=12,
    ))
    styles.add(ParagraphStyle(
        name="ReportSubtitle", parent=styles["Normal"], fontName=regular_font, fontSize=13,
        leading=19, textColor=navy, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="ReportHeading", parent=styles["Heading1"], fontName=bold_font, fontSize=16,
        leading=20, textColor=dark_navy, spaceBefore=10, spaceAfter=9, keepWithNext=True,
    ))
    styles.add(ParagraphStyle(
        name="ReportSubheading", parent=styles["Heading2"], fontName=bold_font, fontSize=11,
        leading=14, textColor=navy, spaceBefore=7, spaceAfter=5, keepWithNext=True,
    ))
    styles.add(ParagraphStyle(
        name="ReportBody", parent=styles["BodyText"], fontName=regular_font, fontSize=9,
        leading=13, textColor=ink, spaceAfter=7,
    ))
    styles.add(ParagraphStyle(
        name="ReportSmall", parent=styles["BodyText"], fontName=regular_font, fontSize=7.5,
        leading=10, textColor=ink,
    ))
    styles.add(ParagraphStyle(
        name="ReportTableHeader", parent=styles["BodyText"], fontName=bold_font, fontSize=7.3,
        leading=9, textColor=colors.white,
    ))
    styles.add(ParagraphStyle(
        name="ReportTableCell", parent=styles["BodyText"], fontName=regular_font, fontSize=7.1,
        leading=9, textColor=ink,
    ))
    styles.add(ParagraphStyle(
        name="ReportCallout", parent=styles["BodyText"], fontName=regular_font, fontSize=9.2,
        leading=14, textColor=dark_navy,
    ))
    styles.add(ParagraphStyle(
        name="ReportCoverNote", parent=styles["BodyText"], fontName=regular_font, fontSize=9,
        leading=14, textColor=muted,
    ))

    def paragraph(value: object, style: str = "ReportBody") -> Paragraph:
        text = escape(str(value if value not in (None, "") else "Not recorded")).replace("\n", "<br/>")
        return Paragraph(text, styles[style])

    def report_table(
        headers: list[str],
        rows: list[list[object]],
        widths: list[float],
    ) -> Table:
        table_rows: list[list[object]] = [
            [paragraph(value, "ReportTableHeader") for value in headers]
        ]
        table_rows.extend([
            [value if isinstance(value, Paragraph) else paragraph(value, "ReportTableCell") for value in row]
            for row in rows
        ])
        table = Table(table_rows, colWidths=widths, repeatRows=1, hAlign="LEFT", splitByRow=1)
        commands = [
            ("BACKGROUND", (0, 0), (-1, 0), navy),
            ("GRID", (0, 0), (-1, -1), 0.35, line),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]
        for row_number in range(2, len(table_rows), 2):
            commands.append(("BACKGROUND", (0, row_number), (-1, row_number), colors.HexColor("#F6F9FC")))
        table.setStyle(TableStyle(commands))
        return table

    page_width, page_height = A4
    margin = 17 * mm
    document = SimpleDocTemplate(
        stream,
        pagesize=A4,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=21 * mm,
        bottomMargin=18 * mm,
        title=f"Capability Needs Analysis Report - {report['agency']['name']}",
        author="CNAMIS - Office of the Public Prosecutor",
        subject="Agency-wide Capability Needs Analysis findings and training priorities",
    )

    def draw_page(canvas, doc) -> None:
        canvas.saveState()
        canvas.setFillColor(dark_navy)
        canvas.rect(0, page_height - 12 * mm, page_width, 12 * mm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont(regular_font, 7.5)
        canvas.drawString(margin, page_height - 8 * mm, "CNAMIS  /  CAPABILITY NEEDS ANALYSIS")
        canvas.setStrokeColor(line)
        canvas.line(margin, 13 * mm, page_width - margin, 13 * mm)
        canvas.setFillColor(muted)
        canvas.setFont(regular_font, 7)
        canvas.drawString(margin, 8.5 * mm, str(report["agency"]["name"]))
        canvas.drawRightString(page_width - margin, 8.5 * mm, f"Planning period 2027–2031  ·  Page {doc.page}")
        canvas.restoreState()

    story: list[object] = [
        Spacer(1, 27 * mm),
        paragraph("AGENCY-WIDE ASSESSMENT", "ReportSubtitle"),
        paragraph("CAPABILITY NEEDS<br/>ANALYSIS REPORT", "ReportTitle"),
        paragraph(report["agency"]["name"], "ReportSubtitle"),
        Spacer(1, 5 * mm),
        paragraph("Overall CNA findings, workforce capability signals and evidence-based development planning", "ReportBody"),
        Spacer(1, 9 * mm),
        report_table(
            ["REPORT PERIOD", "CNA EVIDENCE", "PREPARED"],
            [[
                "2027–2031 plan horizon",
                f"{report['dashboard']['cna_responses']} survey submissions · {report['dashboard']['roster_count']} roster people",
                report["generated_at"],
            ]],
            [48 * mm, 70 * mm, 41 * mm],
        ),
        Spacer(1, 9 * mm),
    ]
    compliance_box = Table(
        [[paragraph(report["compliance_statement"], "ReportCallout")]],
        colWidths=[159 * mm],
    )
    compliance_box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pale_green),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#B8DDC0")),
        ("LEFTPADDING", (0, 0), (-1, -1), 11),
        ("RIGHTPADDING", (0, 0), (-1, -1), 11),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.extend([
        compliance_box,
        Spacer(1, 8 * mm),
        paragraph(
            "Prepared from the supplied OPP CNA survey, Form 3.17 occupancy report and CNAMIS analysis workbook. "
            "This is an evidence-led planning report, not a personnel decision, formal audit opinion or approved training budget.",
            "ReportCoverNote",
        ),
        PageBreak(),
    ])

    dashboard = report["dashboard"]
    score = report["score_summary"]
    overall_mean = score["mean_score"]
    overall_text = "Not assessed" if overall_mean is None else f"{overall_mean:.2f} / 10 ({report['overall_band']})"
    story.extend([
        paragraph("Executive summary", "ReportHeading"),
        paragraph(
            f"This report consolidates {dashboard['cna_responses']} submitted CNA survey records and the "
            f"{dashboard['roster_count']}-person Form 3.17 roster into an agency-level diagnostic and a draft "
            "five-year development programme. It separates survey evidence from interpretation and identifies "
            "where follow-up, management action or institutional strengthening may be more appropriate than training.",
        ),
        paragraph(report["compliance_statement"]),
        paragraph(report["participation_caveat"]),
        report_table(
            ["OVERALL INDICATOR", "FINDING", "INTERPRETATION"],
            [
                ["CNA submissions", dashboard["cna_responses"], f"{report['response_rate']}% of roster count; not all names are confirmed matches."],
                ["Confirmed roster coverage", f"{dashboard['matched_participants']} / {dashboard['roster_count']} ({report['confirmed_coverage']}%)", "Name-level confirmed matches only."],
                ["Overall CNA score", overall_text, f"{score['score_count']} valid item ratings; gap = 10 − score."],
                ["Training-interest categories", len(report["categories"]), "Phrase-matched E7 signals; categories overlap."],
                ["Draft programme", f"{len(report['programme'])} activities", "Existing workbook schedule for 2027–2031; not yet approved."],
                ["Budget", "Not priced" if not report["budget_priced"] else f"PGK {dashboard['estimated_budget_pgk']:,.2f}", "No provider quotes supplied in the source records."],
            ],
            [39 * mm, 47 * mm, 73 * mm],
        ),
        paragraph("Principal conclusions", "ReportSubheading"),
    ])
    findings = [
        f"The combined mean across valid 1–10 institutional ratings is {overall_text}. Item means and gap values are descriptive respondent perceptions, not causal findings.",
        f"The five lowest-scoring survey items are {', '.join(item['item_code'] for item in report['lowest_items']) or 'not available'}; each requires agency review to determine whether training, procedure, management or policy action is suitable.",
        f"Training-interest phrase matches are strongest in {', '.join(item['name'] for item in sorted(report['categories'], key=lambda row: row['respondents_mentioning'], reverse=True)[:3]) or 'no recorded category'}; these counts do not represent validated priorities or staff affected.",
        f"The source plan contains {len(report['programme'])} draft activities across 2027–2031. No budget, provider quotation or implementation outcome is inferred where none was supplied.",
    ]
    for index, finding in enumerate(findings, 1):
        story.append(paragraph(f"{index}. {finding}"))

    story.extend([
        paragraph("1. Purpose, scope and evidence base", "ReportHeading"),
        paragraph(
            "The purpose is to summarize overall agency capability signals, reconcile CNA submissions against the "
            "available establishment roster, review qualification and competency evidence, and present the source "
            "training programme for management review. Analysis is confined to the three supplied workbooks and "
            "current local CNAMIS records. No absent questionnaire response, job standard, corporate KRA, cost or "
            "outcome has been fabricated.",
        ),
        report_table(
            ["SOURCE", "USE IN THIS REPORT", "LIMITATION"],
            [
                ["Office of Public Presecutors (Responses).xlsx", "CNA scores, qualifications, competencies, SPA, expressed interests and agency-process answers.", f"{dashboard['cna_responses']} submissions; survey identity matching is not definitive."],
                ["Form 3.17 - Occupancy Report updated.xlsx", f"{dashboard['roster_count']} distinct named roster people for reconciliation.", "Names are consolidated; possible matches require manual identity checks."],
                ["Capability Needs Analysis Hub.xlsx", f"{len(report['options'])} scenarios and {len(report['programme'])} draft programme activities.", "Schedules are planning options; corporate-plan references and costs require validation."],
            ],
            [48 * mm, 62 * mm, 49 * mm],
        ),
        paragraph("Analytical rules", "ReportSubheading"),
        paragraph(
            "CNA item ratings use the 1–10 scale: 1–3 Poor/Low, 4–8 Average and 9–10 High/Excellent. "
            "Capability gap is calculated as 10 minus the score. SPA ratings remain on the separate 1–5 scale. "
            "Name matching is conservative; potential one-name matches are held for human review. Qualification "
            "comparisons are prompts for manual review, not equivalency decisions. Training-interest counts use "
            "transparent phrase matching and categories may overlap.",
        ),
        paragraph("2. CNA compliance and participation reconciliation", "ReportHeading"),
    ])
    compliance = Table(
        [[paragraph(report["compliance_statement"], "ReportCallout")]],
        colWidths=[159 * mm],
    )
    compliance.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pale_green),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#B8DDC0")),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.extend([
        compliance,
        Spacer(1, 3 * mm),
        paragraph(report["participation_caveat"]),
        report_table(
            ["RECONCILIATION STATUS", "PEOPLE", "SHARE OF ROSTER", "REPORTING TREATMENT"],
            [
                ["Confirmed CNA match", dashboard["matched_participants"], f"{report['confirmed_coverage']}%", "Counted as confirmed participation."],
                ["Possible name match", dashboard["verify_name"], "Pending", "Do not classify until verified against agency identity records."],
                ["No confirmed CNA match", dashboard["not_in_cna"], f"{dashboard['not_in_cna'] / dashboard['roster_count'] * 100:.1f}%" if dashboard["roster_count"] else "—", "Follow up; absence of a match is not proof of intent or refusal."],
                ["Survey submissions", dashboard["cna_responses"], f"{report['response_rate']}% of roster count", "Submission count and confirmed roster matches are different measures."],
            ],
            [38 * mm, 20 * mm, 29 * mm, 72 * mm],
        ),
        paragraph(
            "Compliance conclusion: the agency-level CNA submission and assessment requirements are met on the "
            "evidence available for this review. A complete officer-by-officer participation certificate is not "
            "supported until the two possible matches and the 43 unmatched roster names have been resolved or "
            "documented through the agency’s approved follow-up process.",
            "ReportBody",
        ),
        paragraph("3. Institutional CNA diagnostic", "ReportHeading"),
        paragraph(
            f"Across {score['score_count']} valid ratings, the overall mean is {overall_text}. "
            "The table reports each questionnaire item separately; means are not converted into competency pass/fail results.",
        ),
        report_table(
            ["DIMENSION", "MEAN / 10", "GAP", "VALID RATINGS"],
            [
                [item["name"], f"{item['mean_score']:.2f}" if item["mean_score"] is not None else "Not assessed",
                 f"{10 - item['mean_score']:.2f}" if item["mean_score"] is not None else "—", item["score_count"]]
                for item in report["dimensions"]
            ],
            [91 * mm, 24 * mm, 20 * mm, 24 * mm],
        ),
        paragraph("Full item-level results", "ReportSubheading"),
        report_table(
            ["ITEM", "DIMENSION", "SURVEY STATEMENT", "n", "MEAN", "GAP", "BAND", "REVIEW ROUTE"],
            [
                [item["item_code"], item["dimension"], item["statement"], item["valid_scores"],
                 f"{item['mean_score']:.2f}" if item["mean_score"] is not None else "—",
                 f"{item['gap']:.2f}" if item["gap"] is not None else "—",
                 item["band"], item["review_route"]]
                for item in report["diagnostic"]
            ],
            [12 * mm, 29 * mm, 43 * mm, 8 * mm, 12 * mm, 11 * mm, 18 * mm, 26 * mm],
        ),
        paragraph(
            "Interpretation: lower-rated items are prompts for validation. For example, a low perception rating "
            "about a procedure may indicate that the approved procedure is missing, outdated or poorly communicated; "
            "training should not be selected until the cause is established.",
            "ReportCoverNote",
        ),
        paragraph("4. Workforce qualifications and capability evidence", "ReportHeading"),
        paragraph(
            f"The survey contains {len(report['skills'])} respondent records. {report['qualification_records']} "
            "records include both a self-reported qualification and a stated position requirement for manual "
            f"comparison; {report['competency_records']} record(s) include reported competency text. Approved "
            "competency profiles and qualification equivalency rules were not supplied, so this report makes no "
            "pass/fail or qualification-equivalency determination.",
        ),
        report_table(
            ["WORKFORCE EVIDENCE", "COUNT / DISTRIBUTION", "LIMITATION"],
            [
                ["SPA 4–5 responses", f"{dashboard['spa_4_5_count']} of {dashboard['spa_rated_count']} rated responses", "Separate 1–5 scale; not automatic eligibility for formal training."],
                ["Respondent employment status", _format_frequency(report["employment_summary"]), "Self-reported survey category; confirm against current HR records."],
                ["Respondent divisions", _format_frequency(report["division_summary"][:8]), "Only as reported in CNA responses."],
                ["Respondent grades", _format_frequency(report["grade_summary"][:8]), "Only as reported in CNA responses."],
                ["Competency standard", "Not supplied", "Link approved position profiles before making skills-gap determinations."],
            ],
            [43 * mm, 52 * mm, 69 * mm],
        ),
        paragraph("5. Reported development needs", "ReportHeading"),
        paragraph(
            "The following counts result from phrase matching against the CNA E7 training-interest responses. "
            "A person may appear in multiple categories. The counts are expressions of interest, not confirmed "
            "capability gaps, intervention priorities or the number of staff affected.",
        ),
        report_table(
            ["TRAINING CATEGORY", "RESPONDENTS MENTIONING", "SOURCE / REVIEW ROUTE"],
            [
                [item["name"], f"{item['respondents_mentioning']} of {dashboard['cna_responses']}", item["intervention_route"]]
                for item in report["categories"]
            ],
            [54 * mm, 35 * mm, 75 * mm],
        ),
        paragraph(
            f"On whether the agency has a process for identifying future training needs, "
            f"{dashboard['needs_process_yes']} respondents answered Yes and {dashboard['needs_process_no']} answered No. "
            "These are respondent counts; verify policy and current practice with management.",
        ),
        paragraph("6. Five-year training programme and delivery model", "ReportHeading"),
        paragraph(
            f"The source workbook provides {len(report['options'])} alternative planning options and "
            f"{len(report['programme'])} draft activities for 2027–2031. The schedule below preserves the "
            "existing plan content. It remains a draft for agency prioritisation and approval.",
        ),
        report_table(
            ["YEAR", "ACTIVITIES", "CATEGORIES", "KNOWN BUDGET"],
            [
                [item["year"], item["activity_count"], item["category_count"],
                 "Not priced" if item["budget"] is None else f"PGK {item['budget']:,.2f}"]
                for item in report["annual_summary"]
            ],
            [24 * mm, 36 * mm, 39 * mm, 65 * mm],
        ),
        paragraph("Draft activity schedule", "ReportSubheading"),
        report_table(
            ["YEAR", "CATEGORY", "DRAFT FOCUS", "10:20:70 DELIVERY (FORMAL / SOCIAL / EXPERIENTIAL)", "OWNER / TIMING"],
            [
                [
                    item["year"],
                    item["category"],
                    item["focus"],
                    f"10%: {item['formal_learning']}\n\n20%: {item['social_learning']}\n\n70%: {item['experiential_learning']}",
                    f"{item['owner_role']}\n{item['timing']}",
                ]
                for item in report["programme"]
            ],
            [10 * mm, 31 * mm, 44 * mm, 61 * mm, 18 * mm],
        ),
        paragraph("Learning and cost-efficiency rules", "ReportSubheading"),
        paragraph(
            "The 10:20:70 model is used as an approach mix: 10% formal learning, 20% social learning and "
            "70% experiential learning. These percentages are not shares of the training budget. Classroom "
            "or provider-based formal learning should be considered only when a verified role gap, approved "
            "eligibility, fair access, supervisor support and funding are confirmed. Apply coaching, mentoring, "
            "peer learning, job assignments, supervised practice and other workplace application where they "
            "address the validated need more efficiently.",
        ),
    ])
    story.extend([
        paragraph("Planning options retained from the source workbook", "ReportSubheading"),
    ])
    for option in report["options"]:
        story.append(paragraph(f"{option['code']}. {option['title']}", "ReportSubheading"))
        story.append(paragraph(option["description"]))
        story.append(report_table(
            ["YEAR", "FOCUS", "TARGET COHORT", "WORKPLACE PRACTICE", "PEER / FORMAL LEARNING"],
            [
                [year["year"], year["focus"], year["target_cohort"],
                 year["practice"], f"Peer: {year['peer']}\n\nFormal: {year['formal']}"]
                for year in option["annual_schedule"]
            ],
            [18 * mm, 43 * mm, 37 * mm, 43 * mm, 35 * mm],
        ))

    story.extend([
        paragraph("7. Budget, monitoring and implementation status", "ReportHeading"),
        paragraph(
            "No provider quotations or approved training budget were present in the supplied evidence. "
            "Accordingly, estimated budgets remain unpriced unless a later agency edit provides an amount. "
            "Actual expenditure, participants trained, results and supporting evidence are not inferred.",
        ),
        report_table(
            ["MONITORING INDICATOR", "CURRENT RECORD"],
            [
                ["Planned programme activities", len(report["programme"])],
                ["Activities marked completed", dashboard["completed_activities"]],
                ["Activities with participant, expenditure or outcome evidence recorded", report["monitored_records"]],
                ["Estimated programme budget", "Not priced" if not report["budget_priced"] else f"PGK {dashboard['estimated_budget_pgk']:,.2f}"],
                ["Recorded actual expenditure", "Not recorded" if dashboard["actual_spend_pgk"] is None else f"PGK {dashboard['actual_spend_pgk']:,.2f}"],
                ["Current activity status summary", _format_frequency(list(report["activity_status"].items()))],
            ],
            [89 * mm, 87 * mm],
        ),
        paragraph(
            "Before approval, assign accountable owners, confirm annual targets and indicators, obtain provider "
            "quotes, verify corporate KRA references, establish baselines, and record participant completion "
            "and workplace outcomes in the M&E tracker.",
        ),
        paragraph("8. Recommendations and proposed actions", "ReportHeading"),
        report_table(
            ["ACTION", "PROPOSED LEAD", "EVIDENCE OF COMPLETION"],
            [
                ["Resolve the two possible name matches and follow up on the 43 roster names without a confirmed CNA match; record outcomes.", "HR / agency CNA focal point", "Updated reconciliation and documented follow-up; no identity inferred by this report."],
                ["Review the lowest-rated institutional items and identify whether root causes call for training, policy, management or institutional action.", "Executive / divisional managers", "Validated issue and selected non-training or training response."],
                ["Validate qualification requirements and competency expectations against approved job descriptions and role profiles.", "HR / line managers", "Approved standards linked to positions; documented manual review."],
                ["Confirm the seven expressed-interest categories against operational demand, affected groups, urgency and corporate objectives.", "Agency leadership / planning unit", "Prioritised and approved needs register with rationale."],
                ["Select and approve one five-year option and annual schedule; confirm 2027–2030 KRA links and the 2031 successor-plan priorities.", "Executive / corporate planning", "Signed-off plan and verified corporate-plan references."],
                ["Apply the 10:20:70 mix and formal-learning eligibility controls; obtain provider quotations before committing funds.", "HR / finance / supervisors", "Eligible participant selection, costed proposals and budget authorization."],
                ["Implement quarterly monitoring of participation, expenditure, completion, workplace application and outcomes.", "Programme owners / M&E focal point", "Baselines, targets, evidence and periodic progress reviews entered in CNAMIS."],
            ],
            [69 * mm, 34 * mm, 73 * mm],
        ),
        paragraph("9. Limitations and management caveats", "ReportHeading"),
        paragraph(
            "The corporate plan supplied covers 2025–2030 and is image-based; exact KRA references must be "
            "confirmed against an approved readable plan, and 2031 successor priorities are not available. "
            "The source data do not provide authoritative employment/training eligibility rules, approved "
            "competency standards, qualification equivalency rules, validated weighted priorities, provider "
            "quotes, approved budgets, or training outcomes. Name reconciliation is provisional and cannot "
            "establish intent to participate. Phrase-matched interest categories overlap. This report is "
            "therefore a management-planning evidence product, not an audit, compliance certification for "
            "individual employees, expenditure approval or training entitlement decision.",
        ),
        paragraph("Appendix: analytical reference", "ReportHeading"),
        report_table(
            ["REFERENCE", "APPLICATION"],
            [
                ["CNA rating scale", "1–3 Poor/Low; 4–8 Average; 9–10 High/Excellent."],
                ["Capability gap", "10 − assessment score; calculated for each assessment item."],
                ["SPA scale", "Separate 1–5 rating; not converted to CNA's 1–10 scale."],
                ["10:20:70 model", "10% formal, 20% social and 70% experiential learning; percentages are not budget allocations."],
                ["Prioritisation", "Validate severity, strategic importance, operational requirements, people affected and urgency; no unsupported weighted ranking is calculated."],
                ["Report generation", f"Generated {report['generated_at']} from the current local OPP database snapshot."],
            ],
            [43 * mm, 133 * mm],
        ),
    ])

    document.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    return stream.getvalue()


def _format_frequency(values: list[tuple[str, int]]) -> str:
    return "; ".join(f"{label}: {count}" for label, count in values) if values else "Not recorded"


class CnamisHandler(BaseHTTPRequestHandler):
    server_version = "CNAMISLocal/1.0"
    max_body_size = 256 * 1024

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def send_json(self, status: int, data: object) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def send_bytes(self, status: int, body: bytes, content_type: str, filename: str | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Cache-Control", "no-store")
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> dict[str, object]:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ValueError("Invalid request body size.") from exc
        if length < 0 or length > self.max_body_size:
            raise ValueError("Request body exceeds the allowed size.")
        raw = self.rfile.read(length)
        if not raw:
            return {}
        try:
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("Request body must be valid JSON.") from exc
        if not isinstance(data, dict):
            raise ValueError("Request body must be a JSON object.")
        return data

    def selected_agency(self, query: dict[str, list[str]], connection: sqlite3.Connection) -> int:
        code = query.get("agency", [AGENCY_CODE])[0]
        return agency_id_for(connection, code)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        query = parse_qs(parsed.query)
        if path == "/" or path == "/index.html":
            self.serve_static("index.html")
            return
        if path == "/image/opp-logo.webp":
            self.serve_logo()
            return
        if path.startswith("/static/"):
            self.serve_static(path.removeprefix("/static/"))
            return
        if not path.startswith("/api/"):
            self.send_error(404)
            return

        try:
            with database_session() as connection:
                if path == "/api/health":
                    self.send_json(200, {"ok": True, "mode": "local", "agency": AGENCY_CODE})
                elif path == "/api/agencies":
                    self.send_json(200, list_agencies(connection))
                else:
                    agency_id = self.selected_agency(query, connection)
                    if path == "/api/dashboard":
                        self.send_json(200, dashboard_data(connection, agency_id))
                    elif path == "/api/diagnostic":
                        self.send_json(200, item_analytics(connection, agency_id))
                    elif path == "/api/needs":
                        self.send_json(200, category_analytics(connection, agency_id))
                    elif path == "/api/skills":
                        self.send_json(200, skills_data(connection, agency_id))
                    elif path == "/api/staff":
                        self.send_json(200, staff_data(connection, agency_id, query.get("q", [""])[0]))
                    elif path == "/api/matches":
                        self.send_json(200, possible_matches(connection, agency_id))
                    elif path == "/api/programme":
                        self.send_json(200, programme_data(connection, agency_id))
                    elif path == "/api/options":
                        self.send_json(200, options_data(connection, agency_id))
                    elif path == "/api/methodology":
                        self.send_json(200, methodology_data(connection, agency_id))
                    elif path == "/api/report.cna.pdf":
                        body = export_capability_report(connection, agency_id)
                        agency_code = connection.execute(
                            "SELECT code FROM agencies WHERE id=?",
                            (agency_id,),
                        ).fetchone()["code"]
                        safe_code = re.sub(r"[^A-Za-z0-9_-]", "", agency_code)
                        self.send_bytes(
                            200,
                            body,
                            "application/pdf",
                            f"{safe_code}-Capability-Needs-Analysis-Report.pdf",
                        )
                    elif path == "/api/export.xlsx":
                        body = export_workbook(connection, agency_id)
                        self.send_bytes(
                            200, body,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "CNAMIS-OPP-report.xlsx",
                        )
                    elif path.startswith("/api/respondents/"):
                        response_id = int(path.rsplit("/", 1)[1])
                        record = response_card(connection, response_id, agency_id)
                        self.send_json(200 if record else 404, record or {"error": "Respondent not found."})
                    else:
                        self.send_json(404, {"error": "API route not found."})
        except ValueError as exc:
            self.send_json(400, {"error": str(exc)})
        except (sqlite3.Error, OSError) as exc:
            self.send_json(500, {"error": f"Local data request failed: {exc}"})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            data = self.read_json()
            with database_session() as connection:
                agency_id = agency_id_for(connection, str(data.get("agency", AGENCY_CODE)))
                if parsed.path == "/api/workspace":
                    result = answer_workspace(connection, agency_id, str(data.get("question", "")))
                    self.send_json(200, result)
                elif parsed.path == "/api/matches/verify":
                    self.verify_match(connection, agency_id, data)
                elif parsed.path == "/api/monitoring":
                    self.save_monitoring(connection, agency_id, data)
                else:
                    self.send_json(404, {"error": "API route not found."})
        except ValueError as exc:
            self.send_json(400, {"error": str(exc)})
        except (sqlite3.Error, OSError) as exc:
            self.send_json(500, {"error": f"Local update failed: {exc}"})

    def do_PATCH(self) -> None:
        parsed = urlparse(self.path)
        if not parsed.path.startswith("/api/programme/"):
            self.send_json(404, {"error": "API route not found."})
            return
        try:
            activity_id = unquote(parsed.path.rsplit("/", 1)[1])
            data = self.read_json()
            with database_session() as connection:
                agency_id = agency_id_for(connection, str(data.pop("agency", AGENCY_CODE)))
                self.update_programme(connection, agency_id, activity_id, data)
        except ValueError as exc:
            self.send_json(400, {"error": str(exc)})
        except (sqlite3.Error, OSError) as exc:
            self.send_json(500, {"error": f"Local update failed: {exc}"})

    def serve_static(self, filename: str) -> None:
        allowed = {"index.html", "app.js", "styles.css"}
        if filename not in allowed:
            self.send_error(404)
            return
        file_path = (STATIC_DIR / filename).resolve()
        if file_path.parent != STATIC_DIR.resolve() or not file_path.is_file():
            self.send_error(404)
            return
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        self.send_bytes(200, file_path.read_bytes(), content_type + ("; charset=utf-8" if content_type.startswith("text/") else ""))

    def serve_logo(self) -> None:
        if not LOGO_FILE.is_file():
            self.send_error(404)
            return
        self.send_bytes(200, LOGO_FILE.read_bytes(), "image/webp")

    def verify_match(self, connection: sqlite3.Connection, agency_id: int, data: dict[str, object]) -> None:
        response_id = _positive_int(data.get("response_id"), "response_id")
        staff_id = _positive_int(data.get("staff_id"), "staff_id")
        response = connection.execute(
            "SELECT * FROM responses WHERE id=? AND agency_id=?",
            (response_id, agency_id),
        ).fetchone()
        staff = connection.execute(
            "SELECT * FROM staff WHERE id=? AND agency_id=?",
            (staff_id, agency_id),
        ).fetchone()
        if response is None or staff is None:
            raise ValueError("Select an existing CNA response and OPP roster officer.")
        if staff_id not in json.loads(response["candidate_staff_ids"]):
            raise ValueError("Selected officer was not among the possible name matches for this response.")
        connection.execute(
            "UPDATE responses SET staff_id=?, match_status='Verified' WHERE id=?",
            (staff_id, response_id),
        )
        _log_change(connection, "response", response_id, "verify_name_match", {"staff_id": staff_id})
        self.send_json(200, {"ok": True, "message": f"Verified {response['name']} against {staff['name']}."})

    def update_programme(
        self, connection: sqlite3.Connection, agency_id: int, activity_id: str, data: dict[str, object]
    ) -> None:
        editable = {
            "year", "category", "focus", "formal_learning", "social_learning", "experiential_learning",
            "owner_role", "timing", "corporate_reference", "estimated_budget", "quote_basis",
            "approval_status", "indicator", "status", "outcome_notes",
        }
        unknown = set(data) - editable
        if unknown:
            raise ValueError("Unsupported plan field(s): " + ", ".join(sorted(unknown)))
        current = connection.execute(
            "SELECT * FROM programme WHERE agency_id=? AND activity_id=?",
            (agency_id, activity_id),
        ).fetchone()
        if current is None:
            raise ValueError("Plan activity was not found.")
        values: dict[str, object] = {}
        for key, value in data.items():
            if key == "year":
                year = _positive_int(value, "year")
                if year < 2027 or year > 2031:
                    raise ValueError("Plan year must be between 2027 and 2031.")
                values[key] = year
            elif key == "estimated_budget":
                if value in (None, ""):
                    values[key] = None
                else:
                    try:
                        amount = float(value)
                    except (TypeError, ValueError) as exc:
                        raise ValueError("Estimated budget must be a number in PGK.") from exc
                    if amount < 0:
                        raise ValueError("Estimated budget cannot be negative.")
                    values[key] = amount
            elif key == "approval_status":
                status = str(value).strip()
                if status not in VALID_PLAN_STATUSES:
                    raise ValueError("Choose a supported approval status.")
                values[key] = status
            elif key == "category":
                valid = {
                    row["name"] for row in connection.execute(
                        "SELECT name FROM need_categories WHERE agency_id=?",
                        (agency_id,),
                    )
                }
                if str(value) not in valid:
                    raise ValueError("Select one of the seven existing training categories.")
                values[key] = str(value)
            else:
                string_value = str(value or "").strip()
                if len(string_value) > 5000:
                    raise ValueError(f"{key} is too long.")
                values[key] = string_value
        if not values:
            raise ValueError("No plan changes were provided.")
        columns = ", ".join(f"{key}=?" for key in values)
        connection.execute(
            f"UPDATE programme SET {columns}, updated_at=? WHERE agency_id=? AND activity_id=?",
            (*values.values(), now_utc(), agency_id, activity_id),
        )
        _log_change(connection, "programme", activity_id, "update", values)
        self.send_json(200, {"ok": True, "activity_id": activity_id, "updated": values})

    def save_monitoring(self, connection: sqlite3.Connection, agency_id: int, data: dict[str, object]) -> None:
        activity_id = str(data.get("activity_id", "")).strip()
        programme = connection.execute(
            "SELECT id FROM programme WHERE agency_id=? AND activity_id=?",
            (agency_id, activity_id),
        ).fetchone()
        if programme is None:
            raise ValueError("Select a valid plan activity.")
        editable = {
            "planned_participants", "participants_completed", "actual_expenditure",
            "indicator", "baseline", "target", "actual_result", "evidence",
        }
        values = {key: data.get(key) for key in editable if key in data}
        for key in ("planned_participants", "participants_completed"):
            if key in values:
                values[key] = None if values[key] in (None, "") else _positive_int(values[key], key, allow_zero=True)
        if "actual_expenditure" in values:
            if values["actual_expenditure"] in (None, ""):
                values["actual_expenditure"] = None
            else:
                try:
                    values["actual_expenditure"] = float(values["actual_expenditure"])
                except (TypeError, ValueError) as exc:
                    raise ValueError("Actual expenditure must be a PGK number.") from exc
                if values["actual_expenditure"] < 0:
                    raise ValueError("Actual expenditure cannot be negative.")
        for key in ("indicator", "baseline", "target", "actual_result", "evidence"):
            if key in values:
                values[key] = str(values[key] or "").strip()
                if len(values[key]) > 5000:
                    raise ValueError(f"{key} is too long.")
        if not values:
            raise ValueError("No monitoring fields were provided.")
        fields = ["planned_participants", "participants_completed", "actual_expenditure", "indicator",
                  "baseline", "target", "actual_result", "evidence"]
        current = connection.execute(
            "SELECT * FROM monitoring WHERE programme_id=?",
            (programme["id"],),
        ).fetchone()
        merged = {field: current[field] for field in fields}
        merged.update(values)
        connection.execute(
            """INSERT INTO monitoring
               (programme_id, planned_participants, participants_completed, actual_expenditure,
                indicator, baseline, target, actual_result, evidence, updated_at)
               VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(programme_id) DO UPDATE SET
                planned_participants=excluded.planned_participants,
                participants_completed=excluded.participants_completed,
                actual_expenditure=excluded.actual_expenditure,
                indicator=excluded.indicator, baseline=excluded.baseline, target=excluded.target,
                actual_result=excluded.actual_result, evidence=excluded.evidence, updated_at=excluded.updated_at""",
            (programme["id"], *(merged[field] for field in fields), now_utc()),
        )
        _log_change(connection, "monitoring", activity_id, "update", values)
        self.send_json(200, {"ok": True, "activity_id": activity_id, "updated": values})


def _positive_int(value: object, field: str, allow_zero: bool = False) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field} must be a whole number.") from exc
    if parsed < (0 if allow_zero else 1):
        raise ValueError(f"{field} must be {'zero or greater' if allow_zero else 'greater than zero'}.")
    return parsed


def _log_change(connection: sqlite3.Connection, entity: str, entity_id: object, action: str, changes: object) -> None:
    connection.execute(
        "INSERT INTO audit_log(entity, entity_id, action, changes, created_at) VALUES(?, ?, ?, ?, ?)",
        (entity, str(entity_id), action, json.dumps(changes, ensure_ascii=False), now_utc()),
    )


def main() -> None:
    initialize_database()
    server = ThreadingHTTPServer((HOST, PORT), CnamisHandler)
    print(f"CNAMIS for OPP is ready at http://{HOST}:{PORT}")
    print(f"Local database: {db_path()}")
    print("Press Ctrl+C to stop the server.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping CNAMIS.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
