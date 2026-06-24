from __future__ import annotations

import argparse
import csv
import io
import json
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any


REQUIRED_COLUMNS = (
    "Action Items",
    "Status",
    "Owner",
    "Priority",
    "Notes",
    "Modified",
    "Due Date",
    "TestRail/JIRA Link",
)

ACTIVE_STATUSES = {
    "in progress": "In Progress",
    "on hold": "On Hold",
    "ready for uat": "Ready for UAT",
    "client review": "Client Review",
}
KNOWN_STATUSES = {**ACTIVE_STATUSES, "completed": "Completed"}
PRIORITY_ORDER = {
    "Critical": 0,
    "High": 1,
    "Normal": 2,
    "Low": 3,
    "Unassigned": 4,
}
PRIORITY_WORDS = {
    "critical": "Critical",
    "high": "High",
    "normal": "Normal",
    "low": "Low",
}
DATED_UPDATE_RE = re.compile(
    r"(?im)^[ \t]*(?P<date>\d{1,2}/\d{1,2}(?:/\d{2,4})?)[ \t]*"
    r"(?:-|–|—|:)[ \t]*"
)
SENTENCE_END_RE = re.compile(r"(?<=[.!?])\s+")
LEADING_FILLER_PATTERNS = (
    re.compile(r"(?i)^(?:as of|on)\s+[^,.;:]{1,30}[,;:]\s*"),
    re.compile(r"(?i)^(?:during|following)\s+(?:our|the)\s+call[,;:]?\s*"),
    re.compile(r"(?i)^on\s+(?:our|the)\s+call[,;:]?\s*"),
)
FIRST_PERSON_REPLACEMENTS = (
    (re.compile(r"(?i)^we\s+are\b"), "The team is"),
    (re.compile(r"(?i)^we're\b"), "The team is"),
    (re.compile(r"(?i)^we\s+have\b"), "The team has"),
    (re.compile(r"(?i)^we\s+will\b"), "The team will"),
    (re.compile(r"(?i)^we\s+need\s+to\b"), "The team needs to"),
    (re.compile(r"(?i)^we\b"), "The team"),
    (re.compile(r"(?i)^i\s+am\b"), "The team is"),
    (re.compile(r"(?i)^i\s+will\b"), "The team will"),
    (re.compile(r"(?i)^i\b"), "The team"),
    (re.compile(r"(?i)^our\s+team\b"), "The team"),
)


class AdiExportError(ValueError):
    """Raised when the ADI CSV or base dashboard data cannot be processed."""


@dataclass(frozen=True)
class ParsedAdiItem:
    raid_item: dict[str, str]
    due_date: date | None
    warnings: list[str]


def clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def collapse_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", clean(value)).strip()


def truncate_cleanly(value: str, limit: int) -> str:
    text = collapse_whitespace(value)
    if len(text) <= limit:
        return text
    if limit <= 1:
        return text[:limit]
    shortened = text[:limit].rsplit(" ", 1)[0].rstrip(" ,;:-")
    if not shortened:
        shortened = text[: limit - 1].rstrip()
    return shortened.rstrip(".") + "."


def extract_latest_update(notes: str) -> str:
    """Return only the first/top update, without its date prefix."""
    text = clean(notes).replace("\r\n", "\n").replace("\r", "\n")
    if not text:
        return ""

    dated_updates = list(DATED_UPDATE_RE.finditer(text))
    if dated_updates:
        first = dated_updates[0]
        next_start = dated_updates[1].start() if len(dated_updates) > 1 else len(text)
        text = text[first.end() : next_start]

    return collapse_whitespace(text)


def rewrite_common_project_status(title: str, update: str) -> str | None:
    """Return a confident rewrite for common project-management note patterns."""
    title_text = collapse_whitespace(title)
    combined = f"{title_text} {update}".casefold()

    if (
        "chart of accounts" in combined
        and re.search(r"(?i)\b(?:we discussed|it was discussed)\b", update)
        and re.search(r"(?i)\b(?:confirmed|feedback|keep the gl|coa)\b", update)
    ):
        return (
            "Danvers is updating the chart of accounts structure based on the "
            "latest project feedback."
        )

    if (
        "munis" in combined
        and re.search(r"(?i)^Cogsdale requested (?:a )?meeting\b", update)
        and re.search(r"(?i)\bopen questions?\b", update)
    ):
        return (
            "Cogsdale is coordinating with Danvers to resolve open Munis "
            "integration design questions."
        )

    fixed_assets_match = re.search(
        r"(?i)^(?:we|the team)\s+confirmed\s+Fixed Assets as a dimension\b",
        update,
    )
    if fixed_assets_match:
        return "The team confirmed Fixed Assets as a dimension."

    if (
        re.search(r"(?i)^Danvers is targeting (?:to deliver|delivery of)\b", update)
        and re.search(r"(?i)\bFinance\b.*\btemplate\b", update)
    ):
        return (
            "Danvers is targeting completion of the finance data preparation "
            "template."
        )

    discussed_match = re.match(
        r"(?i)^(?:on (?:our|the) call[,;]?\s*)?"
        r"(?:we discussed|it was discussed)(?:\s+that)?\s+(.+)$",
        update,
    )
    if discussed_match:
        discussed_text = discussed_match.group(1).strip()
        first_sentence = SENTENCE_END_RE.split(discussed_text, maxsplit=1)[0].strip()
        if re.match(
            r"(?i)^(?:we|the team|Danvers|Cogsdale|[A-Z][a-z]+)\s+"
            r"(?:is|are|has|have|will|confirmed|agreed|approved|requested)\b",
            first_sentence,
        ):
            return neutralize_first_person(first_sentence)
        return (
            f"The team is reviewing {title_text or 'this item'} based on the "
            "latest project discussion."
        )

    return None


def neutralize_first_person(value: str) -> str:
    summary = collapse_whitespace(value)
    for pattern, replacement in FIRST_PERSON_REPLACEMENTS:
        if pattern.search(summary):
            return pattern.sub(replacement, summary, count=1)
    return summary


def summarize_current_status(title: str, notes: str) -> str:
    """Create a concise, rule-based current-status sentence for the dashboard."""
    update = extract_latest_update(notes)
    if not update:
        return "No recent update provided."

    confident_rewrite = rewrite_common_project_status(title, update)
    summary = confident_rewrite or update
    for pattern in LEADING_FILLER_PATTERNS:
        summary = pattern.sub("", summary).strip()
    summary = neutralize_first_person(summary)

    summary = re.sub(r"(?i)^(?:update|status)\s*:\s*", "", summary).strip()
    summary = re.sub(r"^[\-–—:;,\s]+", "", summary)
    summary = collapse_whitespace(summary)

    sentences = [part.strip() for part in SENTENCE_END_RE.split(summary) if part.strip()]
    if sentences:
        summary = sentences[0]

    if not summary:
        summary = f"Work on {clean(title) or 'this item'} is being reviewed."
    if summary[0].islower():
        summary = summary[0].upper() + summary[1:]
    summary = truncate_cleanly(summary, 140)
    if summary[-1] not in ".!?":
        summary += "."
    return summary


def build_description(title: str, status_summary: str, limit: int = 220) -> str:
    label = "\nStatus: "
    status = collapse_whitespace(status_summary)
    minimum_status_space = min(40, len(status))
    heading_limit = max(1, limit - len(label) - minimum_status_space)
    heading = truncate_cleanly(title, heading_limit)
    status = collapse_whitespace(status_summary)
    prefix = f"{heading}{label}"
    available = max(1, limit - len(prefix))
    return prefix + truncate_cleanly(status, available)


def normalize_priority(value: str) -> tuple[str, bool]:
    text = clean(value)
    if not text:
        return "Unassigned", False

    words = re.findall(r"[A-Za-z]+", text.lower())
    matches = [PRIORITY_WORDS[word] for word in words if word in PRIORITY_WORDS]
    if len(set(matches)) == 1:
        return matches[0], True
    return "Unassigned", False


def normalize_status(value: str) -> tuple[str, bool]:
    text = clean(value)
    normalized_key = re.sub(r"\s+", " ", text).casefold()
    if normalized_key in KNOWN_STATUSES:
        return KNOWN_STATUSES[normalized_key], True
    return text or "Unrecognized", False


def parse_date(value: str) -> date | None:
    text = clean(value)
    if not text:
        return None

    normalized = text.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(normalized).date()
    except ValueError:
        pass

    formats = (
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y-%m-%d",
        "%B %d, %Y",
        "%b %d, %Y",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M",
    )
    for date_format in formats:
        try:
            return datetime.strptime(text, date_format).date()
        except ValueError:
            continue
    return None


def normalized_date_text(value: str) -> str:
    parsed = parse_date(value)
    return parsed.isoformat() if parsed else clean(value)


def read_csv_text(csv_text: str) -> list[dict[str, str]]:
    reader = csv.DictReader(io.StringIO(csv_text, newline=""))
    if reader.fieldnames is None:
        raise AdiExportError("The ADI CSV does not contain a header row.")

    normalized_headers = [clean(header) for header in reader.fieldnames]
    duplicate_headers = {
        header for header in normalized_headers if normalized_headers.count(header) > 1
    }
    if duplicate_headers:
        raise AdiExportError(
            "The ADI CSV has duplicate columns after trimming spaces: "
            + ", ".join(sorted(duplicate_headers))
        )

    missing = [column for column in REQUIRED_COLUMNS if column not in normalized_headers]
    if missing:
        raise AdiExportError(
            "The ADI CSV is missing required column(s): " + ", ".join(missing)
        )

    reader.fieldnames = normalized_headers
    return [
        {clean(key): clean(value) for key, value in row.items() if key is not None}
        for row in reader
        if any(clean(value) for value in row.values())
    ]


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        raise AdiExportError(f"ADI CSV file was not found: {path}")
    if not path.is_file():
        raise AdiExportError(f"ADI input path is not a file: {path}")

    try:
        csv_text = path.read_text(encoding="utf-8-sig")
        return read_csv_text(csv_text)
    except UnicodeDecodeError as exc:
        raise AdiExportError(
            "The ADI CSV could not be read as UTF-8. Re-save it as CSV UTF-8 in Excel."
        ) from exc
    except OSError as exc:
        raise AdiExportError(f"Could not read ADI CSV '{path}': {exc}") from exc


def parse_row(row: dict[str, str], row_number: int, today: date) -> ParsedAdiItem:
    title = clean(row["Action Items"]) or f"Untitled ADI item (CSV row {row_number})"
    notes = clean(row["Notes"])
    latest_update = extract_latest_update(notes)
    status_summary = summarize_current_status(title, notes)
    owner = clean(row["Owner"])
    due_text = clean(row["Due Date"])
    due_date = parse_date(due_text)
    priority, recognized_priority = normalize_priority(row["Priority"])
    status, recognized_status = normalize_status(row["Status"])

    description = build_description(title, status_summary)

    warnings = []
    prefix = f"CSV row {row_number} - {title}"
    if not owner:
        warnings.append(f"{prefix}: missing owner.")
    if not due_text:
        warnings.append(f"{prefix}: missing due date.")
    elif due_date is None:
        warnings.append(f"{prefix}: due date '{due_text}' was not recognized.")
    if priority in {"Critical", "High"} and status in ACTIVE_STATUSES.values():
        warnings.append(f"{prefix}: {priority} priority active item requires PM review.")
    if due_date is not None and due_date < today and status in ACTIVE_STATUSES.values():
        warnings.append(f"{prefix}: item is overdue ({due_date.isoformat()}).")
    if status == "On Hold":
        warnings.append(f"{prefix}: item is On Hold.")
    if status == "Client Review" and not due_text:
        warnings.append(f"{prefix}: Client Review item has no due date.")
    if not notes:
        warnings.append(f"{prefix}: missing latest update in Notes.")
    if not recognized_priority:
        warnings.append(
            f"{prefix}: priority '{clean(row['Priority']) or '(blank)'}' is unrecognized; "
            "using Unassigned."
        )
    if not recognized_status:
        warnings.append(
            f"{prefix}: status '{clean(row['Status']) or '(blank)'}' is unrecognized "
            "and will not be included in the dashboard."
        )

    raid_item = {
        "type": "Action",
        "title": title,
        "latest_update": latest_update,
        "description": description,
        "priority": priority,
        "status": status,
        "assigned": owner,
        "due_date": due_date.isoformat() if due_date else due_text,
        "last_updated": normalized_date_text(row["Modified"]),
        "source": "ADI List",
        "reference_link": clean(row["TestRail/JIRA Link"]),
        "attachments": clean(row.get("Attachments", "")),
    }
    return ParsedAdiItem(raid_item, due_date, warnings)


def load_base_dashboard(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise AdiExportError(f"Base dashboard JSON file was not found: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise AdiExportError(
            f"Base dashboard JSON is invalid near line {exc.lineno}: {exc.msg}."
        ) from exc
    if not isinstance(data, dict):
        raise AdiExportError("Base dashboard JSON must contain a JSON object.")
    for section in ("project_summary", "deliverables"):
        if section not in data:
            raise AdiExportError(
                f"Base dashboard JSON is missing required section: {section}."
            )
    return data


def rank_key(item: ParsedAdiItem) -> tuple[int, bool, date]:
    return (
        PRIORITY_ORDER[item.raid_item["priority"]],
        item.due_date is None,
        item.due_date or date.max,
    )


def build_dashboard_data(
    rows: list[dict[str, str]],
    base_dashboard: dict[str, Any],
    limit: int,
    today: date,
) -> tuple[dict[str, Any], list[str], int]:
    parsed_items = [
        parse_row(row, row_number=index + 2, today=today)
        for index, row in enumerate(rows)
    ]
    warnings = [warning for item in parsed_items for warning in item.warnings]
    active_items = [
        item
        for item in parsed_items
        if item.raid_item["status"] in ACTIVE_STATUSES.values()
    ]
    active_items.sort(key=rank_key)
    selected = active_items[:limit]

    output = dict(base_dashboard)
    output["raid_items"] = [item.raid_item for item in selected]
    output["pm_review_warnings"] = warnings
    output["adi_import_summary"] = {
        "source": "ADI List CSV export",
        "rows_read": len(rows),
        "active_rows": len(active_items),
        "dashboard_rows": len(selected),
        "limit": limit,
    }
    return output, warnings, len(active_items)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert an exported ADI List CSV into dashboard JSON data."
    )
    parser.add_argument("--input", type=Path, required=True, help="Path to the ADI CSV.")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("sample_inputs/dashboard_from_adi.json"),
        help="Output dashboard JSON path.",
    )
    parser.add_argument(
        "--base",
        type=Path,
        default=Path("sample_inputs/dashboard_sample.json"),
        help="Dashboard JSON whose project summary and deliverables should be preserved.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Maximum active ADI items to include. Default: 5.",
    )
    return parser.parse_args()


def main() -> int:
    arguments = parse_args()
    try:
        if arguments.limit < 1:
            raise AdiExportError("--limit must be at least 1.")
        rows = read_csv_rows(arguments.input)
        base_dashboard = load_base_dashboard(arguments.base)
        output, warnings, active_count = build_dashboard_data(
            rows, base_dashboard, arguments.limit, date.today()
        )
        arguments.output.parent.mkdir(parents=True, exist_ok=True)
        arguments.output.write_text(
            json.dumps(output, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
    except (AdiExportError, OSError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2

    print(
        f"Created {arguments.output} with "
        f"{len(output['raid_items'])} of {active_count} active ADI items."
    )
    if warnings:
        print(f"\nPM review warnings ({len(warnings)}):")
        for warning in warnings:
            print(f"- {warning}")
    else:
        print("No PM review warnings.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
