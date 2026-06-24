from __future__ import annotations

import re
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

from docx import Document


MAX_ITEMS = 5
MAX_LOW_CONFIDENCE_ITEMS = 10
MAX_NARRATIVE_SENTENCES = 3
INSUFFICIENT_CONTENT_WARNING = (
    "The transcript did not contain enough high-confidence project status "
    "content to generate a reliable narrative."
)

PROJECT_TERMS = {
    "business central": "Business Central",
    "bc": "Business Central",
    "power bi": "Power BI",
    "chart of accounts": "chart of accounts",
    "coa": "chart of accounts",
    "dimensions": "dimensions",
    "fixed assets": "fixed assets",
    "uat": "UAT",
    "test cases": "test cases",
    "training plan": "training plan",
    "configuration": "configuration",
    "data migration": "data migration",
    "data conversion": "data conversion",
    "master data": "master data",
    "balances": "balances",
    "test environment": "test environment",
    "production": "production environment",
    "integration": "integration",
    "interface": "interface design",
    "payroll": "payroll",
    "munis": "Munis integration",
    "northstar": "NorthStar",
    "smartconnect": "SmartConnect",
    "go-live": "go-live",
    "go live": "go-live",
    "cutover": "cutover",
    "hypercare": "hypercare",
    "validation": "validation",
    "signoff": "signoff",
    "sign off": "signoff",
    "deliverable": "deliverables",
}
PROJECT_IMPACT_TERMS = (
    "issue",
    "risk",
    "blocker",
    "blocked",
    "decision",
    "dependency",
    "action item",
    "client needs to",
    "waiting on client",
    "due date",
    "target date",
    "delay",
    "impact",
    "cannot move forward",
)
CATEGORY_KEYWORDS = {
    "accomplishments": (
        "completed",
        "finalized",
        "confirmed",
        "approved",
        "resolved",
        "delivered",
        "validated",
    ),
    "upcoming_focus_areas": (
        "next step",
        "next steps",
        "will continue",
        "continue working",
        "working on",
        "plan to",
        "plans to",
        "targeting",
        "prepare",
        "needs to finalize",
        "needs to complete",
    ),
    "blockers_or_concerns": (
        "blocked",
        "delayed",
        "waiting on client",
        "waiting on",
        "risk",
        "issue",
        "concern",
        "dependency",
        "behind",
        "cannot move forward",
        "may impact",
        "may delay",
    ),
    "decisions_needed": (
        "decision",
        "need to confirm",
        "needs to confirm",
        "need confirmation",
        "waiting for approval",
        "determine whether",
        "choose",
        "sign off",
        "signoff",
        "needs to finalize",
    ),
    "possible_raid_items_for_review": (
        "action item",
        "follow up",
        "follow-up",
        "due date",
        "assigned to",
        "needs to",
        "responsible for",
        "open item",
        "create a",
        "confirm whether",
    ),
}
CATEGORY_THRESHOLDS = {
    "accomplishments": 3,
    "upcoming_focus_areas": 3,
    "blockers_or_concerns": 4,
    "decisions_needed": 4,
    "possible_raid_items_for_review": 4,
}
CATEGORY_LABELS = {
    "accomplishments": "Accomplishment",
    "upcoming_focus_areas": "Upcoming focus area",
    "blockers_or_concerns": "Blocker or concern",
    "decisions_needed": "Decision needed",
    "possible_raid_items_for_review": "Possible RAID/ADI item",
}

VTT_TIMESTAMP_RE = re.compile(
    r"^\s*(?:\d{1,2}:)?\d{2}:\d{2}[.,]\d{3}\s*-->\s*"
    r"(?:\d{1,2}:)?\d{2}:\d{2}[.,]\d{3}.*$"
)
SPEAKER_LABEL_RE = re.compile(
    r"^\s*(?:\[[^\]]+\]|[A-Z][A-Za-z .'-]{0,40}):\s*"
)
SPEAKER_TIMESTAMP_RE = re.compile(
    r"\b[A-Z][A-Za-z'-]+(?:\s+[A-Z][A-Za-z'-]+){1,2}\s+"
    r"\d{1,2}:\d{2}(?::\d{2})?\.?\s*"
)
STANDALONE_TIMESTAMP_RE = re.compile(
    r"(?<!\w)\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d{1,3})?\.?(?!\w)"
)
LEADING_FILLER_RE = re.compile(
    r"(?i)^(?:(?:yeah|right|okay|ok|obviously|so|well|and|um|uh)"
    r"[,.!?;:\s]+)+"
)
SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|\n+")
FILLER_ONLY_RE = re.compile(
    r"^(?:okay|ok|yes|yeah|yep|no|right|great|thanks|thank you|"
    r"sounds good|all right|um|uh|hello|hi|good morning|good afternoon)[.!?]*$",
    re.IGNORECASE,
)
MEETING_HOUSEKEEPING_RE = re.compile(
    r"\b(?:reviewed? the agenda|meeting agenda|round of introductions|"
    r"introduced themselves|recording (?:has )?started|can you hear me|"
    r"general conversation|sorry to interrupt)\b",
    re.IGNORECASE,
)
PERSONAL_OR_LOGISTICS_RE = re.compile(
    r"\b(?:pto|vacation|doctor(?:'s)? appointment|personal appointment|"
    r"out of office|oof|travel(?:ing)?|calendar availability|"
    r"I won'?t be here|I won'?t see you|out next week|away next week|"
    r"schedule (?:the|our|a) next meeting|reschedule (?:the|our|a) meeting|"
    r"meeting (?:time|invite)|session next week may move|"
    r"(?:move|schedule|reschedule).{0,30}(?:meeting|session)|"
    r"(?:meeting|session).{0,30}(?:move|schedule|reschedule)|weather)\b",
    re.IGNORECASE,
)
UNRELATED_EMPLOYEE_RE = re.compile(
    r"\b(?:other employee|another employee|one of my (?:other )?employees|"
    r"employee (?:issue|problem)|hr issue|personnel issue)\b",
    re.IGNORECASE,
)
PROJECT_IMPACT_RE = re.compile(
    r"\b(?:delay|delayed|impact|risk|blocked|blocker|dependency|"
    r"cannot move forward|schedule impact|project impact)\b",
    re.IGNORECASE,
)
STOP_WORDS = {
    "about",
    "after",
    "again",
    "also",
    "and",
    "are",
    "been",
    "before",
    "being",
    "client",
    "for",
    "from",
    "have",
    "into",
    "next",
    "project",
    "status",
    "team",
    "that",
    "the",
    "their",
    "this",
    "through",
    "will",
    "with",
    "work",
}


class MeetingTranscriptError(ValueError):
    """Raised when an uploaded meeting transcript cannot be read."""


def collapse_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def clean_transcript_artifacts(text: str) -> str:
    cleaned = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    cleaned = re.sub(
        r"(?m)^\s*(?:\d{1,2}:)?\d{2}:\d{2}[.,]\d{3}\s*-->.*$",
        " ",
        cleaned,
    )
    cleaned = SPEAKER_TIMESTAMP_RE.sub(" ", cleaned)
    cleaned = STANDALONE_TIMESTAMP_RE.sub(" ", cleaned)
    cleaned = re.sub(r"(?i)\b(if it'?s|you know|I mean)\s+\1\b", r"\1", cleaned)
    cleaned = collapse_whitespace(cleaned)
    cleaned = LEADING_FILLER_RE.sub("", cleaned)
    cleaned = re.sub(
        r"(?i)(?<=[.!?])\s+(?:(?:yeah|right|okay|ok|obviously|so|well)"
        r"[,.!?;:\s]+)+",
        " ",
        cleaned,
    )
    cleaned = re.sub(r"\s+([,.!?;:])", r"\1", cleaned)
    cleaned = re.sub(r"(?:\.\s*){2,}", ". ", cleaned)
    return collapse_whitespace(cleaned).strip(" ,;:-")


def read_transcript(file_bytes: bytes, file_name: str) -> str:
    suffix = Path(file_name).suffix.casefold()
    if suffix in {".txt", ".vtt"}:
        for encoding in ("utf-8-sig", "utf-8", "cp1252"):
            try:
                return file_bytes.decode(encoding)
            except UnicodeDecodeError:
                continue
        raise MeetingTranscriptError(
            "The transcript could not be decoded. Save it as UTF-8 text and try again."
        )
    if suffix == ".docx":
        try:
            document = Document(BytesIO(file_bytes))
        except Exception as exc:
            raise MeetingTranscriptError(
                "The Word transcript could not be read. Confirm it is a valid .docx file."
            ) from exc
        return "\n".join(paragraph.text for paragraph in document.paragraphs)
    raise MeetingTranscriptError(
        "Unsupported transcript type. Upload a TXT, VTT, or DOCX file."
    )


def clean_transcript_text(text: str, is_vtt: bool = False) -> str:
    cleaned_lines = []
    previous_line = ""
    for raw_line in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw_line.strip()
        if not line:
            continue
        if is_vtt and (
            line.upper() == "WEBVTT"
            or VTT_TIMESTAMP_RE.match(line)
            or line.isdigit()
            or line.upper().startswith(("NOTE", "STYLE", "REGION"))
        ):
            continue
        line = re.sub(r"<v(?:\.[^ >]+)?(?:\s+[^>]+)?>", "", line, flags=re.IGNORECASE)
        line = re.sub(r"</?[^>]+>", "", line)
        line = SPEAKER_LABEL_RE.sub("", line)
        line = clean_transcript_artifacts(line)
        if not line or FILLER_ONLY_RE.match(line):
            continue
        if line.casefold() == previous_line.casefold():
            continue
        cleaned_lines.append(line)
        previous_line = line
    return "\n".join(cleaned_lines)


def split_sentences(text: str) -> list[str]:
    sentences = []
    seen = set()
    for part in SENTENCE_SPLIT_RE.split(text):
        sentence = collapse_whitespace(part).strip(" -")
        if (
            len(sentence) < 12
            or FILLER_ONLY_RE.match(sentence)
            or MEETING_HOUSEKEEPING_RE.search(sentence)
        ):
            continue
        if sentence[-1] not in ".!?":
            sentence += "."
        key = sentence.casefold()
        if key not in seen:
            sentences.append(sentence)
            seen.add(key)
    return sentences


def has_project_anchor(text: str, project_context: dict[str, Any] | None = None) -> bool:
    lowered = text.casefold()
    if any(_phrase_present(lowered, term) for term in PROJECT_TERMS):
        return True
    return bool(_context_matches(text, project_context))


def is_irrelevant_transcript_content(text: str) -> bool:
    content = collapse_whitespace(text)
    if not content or FILLER_ONLY_RE.match(content) or MEETING_HOUSEKEEPING_RE.search(content):
        return True
    personal_or_unrelated = bool(
        PERSONAL_OR_LOGISTICS_RE.search(content) or UNRELATED_EMPLOYEE_RE.search(content)
    )
    if not personal_or_unrelated:
        return False
    return not (PROJECT_IMPACT_RE.search(content) and has_project_anchor(content))


def build_context_chunks(sentences: list[str]) -> list[str]:
    chunks = []
    seen = set()
    for start in range(len(sentences)):
        parts = []
        for end in range(start, min(len(sentences), start + 3)):
            sentence = sentences[end]
            if is_irrelevant_transcript_content(sentence):
                if not parts:
                    continue
                break
            parts.append(sentence)
            chunk = collapse_whitespace(" ".join(parts))
            key = chunk.casefold()
            if key not in seen:
                chunks.append(chunk)
                seen.add(key)
    return chunks


def _phrase_present(lowered_text: str, phrase: str) -> bool:
    return bool(re.search(rf"(?<!\w){re.escape(phrase.casefold())}(?!\w)", lowered_text))


def _meaningful_tokens(value: str) -> set[str]:
    return {
        token
        for token in re.findall(r"[a-z0-9]+", value.casefold())
        if len(token) >= 3 and token not in STOP_WORDS
    }


def _context_values(project_context: dict[str, Any] | None) -> list[str]:
    if not project_context:
        return []
    values = []
    for key in ("project_name", "project_status_narrative"):
        value = project_context.get(key, "")
        if value:
            values.append(collapse_whitespace(str(value)))
    for key in ("adi_item_titles", "adi_titles", "clarizen_deliverable_names"):
        for value in project_context.get(key, []) or []:
            if value:
                values.append(collapse_whitespace(str(value)))
    return values


def _context_matches(text: str, project_context: dict[str, Any] | None) -> list[str]:
    lowered = text.casefold()
    text_tokens = _meaningful_tokens(text)
    matches = []
    for value in _context_values(project_context):
        value_lower = value.casefold()
        if len(value_lower) >= 5 and value_lower in lowered:
            matches.append(value)
            continue
        overlap = text_tokens & _meaningful_tokens(value)
        if len(overlap) >= 2:
            matches.append(value)
    return matches


def score_project_relevance(
    text: str, project_context: dict[str, Any] | None = None
) -> int:
    content = collapse_whitespace(text)
    if not content:
        return 0
    lowered = content.casefold()
    score = 0
    matched_terms = [term for term in PROJECT_TERMS if _phrase_present(lowered, term)]
    score += min(6, len(matched_terms) * 2)
    impact_matches = [term for term in PROJECT_IMPACT_TERMS if term in lowered]
    score += min(3, len(impact_matches))
    score += min(4, len(_context_matches(content, project_context)) * 2)
    if any(
        keyword in lowered
        for keyword in (
            "completed",
            "finalized",
            "confirmed",
            "approved",
            "validated",
            "will continue",
            "plan to",
            "needs to",
            "need to confirm",
            "follow up",
            "follow-up",
            "create a",
            "confirm whether",
            "responsible for",
        )
    ):
        score += 1
    if is_irrelevant_transcript_content(content):
        score -= 5
    return max(0, score)


def keyword_matches(text: str, keywords: Iterable[str]) -> list[str]:
    lowered = text.casefold()
    return [keyword for keyword in keywords if keyword in lowered]


def _truncate_bullet(value: str, limit: int = 160) -> str:
    text = collapse_whitespace(value).strip(" -")
    if len(text) > limit:
        shortened = text[:limit].rsplit(" ", 1)[0].rstrip(" ,;:-")
        text = shortened or text[:limit]
    text = text.rstrip(".!?")
    return text + "." if text else ""


def _best_project_sentence(text: str, category: str) -> str:
    sentences = [
        collapse_whitespace(part)
        for part in re.split(r"(?<=[.!?])\s+", text)
        if collapse_whitespace(part)
    ]
    if not sentences:
        return collapse_whitespace(text)
    keywords = CATEGORY_KEYWORDS.get(category, ())
    return max(
        sentences,
        key=lambda sentence: (
            score_project_relevance(sentence),
            len(keyword_matches(sentence, keywords)),
            len(sentence),
        ),
    )


def identify_project_topics(text: str) -> list[str]:
    lowered = clean_transcript_artifacts(text).casefold()
    topics = []

    def add(topic: str) -> None:
        if topic not in topics:
            topics.append(topic)

    if "test" in lowered and any(
        term in lowered
        for term in ("production", "prod", "master data", "balance", "data")
    ):
        add("data_environment")
    if "power bi" in lowered and any(
        term in lowered for term in ("business central", "configuration")
    ) or ("power bi" in lowered and _phrase_present(lowered, "bc")):
        add("power_bi_business_central")
    if "shell" in lowered and any(
        term in lowered for term in ("chart of accounts", "old coa", "legacy coa")
    ):
        add("shell_company_chart_of_accounts")
    if "fixed assets" in lowered:
        add("fixed_assets")
    if "custom report" in lowered or "reporting needs" in lowered:
        add("custom_reports")
    if "payroll" in lowered:
        add("payroll_integration")
    if "integration" in lowered or "interface" in lowered:
        add("integration")
    if "data migration" in lowered or "data conversion" in lowered:
        add("data_migration")
    if ("chart of accounts" in lowered or _phrase_present(lowered, "coa")) and not any(
        topic == "shell_company_chart_of_accounts" for topic in topics
    ):
        add("chart_of_accounts")
    if "uat" in lowered or "test case" in lowered:
        add("uat_testing")
    if (
        "validation" in lowered
        and "business central" not in lowered
        and not _phrase_present(lowered, "bc")
        and not any(
            topic in topics
            for topic in (
                "uat_testing",
                "data_environment",
                "fixed_assets",
                "business_central",
                "power_bi_business_central",
            )
        )
    ):
        add("validation")
    if "master data" in lowered and "data_environment" not in topics:
        add("master_data")
    if "power bi" in lowered and "power_bi_business_central" not in topics:
        add("power_bi")
    if ("business central" in lowered or _phrase_present(lowered, "bc")) and not any(
        topic in topics
        for topic in ("power_bi_business_central", "shell_company_chart_of_accounts")
    ):
        add("business_central")
    if "configuration" in lowered and not any(
        topic in topics
        for topic in ("fixed_assets", "power_bi_business_central", "business_central")
    ):
        add("configuration")
    return topics


def _topic_label(topic: str) -> str:
    return {
        "data_environment": "master data and balance movement into the test environment",
        "power_bi_business_central": "Power BI and Business Central configuration",
        "shell_company_chart_of_accounts": "the legacy chart of accounts",
        "fixed_assets": "fixed assets configuration",
        "custom_reports": "custom reporting requirements",
        "payroll_integration": "payroll integration",
        "integration": "integration design",
        "data_migration": "data migration and conversion",
        "chart_of_accounts": "chart of accounts configuration",
        "uat_testing": "UAT preparation and testing",
        "validation": "data validation",
        "master_data": "master data requirements",
        "power_bi": "Power BI configuration",
        "business_central": "Business Central configuration",
        "configuration": "configuration activities",
    }.get(topic, "project work")


def _canonical_topic_bullet(topic: str, category: str) -> str:
    if category == "accomplishments":
        return _truncate_bullet(
            f"Completed the current {_topic_label(topic)} work for this project phase"
        )
    if category == "blockers_or_concerns":
        return _truncate_bullet(
            f"Resolve the open dependency affecting {_topic_label(topic)} to avoid further project impact"
        )

    if topic == "data_environment":
        if category == "possible_raid_items_for_review":
            return (
                "Confirm ownership and timing for refreshing the test environment "
                "with current master data and balances."
            )
        return (
            "Confirm the approach for moving master data and balances from "
            "production into the test environment."
        )
    if topic == "power_bi_business_central":
        return (
            "Review Power BI and Business Central configuration once the test "
            "environment is updated."
        )
    if topic == "shell_company_chart_of_accounts":
        return (
            "Determine whether a Business Central shell company is needed to "
            "retain the legacy chart of accounts."
        )
    if topic == "fixed_assets":
        return (
            "Confirm the fixed assets configuration approach and remaining data "
            "needs before testing begins."
        )
    if topic == "custom_reports":
        return "Review reporting needs and determine whether any custom reports are required."
    if topic == "payroll_integration":
        return (
            "Update the payroll integration document and confirm any remaining "
            "open questions."
        )
    if topic == "integration":
        return "Confirm the integration approach and resolve remaining design questions."
    if topic == "data_migration":
        return "Finalize the data migration and conversion approach before testing begins."
    if topic == "chart_of_accounts":
        return "Finalize the chart of accounts structure and confirm remaining mapping decisions."
    if topic == "uat_testing":
        return "Prepare the remaining test cases and validation activities for UAT."
    if topic == "validation":
        return "Validate the remaining project data and confirm readiness for testing."
    if topic == "master_data":
        return "Review master data requirements and confirm remaining validation needs."
    if topic == "power_bi":
        return "Review Power BI configuration and confirm remaining reporting requirements."
    if topic == "business_central":
        return "Continue Business Central configuration and confirm remaining setup requirements."
    if topic == "configuration":
        return "Continue configuration activities and confirm remaining setup requirements."
    return ""


def rewrite_topic_as_project_bullet(
    topic: str, source_chunks: list[str], category: str
) -> str:
    del source_chunks
    return _canonical_topic_bullet(topic, category)


def rewrite_as_project_bullet(text: str, category: str) -> str:
    cleaned = clean_transcript_artifacts(text)
    topics = identify_project_topics(cleaned)
    if not topics:
        return ""
    return rewrite_topic_as_project_bullet(topics[0], [cleaned], category)


def is_high_quality_project_bullet(text: str) -> bool:
    bullet = collapse_whitespace(text)
    lowered = bullet.casefold()
    if not bullet or len(re.findall(r"\b[\w'-]+\b", bullet)) < 8:
        return False
    if SPEAKER_TIMESTAMP_RE.search(bullet) or STANDALONE_TIMESTAMP_RE.search(bullet):
        return False
    if bullet.endswith("?") or "?" in bullet:
        return False
    if re.match(r"(?i)^(?:yeah|yep|right|okay|sorry|address)\b", bullet):
        return False
    if any(
        phrase in lowered
        for phrase in (
            "can you go back",
            "you said",
            "i was thinking",
            "i don't know",
            "kind of",
            " like ",
            " ah ",
            "follow up on you said",
            "confirm the plan to get that",
            "follow up on i ",
            "where we're going to be",
            "determine what to do with that",
            "review if it's making sense",
        )
    ):
        return False
    if re.search(r"\b(?:that|it|this|there)\s+(?:into|from|with|to|is|was)\b", lowered):
        return False
    if not identify_project_topics(bullet):
        return False
    action_or_status_verbs = (
        "confirm",
        "continue",
        "completed",
        "determine",
        "finalize",
        "prepare",
        "resolve",
        "review",
        "update",
        "validate",
    )
    return any(re.search(rf"\b{verb}\w*\b", lowered) for verb in action_or_status_verbs)


def _reason_for_item(
    text: str,
    category_matches: list[str],
    project_context: dict[str, Any] | None,
) -> str:
    lowered = text.casefold()
    anchors = [label for term, label in PROJECT_TERMS.items() if _phrase_present(lowered, term)]
    context_matches = _context_matches(text, project_context)
    reasons = []
    if anchors:
        reasons.append("project terms: " + ", ".join(dict.fromkeys(anchors)))
    if context_matches:
        reasons.append("matches uploaded project context")
    if category_matches:
        reasons.append("status signal: " + ", ".join(category_matches[:2]))
    return "; ".join(reasons) or "limited project context"


def _chunks_are_similar(first: str, second: str) -> bool:
    def normalized_tokens(value: str) -> set[str]:
        normalized = value.casefold()
        normalized = re.sub(r"\bprod(?:uction)?\b", "production", normalized)
        normalized = re.sub(r"\btesting?\b", "test", normalized)
        normalized = re.sub(r"\bmaster data\b|\bbalances?\b", "data", normalized)
        normalized = re.sub(r"[^a-z0-9\s]", " ", normalized)
        ignored = {
            "any",
            "confirm",
            "determine",
            "follow",
            "needed",
            "needs",
            "plan",
            "remaining",
            "review",
            "the",
            "to",
            "whether",
        }
        return {
            token
            for token in normalized.split()
            if len(token) >= 3 and token not in ignored
        }

    def concept_key(value: str) -> str:
        lowered = value.casefold()
        if (
            "test" in lowered
            and any(
                term in lowered for term in ("master data", "balance", "prod", "production")
            )
            and any(
                term in lowered
                for term in ("move", "moving", "refresh", "production", "test environment")
            )
        ):
            return "data-to-test"
        if "shell" in lowered and any(
            term in lowered for term in ("chart of accounts", "coa")
        ):
            return "bc-shell-chart-of-accounts"
        if "payroll" in lowered and "document" in lowered:
            return "payroll-integration-document"
        if "fixed assets" in lowered and "list" in lowered:
            return "fixed-assets-list"
        if "custom report" in lowered or "reporting needs" in lowered:
            return "custom-reporting-needs"
        return ""

    first_concept = concept_key(first)
    second_concept = concept_key(second)
    if first_concept and first_concept == second_concept:
        return True

    first_tokens = normalized_tokens(first)
    second_tokens = normalized_tokens(second)
    if not first_tokens or not second_tokens:
        return False
    overlap = len(first_tokens & second_tokens)
    return (
        overlap / min(len(first_tokens), len(second_tokens)) >= 0.75
        or overlap / len(first_tokens | second_tokens) >= 0.65
    )


def categorize_chunks(
    chunks: list[str], project_context: dict[str, Any] | None = None
) -> tuple[dict[str, list[dict[str, Any]]], list[dict[str, Any]]]:
    categories: dict[str, list[dict[str, Any]]] = {
        category: [] for category in CATEGORY_KEYWORDS
    }
    low_confidence = []

    topic_chunks: dict[str, list[tuple[int, str]]] = {}
    for chunk in chunks:
        relevance = score_project_relevance(chunk, project_context)
        for topic in identify_project_topics(chunk):
            topic_chunks.setdefault(topic, []).append((relevance, chunk))

    decision_re = re.compile(
        r"\b(?:decision|decide|need(?:s)? to confirm|need(?:s)? confirmation|confirm whether|approval|"
        r"sign[ -]?off|determine whether|up in the air|uncertain)\b",
        re.IGNORECASE,
    )
    tracking_re = re.compile(
        r"\b(?:action item|follow up|follow-up|assigned to|owner|ownership|due date|responsible for|"
        r"needs tracking)\b",
        re.IGNORECASE,
    )
    blocker_re = re.compile(
        r"\b(?:blocked|delay(?:ed)?|risk|dependency|waiting on|missing data|"
        r"cannot move forward|may impact|schedule impact)\b",
        re.IGNORECASE,
    )
    completion_re = re.compile(
        r"\b(?:completed|finalized|confirmed|approved|resolved|delivered|validated)\b",
        re.IGNORECASE,
    )

    for topic, scored_chunks in topic_chunks.items():
        minimum_topic_count = min(
            len(identify_project_topics(chunk)) for _, chunk in scored_chunks
        )
        focused_chunks = [
            pair
            for pair in scored_chunks
            if len(identify_project_topics(pair[1])) == minimum_topic_count
        ]
        focused_chunks.sort(key=lambda pair: (len(pair[1]), -pair[0]))
        supporting_chunks = []
        for _, chunk in focused_chunks:
            if not any(_chunks_are_similar(chunk, existing) for existing in supporting_chunks):
                supporting_chunks.append(chunk)
            if len(supporting_chunks) == 3:
                break
        combined_raw = collapse_whitespace(" ".join(supporting_chunks))
        strongest_relevance = max(
            relevance
            for relevance, chunk in focused_chunks
            if chunk in supporting_chunks
        )

        if topic in {"shell_company_chart_of_accounts", "custom_reports"}:
            category = "decisions_needed"
        elif decision_re.search(combined_raw):
            category = "decisions_needed"
        elif tracking_re.search(combined_raw):
            category = "possible_raid_items_for_review"
        elif blocker_re.search(combined_raw):
            category = "blockers_or_concerns"
        elif completion_re.search(combined_raw):
            category = "accomplishments"
        else:
            category = "upcoming_focus_areas"

        matches = keyword_matches(combined_raw, CATEGORY_KEYWORDS[category])
        category_signal = (
            2
            if category
            in {
                "decisions_needed",
                "possible_raid_items_for_review",
                "blockers_or_concerns",
            }
            else max(1, min(2, len(matches)))
        )
        confidence = strongest_relevance + category_signal
        bullet = rewrite_topic_as_project_bullet(
            topic, supporting_chunks, category
        )
        item = {
            "text": bullet,
            "raw_text": combined_raw,
            "topic": topic,
            "category": category,
            "confidence_score": confidence,
            "reason": _reason_for_item(combined_raw, matches, project_context),
        }
        threshold = CATEGORY_THRESHOLDS[category]
        if (
            confidence < threshold
            or not has_project_anchor(combined_raw, project_context)
            or not is_high_quality_project_bullet(bullet)
        ):
            low_confidence.append(item)
            continue
        categories[category].append(item)

    for category, items in categories.items():
        items.sort(key=lambda item: (-item["confidence_score"], len(item["text"])))
        categories[category] = items[:MAX_ITEMS]
    low_confidence.sort(key=lambda item: (-item["confidence_score"], len(item["text"])))
    return categories, low_confidence[:MAX_LOW_CONFIDENCE_ITEMS]


def _topics_from_items(
    items: list[dict[str, Any]], project_context: dict[str, Any] | None
) -> list[str]:
    topics = []
    for item in items:
        lowered = item["text"].casefold()
        for term, label in PROJECT_TERMS.items():
            if _phrase_present(lowered, term) and label not in topics:
                topics.append(label)
        for context_value in _context_matches(item["text"], project_context):
            if len(context_value) <= 65 and context_value not in topics:
                topics.append(context_value)
    return topics[:3]


def _join_topics(topics: list[str], fallback: str) -> str:
    if not topics:
        return fallback
    if len(topics) == 1:
        return topics[0]
    if len(topics) == 2:
        return f"{topics[0]} and {topics[1]}"
    return f"{topics[0]}, {topics[1]}, and {topics[2]}"


def _bullet_to_narrative_clause(value: str, limit: int = 105) -> str:
    text = collapse_whitespace(value).rstrip(".!?")
    replacements = (
        (r"(?i)^Confirm\b", "confirming"),
        (r"(?i)^Follow up on\b", "following up on"),
        (r"(?i)^Follow up to\b", "following up to"),
        (r"(?i)^Determine\b", "determining"),
        (r"(?i)^Review\b", "reviewing"),
        (r"(?i)^Finalize\b", "finalizing"),
        (r"(?i)^Validate\b", "validating"),
        (r"(?i)^Update\b", "updating"),
        (r"(?i)^Prepare\b", "preparing"),
        (r"(?i)^Complete\b", "completing"),
        (r"(?i)^Resolve\b", "resolving"),
        (r"(?i)^Address\b", "addressing"),
    )
    for pattern, replacement in replacements:
        if re.search(pattern, text):
            text = re.sub(pattern, replacement, text, count=1)
            break
    if text and text[0].isupper():
        text = text[0].lower() + text[1:]
    if len(text) > limit:
        text = text[:limit].rsplit(" ", 1)[0].rstrip(" ,;:-")
    return text


def build_narrative(
    categories: dict[str, list[dict[str, Any]]],
    project_context: dict[str, Any] | None = None,
) -> str:
    narrative = []
    accomplishments = categories["accomplishments"]
    upcoming = categories["upcoming_focus_areas"]
    blockers = categories["blockers_or_concerns"]
    decisions = categories["decisions_needed"]
    raid_items = categories["possible_raid_items_for_review"]

    all_primary_items = accomplishments + upcoming + blockers + decisions
    all_topics = _join_topics(
        _topics_from_items(all_primary_items, project_context),
        "current project priorities",
    )

    if accomplishments:
        narrative.append(f"The team made progress this period across {all_topics}.")
    else:
        narrative.append(f"The team is working through {all_topics}.")

    follow_up_items = upcoming + decisions + raid_items
    follow_up_clauses = []
    for item in follow_up_items:
        clause = _bullet_to_narrative_clause(item["text"])
        if clause and not any(
            _chunks_are_similar(clause, existing) for existing in follow_up_clauses
        ):
            follow_up_clauses.append(clause)
        if len(follow_up_clauses) == 2:
            break
    if follow_up_clauses:
        if len(follow_up_clauses) == 1:
            follow_up_text = follow_up_clauses[0]
        else:
            follow_up_text = f"{follow_up_clauses[0]} and {follow_up_clauses[1]}"
        narrative.append(f"Key follow-ups include {follow_up_text}.")

    if blockers and len(narrative) < MAX_NARRATIVE_SENTENCES:
        topics = _join_topics(
            _topics_from_items(blockers, project_context),
            "open project items",
        )
        narrative.append(
            f"Project-impacting concerns related to {topics} require follow-up."
        )
    elif decisions and not follow_up_clauses and len(narrative) < MAX_NARRATIVE_SENTENCES:
        topics = _join_topics(
            _topics_from_items(decisions, project_context),
            "open project decisions",
        )
        narrative.append(f"Open decisions related to {topics} require confirmation.")
    return " ".join(narrative[:MAX_NARRATIVE_SENTENCES])


def parse_transcript_text(
    text: str,
    is_vtt: bool = False,
    project_context: dict[str, Any] | None = None,
) -> dict[str, object]:
    project_context = dict(project_context or {})
    cleaned_text = clean_transcript_text(text, is_vtt=is_vtt)
    sentences = split_sentences(cleaned_text)
    chunks = build_context_chunks(sentences)
    categories, low_confidence = categorize_chunks(chunks, project_context)
    warnings = []

    narrative_groups = sum(
        bool(categories[name])
        for name in (
            "accomplishments",
            "upcoming_focus_areas",
            "blockers_or_concerns",
            "decisions_needed",
        )
    )
    narrative = build_narrative(categories, project_context)
    narrative_sentence_count = len(
        [part for part in re.split(r"(?<=[.!?])\s+", narrative) if part]
    )
    if narrative_groups < 2 or narrative_sentence_count < 2:
        narrative = ""
        warnings.append(INSUFFICIENT_CONTENT_WARNING)
    if sentences and not any(categories.values()):
        warnings.append(
            "Project-related transcript content was not confident enough to show "
            "in the main review categories."
        )

    return {
        "suggested_project_status_narrative": narrative,
        **categories,
        "low_confidence_items": low_confidence,
        "transcript_warnings": warnings,
    }


def parse_meeting_transcript(
    file_bytes: bytes,
    file_name: str,
    project_context: dict[str, Any] | None = None,
) -> dict[str, object]:
    project_context = dict(project_context or {})
    text = read_transcript(file_bytes, file_name)
    result = parse_transcript_text(
        text,
        is_vtt=file_name.casefold().endswith(".vtt"),
        project_context=project_context,
    )
    for category in CATEGORY_KEYWORDS:
        result[category] = [
            {
                **item,
                "source_file": file_name,
                "source_files": [file_name],
            }
            for item in result[category]
        ]
    result["low_confidence_items"] = [
        {
            **item,
            "source_file": file_name,
            "source_files": [file_name],
        }
        for item in result["low_confidence_items"]
    ]
    result["transcript_warnings"] = [
        f"{file_name}: {warning}" for warning in result["transcript_warnings"]
    ]
    return result


def _merge_sourced_items(
    results: list[dict[str, object]],
    category: str,
    limit: int,
) -> list[dict[str, Any]]:
    candidates = [
        dict(item)
        for result in results
        for item in result.get(category, [])
    ]
    candidates.sort(
        key=lambda item: (-int(item.get("confidence_score", 0)), len(item.get("text", "")))
    )
    merged: list[dict[str, Any]] = []
    for candidate in candidates:
        duplicate = next(
            (
                item
                for item in merged
                if _chunks_are_similar(candidate.get("text", ""), item.get("text", ""))
            ),
            None,
        )
        if duplicate is not None:
            sources = list(duplicate.get("source_files", []))
            for source in candidate.get("source_files", []):
                if source and source not in sources:
                    sources.append(source)
            duplicate["source_files"] = sources
            duplicate["source_file"] = ", ".join(sources)
            continue
        sources = list(candidate.get("source_files", []))
        if not sources and candidate.get("source_file"):
            sources = [candidate["source_file"]]
        candidate["source_files"] = sources
        candidate["source_file"] = ", ".join(sources)
        merged.append(candidate)
    return merged[:limit]


def _deduplicate_across_categories(
    categories: dict[str, list[dict[str, Any]]],
) -> dict[str, list[dict[str, Any]]]:
    priority = {
        "decisions_needed": 0,
        "possible_raid_items_for_review": 1,
        "upcoming_focus_areas": 2,
        "blockers_or_concerns": 3,
        "accomplishments": 4,
    }
    candidates = [
        item
        for category, items in categories.items()
        for item in items
        if item.get("text")
    ]
    candidates.sort(
        key=lambda item: (
            priority.get(item.get("category", ""), 99),
            -int(item.get("confidence_score", 0)),
            len(item.get("text", "")),
        )
    )
    selected: list[dict[str, Any]] = []
    for candidate in candidates:
        duplicate = next(
            (
                item
                for item in selected
                if (
                    candidate.get("topic")
                    and candidate.get("topic") == item.get("topic")
                )
                or _chunks_are_similar(
                    candidate.get("text", ""), item.get("text", "")
                )
            ),
            None,
        )
        if duplicate is not None:
            sources = list(duplicate.get("source_files", []))
            for source in candidate.get("source_files", []):
                if source and source not in sources:
                    sources.append(source)
            duplicate["source_files"] = sources
            duplicate["source_file"] = ", ".join(sources)
            continue
        selected.append(candidate)

    resolved = {category: [] for category in CATEGORY_KEYWORDS}
    for item in selected:
        category = item.get("category", "")
        if category in resolved and len(resolved[category]) < MAX_ITEMS:
            resolved[category].append(item)
    return resolved


def combine_transcript_results(
    results: list[dict[str, object]],
    project_context: dict[str, Any] | None = None,
    additional_warnings: list[str] | None = None,
) -> dict[str, object]:
    project_context = dict(project_context or {})
    categories = {
        category: _merge_sourced_items(results, category, MAX_ITEMS)
        for category in CATEGORY_KEYWORDS
    }
    categories = _deduplicate_across_categories(categories)
    low_confidence = _merge_sourced_items(
        results, "low_confidence_items", MAX_LOW_CONFIDENCE_ITEMS
    )
    warnings = list(additional_warnings or [])
    for result in results:
        for warning in result.get("transcript_warnings", []):
            if warning not in warnings:
                warnings.append(warning)

    narrative_groups = sum(
        bool(categories[name])
        for name in (
            "accomplishments",
            "upcoming_focus_areas",
            "blockers_or_concerns",
            "decisions_needed",
        )
    )
    narrative = build_narrative(categories, project_context)
    narrative_sentence_count = len(
        [part for part in re.split(r"(?<=[.!?])\s+", narrative) if part]
    )
    if narrative_groups < 2 or narrative_sentence_count < 2:
        narrative = ""
        if INSUFFICIENT_CONTENT_WARNING not in warnings:
            warnings.append(INSUFFICIENT_CONTENT_WARNING)

    return {
        "suggested_project_status_narrative": narrative,
        **categories,
        "low_confidence_items": low_confidence,
        "transcript_warnings": warnings,
    }


def parse_meeting_transcripts(
    transcripts: Iterable[tuple[bytes, str]],
    project_context: dict[str, Any] | None = None,
) -> dict[str, object]:
    project_context = dict(project_context or {})
    results = []
    warnings = []
    for file_bytes, file_name in transcripts:
        try:
            results.append(
                parse_meeting_transcript(
                    file_bytes,
                    file_name,
                    project_context=project_context,
                )
            )
        except MeetingTranscriptError as error:
            warnings.append(f"{file_name}: could not be parsed. {error}")
    return combine_transcript_results(
        results,
        project_context=project_context,
        additional_warnings=warnings,
    )
