from __future__ import annotations

from datetime import date, datetime
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError


def current_report_date() -> date:
    """Return today's Eastern date, with a local-date fallback on Windows."""
    try:
        return datetime.now(ZoneInfo("America/New_York")).date()
    except ZoneInfoNotFoundError:
        return date.today()
