"""Report request routing and simulator source-link presentation."""

from __future__ import annotations

import re
import json


def parse_report_tool_result(texts: list[str]) -> dict:
    if not texts:
        raise ValueError("Report tool returned no text content.")
    # The SDK may append a structured-content mirror after the text result.
    report = json.loads(texts[0])
    if not isinstance(report, dict):
        raise ValueError("Report tool returned an invalid result.")
    return report


def report_type_for(question: str) -> str | None:
    text = question.lower()
    if not re.search(r"\b(report|spreadsheet|excel|workbook)\b", text):
        return None
    if "joint commission" in text or re.search(r"\bjc\b", text) or "readiness" in text:
        return "jc_readiness"
    if "vendor" in text and ("weekly" in text or "commitment" in text):
        return "weekly_vendor_status"
    return None


def is_download_followup(question: str) -> bool:
    text = question.lower().strip(" .?!")
    return bool(
        re.search(r"\b(download|link|excel|spreadsheet)\b", text)
        and re.fullmatch(
            r"(?:yes|please|give|me|the|a|an|download|link|excel|spreadsheet|file|"
            r"report|it|that|this|my|where|is|can|you|share|send|show|again|to|"
            r"get|for|weekly|vendor|status|and|commitment|\s|,|'s)+",
            text,
        )
    )


def localize_source_links(answer: str, index: dict) -> str:
    replacements = {}
    for cid, (kind, record) in index.items():
        url = record.get("url")
        if isinstance(url, str) and url:
            replacements[url] = f"/citation/{kind}/{cid}"
    if not replacements:
        return answer
    pattern = "|".join(re.escape(url) for url in sorted(replacements, key=len, reverse=True))
    return re.sub(pattern, lambda match: replacements[match.group(0)], answer)
