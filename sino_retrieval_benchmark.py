#!/usr/bin/env python3
"""
Sino Voicebot FAQ retrieval benchmark tool.

Calls the setsail multi-graph "text_retrieval_test" endpoint for a batch of
test queries, matches the retrieved FAQ_ID(s) against expected ground-truth
answers, and writes a scored results workbook (per-query detail + summary
metrics, overall and per FAQ category/path).

Usage:
    python3 sino_retrieval_benchmark.py --testcases testcases.xlsx --output results.xlsx

    # Inspect the raw API response shape before running a full batch:
    python3 sino_retrieval_benchmark.py --probe-query "How do I apply for a mortgage?" --probe-path FAQ_Mortgage

See README_benchmark.md for full documentation.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import requests

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Missing dependency 'openpyxl'. Install with: pip install openpyxl", file=sys.stderr)
    raise

DEFAULT_API_URL = "https://dev.setsailapi.com/nlp/llmGraph/graph/multi/run"
DEFAULT_PROJECT_ID = "public"
DEFAULT_CMS_PROJECT_ID = "sinoexvoicebot-52ynuou"
DEFAULT_GRAPH_NAME = "text_retrieval_test"
ID_KEY_NAME = "FAQ_ID"
DISPLAY_COLUMNS = ["FAQ_ID", "FAQ"]
RETRIEVED_TEXT_KEY_NAME = "FAQ_answer_listening_friendly"
VALID_PATHS = ("FAQ_Mortgage", "FAQ_Leasing")

TESTCASE_QUERY_COLS = ["Query", "Question", "query", "question", "Testing Query"]
TESTCASE_EXPECTED_COLS = [
    "Expected_FAQ_ID", "Expected FAQ_ID", "FAQ_ID", "expected_faq_id",
    "FAQ Reference No.", "FAQ Reference No",
]
TESTCASE_PATH_COLS = ["Path", "Category", "path"]
TESTCASE_LANG_COLS = ["Language", "Lang", "query_lang", "language"]
TESTCASE_NOTES_COLS = ["Notes", "notes"]

# Lets a "Category" column (e.g. Sino's own benchmark template, which uses
# "Leasing"/"Mortgage" rather than the API's FAQ_Leasing/FAQ_Mortgage path
# names) resolve straight to a valid Path. Extend this if new categories
# are added on the CMS side.
CATEGORY_TO_PATH = {
    "leasing": "FAQ_Leasing",
    "mortgage": "FAQ_Mortgage",
}

# When a workbook has no single "Test Cases" sheet, each sheet is treated as
# its own set of test cases and its query_lang is inferred from a language
# suffix in the sheet name (e.g. "Benchmark Template cn" -> sc). Add
# entries here if other sheet-name suffixes/languages are used.
SHEET_NAME_LANG_SUFFIXES = {
    "zh": "zh",
    "en": "en",
    "cn": "sc",
}

# Rows whose query starts with an "EG:" / "EG -" marker are template
# instructional examples, not real test cases, and are skipped.
EG_PREFIX_RE = re.compile(r"^\s*eg\s*[:\-]", re.IGNORECASE)


# --------------------------------------------------------------------------- #
# Data model
# --------------------------------------------------------------------------- #

@dataclass
class TestCase:
    row_num: int
    query: str
    expected_ids: list[str]
    path: str
    lang: str
    notes: str = ""
    sheet: str = ""


@dataclass
class TestResult:
    case: TestCase
    retrieved: list[dict[str, Any]] = field(default_factory=list)
    rank: Optional[int] = None
    latency_ms: Optional[float] = None
    http_status: Optional[int] = None
    error: str = ""
    raw_response: Any = None
    exact_match: Optional[bool] = None
    confidence_score: Optional[float] = None
    direct_answer: Optional[bool] = None

    @property
    def top1_correct(self) -> bool:
        return self.rank == 1

    @property
    def reciprocal_rank(self) -> float:
        return 1.0 / self.rank if self.rank else 0.0

    def hit_at(self, k: int) -> bool:
        return self.rank is not None and self.rank <= k

    def retrieved_id(self, position: int) -> str:
        if position < len(self.retrieved):
            return str(self.retrieved[position].get(ID_KEY_NAME, ""))
        return ""

    def retrieved_response(self, position: int) -> str:
        if position < len(self.retrieved):
            return str(self.retrieved[position].get(RETRIEVED_TEXT_KEY_NAME, ""))
        return ""

    @property
    def matched_id(self) -> str:
        """The FAQ the API actually returned (its top pick), regardless of whether
        it happens to be correct — always position 0, never tied to `rank`."""
        return self.retrieved_id(0)

    @property
    def matched_response(self) -> str:
        return self.retrieved_response(0)

    @property
    def other_suggestions(self) -> list[dict[str, str]]:
        """Every other candidate the API returned, in rank order, excluding the
        top pick already shown as matched_id/matched_response."""
        return [
            {"faq_id": str(doc.get(ID_KEY_NAME, "")), "response": str(doc.get(RETRIEVED_TEXT_KEY_NAME, ""))}
            for doc in self.retrieved[1:]
        ]


# --------------------------------------------------------------------------- #
# Response parsing
#
# Confirmed against a live call: the API does NOT return a ranked list of
# retrieved documents. It resolves the query to a single matched FAQ and
# hands back:
#
#   response_json["entity"][id_key_name]  -> the matched FAQ_ID
#   response_json["bot_responses"][0]     -> {"answer": <response text>,
#                                              "confidence_score": <float>, ...}
#   response_json["similar_ids"]          -> other candidate FAQ_IDs, matched
#                                             one usually first; no answer text
#                                             is returned for these
#   response_json["exact_match"]          -> whether the API considers this an
#                                             exact match
#
# extract_match_data() normalizes this into the same list[{id_key: ...,
# response_key: ...}] shape used everywhere else (matched item first, answer
# text only available for it), so scoring/rank/recall@k keep working
# unchanged. A generic list-of-documents search is kept as a fallback for
# other graph configs that really do return a document list, and
# --response-list-path overrides both for anything else. --save-raw keeps
# every raw response for offline inspection.
# --------------------------------------------------------------------------- #

LIKELY_LIST_KEYS = [
    "documents", "retrieved_documents", "retrieval", "retrieved",
    "results", "data", "output", "response", "hits", "items",
]


def _get_by_dotted_path(obj: Any, dotted_path: str) -> Any:
    cur = obj
    for part in dotted_path.split("."):
        if isinstance(cur, dict):
            cur = cur.get(part)
        elif isinstance(cur, list) and part.isdigit():
            idx = int(part)
            cur = cur[idx] if idx < len(cur) else None
        else:
            return None
    return cur


def _find_id_list(obj: Any, id_key: str, _depth: int = 0) -> Optional[list[dict]]:
    """Recursively search for the first list of dicts that all contain id_key."""
    if _depth > 8:
        return None
    if isinstance(obj, list) and obj and all(isinstance(x, dict) for x in obj):
        if all(id_key in x for x in obj):
            return obj
    if isinstance(obj, dict):
        for key in LIKELY_LIST_KEYS:
            if key in obj:
                found = _find_id_list(obj[key], id_key, _depth + 1)
                if found is not None:
                    return found
        for v in obj.values():
            found = _find_id_list(v, id_key, _depth + 1)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for v in obj:
            found = _find_id_list(v, id_key, _depth + 1)
            if found is not None:
                return found
    return None


def _extract_bot_response(response_json: dict) -> tuple[str, Optional[float]]:
    """Pull the matched FAQ's answer text + confidence score out of the response."""
    bot_responses = response_json.get("bot_responses")
    if isinstance(bot_responses, list) and bot_responses and isinstance(bot_responses[0], dict):
        first = bot_responses[0]
        text = first.get("answer") or first.get("model_output") or first.get("raw_ans") or ""
        return str(text), first.get("confidence_score")
    # Redundant fallbacks seen carrying the same text in the same response.
    actions = response_json.get("actions")
    if isinstance(actions, list) and actions and isinstance(actions[0], dict):
        data = actions[0].get("data")
        message = data.get("message") if isinstance(data, dict) else None
        if isinstance(message, dict) and message.get("content"):
            return str(message["content"]), None
    return "", None


def has_direct_match(response_json: Any, id_key: str = ID_KEY_NAME) -> Optional[bool]:
    """Did the API resolve the query to a FAQ, or only offer similar_ids?

    A populated `entity[id_key]` means the API committed to an answer. An
    absent or empty one means it fell back to suggestions — the caller would
    have been offered choices rather than given an answer. That distinction is
    invisible in Confidence_Score, which came back as 1 on 894 of 905 rows in a
    real run including rows with no matched entity at all.

    Returns None when the response shape is not the entity/similar_ids one.
    """
    if not isinstance(response_json, dict):
        return None
    if "entity" not in response_json and "similar_ids" not in response_json:
        return None
    entity = response_json.get("entity")
    value = entity.get(id_key) if isinstance(entity, dict) else None
    return bool(value is not None and str(value).strip())


def extract_match_data(
    response_json: Any,
    id_key: str = ID_KEY_NAME,
    response_key: str = RETRIEVED_TEXT_KEY_NAME,
    override_path: Optional[str] = None,
) -> tuple[list[dict], Optional[bool], Optional[float]]:
    """Returns (retrieved, exact_match, confidence_score). See module note above."""
    exact_match = response_json.get("exact_match") if isinstance(response_json, dict) else None
    matched_response, confidence = (
        _extract_bot_response(response_json) if isinstance(response_json, dict) else ("", None)
    )

    if override_path:
        candidate = _get_by_dotted_path(response_json, override_path)
        if isinstance(candidate, list):
            return candidate, exact_match, confidence
        raise ValueError(
            f"--response-list-path {override_path!r} did not resolve to a list "
            f"(got {type(candidate).__name__}). Use --probe-query to inspect the real shape."
        )

    if isinstance(response_json, list):
        return response_json, exact_match, confidence
    if not isinstance(response_json, dict):
        return [], exact_match, confidence

    entity = response_json.get("entity")
    matched_id = entity.get(id_key) if isinstance(entity, dict) else None
    similar_ids = response_json.get("similar_ids")

    ordered_ids: list[str] = []
    seen: set[str] = set()

    def _add(value: Any) -> None:
        """Append a candidate, skipping blanks and format-duplicates.

        Blank guard: when the API has no confident match it can return an
        `entity` whose FAQ_ID is present but empty. Appending that put a
        phantom candidate at position 0, which pushed every real candidate
        down one rank and cost those rows their Top-1 (seen on 11 rows of a
        real 905-row run, where the correct FAQ sat at Suggested_2 with
        Suggested_1 blank).

        Dedupe is on the normalized id, so 'FAQ_001' arriving in similar_ids
        after entity returned '1' is not counted as a second candidate.
        """
        text = "" if value is None else str(value).strip()
        if not text:
            return
        key = _normalize_faq_id(text)
        if key in seen:
            return
        seen.add(key)
        ordered_ids.append(text)

    _add(matched_id)
    if isinstance(similar_ids, list):
        for sid in similar_ids:
            _add(sid)

    if ordered_ids:
        retrieved = [
            {id_key: fid, response_key: matched_response if i == 0 else ""}
            for i, fid in enumerate(ordered_ids)
        ]
        return retrieved, exact_match, confidence

    # Fallback for other graph configs that return an actual document list.
    found = _find_id_list(response_json, id_key)
    return (found or []), exact_match, confidence


# --------------------------------------------------------------------------- #
# API call
# --------------------------------------------------------------------------- #

def build_request_body(
    query: str,
    path: str,
    lang: str,
    project_id: str,
    cms_project_id: str,
    graph_name: str,
) -> dict:
    return {
        "project_id": project_id,
        "cms_project_id": cms_project_id,
        "single_graph_name": graph_name,
        "path": path,
        "id_key_name": ID_KEY_NAME,
        "display_columns": DISPLAY_COLUMNS,
        "retrieved_text_key_name": RETRIEVED_TEXT_KEY_NAME,
        "query_lang": lang,
        "messages": [{"type": "human", "content": query}],
        "stream": False,
        "slim_response": True,
    }


def call_api(
    session: requests.Session,
    api_url: str,
    body: dict,
    headers: dict,
    timeout: float,
    retries: int,
    backoff_base: float = 1.5,
) -> tuple[Optional[dict], Optional[int], float, str]:
    """Returns (response_json, http_status, latency_ms, error_message)."""
    last_error = ""
    for attempt in range(retries + 1):
        start = time.monotonic()
        try:
            resp = session.post(api_url, json=body, headers=headers, timeout=timeout)
            latency_ms = (time.monotonic() - start) * 1000
            if resp.status_code >= 500 and attempt < retries:
                last_error = f"HTTP {resp.status_code} (retrying)"
                time.sleep(backoff_base ** attempt)
                continue
            try:
                data = resp.json()
            except ValueError:
                return None, resp.status_code, latency_ms, f"Non-JSON response: {resp.text[:300]!r}"
            if resp.status_code >= 400:
                return data, resp.status_code, latency_ms, f"HTTP {resp.status_code}: {resp.text[:300]}"
            return data, resp.status_code, latency_ms, ""
        except requests.RequestException as exc:
            latency_ms = (time.monotonic() - start) * 1000
            last_error = str(exc)
            if attempt < retries:
                time.sleep(backoff_base ** attempt)
                continue
            return None, None, latency_ms, last_error
    return None, None, 0.0, last_error


# --------------------------------------------------------------------------- #
# Test case loading
# --------------------------------------------------------------------------- #

def _pick_column(header_row: list[str], candidates: list[str]) -> Optional[int]:
    normalized = [h.strip() if isinstance(h, str) else h for h in header_row]
    for cand in candidates:
        if cand in normalized:
            return normalized.index(cand)
    return None


def _infer_lang_from_sheet_name(sheet_name: str, default_lang: str) -> str:
    """E.g. 'Benchmark Template cn' -> sc, via SHEET_NAME_LANG_SUFFIXES."""
    tokens = re.findall(r"[A-Za-z]+", sheet_name.lower())
    for token in reversed(tokens):
        if token in SHEET_NAME_LANG_SUFFIXES:
            return SHEET_NAME_LANG_SUFFIXES[token]
    return default_lang


def _resolve_path(raw_value: Any, default_path: Optional[str]) -> Optional[str]:
    if raw_value is None or str(raw_value).strip() == "":
        return default_path
    key = str(raw_value).strip()
    if key in VALID_PATHS:
        return key
    return CATEGORY_TO_PATH.get(key.lower().replace(" ", "_"), key)


def _load_sheet(
    wb, sheet_name: str, source_path: Path, default_path: Optional[str], default_lang: str
) -> list[TestCase]:
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = list(rows[0])

    query_idx = _pick_column(header, TESTCASE_QUERY_COLS)
    expected_idx = _pick_column(header, TESTCASE_EXPECTED_COLS)
    path_idx = _pick_column(header, TESTCASE_PATH_COLS)
    lang_idx = _pick_column(header, TESTCASE_LANG_COLS)
    notes_idx = _pick_column(header, TESTCASE_NOTES_COLS)

    if query_idx is None or expected_idx is None:
        return []

    cases: list[TestCase] = []
    for row_num, row in enumerate(rows[1:], start=2):
        query = row[query_idx] if query_idx < len(row) else None
        if query is None or str(query).strip() == "":
            continue
        query_str = str(query).strip()
        if EG_PREFIX_RE.match(query_str):
            continue

        expected_raw = row[expected_idx] if expected_idx < len(row) else None
        expected_ids = [e.strip() for e in str(expected_raw or "").split(",") if e.strip()]
        if not expected_ids:
            raise ValueError(
                f"Row {row_num} in sheet {sheet_name!r} of {source_path} has a query but no "
                f"expected FAQ ID."
            )

        raw_path = row[path_idx] if path_idx is not None and path_idx < len(row) else None
        row_path = _resolve_path(raw_path, default_path)
        if not row_path:
            raise ValueError(
                f"Row {row_num} in sheet {sheet_name!r} of {source_path} has no Path/Category "
                f"and no --default-path was given."
            )
        if row_path not in VALID_PATHS:
            raise ValueError(
                f"Row {row_num} in sheet {sheet_name!r} of {source_path} has Path/Category="
                f"{raw_path!r}, which doesn't map to one of {VALID_PATHS}. Add a mapping in "
                f"CATEGORY_TO_PATH or fix the sheet."
            )

        row_lang = None
        if lang_idx is not None and lang_idx < len(row):
            row_lang = row[lang_idx]
        row_lang = str(row_lang).strip() if row_lang else default_lang

        row_notes = ""
        if notes_idx is not None and notes_idx < len(row):
            row_notes = str(row[notes_idx] or "")

        cases.append(
            TestCase(
                row_num=row_num,
                query=query_str,
                expected_ids=expected_ids,
                path=row_path,
                lang=row_lang,
                notes=row_notes,
                sheet=sheet_name,
            )
        )
    return cases


def load_testcases(path: Path, default_path: Optional[str], default_lang: str) -> list[TestCase]:
    """
    Supports two workbook layouts:

    1. A single "Test Cases" sheet (see sino_benchmark_testcases_TEMPLATE.xlsx):
       Query / Expected_FAQ_ID / Path / Language / Notes columns.

    2. Sino's own multi-language benchmark template: one sheet per language
       (e.g. "Benchmark Template zh/en/cn") with No. / Testing Query /
       Category / FAQ Reference No. columns. query_lang is inferred per
       sheet from its name (see SHEET_NAME_LANG_SUFFIXES), Category is
       mapped to a Path via CATEGORY_TO_PATH, and "EG:"-prefixed example
       rows are skipped automatically. All sheets with a recognizable
       header are combined into one test run.
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    if "Test Cases" in wb.sheetnames:
        return _load_sheet(wb, "Test Cases", path, default_path, default_lang)

    cases: list[TestCase] = []
    for sheet_name in wb.sheetnames:
        sheet_lang = _infer_lang_from_sheet_name(sheet_name, default_lang)
        cases.extend(_load_sheet(wb, sheet_name, path, default_path, sheet_lang))

    if not cases:
        raise ValueError(
            f"Could not find a usable test-case sheet in {path}. Expected either a sheet named "
            f"'Test Cases', or one or more sheets with a query column (one of "
            f"{TESTCASE_QUERY_COLS}) and an expected-answer column (one of "
            f"{TESTCASE_EXPECTED_COLS}). Sheets found: {wb.sheetnames}"
        )
    return cases


# --------------------------------------------------------------------------- #
# Scoring
# --------------------------------------------------------------------------- #

_ID_PREFIX_RE = re.compile(r"^(?:faq|q)[\s\-_]*0*(\d+)$", re.IGNORECASE)


def _normalize_faq_id(value: Any) -> str:
    """
    'Q2', '2', 'FAQ_002' and 'faq-2' must all compare equal. The CMS's real
    FAQ_ID values are usually bare numbers (confirmed against a live call, e.g.
    entity.FAQ_ID = "2"), Sino's benchmark template labels them "FAQ Reference
    No." as "Q2", and the API has been observed returning a zero-padded
    "FAQ_001" form on some rows (real 905-row run, row 18: similar_ids came back
    as ["FAQ_001", "FAQ_005"] while every other row returned bare integers).
    Strip an optional FAQ/Q prefix, any separator and leading zeros; also
    normalize a bare zero-padded number so "001" == "1".
    """
    text = str(value).strip()
    m = _ID_PREFIX_RE.match(text)
    if m:
        return m.group(1)
    if text.isdigit():
        return str(int(text))
    return text.lower()


def load_faq_questions(path: Path) -> dict[str, dict[str, str]]:
    """Load the FAQ workbook as {lang: {normalized_faq_id: question}}.

    Used only to annotate and sanity-check the run — never to score. Expects
    the FAQ_Leasing.xlsx shape: one sheet per language, with FAQ_ID and FAQ
    columns. Sheet names map to the API's query_lang codes via
    SHEET_NAME_LANG_SUFFIXES (zh -> zh, en -> en, cn -> sc).
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out: dict[str, dict[str, str]] = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [str(c).strip() if c is not None else "" for c in rows[0]]
        try:
            id_idx = header.index(ID_KEY_NAME)
            q_idx = header.index("FAQ")
        except ValueError:
            continue
        lang = SHEET_NAME_LANG_SUFFIXES.get(sheet_name.strip().lower(), sheet_name.strip().lower())
        table = out.setdefault(lang, {})
        for row in rows[1:]:
            if id_idx >= len(row) or q_idx >= len(row) or row[id_idx] is None:
                continue
            table[_normalize_faq_id(row[id_idx])] = str(row[q_idx] or "")
    wb.close()
    return out


def faq_question(faq_map: Optional[dict], lang: str, faq_id: Any) -> str:
    if not faq_map or faq_id in (None, ""):
        return ""
    key = _normalize_faq_id(faq_id)
    table = faq_map.get(lang) or {}
    if key in table:
        return table[key]
    for other in faq_map.values():          # fall back to any language
        if key in other:
            return other[key]
    return "<not in FAQ file>"


def check_expected_ids(cases: list[TestCase], faq_map: dict) -> list[str]:
    """Report expected FAQ IDs that do not exist in the FAQ file.

    Catches a label pointing at an ID the FAQ set no longer has. It cannot
    catch a label that drifted onto a DIFFERENT still-valid ID, so also read
    the Expected_FAQ_Question column in the output: if it has nothing to do
    with the query, that label has drifted even though check_expected_ids
    found nothing wrong with it.
    """
    known = set()
    for table in faq_map.values():
        known |= set(table)
    missing = []
    for case in cases:
        for eid in case.expected_ids:
            if _normalize_faq_id(eid) not in known:
                missing.append(f"row {case.row_num} ({case.sheet or case.path}): {eid}")
    return missing


def score_result(case: TestCase, retrieved: list[dict]) -> Optional[int]:
    expected_set = {_normalize_faq_id(e) for e in case.expected_ids}
    for i, doc in enumerate(retrieved):
        doc_id = _normalize_faq_id(doc.get(ID_KEY_NAME, ""))
        if doc_id in expected_set:
            return i + 1
    return None


# --------------------------------------------------------------------------- #
# Output workbook
# --------------------------------------------------------------------------- #

def write_results(
    results: list[TestResult],
    output_path: Path,
    top_ks: list[int],
    max_retrieved_cols: int = 5,
    save_raw: bool = False,
    faq_map: Optional[dict] = None,
):
    wb = openpyxl.Workbook()

    # ---- Details sheet ----
    ws = wb.active
    ws.title = "Details"
    header = [
        "Sheet", "Row", "Path", "Language", "Query", "Expected_FAQ_ID",
        "Top1_Correct", "Rank_of_Expected", "Reciprocal_Rank",
    ]
    header += [f"Hit@{k}" for k in top_ks]
    header += ["Matched_FAQ_ID", "Matched_Response", "API_Exact_Match", "Confidence_Score",
               "Direct_Answer"]
    for i in range(1, max_retrieved_cols + 1):
        header += [f"Suggested_{i}_FAQ_ID", f"Suggested_{i}_Response"]
    header += ["Latency_ms", "HTTP_Status", "Error", "Notes"]
    if faq_map:
        header += ["Expected_FAQ_Question", "Matched_FAQ_Question"]
    if save_raw:
        header.append("Raw_Response")
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    for r in results:
        row = [
            r.case.sheet, r.case.row_num, r.case.path, r.case.lang, r.case.query,
            ", ".join(r.case.expected_ids),
            "Y" if r.top1_correct else "N",
            r.rank if r.rank else "",
            round(r.reciprocal_rank, 4),
        ]
        row += ["Y" if r.hit_at(k) else "N" for k in top_ks]
        row += [
            r.matched_id,
            r.matched_response,
            "" if r.exact_match is None else ("Y" if r.exact_match else "N"),
            round(r.confidence_score, 4) if r.confidence_score is not None else "",
            "" if r.direct_answer is None else ("Y" if r.direct_answer else "N"),
        ]
        for i in range(max_retrieved_cols):
            row += [r.retrieved_id(i), r.retrieved_response(i)]
        row += [
            round(r.latency_ms, 1) if r.latency_ms is not None else "",
            r.http_status if r.http_status is not None else "",
            r.error,
            r.case.notes,
        ]
        if faq_map:
            row += [
                " | ".join(faq_question(faq_map, r.case.lang, e) for e in r.case.expected_ids),
                faq_question(faq_map, r.case.lang, r.matched_id),
            ]
        if save_raw:
            row.append(json.dumps(r.raw_response, ensure_ascii=False)[:32000])
        ws.append(row)

    col_widths = (
        [20, 6, 14, 10, 45, 16]  # Sheet, Row, Path, Language, Query, Expected_FAQ_ID
        + [12, 16, 16]  # Top1_Correct, Rank_of_Expected, Reciprocal_Rank
        + [8] * len(top_ks)  # Hit@k...
        + [16, 45, 14, 14, 14]  # Matched_FAQ_ID, Matched_Response, API_Exact_Match, Confidence_Score, Direct_Answer
        + [16, 45] * max_retrieved_cols  # Suggested_i_FAQ_ID, Suggested_i_Response
        + [12, 12, 40, 25]  # Latency_ms, HTTP_Status, Error, Notes
        + ([45, 45] if faq_map else [])  # Expected_FAQ_Question, Matched_FAQ_Question
    )
    for i, col_width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = col_width

    # ---- Summary sheet ----
    ws2 = wb.create_sheet("Summary")
    ws2.append(["Metric", "Overall"] + list(VALID_PATHS))
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    def subset(path: Optional[str]):
        return [r for r in results if path is None or r.case.path == path]

    groups = [None] + list(VALID_PATHS)

    def metric_row(label, fn):
        ws2.append([label] + [fn(subset(g)) for g in groups])

    def pct(n, d):
        return round(100.0 * n / d, 1) if d else 0.0

    metric_row("Test Cases", lambda rs: len(rs))
    metric_row("Errors", lambda rs: sum(1 for r in rs if r.error))
    metric_row("Scored (no error)", lambda rs: sum(1 for r in rs if not r.error))
    metric_row(
        "Top-1 Accuracy (%)",
        lambda rs: pct(sum(1 for r in rs if not r.error and r.top1_correct), sum(1 for r in rs if not r.error)),
    )
    for k in top_ks:
        metric_row(
            f"Recall@{k} (%)",
            lambda rs, k=k: pct(sum(1 for r in rs if not r.error and r.hit_at(k)), sum(1 for r in rs if not r.error)),
        )
    metric_row(
        "MRR",
        lambda rs: round(
            sum(r.reciprocal_rank for r in rs if not r.error) / max(1, sum(1 for r in rs if not r.error)), 4
        ),
    )
    # The API reporting exact_match while the scorer still calls the row
    # wrong can mean the expected FAQ IDs have drifted from the FAQ set — but
    # it isn't the only explanation, and on a real 905-row run investigated
    # directly, fixing the two ID-normalization bugs below flipped zero rows
    # from wrong to correct. Treat a high number here as a prompt to check
    # with --faq-file, not as proof either way.
    metric_row(
        "Exact match but scored wrong (%)",
        lambda rs: pct(
            sum(1 for r in rs if not r.error and r.exact_match and not r.top1_correct),
            sum(1 for r in rs if not r.error and r.exact_match),
        ),
    )
    # How often the API committed to an answer at all, as opposed to handing
    # back suggestions for the caller to choose from.
    metric_row(
        "Direct answer returned (%)",
        lambda rs: pct(
            sum(1 for r in rs if not r.error and r.direct_answer),
            sum(1 for r in rs if not r.error and r.direct_answer is not None),
        ),
    )
    metric_row(
        "Avg Latency (ms)",
        lambda rs: round(
            sum(r.latency_ms for r in rs if r.latency_ms is not None)
            / max(1, sum(1 for r in rs if r.latency_ms is not None)),
            1,
        ),
    )
    languages = sorted({r.case.lang for r in results})
    if len(languages) > 1:
        ws2.append([])
        ws2.append(["By Language", "Test Cases", "Errors", "Top-1 Accuracy (%)", "MRR"])
        for cell in ws2[ws2.max_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        for lang in languages:
            lang_results = [r for r in results if r.case.lang == lang]
            lang_scored = [r for r in lang_results if not r.error]
            ws2.append(
                [
                    lang,
                    len(lang_results),
                    len(lang_results) - len(lang_scored),
                    pct(sum(1 for r in lang_scored if r.top1_correct), len(lang_scored)),
                    round(sum(r.reciprocal_rank for r in lang_scored) / max(1, len(lang_scored)), 4),
                ]
            )

    ws2.append([])
    ws2.append([f"Generated: {datetime.now().isoformat(timespec='seconds')}"])

    for i, w in enumerate([24, 12, 16, 16], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_path)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def parse_headers(header_args: list[str]) -> dict:
    headers = {}
    for h in header_args or []:
        if ":" not in h:
            raise ValueError(f"--header value {h!r} must be in 'Key: Value' form.")
        key, _, value = h.partition(":")
        headers[key.strip()] = value.strip()
    return headers


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--testcases", type=Path, help="Input .xlsx of test cases.")
    ap.add_argument("--output", type=Path, help="Output .xlsx path (default: results_<timestamp>.xlsx).")
    ap.add_argument("--api-url", default=DEFAULT_API_URL)
    ap.add_argument("--project-id", default=DEFAULT_PROJECT_ID)
    ap.add_argument("--cms-project-id", default=DEFAULT_CMS_PROJECT_ID)
    ap.add_argument("--graph-name", default=DEFAULT_GRAPH_NAME)
    ap.add_argument("--default-path", choices=VALID_PATHS, help="Path to use for rows without one.")
    ap.add_argument("--default-lang", default="en", help="query_lang to use for rows without one.")
    ap.add_argument("--top-k", default="1,3,5", help="Comma-separated k values for Recall@k, e.g. '1,3,5'.")
    ap.add_argument("--timeout", type=float, default=30.0)
    ap.add_argument("--delay", type=float, default=0.2, help="Seconds to sleep between requests.")
    ap.add_argument("--retries", type=int, default=2, help="Retries on network error / HTTP 5xx.")
    ap.add_argument("--header", action="append", help="Extra request header 'Key: Value'. Repeatable.")
    ap.add_argument("--response-list-path", help="Dotted path to the retrieved-documents list in the response, e.g. 'data.documents'. Overrides auto-detection.")
    ap.add_argument("--max-retrieved-cols", type=int, default=5, help="How many retrieved ranks to include as columns.")
    ap.add_argument("--faq-file", type=Path, help="FAQ workbook (e.g. FAQ_Leasing.xlsx) used to annotate results with the expected/matched FAQ question and to warn about expected IDs that do not exist. Never used for scoring.")
    ap.add_argument("--limit", type=int, help="Only run the first N test cases (smoke test).")
    ap.add_argument("--save-raw", action="store_true", help="Include the full raw JSON response per row in the output.")
    ap.add_argument("--verbose", action="store_true")
    ap.add_argument("--probe-query", help="Send a single ad-hoc query and pretty-print the raw response, then exit.")
    ap.add_argument("--probe-path", choices=VALID_PATHS, help="Path to use with --probe-query.")
    ap.add_argument("--probe-lang", default="en", help="query_lang to use with --probe-query.")
    args = ap.parse_args()

    headers = parse_headers(args.header)
    session = requests.Session()

    if args.probe_query:
        if not args.probe_path:
            ap.error("--probe-path is required with --probe-query.")
        body = build_request_body(
            args.probe_query, args.probe_path, args.probe_lang,
            args.project_id, args.cms_project_id, args.graph_name,
        )
        print("Request body:")
        print(json.dumps(body, indent=2, ensure_ascii=False))
        data, status, latency_ms, error = call_api(
            session, args.api_url, body, headers, args.timeout, args.retries
        )
        print(f"\nHTTP {status}  ({latency_ms:.0f} ms)")
        if error:
            print(f"Error: {error}")
        print("\nRaw response:")
        print(json.dumps(data, indent=2, ensure_ascii=False) if data is not None else "<no JSON body>")
        if data is not None:
            found, exact_match, confidence = extract_match_data(
                data, ID_KEY_NAME, RETRIEVED_TEXT_KEY_NAME, args.response_list_path
            )
            print(f"\nAPI exact_match: {exact_match}   confidence_score: {confidence}")
            print(f"\nMatched FAQ + candidates ({len(found)} item(s), matched one first):")
            print(json.dumps(found, indent=2, ensure_ascii=False))
            if not found:
                print(
                    "\nNo FAQ_ID could be extracted (no entity.FAQ_ID, similar_ids, or matching "
                    "document list). Inspect the raw response above and pass --response-list-path "
                    "if this graph config returns a different shape."
                )
        return

    if not args.testcases:
        ap.error("--testcases is required (or use --probe-query to inspect the API first).")

    top_ks = sorted({int(k) for k in args.top_k.split(",") if k.strip()})
    output_path = args.output or Path(f"results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    print(f"Loading test cases from {args.testcases} ...")
    cases = load_testcases(args.testcases, args.default_path, args.default_lang)
    if args.limit:
        cases = cases[: args.limit]
    print(f"Loaded {len(cases)} test case(s).")

    faq_map = None
    if args.faq_file:
        faq_map = load_faq_questions(args.faq_file)
        total = sum(len(t) for t in faq_map.values())
        print(f"Loaded {total} FAQ question(s) from {args.faq_file} "
              f"across {len(faq_map)} language(s): {', '.join(sorted(faq_map))}.")
        missing = check_expected_ids(cases, faq_map)
        if missing:
            print(f"\nWARNING: {len(missing)} expected FAQ ID(s) do not exist in {args.faq_file}.")
            for line in missing[:20]:
                print(f"  {line}")
            if len(missing) > 20:
                print(f"  ... and {len(missing) - 20} more")
            print("  Your test-case labels and your FAQ file are out of sync. Fix that before "
                  "reading any accuracy number from this run.\n")

    results: list[TestResult] = []
    for i, case in enumerate(cases, start=1):
        body = build_request_body(
            case.query, case.path, case.lang,
            args.project_id, args.cms_project_id, args.graph_name,
        )
        data, status, latency_ms, error = call_api(
            session, args.api_url, body, headers, args.timeout, args.retries
        )
        result = TestResult(case=case, http_status=status, latency_ms=latency_ms, error=error, raw_response=data)
        if data is not None and not error:
            try:
                retrieved, exact_match, confidence = extract_match_data(
                    data, ID_KEY_NAME, RETRIEVED_TEXT_KEY_NAME, args.response_list_path
                )
            except ValueError as exc:
                retrieved, exact_match, confidence = [], None, None
                result.error = str(exc)
            result.retrieved = retrieved
            result.exact_match = exact_match
            result.confidence_score = confidence
            result.direct_answer = has_direct_match(data, ID_KEY_NAME)
            result.rank = score_result(case, retrieved)
        if result.error:
            print(f"[{i}/{len(cases)}] row {case.row_num} ({case.path}): ERROR: {result.error}")
        elif args.verbose:
            print(f"[{i}/{len(cases)}] row {case.row_num} ({case.path}): "
                  f"rank={result.rank} top1={'Y' if result.top1_correct else 'N'} "
                  f"({result.latency_ms:.0f} ms)")
        results.append(result)
        if args.delay and i < len(cases):
            time.sleep(args.delay)

    write_results(results, output_path, top_ks, args.max_retrieved_cols, args.save_raw, faq_map)
    print(f"\nWrote results to {output_path}")

    scored = [r for r in results if not r.error]
    if scored:
        top1 = 100.0 * sum(1 for r in scored if r.top1_correct) / len(scored)
        mrr = sum(r.reciprocal_rank for r in scored) / len(scored)
        print(f"Top-1 accuracy: {top1:.1f}%   MRR: {mrr:.4f}   "
              f"(scored {len(scored)}/{len(results)}, {len(results) - len(scored)} error(s))")

        exact = [r for r in scored if r.exact_match]
        if exact:
            wrong = sum(1 for r in exact if not r.top1_correct)
            share = 100.0 * wrong / len(exact)
            if share >= 20.0:
                print(
                    f"\nNOTE: the API reported an exact match on {len(exact)} row(s), but "
                    f"{wrong} of them ({share:.0f}%) still scored wrong.\n"
                    f"  This can mean the expected FAQ IDs have drifted from the current FAQ set —\n"
                    f"  but it can just as easily mean the bot is matching confidently to the wrong\n"
                    f"  FAQ. Don't assume either cause: pass --faq-file to see the expected and\n"
                    f"  matched question text side by side for these rows before drawing a\n"
                    f"  conclusion from this number."
                )


if __name__ == "__main__":
    main()
