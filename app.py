#!/usr/bin/env python3
"""
Local web UI for the Sino Voicebot FAQ retrieval benchmark tool.

Runs entirely on your machine and calls dev.setsailapi.com server-side
(the same way sino_retrieval_benchmark.py does) — this has to be a locally
run app rather than a hosted page because browsers block a page on one
origin from calling an API on another origin unless that API opts in via
CORS, which this one does not.

Usage:
    pip install -r requirements.txt
    python3 app.py
    # then open http://127.0.0.1:5000 in a browser
"""

from __future__ import annotations

import io
import re
import tempfile
import uuid
from pathlib import Path

import requests
from flask import Flask, redirect, render_template, request, send_file, url_for

from sino_retrieval_benchmark import (
    DEFAULT_API_URL,
    DEFAULT_CMS_PROJECT_ID,
    DEFAULT_GRAPH_NAME,
    DEFAULT_PROJECT_ID,
    ID_KEY_NAME,
    RETRIEVED_TEXT_KEY_NAME,
    VALID_PATHS,
    TestCase,
    TestResult,
    build_request_body,
    call_api,
    extract_match_data,
    load_testcases,
    parse_headers,
    score_result,
    write_results,
)

app = Flask(__name__)

RESULTS_DIR = Path(tempfile.gettempdir()) / "sino_benchmark_web_results"
RESULTS_DIR.mkdir(exist_ok=True)
SAFE_FILENAME_RE = re.compile(r"^[A-Za-z0-9_.-]+$")

DEFAULTS = {
    "api_url": DEFAULT_API_URL,
    "project_id": DEFAULT_PROJECT_ID,
    "cms_project_id": DEFAULT_CMS_PROJECT_ID,
    "graph_name": DEFAULT_GRAPH_NAME,
}


def form_settings(form) -> dict:
    return {
        "api_url": form.get("api_url", "").strip() or DEFAULT_API_URL,
        "project_id": form.get("project_id", "").strip() or DEFAULT_PROJECT_ID,
        "cms_project_id": form.get("cms_project_id", "").strip() or DEFAULT_CMS_PROJECT_ID,
        "graph_name": form.get("graph_name", "").strip() or DEFAULT_GRAPH_NAME,
        "headers_raw": form.get("headers_raw", "").strip(),
        "timeout": float(form.get("timeout") or 30),
        "retries": int(form.get("retries") or 2),
        "response_list_path": form.get("response_list_path", "").strip() or None,
    }


def parse_headers_raw(headers_raw: str) -> dict:
    if not headers_raw:
        return {}
    lines = [line for line in headers_raw.splitlines() if line.strip()]
    return parse_headers(lines)


@app.route("/", methods=["GET"])
def index():
    return render_template(
        "index.html",
        defaults=DEFAULTS,
        valid_paths=VALID_PATHS,
        probe=None,
        batch=None,
    )


@app.route("/probe", methods=["POST"])
def probe():
    settings = form_settings(request.form)
    query = request.form.get("query", "").strip()
    path = request.form.get("path", "").strip()
    lang = request.form.get("lang", "en").strip() or "en"
    expected_faq_id = request.form.get("expected_faq_id", "").strip()
    expected_ids = [e.strip() for e in expected_faq_id.split(",") if e.strip()]

    probe_result = {
        "query": query, "path": path, "lang": lang, "error": None,
        "expected_faq_id": expected_faq_id, "expected_ids": expected_ids,
    }
    if not query or path not in VALID_PATHS:
        probe_result["error"] = "A query and a valid Path are required."
        return render_template(
            "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=probe_result, batch=None
        )

    try:
        headers = parse_headers_raw(settings["headers_raw"])
    except ValueError as exc:
        probe_result["error"] = str(exc)
        return render_template(
            "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=probe_result, batch=None
        )

    body = build_request_body(
        query, path, lang,
        settings["project_id"], settings["cms_project_id"], settings["graph_name"],
    )
    session = requests.Session()
    data, status, latency_ms, error = call_api(
        session, settings["api_url"], body, headers, settings["timeout"], settings["retries"]
    )
    retrieved, exact_match, confidence = [], None, None
    rank = None
    if data is not None:
        try:
            retrieved, exact_match, confidence = extract_match_data(
                data, ID_KEY_NAME, RETRIEVED_TEXT_KEY_NAME, override_path=settings["response_list_path"]
            )
        except ValueError as exc:
            error = error or str(exc)
        if expected_ids:
            temp_case = TestCase(row_num=0, query=query, expected_ids=expected_ids, path=path, lang=lang)
            rank = score_result(temp_case, retrieved)

    probe_result.update(
        {
            "request_body": body,
            "status": status,
            "latency_ms": latency_ms,
            "error": error,
            "raw_response": data,
            "retrieved": retrieved,
            "exact_match": exact_match,
            "confidence": confidence,
            "matched_id": retrieved[0][ID_KEY_NAME] if retrieved else "",
            "matched_response": retrieved[0][RETRIEVED_TEXT_KEY_NAME] if retrieved else "",
            "other_suggestions": retrieved[1:],
            "rank": rank,
        }
    )
    return render_template(
        "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=probe_result, batch=None
    )


@app.route("/batch", methods=["POST"])
def batch():
    settings = form_settings(request.form)
    default_path = request.form.get("default_path", "").strip() or None
    default_lang = request.form.get("default_lang", "en").strip() or "en"
    delay = float(request.form.get("delay") or 0.2)
    top_k_raw = request.form.get("top_k", "1,3,5").strip() or "1,3,5"
    limit_raw = request.form.get("limit", "").strip()
    limit = int(limit_raw) if limit_raw else None

    batch_result = {"error": None}
    upload = request.files.get("testcases")
    if not upload or not upload.filename:
        batch_result["error"] = "Please choose a test-case .xlsx file."
        return render_template(
            "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=None, batch=batch_result
        )

    try:
        headers = parse_headers_raw(settings["headers_raw"])
    except ValueError as exc:
        batch_result["error"] = str(exc)
        return render_template(
            "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=None, batch=batch_result
        )

    run_id = uuid.uuid4().hex[:12]
    upload_path = RESULTS_DIR / f"upload_{run_id}.xlsx"
    upload.save(upload_path)

    try:
        top_ks = sorted({int(k) for k in top_k_raw.split(",") if k.strip()})
        cases = load_testcases(upload_path, default_path, default_lang)
    except ValueError as exc:
        batch_result["error"] = str(exc)
        return render_template(
            "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=None, batch=batch_result
        )
    finally:
        upload_path.unlink(missing_ok=True)

    if limit:
        cases = cases[:limit]

    session = requests.Session()
    results: list[TestResult] = []
    import time as _time

    for i, case in enumerate(cases):
        req_body = build_request_body(
            case.query, case.path, case.lang,
            settings["project_id"], settings["cms_project_id"], settings["graph_name"],
        )
        data, status, latency_ms, error = call_api(
            session, settings["api_url"], req_body, headers, settings["timeout"], settings["retries"]
        )
        result = TestResult(case=case, http_status=status, latency_ms=latency_ms, error=error, raw_response=data)
        if data is not None and not error:
            try:
                retrieved, exact_match, confidence = extract_match_data(
                    data, ID_KEY_NAME, RETRIEVED_TEXT_KEY_NAME, override_path=settings["response_list_path"]
                )
            except ValueError as exc:
                retrieved, exact_match, confidence = [], None, None
                result.error = str(exc)
            result.retrieved = retrieved
            result.exact_match = exact_match
            result.confidence_score = confidence
            result.rank = score_result(case, retrieved)
        results.append(result)
        if delay and i < len(cases) - 1:
            _time.sleep(delay)

    output_filename = f"results_{run_id}.xlsx"
    write_results(results, RESULTS_DIR / output_filename, top_ks)

    scored = [r for r in results if not r.error]
    summary = {
        "total": len(results),
        "errors": len(results) - len(scored),
        "top1_pct": round(100.0 * sum(1 for r in scored if r.top1_correct) / len(scored), 1) if scored else 0.0,
        "mrr": round(sum(r.reciprocal_rank for r in scored) / len(scored), 4) if scored else 0.0,
    }

    batch_result.update(
        {
            "results": results,
            "summary": summary,
            "top_ks": top_ks,
            "download_filename": output_filename,
        }
    )
    return render_template(
        "index.html", defaults=DEFAULTS, valid_paths=VALID_PATHS, probe=None, batch=batch_result
    )


@app.route("/download/<filename>")
def download(filename: str):
    if not SAFE_FILENAME_RE.match(filename):
        return "Invalid filename.", 400
    file_path = RESULTS_DIR / filename
    if not file_path.is_file():
        return "File not found (results are temporary and may have been cleared).", 404
    return send_file(file_path, as_attachment=True, download_name="benchmark_results.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
