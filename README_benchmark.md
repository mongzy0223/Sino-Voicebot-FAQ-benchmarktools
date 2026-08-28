# Sino Voicebot FAQ Retrieval Benchmark

A command-line tool (and optional local web UI) that batch-tests the Sino
voicebot's FAQ text-retrieval API (`text_retrieval_test` graph on
`dev.setsailapi.com`) against a set of expected question/answer pairs, and
scores the results (Top-1 accuracy, Recall@k, MRR).

## What it calls

```
POST https://dev.setsailapi.com/nlp/llmGraph/graph/multi/run
```

with a body of the form:

```json
{
  "project_id": "public",
  "cms_project_id": "sinoexvoicebot-52ynuou",
  "single_graph_name": "text_retrieval_test",
  "path": "FAQ_Mortgage",
  "id_key_name": "FAQ_ID",
  "display_columns": ["FAQ_ID", "FAQ"],
  "retrieved_text_key_name": "FAQ_answer_listening_friendly",
  "query_lang": "en",
  "messages": [{"type": "human", "content": "<query>"}],
  "stream": false,
  "slim_response": true
}
```

`path` is `FAQ_Mortgage` or `FAQ_Leasing`, set per test case.

> **Note on response parsing.** This tool was built without the ability to
> make a live call to `dev.setsailapi.com` (network policy in the build
> environment blocked it), so the exact JSON shape of a `slim_response`
> reply was not verified directly. `extract_retrieved_list()` in
> `sino_retrieval_benchmark.py` handles this by recursively searching the
> response for a list of objects that each carry an `FAQ_ID` field, which
> should work for most reasonable response shapes. **Before running a full
> batch, use `--probe-query` (below) to print the real raw response** and
> confirm the auto-detected list looks right. If it doesn't, pass
> `--response-list-path` with the dotted key path to the correct list (see
> Troubleshooting).

## Setup

```bash
pip install -r requirements.txt
```

Requires Python 3.9+, `requests`, `openpyxl`, and (for the web UI) `flask`.

If the API requires an API key or auth header, pass it with `--header`,
e.g. `--header "Authorization: Bearer <token>"` (repeatable for multiple
headers).

## Web UI (optional)

Prefer a browser to the command line? Run:

```bash
python3 app.py
```

then open **http://127.0.0.1:5000**. It has the same two workflows as the
CLI — a single-query probe box (to check the API and response shape first)
and a batch runner (upload a test-case `.xlsx`, get a scored summary,
per-row results table, and a results file to download) — plus an
"Advanced settings" panel per form for the API URL, project IDs, auth
headers, timeout/retries, and a response-list-path override.

This has to run locally rather than as a hosted page: browsers block a web
page from calling an API on a different origin unless that API opts in via
CORS, and `dev.setsailapi.com` doesn't. Running it with `python3 app.py`
calls the API from your machine directly, the same way the CLI does — no
request ever goes through a third-party server, and nothing is uploaded
anywhere.

## 1. Inspect the live API first

Before running a full benchmark, send one ad-hoc query and inspect the raw
response:

```bash
python3 sino_retrieval_benchmark.py \
  --probe-query "What is the current mortgage interest rate?" \
  --probe-path FAQ_Mortgage
```

This prints the exact request body, HTTP status, the full raw JSON
response, and what the tool auto-detected as the ranked list of retrieved
FAQs. Confirm the auto-detected list matches the real ranked results
before moving on — see Troubleshooting if it doesn't.

## 2. Prepare test cases

Start from `sino_benchmark_testcases_TEMPLATE.xlsx`. It has a `Test Cases`
sheet with these columns:

| Column | Required | Description |
|---|---|---|
| `Query` | Yes | The exact question text sent to the API. |
| `Expected_FAQ_ID` | Yes | The ground-truth `FAQ_ID`. Comma-separate multiple acceptable IDs, e.g. `MTG-002, MTG-003`. |
| `Path` | Yes* | `FAQ_Mortgage` or `FAQ_Leasing`. *Can be omitted if you pass `--default-path`. |
| `Language` | No | `query_lang` sent to the API (e.g. `en`, `zh`, `sc`). Defaults to `en`, or `--default-lang`. |
| `Notes` | No | Free text, passed through to the results file. |

Delete the example rows before running a real benchmark. An `Instructions`
sheet in the template repeats this reference.

**Sino's own multi-language benchmark template is also accepted directly** —
no conversion needed. That format has one sheet per language (e.g.
`Benchmark Template zh` / `en` / `cn`) with `No.` / `Testing Query` /
`Category` / `FAQ Reference No.` columns, and the tool handles it
automatically:

- Every sheet with a recognizable header is loaded and combined into one run.
- `query_lang` is inferred per sheet from a language suffix in its name
  (`zh` → `zh`, `en` → `en`, `cn` → `sc` — edit
  `SHEET_NAME_LANG_SUFFIXES` in `sino_retrieval_benchmark.py` if your sheet
  names or language codes differ).
- `Category` values (`Leasing`, `Mortgage`) are mapped to the API's `Path`
  values (`FAQ_Leasing`, `FAQ_Mortgage`) via `CATEGORY_TO_PATH` — extend
  that dict if a new category is added on the CMS side.
- Rows whose query starts with `EG:` / `EG -` (the template's instructional
  examples) are skipped automatically.
- The results file's `Details` sheet gets a `Sheet` column so you can see
  which language sheet each row came from, and the `Summary` sheet gets a
  "By Language" breakdown (Top-1 accuracy / MRR per language) whenever a
  run spans more than one language.

This assumes the CMS's real `FAQ_ID` values match the `FAQ Reference No.`
values in the sheet (e.g. `Q1`, `Q40`) — confirm this with `--probe-query`
against a query you know the expected ID for before trusting a full run's
scores.

## 3. Run the benchmark

```bash
python3 sino_retrieval_benchmark.py \
  --testcases sino_benchmark_testcases_TEMPLATE.xlsx \
  --output results.xlsx
```

Useful flags:

- `--limit N` — only run the first N test cases (smoke test).
- `--delay 0.2` — seconds to sleep between requests (default `0.2`, be a
  reasonable API citizen).
- `--retries 2` — retries on network errors / HTTP 5xx, with backoff.
- `--timeout 30` — per-request timeout in seconds.
- `--save-raw` — include the full raw JSON response for every row in the
  output (useful for debugging, makes the file much larger).
- `--verbose` — print per-row rank/latency as it runs.
- `--top-k 1,3,5` — which Recall@k values to compute.
- `--default-path` / `--default-lang` — fallback values for test-case rows
  that omit `Path` / `Language`.
- `--response-list-path data.documents` — override auto-detection of the
  retrieved-documents list in the response (see Troubleshooting).

## Output

`results.xlsx` has two sheets:

**Details** — one row per test case: the query, expected FAQ_ID, Top-1
correct (Y/N), the rank at which the expected answer was found (blank if
not found), reciprocal rank, Hit@k for each configured k, the
**Matched_FAQ_ID** / **Matched_Response** (the retrieved FAQ that matched
the expected answer, and its actual answer text), the full ranked
**Suggested_1..N_FAQ_ID** / **Suggested_i_Response** pairs (every candidate
the API returned, in rank order — this includes the matched one at its
rank, plus the alternatives), latency, HTTP status, and any error.

**Summary** — aggregate metrics, both overall and split by `FAQ_Mortgage` /
`FAQ_Leasing`:

- **Top-1 Accuracy** — % of queries where the expected FAQ was the #1
  result.
- **Recall@k** — % of queries where the expected FAQ appeared anywhere in
  the top k results.
- **MRR** (Mean Reciprocal Rank) — average of `1/rank` across all queries
  (0 if never found).
- **Avg Latency** — average API response time in ms.

Rows that errored (network failure, non-2xx, unparseable response) are
counted separately and excluded from accuracy/recall/MRR so a handful of
transient failures don't skew the score — check the `Error` column and
re-run if the error count is non-trivial.

## Troubleshooting

**Auto-detected retrieved list is empty or wrong.** Run `--probe-query`
and read the raw response. Find the key path to the list of retrieved FAQ
objects (each should contain an `FAQ_ID` field) and pass it with
`--response-list-path`, using dots to walk into nested objects, e.g.:

```bash
--response-list-path output.retrieved_documents
```

**All rows show `HTTP None` with a connection error.** Check network
access to `dev.setsailapi.com` and any required auth headers. If you're
behind a proxy, make sure `HTTPS_PROXY`/`HTTP_PROXY` are set in your shell
as needed — `requests` picks them up automatically.

**A `Path` value fails validation.** Only `FAQ_Mortgage` and `FAQ_Leasing`
are accepted (matching the CMS project's two FAQ sets). Check for typos or
stray whitespace in the test case sheet.

**Expected answer never found even though it looks right.** ID matching is
case-insensitive but exact otherwise — check for stray spaces or a
mismatched FAQ_ID between your test sheet and the actual CMS content.
