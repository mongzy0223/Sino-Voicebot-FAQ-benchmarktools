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

### How the response is read

Confirmed against a live call — the API does **not** return a ranked list
of retrieved documents. It resolves the query to a single matched FAQ and
hands back:

```jsonc
{
  "exact_match": true,
  "similar_ids": ["2", "46", "19"],       // candidate FAQ_IDs, matched one usually first
  "entity": {"FAQ_ID": "2", "lang": "en"}, // the matched FAQ_ID
  "bot_responses": [
    {"answer": "...", "confidence_score": 1, ...}  // that FAQ's answer text
  ],
  // "actions"[0].data.message.content and "message.series"[0].text carry
  // the same answer text redundantly
}
```

`extract_match_data()` in `sino_retrieval_benchmark.py` reads
`entity.FAQ_ID` as the matched FAQ, `bot_responses[0].answer` as its
response text, and `similar_ids` as the other candidate IDs (no response
text is returned for those — only their IDs). `exact_match` and
`confidence_score` are surfaced as their own columns. A generic
list-of-documents search is kept as a fallback for other graph configs
that might return an actual document list, and `--response-list-path`
overrides both if a graph config returns something else entirely — use
`--probe-query` (below) to check.

**The CMS's real `FAQ_ID` values are bare numbers** (e.g. `"2"`, not
`"Q2"`). If your test-case sheet uses `Q`-prefixed reference numbers (like
Sino's own benchmark template's `FAQ Reference No.` column), the tool
strips a leading `Q`/`q` before comparing IDs so `Q2` matches the API's
`2` — see `_normalize_faq_id()`. IDs that don't look like `Q<number>` are
compared as-is (case-insensitive), so the original `MTG-001`-style
template is unaffected.

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
CLI — a single-query probe box (query + Path + Language, plus an optional
Expected FAQ_ID field so a single probe also shows Correct/Rank against a
known answer, matching your test-case sheet's `FAQ Reference No.`) and a
batch runner (upload a test-case `.xlsx`, get a scored summary, per-row
results table, and a results file to download) — plus an "Advanced
settings" panel per form for the API URL, project IDs, auth headers,
timeout/retries, and a response-list-path override. The probe result shows
the matched FAQ and other suggested FAQs in a plain results view; the raw
request/response JSON is tucked into a collapsed "Show raw request/response
(debug)" section if you need it.

The batch runner runs in the background and shows a live progress bar
(`current / total`, updated roughly twice a second) while the "Run
benchmark" button is disabled — you can't accidentally start a second run
on top of one still in progress. The page moves to the results view
automatically once it finishes.

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
**Matched_FAQ_ID** / **Matched_Response** (the FAQ the API actually
matched — `entity.FAQ_ID` — and its answer text from `bot_responses`),
**API_Exact_Match** / **Confidence_Score** (the API's own verdict — useful
to tell "wrong FAQ" apart from "the API fell back / wasn't confident"),
the full ranked **Suggested_1..N_FAQ_ID** / **Suggested_i_Response** pairs
(the matched FAQ plus every other candidate ID from `similar_ids`, in
order — response text is only available for the matched one, since that's
all the API returns), latency, HTTP status, and any error.

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

**Matched FAQ / candidates come back empty.** Run `--probe-query` and read
the raw response. Normally `entity.FAQ_ID` and `similar_ids` are all the
tool needs — if your graph config returns something else entirely (no
`entity`, no `similar_ids`, no document list), pass `--response-list-path`
with the dotted key path to whatever list of `{FAQ_ID: ...}` objects it
does return, e.g. `--response-list-path output.retrieved_documents`.

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
