# CRIDA DACP Chatbot Testing

Evaluation framework for scoring chatbot responses against a curated golden dataset of agronomic recommendations for paddy/rice cultivation in Chhattisgarh districts.

Testing is restricted to three fieldwork districts: **Balrampur, Surajpur, and Jashpur**.

---

## Directory Structure

```
chatbot_testing/
├── code/
│   ├── score_chatbot.py          # Fuzzy/semantic keyword scoring (primary)
│   ├── score_chatbot_rrf.py      # Reciprocal Rank Fusion (RRF) scoring
│   ├── score_chatbot_llm.py      # LLM-as-judge scoring via Claude API
│   └── rebuild_golden_dataset.py # Golden dataset preprocessing script
├── data/
│   ├── Chatbot_Query_Bank_Surajpur_Balrampur_Jashpur_v2.xlsx  # Current query bank (28 rows)
│   ├── Prompt Question Bank.xlsx # Legacy query bank (kept for reference)
│   └── golden_dataset_v2.xlsx    # Curated reference dataset (one sheet per district)
└── output/
    ├── Chatbot_Scoring_Results_v5.xlsx   # Fuzzy/semantic scoring results (current)
    ├── Chatbot_Scoring_Results_rrf.xlsx  # RRF scoring results
    └── Chatbot_Scoring_Results_llm.xlsx  # LLM scoring results
```

---

## Data

### `Chatbot_Query_Bank_Surajpur_Balrampur_Jashpur_v2.xlsx`
The current query bank used for evaluation. 28 rows covering Balrampur, Surajpur, and Jashpur districts. Key columns:

| Column | Description |
|--------|-------------|
| A — Q# | Row identifier |
| B — District | Balrampur, Surajpur, or Jashpur |
| C — Subdistrict | Block/tehsil name |
| D — Land Type (local term) | Hindi land classification (e.g. Mal, Marhan, Gabhar, Tikra) |
| E — Land Type (meaning) | English equivalent (upland, midland, lowland) |
| F — Irrigation | Whether irrigation is available |
| G — Scenario this maps to | Backend-detected scenario label (not in the query); used to select the correct golden dataset row and chatbot response column |
| J — English Query | The query submitted to the chatbot |
| L — English response (normal onset) | Chatbot response for normal onset scenario |
| M–Q | Chatbot responses for 2-, 4-, 6-, 8-week delay and early onset scenarios |

### `golden_dataset_v2.xlsx`
Reference dataset with one sheet per Chhattisgarh district. Each row represents a specific agronomy scenario (crop, land type, irrigation, monsoon condition). Key columns include expected seed varieties, farming practices, fertilizer doses, chemicals, and infrastructure recommendations.

Built from a raw source using `rebuild_golden_dataset.py`, which filters to Standard input type, paddy/rice rows, and monsoon onset scenarios only.

---

## Code

### `score_chatbot.py` — Fuzzy/Semantic Scoring

For each row in the query bank:

1. Reads col G (`Scenario this maps to`) to select the correct chatbot response column (col L for normal onset, cols M–P for delayed scenarios).
2. Maps the scenario label to a specific golden dataset row filter, so the keyword denominator reflects only that scenario's expected recommendations — not a union across all scenarios.
3. Scores the chatbot response by percentage of reference keywords captured, using a three-tier cascade: exact string match → fuzzy partial ratio (threshold 75) → semantic cosine similarity (threshold 0.60).

Outputs `output/Chatbot_Scoring_Results_v5.xlsx`.

**Dependencies:** `pandas`, `openpyxl`, `rapidfuzz`, `indic-transliteration`, `sentence-transformers` (optional)

```bash
pip install pandas openpyxl rapidfuzz indic-transliteration
python3 code/score_chatbot.py
```

### `score_chatbot_rrf.py` — RRF Scoring

Same structure as `score_chatbot.py` but uses Reciprocal Rank Fusion to combine exact, fuzzy, and semantic match signals into a single ranked score rather than a cascade. Outputs `output/Chatbot_Scoring_Results_rrf.xlsx`.

```bash
python3 code/score_chatbot_rrf.py
```

### `score_chatbot_llm.py` — LLM-as-Judge Scoring

Uses the Claude API to judge keyword coverage. For each keyword, the LLM decides whether the chatbot response conveys the meaning, allowing partial credit for paraphrasing and summaries. Requires an `ANTHROPIC_API_KEY` environment variable.

Outputs `output/Chatbot_Scoring_Results_llm.xlsx`.

```bash
export ANTHROPIC_API_KEY=<your key>
python3 code/score_chatbot_llm.py
```

### `rebuild_golden_dataset.py`
Preprocesses the raw golden dataset. Filters each district sheet to Standard input type rows, paddy/rice crops, and monsoon onset scenarios. Trims to essential columns and applies Excel styling. Run this if the source `golden_dataset.xlsx` changes.

```bash
python3 code/rebuild_golden_dataset.py  # run from data/ directory
```

---

## Output

All output files contain a **Summary** sheet and a **Standard Scores** sheet.

| File | Method | Scored | Avg Score | ≥60% |
|------|--------|--------|-----------|------|
| `Chatbot_Scoring_Results_v5.xlsx` | Fuzzy/semantic (scenario-filtered) | 25/28 | 22.2% | 2 |
| `Chatbot_Scoring_Results_rrf.xlsx` | RRF | 26/28 | 17.8% | 0 |
| `Chatbot_Scoring_Results_llm.xlsx` | LLM-as-judge | 26/28 | 24.0% | 0 |

Score color coding: green ≥ 60%, yellow 30–59%, red < 30%, grey = N/A.

The **Standard Scores** sheet includes `Scenario` and `Response Col Used` diagnostic columns showing which golden row and chatbot response column were matched for each query.

---

## Scoring Logic

### Scenario-specific keyword filtering
A key design choice: the scoring denominator uses only the golden dataset row that matches the query's specific scenario (normal onset, delayed 2 weeks, etc.), not a union across all scenario rows. Without this, the denominator inflates 3–5× and realistic chatbot responses can only cover ~20% of the pooled keywords.

Col G in the query bank (`Scenario this maps to`) identifies the correct scenario. This is mapped to a substring filter applied to the golden dataset's `Specific Scenario` column before extracting keywords.

### Land type mapping
Local Hindi terms are normalized before lookup: Marhan/Tikra → upland, Mal → midland, Gabhar → lowland. The Jashpur irrigated golden dataset uses "Low land" (two words) and is matched accordingly.

### Known N/A cases (3 rows)
Three queries have no matching golden dataset row — these are genuine gaps in the golden dataset, not scoring failures:
- **Q7**: Balrampur upland, delayed 6–8 weeks — no golden row for this combination
- **Q13**: Surajpur upland, delayed 4–6 weeks — no separate golden row (Surajpur upland only has a combined "delayed 2–4 weeks" row)
- **Q24**: Balrampur, normal onset — Balrampur has no baseline normal-onset golden row (only delayed and dry-spell rows)

### Known score ceiling for Surajpur upland
The Surajpur upland golden dataset rows include ~35 non-paddy crop variety keywords (pigeonpea, maize, urd, groundnut) that the chatbot explicitly declines to advise on. Paddy-focused responses cannot realistically exceed ~50% on these rows regardless of quality.
