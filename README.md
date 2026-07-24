# CRIDA DACP Chatbot Testing

Evaluation framework for scoring chatbot responses against a curated golden dataset of agronomic recommendations for paddy/rice cultivation in Chhattisgarh districts.

---

## Directory Structure

```
chatbot_testing/
├── code/
│   ├── score_chatbot.py          # Main scoring script
│   └── rebuild_golden_dataset.py # Golden dataset preprocessing script
├── data/
│   ├── Prompt Question Bank.xlsx # Prompts, chatbot responses, and metadata
│   └── golden_dataset_v2.xlsx    # Curated reference dataset (one sheet per district)
└── output/
    └── Chatbot_Scoring_Results_v4.xlsx  # Scored results
```

---

## Data

### `Prompt Question Bank.xlsx`
Contains the prompts sent to the chatbot and its recorded responses. Key columns:
- **Sr. No.** — row identifier
- **Prompt Type** — `Standard prompt` or `Adversarial Prompt`
- **Prompt Questions** — the query submitted to the chatbot
- **Chatbot response** — the chatbot's answer (filled in manually)
- **Key Inputs** — sub-district, land type, and irrigation context extracted from the prompt
- **Flags for incorrect prompts** — expected behavior for adversarial prompts

### `golden_dataset_v2.xlsx`
Reference dataset with one sheet per Chhattisgarh district. Each row represents a specific agronomy scenario (crop, land type, irrigation availability, monsoon condition). Key columns include expected seed varieties, farming practices, fertilizer doses, chemicals, and infrastructure recommendations.

Built from a raw `golden_dataset.xlsx` using `rebuild_golden_dataset.py`, which filters to Standard input type, paddy/rice rows, and monsoon onset scenarios only.

---

## Code

### `score_chatbot.py`
Main evaluation script. For each prompt in the question bank:

- **Standard prompts** — resolves sub-district to district, looks up matching golden dataset rows, extracts expected keywords, and scores the chatbot response by percentage of keywords captured. Matching uses exact string match, fuzzy match (rapidfuzz), and optional semantic similarity (sentence-transformers).
- **Adversarial prompts** — checks whether the chatbot correctly flagged an invalid or out-of-scope query instead of answering it directly.

Outputs a color-coded Excel file with three sheets: Summary, Standard Scores, and Adversarial Scores.

**Dependencies:** `pandas`, `openpyxl`, `rapidfuzz`, `indic-transliteration`, `sentence-transformers` (optional)

**Run:**
```bash
pip install pandas openpyxl rapidfuzz indic-transliteration
python3 code/score_chatbot.py
```

### `rebuild_golden_dataset.py`
Preprocesses the raw golden dataset. Filters each district sheet to Standard input type rows, paddy/rice crops, and monsoon onset scenarios. Trims to essential columns, formats keyword cells as bullet lists, and applies Excel styling. Run this if the source `golden_dataset.xlsx` changes.

**Run** (from `data/` directory):
```bash
python3 ../code/rebuild_golden_dataset.py
```

---

## Output

### `Chatbot_Scoring_Results_v4.xlsx`

| Sheet | Contents |
|---|---|
| **Summary** | Aggregate counts — total prompts, average score, pass/fail breakdown |
| **Standard Scores** | Per-prompt keyword scores with matched/missed keyword lists and notes |
| **Adversarial Scores** | Per-prompt PASS/FAIL indicating whether the chatbot correctly refused or flagged bad input |

Score color coding: green >= 60%, yellow 30–59%, red < 30%, grey = N/A (unevaluable).

---

## Scoring Logic

- **Keyword match** uses a three-tier cascade: exact string → fuzzy partial ratio (threshold 75) → semantic cosine similarity (threshold 0.60).
- **Adversarial pass** requires the chatbot response to contain a flagging phrase (e.g., "does not exist", "please clarify") rather than a direct answer.
- Prompts with missing key inputs (sub-district/land type/irrigation) where the chatbot correctly asks for clarification are marked `N/A` with a "Correct behaviour" note.
