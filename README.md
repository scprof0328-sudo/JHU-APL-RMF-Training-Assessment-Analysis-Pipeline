# RMF Training Assessment – Data Preparation

A single‑file Python script that converts the six Kahoot/Qualtrics export Excel files (pre‑ and post‑training for each of the three RMF sessions) into a set of CSV files ready for Power BI. All column names are plain English and the calculations follow the **RMF Training Assessment Aspire** specification.

---

## Table of Contents
1. Quick start  
2. Requirements  
3. Repository layout  
4. Running the script  
5. Outputs produced  
6. Sensitive data handling  
7. Configuration & customization  
8. Metrics reference  
9. Troubleshooting  
10. Contributing  
11. License  
12. Contact  

---

## Quick start
```bash
# Clone the repository
git clone https://github.com/your-org/rmf-training-assessment.git
cd rmf-training-assessment

# (Optional) create an isolated Python environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Place the raw Excel exports in the folder RMF_Raw_Data/
# (the filenames must contain the word "kahoot", e.g. RMF01_Pre_Kahoot.xlsx)

# Execute the script
python data_prep.py
```

When the script finishes, the folder `RMF_Analyzed_Data/` will contain all output files.

---

## Requirements
| Package | Minimum version |
|---------|-----------------|
| Python | 3.9 |
| pandas | 1.5 |
| openpyxl | 3.1 |
| xlrd | 2.0 (only needed for legacy `.xls` files) |

All required packages are listed in `requirements.txt`. Install them with `pip install -r requirements.txt`.

---

## Repository layout
```
rmf-training-assessment/
├─ data_prep.py                # main script
├─ requirements.txt
├─ .gitignore                  # excludes raw and generated data
├─ README.md
├─ RMF_Raw_Data/               # <‑‑ place your raw *.xlsx / *.xls files here
└─ RMF_Analyzed_Data/          # generated CSVs (created automatically)
```

**Important:** The raw Excel files and the generated CSVs contain participant identifiers and free‑text comments. Both directories are ignored by Git (see `.gitignore`).

---

## Running the script
The script does not accept command‑line arguments; it expects the default folder structure described above. If you need to change the input or output location, edit the constants near the top of `data_prep.py`:

```python
BASE_DIR = pathlib.Path(__file__).parent
RAW_DIR  = BASE_DIR / "RMF_Raw_Data"
OUT_DIR  = BASE_DIR / "RMF_Analyzed_Data"
```

The script will:

1. Locate all Excel files in `RMF_Raw_Data/` whose names contain “kahoot”.  
2. Load each workbook (supporting both `.xlsx` and `.xls`).  
3. Add helper columns (`is_correct`, numeric conversions, topic extraction, Likert mapping).  
4. Aggregate knowledge scores, Likert results, and free‑text statistics.  
5. Write seven CSV/text files to `RMF_Analyzed_Data/`.

---

## Outputs produced
All files are written to `RMF_Analyzed_Data/`.

| File | Description |
|------|-------------|
| `question_list.txt` | List of every distinct question text (for manual verification). |
| `RMF_pre_summary.csv` | Knowledge‑level aggregates for the pre‑training data. |
| `RMF_post_summary.csv` | Same as above, for post‑training data. |
| `RMF_wide_summary.csv` | Side‑by‑side pre / post view for each participant (score change, percentiles, etc.). |
| `RMF_metrics_summary.csv` | One‑row‑per‑metric table ready for KPI visualisation. |
| `RMF_participant_summary.csv` | Participant‑level table with all Likert scores, optional free‑text columns, and derived flags. |
| `open_text_top_words.csv` | Top‑20 most frequent non‑stop‑words from any open‑text comment column (if present). |

Numeric columns are rounded to two decimal places where appropriate.

---

## Sensitive data handling
* The raw export files contain personally identifiable information (player IDs, free‑text comments). **Do not commit** `RMF_Raw_Data/` to any public repository.
* The generated CSVs also contain participant identifiers; they are reproducible from the raw data, so they are excluded from version control as well.
* The repository’s `.gitignore` already lists both directories:

```gitignore
# Raw exports (contain participant identifiers)
RMF_Raw_Data/
# Generated CSVs – can be recreated from the raw data
RMF_Analyzed_Data/
```

If you need to share the data with others, consider pseudonymising the `Player` column before committing:

```python
import hashlib
df["Player"] = df["Player"].apply(
    lambda x: hashlib.sha256(str(x).encode()).hexdigest()[:10]
)
```

Place the snippet after the DataFrame is loaded (`pd.read_excel`).

---

## Configuration & customization
| Aspect | Location in script | How to modify |
|--------|-------------------|---------------|
| Target proficiency score (points) | `TARGET_SCORE = 6000` | Change the integer value. |
| Domain‑level proficiency threshold | `target_proficiency = 80` | Adjust the percentage. |
| Likert text‑to‑numeric mapping | `likert_map = {...}` | Add, remove, or rename entries. |
| Topic extraction (keyword → topic) | `extract_topic()` | Extend the `topic_map` dictionary. |
| Keyword sets for Likert metrics (role clarity, confidence, etc.) | Lists such as `role_clarity_keywords`, `confidence_questions` | Edit or add keywords to match your questionnaire wording. |
| Number of lowest‑scoring knowledge gaps reported | `bottom_n = 5` | Change the integer. |
| Stop‑words for open‑text frequency | `stop_words` set | Add or remove stop‑words. |

The script is deliberately simple and avoids external configuration files. If you anticipate many environment‑specific tweaks, you may refactor the constants into a `config.yaml` and load them with a YAML parser.

---

## Metrics reference
Below is a concise description of every metric that appears in `RMF_metrics_summary.csv`.

| Metric (as appears) | Meaning |
|---------------------|---------|
| Average knowledge score before training | Mean of `final_score_pre` across participants. |
| Average knowledge score after training | Mean of `final_score_post`. |
| Average percentage‑point improvement | `(avg_after – avg_before)` expressed in points. |
| Percentage of participants whose scores improved | Share of participants with `score_change > 0`. |
| Percentage of participants meeting target proficiency **before** training | Share where `final_score_pre >= TARGET_SCORE`. |
| Percentage of participants meeting target proficiency **after** training | Share where `final_score_post >= TARGET_SCORE`. |
| Knowledge percentage correct – *{topic}* before training | Domain‑level correctness (%) for each topic (pre). |
| Knowledge percentage correct – *{topic}* after training | Same as above (post). |
| Knowledge improvement – *{topic}* | Difference between post and pre percentages. |
| Average *{friendly_name}* before training | Mean Likert (1‑4) for the given question group (pre). |
| Average *{friendly_name}* after training | Same as above (post). |
| Percent increase in average *{friendly_name}* | `(after – before) / before × 100`. |
| Percent of participants rating *{friendly_name}* as Agree or Strongly Agree (before) | Share of Likert ≥ 3 (pre). |
| Percent of participants rating *{friendly_name}* as Agree or Strongly Agree (after) | Same as above (post). |
| Target confidence score after training | Benchmark = `ceil(avg_conf_post × 2) / 2`, rounded to the nearest 0.5. |
| Percent of participants meeting the confidence target after training | Share where post‑confidence ≥ target. |
| Lowest‑scoring question – ‘{question}’ | Percentage correct for the five worst knowledge items. |
| Domain difficulty – *{topic}* | Overall % correct per topic, sorted low → high. |
| Hardest/Easiest domain | Top‑3 hardest and easiest topics. |
| Hardest knowledge question – session *x* (pre/post) | Question with the lowest % correct for that session/timepoint. |
| Easiest knowledge question – session *x* (pre/post) | Question with the highest % correct for that session/timepoint. |
| Percent of participants whose role clarity rating improved | Share where post > pre for role‑clarity Likert. |
| Average confidence change (after – before) | Mean of (post – pre) confidence Likert per participant. |
| Percent of participants who identified at least one planned change | Based on optional free‑text column “Planned Change”. |
| Percent of participants meeting target (*{domain}*) | For each topic, share of participants with % correct ≥ `target_proficiency`. |

All numeric values are rounded to two decimal places.

---

## Troubleshooting

| Symptom | Likely cause | Remedy |
|---------|--------------|--------|
| `FileNotFoundError: No Kahoot/Qualtrics Excel files were found.` | No files in `RMF_Raw_Data/` or filenames do not contain the word “kahoot”. | Rename files to include “kahoot” (case‑insensitive) and ensure they are placed in the correct folder. |
| `zipfile.BadZipFile` | Corrupt or non‑Excel `.xlsx` file. | Re‑export the file from Kahoot/Qualtrics or delete the problematic file. |
| `KeyError: 'Correct / Incorrect'` | Column name differs in your export. | Update every occurrence of `"Correct / Incorrect"` in the script to match the actual column header. |
| `ImportError: cannot import name 'xlrd'` | Using pandas 2.x which dropped legacy `.xls` support. | Ensure input files are `.xlsx`. If you must read `.xls`, install an older `xlrd` (e.g., `pip install xlrd==1.2.0`). |
| `MemoryError` while loading a file | Very large `.xls` workbook. | Convert the file to `.xlsx` or run the script on a machine with more RAM. |
| Metric value appears as `None` | No rows matched the supplied keyword list for that Likert item. | Verify the wording of the question in the raw data and adjust the keyword list accordingly. |
| Output CSVs missing expected columns | Raw export lacks one of the expected columns (e.g., `Answer Time (seconds)`). | The script silently skips missing columns; add the column to your export or remove its reference from the `keep` list. |

---

## Contributing
1. Fork the repository.  
2. Create a feature branch (`git checkout -b my-feature`).  
3. Make changes, adhering to PEP 8 style and adding docstrings where appropriate.  
4. Run the test suite (if applicable) with `pytest`.  
5. Submit a Pull Request with a clear description of the change.

### Testing (optional)
The repository includes a minimal test suite under `tests/`. To run:

```bash
pip install pytest
pytest
```

The tests use synthetic data and do not require any real participant data.

### Code quality
We use `ruff` (or `flake8`) for linting. Run locally:

```bash
pip install ruff
ruff check .
```

---

## License
This project is licensed under the MIT License. See the `LICENSE` file for the full text.

```
MIT License

Copyright (c) 2024 <Your Organization>

Permission is hereby granted, free of charge, to any person obtaining a copy
...
```

Replace `<Your Organization>` with the appropriate holder.

---

## Contact
- **Primary maintainer:** Shaunak Chittimalla - shaunakprof0328@gmail.com
