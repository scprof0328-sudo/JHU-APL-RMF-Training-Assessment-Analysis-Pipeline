#!/usr/bin/env python3
"""
data_prep.py
------------

Reads the six Kahoot/Qualtrics export Excel files (pre‑ and post‑ for each
training session), adds helper columns, calculates every metric required by the
“RMF Training Assessment Aspire” specification, and writes a set of CSV files
ready for Power BI.

All output files are written to the folder *RMF_Analyzed_Data* and use only
plain‑English wording – no acronyms or shorthand.
"""

import pathlib
import re
import math                     # used for confidence‑target rounding
import zipfile                  # validates .xlsx files
import pandas as pd

# ---------------------------------------------------------
# 0️⃣  Helper – does a column contain ANY of a list of keywords?
# ---------------------------------------------------------
def contains_any(series: pd.Series, keywords: list[str]) -> pd.Series:
    """
    Returns True for rows where the string in `series` contains *any* keyword
    (case‑insensitive).  Keywords are escaped so characters such as “/” or “‑”
    are safe.
    """
    pattern = "|".join([re.escape(k) for k in keywords])
    return series.str.contains(pattern, case=False, na=False)

# -----------------------------------------------------------------
# 1️⃣  Paths
# -----------------------------------------------------------------
BASE_DIR = pathlib.Path(__file__).parent                 # …/RMFProjectASPIRE
RAW_DIR  = BASE_DIR / "RMF_Raw_Data"                     # raw Kahoot/Qualtrics exports
OUT_DIR  = BASE_DIR / "RMF_Analyzed_Data"                # where CSVs will go
OUT_DIR.mkdir(parents=True, exist_ok=True)               # create if missing

# -----------------------------------------------------------------
# 2️⃣  Helper: parse filename → session + timepoint
# -----------------------------------------------------------------
def parse_filename(fname: str) -> tuple[str, str]:
    """
    Expected patterns (case‑insensitive):
        RMF01_Post_Kahoot.xlsx → session='01', timepoint='post'
        RMF02_Pre_Kahoot.xlsx   → session='02', timepoint='pre'
    """
    fname = fname.lower()
    tp = "pre" if "pre" in fname else "post" if "post" in fname else None
    if tp is None:
        raise ValueError(f"Cannot determine pre/post from '{fname}'")
    m = re.search(r"rmf(\d{2})", fname)
    if not m:
        raise ValueError(f"Cannot extract session number from '{fname}'")
    return m.group(1), tp

# -----------------------------------------------------------------
# 3️⃣  OPTIONAL – simple topic extraction from question text
# -----------------------------------------------------------------
def extract_topic(question: str) -> str:
    """Very small keyword‑based mapping – extend if you need finer‑grained groups."""
    topic_map = {
        "security categorization": "Security Categorization",
        "rmf roles": "RMF Roles",
        "fips 199": "FIPS‑199",
        "system description": "System Description",
    }
    q_low = question.lower()
    for kw, t in topic_map.items():
        if kw in q_low:
            return t
    return "Other"

# -----------------------------------------------------------------
# 4️⃣  Validate a .xlsx file (quick zip‑header test)
# -----------------------------------------------------------------
def is_valid_xlsx(path: pathlib.Path) -> bool:
    """True if the file can be opened as a zip archive (the OOXML .xlsx format)."""
    try:
        with zipfile.ZipFile(path, "r") as _:
            return True
    except zipfile.BadZipFile:
        return False

# -----------------------------------------------------------------
# 5️⃣  Gather only the real Excel files we want to process
# -----------------------------------------------------------------
excel_paths = [
    p
    for p in RAW_DIR.rglob("*")
    if p.suffix.lower() in {".xlsx", ".xls"}               # keep only Excel extensions
    and "kahoot" in p.name.lower()                         # original filter
    and not p.name.startswith("~$")                        # skip Excel’s temporary lock files
    and (p.suffix.lower() != ".xlsx" or is_valid_xlsx(p)) # sanity‑check .xlsx files
]

if not excel_paths:
    raise FileNotFoundError(
        "❌  No Kahoot/Qualtrics Excel files were found. "
        "Make sure they live inside RMF_Raw_Data, contain the word "
        "'kahoot' (any case) in the filename, and have .xlsx/.xls extensions."
    )

# -----------------------------------------------------------------
# 6️⃣  Process each file (gracefully skip a broken workbook)
# -----------------------------------------------------------------
all_frames = []

print("\n🔍 Looking for Kahoot export files in:", RAW_DIR)
for p in RAW_DIR.rglob("*"):
    print("  •", p.name)

for excel_path in excel_paths:
    sess, tp = parse_filename(excel_path.name)

    # --------------------------------------------------------
    #   Load the sheet – with graceful fallback
    # --------------------------------------------------------
    try:
        if excel_path.suffix.lower() == ".xlsx":
            df = pd.read_excel(excel_path, engine="openpyxl")
        else:   # .xls
            df = pd.read_excel(excel_path, engine="xlrd")
    except Exception as e:
        print(f"⚠️  Skipping {excel_path.name}: {e}")
        continue

    # --------------------------------------------------------
    #   Tag the DataFrame
    # --------------------------------------------------------
    df["session"]      = sess
    df["timepoint"]    = tp
    df["session_name"] = f"Training {sess}"          # friendly label

    # --------------------------------------------------------
    #   Helper columns   (only modified part)
    # --------------------------------------------------------
    # 1️⃣  Boolean correct/incorrect – **nullable Boolean**.
    #    This lets us assign `pd.NA` later for Likert rows.
    if "Correct / Incorrect" in df.columns:
        df["is_correct"] = df["Correct / Incorrect"].str.lower() == "correct"
        df["is_correct"] = df["is_correct"].astype("boolean")   # nullable bool
    else:
        df["is_correct"] = pd.Series(pd.NA, dtype="boolean")

    # 2️⃣  Force numeric types (quietly coerce non‑numeric to NaN)
    for col in ["Score (points)", "Current Total Score (points)", "Answer Time (seconds)"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 3️⃣  Question number (strip any non‑numeric characters)
    if "Question Number" in df.columns:
        df["question_number"] = pd.to_numeric(
            df["Question Number"].astype(str).str.extract(r"(\d+)")[0],
            errors="coerce",
        )
    else:
        df["question_number"] = pd.NA

    # 4️⃣  Topic extraction (optional – used for knowledge‑by‑topic)
    if "Question" in df.columns:
        df["topic"] = df["Question"].astype(str).apply(extract_topic)
    else:
        df["topic"] = "Other"

    # 5️⃣  Likert‑scale mapping (only for rows where the Likert text appears)
    #     The survey you described uses four choices only.
    likert_map = {
        "strongly disagree": 1,
        "disagree": 2,
        "agree": 3,
        "strongly agree": 4,
    }
    if "Answer" in df.columns:
        df["likert_score"] = df["Answer"].astype(str).str.lower().map(likert_map)

        # ---- IMPORTANT ----
        # Any row that got a numeric likert_score is a **Likert** item,
        # not a knowledge (multiple‑choice) question.  Mark its `is_correct`
        # as missing so it is ignored by every knowledge‑only aggregation.
        df.loc[df["likert_score"].notna(), "is_correct"] = pd.NA
    else:
        df["likert_score"] = pd.NA

    # --------------------------------------------------------
    #   Keep only the columns we really need for later aggregation
    # --------------------------------------------------------
    # Add any participant‑metadata columns you need here.  If a column does not
    # exist in a particular export it will simply be ignored.
    keep = [
        "Player",                # participant identifier
        "Role",                  # risk‑management role
        "Team",                  # functional area / organization
        "RMF Experience",        # prior level of risk‑management experience
        "session",
        "session_name",
        "timepoint",
        "question_number",
        "topic",
        "Question",                # <-- KEEP THIS ONE!
        "Correct / Incorrect",
        "is_correct",
        "Score (points)",
        "Current Total Score (points)",
        "Answer Time (seconds)",
        "Answer",                  # original answer text (optional, for audits)
        "likert_score",
        # ---- optional free‑text fields (add if they exist) ----
        "Intended Application",   # e.g. “Will the training improve my work?”
        "Open Text Feedback",     # any open‑ended comment column
    ]
    df = df[[c for c in keep if c in df.columns]]
    all_frames.append(df)

# -----------------------------------------------------------------
# 7️⃣  Concatenate everything into ONE big DataFrame
# -----------------------------------------------------------------
big_df = pd.concat(all_frames, ignore_index=True)

# ---------------------------------------------------------
# 1️⃣  Dump every unique question text to a file for inspection
# ---------------------------------------------------------
question_list_path = OUT_DIR / "question_list.txt"
unique_questions = big_df["Question"].dropna().unique()
unique_questions_sorted = sorted(unique_questions, key=lambda x: x.lower())

with open(question_list_path, "w", encoding="utf-8") as f:
    for i, q in enumerate(unique_questions_sorted, 1):
        f.write(f"{i:03d}: {q}\n\n")
print(f"📝  All distinct question texts written to → {question_list_path}")

# -----------------------------------------------------------------
# 8️⃣  Summarise knowledge per participant / session / timepoint
# -----------------------------------------------------------------
def summarise_knowledge(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregates **knowledge (multiple‑choice) rows only** – because `is_correct`
    is NA for Likert items, `count()` automatically ignores them.
    """
    agg = (
        df.groupby(["Player", "session", "session_name", "timepoint"], as_index=False)
        .agg(
            final_score=("Current Total Score (points)", "max"),
            total_questions=("is_correct", "count"),       # counts only non‑NA (knowledge rows)
            num_correct=("is_correct", "sum"),            # NA treated as 0
            avg_answer_time_sec=("Answer Time (seconds)", "mean"),
        )
    )
    agg["pct_correct"] = (agg["num_correct"] / agg["total_questions"] * 100).round(2)
    cols = [
        "Player",
        "session",
        "session_name",
        "timepoint",
        "final_score",
        "total_questions",
        "num_correct",
        "pct_correct",
        "avg_answer_time_sec",
    ]
    return agg[cols]

knowledge_summary = summarise_knowledge(big_df)

# -----------------------------------------------------------------
# 9️⃣  Create a wide table (before‑training vs after‑training side‑by‑side)
# -----------------------------------------------------------------
wide = (
    knowledge_summary.pivot_table(
        index=["Player", "session", "session_name"],
        columns="timepoint",
        values=[
            "final_score",
            "total_questions",
            "num_correct",
            "pct_correct",
            "avg_answer_time_sec",
        ],
        aggfunc="first",
    )
    .reset_index()
)

wide.columns = [f"{val}_{tp}" if tp else val for val, tp in wide.columns]

# -----------------------------------------------------------------
# 🔟  Add derived columns to the wide table
# -----------------------------------------------------------------
wide["score_change"] = wide["final_score_post"] - wide["final_score_pre"]
wide["improved_flag"] = (wide["score_change"] > 0).astype(int)

wide["percentile_before"] = (
    wide["final_score_pre"].rank(pct=True, method="average") * 100
).round(2)
wide["percentile_after"] = (
    wide["final_score_post"].rank(pct=True, method="average") * 100
).round(2)

# -----------------------------------------------------------------
# 1️⃣0️⃣  PROJECT‑LEVEL METRICS – build a summary table
# -----------------------------------------------------------------
metric_rows = []

# ---------------------------------------------------------
# 10.1  Knowledge (score) metrics
# ---------------------------------------------------------
avg_before_score = wide["final_score_pre"].mean()
avg_after_score  = wide["final_score_post"].mean()
avg_improvement  = (avg_after_score - avg_before_score).round(2)
pct_improved     = wide["improved_flag"].mean() * 100

metric_rows.append({"Metric": "Average knowledge score before training",   "Value": round(avg_before_score, 2)})
metric_rows.append({"Metric": "Average knowledge score after training",    "Value": round(avg_after_score, 2)})
metric_rows.append({"Metric": "Average percentage‑point improvement",      "Value": round(avg_improvement, 2)})
metric_rows.append({"Metric": "Percentage of participants whose scores improved", "Value": round(pct_improved, 2)})

# ---------------------------------------------------------
# 10.2  Target‑proficiency metric (now 6 000 points)
# ---------------------------------------------------------
TARGET_SCORE = 6000          # <<< UPDATED TO 6 000
wide["meets_target_before"] = (wide["final_score_pre"]  >= TARGET_SCORE).astype(int)
wide["meets_target_after"]  = (wide["final_score_post"] >= TARGET_SCORE).astype(int)

pct_target_before = wide["meets_target_before"].mean() * 100
pct_target_after  = wide["meets_target_after"].mean() * 100

metric_rows.append({"Metric": "Percentage of participants meeting target proficiency before training", "Value": round(pct_target_before, 2)})
metric_rows.append({"Metric": "Percentage of participants meeting target proficiency after training",  "Value": round(pct_target_after, 2)})

# ---------------------------------------------------------
# 10.3  Knowledge improvement **by topic**
# ---------------------------------------------------------
if "topic" in big_df.columns:
    def pct_correct_by_topic(df_subset):
        return (
            df_subset.groupby("topic")["is_correct"]
            .mean()
            .multiply(100)
            .round(2)
            .reset_index(name="pct_correct")
        )

    topic_before = pct_correct_by_topic(big_df[big_df["timepoint"] == "pre"])
    topic_after  = pct_correct_by_topic(big_df[big_df["timepoint"] == "post"])
    topic_merge = pd.merge(topic_before, topic_after, on="topic", suffixes=("_before", "_after"))
    topic_merge["improvement"] = (topic_merge["pct_correct_after"] - topic_merge["pct_correct_before"]).round(2)

    for _, row in topic_merge.iterrows():
        metric_rows.append({"Metric": f"Knowledge percentage correct – {row['topic']} before training", "Value": row["pct_correct_before"]})
        metric_rows.append({"Metric": f"Knowledge percentage correct – {row['topic']} after training",  "Value": row["pct_correct_after"]})
        metric_rows.append({"Metric": f"Knowledge improvement – {row['topic']}",                     "Value": row["improvement"]})

# ---------------------------------------------------------
# 10.4  Likert‑scale helper – tolerant matching
# ---------------------------------------------------------
def likert_metrics(df: pd.DataFrame,
                   keywords: list[str],
                   friendly_name: str) -> pd.Series | None:
    """
    df               – the long‑format data (big_df)
    keywords         – list of words/phrases that should appear in the Question.
                       The row matches when any keyword is present.
    friendly_name    – label that will appear in the final metrics CSV.
    Returns the Series of after‑training numeric likert scores (or None).
    """
    mask = contains_any(df["Question"], keywords)
    subset = df[mask].copy()

    if subset.empty:
        print(f"⚠️  No rows found for Likert question containing any of {keywords}. Skipping.")
        return None

    subset = subset.dropna(subset=["likert_score"])

    before = subset[subset["timepoint"] == "pre"]
    after  = subset[subset["timepoint"] == "post"]

    avg_before = before["likert_score"].mean()
    avg_after  = after["likert_score"].mean()
    pct_increase = ((avg_after - avg_before) / avg_before * 100) if avg_before else None

    high_before = before["likert_score"].isin([3, 4]).mean() * 100   # Agree / Strongly Agree
    high_after  = after["likert_score"].isin([3, 4]).mean() * 100

    metric_rows.append({"Metric": f"Average {friendly_name} before training",   "Value": round(avg_before, 2) if pd.notna(avg_before) else None})
    metric_rows.append({"Metric": f"Average {friendly_name} after training",    "Value": round(avg_after, 2) if pd.notna(avg_after) else None})
    metric_rows.append({"Metric": f"Percent increase in average {friendly_name}", "Value": round(pct_increase, 2) if pct_increase is not None else None})
    metric_rows.append({"Metric": f"Percent of participants rating {friendly_name} as Agree or Strongly Agree before training", "Value": round(high_before, 2)})
    metric_rows.append({"Metric": f"Percent of participants rating {friendly_name} as Agree or Strongly Agree after training",  "Value": round(high_after, 2)})

    return after["likert_score"]

# -----------------------------------------------------------------
# 10.5  Role‑clarity (optional – keep if you have such a question)
# -----------------------------------------------------------------
role_clarity_keywords = ["role", "responsibilities"]
likert_metrics(big_df, role_clarity_keywords, "role clarity")

# -----------------------------------------------------------------
# 10.6  Confidence – three statements you gave
# -----------------------------------------------------------------
confidence_questions = [
    ["map security", "privacy requirements", "NIST 800‑53"],          # mapping confidence
    ["detail/describe the key RMF roles", "Authorizing Official"], # role‑clarity confidence
    ["compose each section", "system description"]                  # system‑description confidence
]

conf_post_scores = []
for kwlist in confidence_questions:
    scores = likert_metrics(big_df, kwlist, "confidence")
    if scores is not None:
        conf_post_scores.append(scores)

# -----------------------------------------------------------------
# 10.7  Practical application (training will improve work)
# -----------------------------------------------------------------
practical_keywords = ["improve"]
likert_metrics(big_df, practical_keywords, "practical application")

# -----------------------------------------------------------------
# 10.8  Training usefulness / relevance
# -----------------------------------------------------------------
usefulness_keywords = ["useful"]
likert_metrics(big_df, usefulness_keywords, "training usefulness")

# -----------------------------------------------------------------
# 10.9  Confidence target score – compute a realistic benchmark
# -----------------------------------------------------------------
if conf_post_scores:
    all_conf_post = pd.concat(conf_post_scores)
    avg_conf_post = all_conf_post.mean()
    target_conf = (math.ceil(avg_conf_post * 2) / 2)   # round up to nearest 0.5
    target_conf = round(target_conf, 2)

    metric_rows.append({
        "Metric": "Target confidence score after training",
        "Value": float(target_conf)
    })

    # Keep only the three confidence questions for target‑confidence calculation
    conf_long = big_df[
        big_df["Question"].apply(
            lambda q: any(
                any(kw.lower() in (q or "").lower() for kw in snippet)
                for snippet in confidence_questions
            )
        )
    ].copy()
    conf_long = conf_long.dropna(subset=["likert_score"])

    conf_avg_per_part = (
        conf_long[conf_long["timepoint"] == "post"]
        .groupby("Player")["likert_score"]
        .mean()
        .reset_index(name="avg_conf_post")
    )
    conf_avg_per_part["meets_target"] = (conf_avg_per_part["avg_conf_post"] >= target_conf).astype(int)

    pct_meet_conf = conf_avg_per_part["meets_target"].mean() * 100
    metric_rows.append({
        "Metric": "Percent of participants meeting the confidence target after training",
        "Value": round(pct_meet_conf, 2)
    })
else:
    print("⚠️  No confidence items were found – cannot compute confidence target.")

# -----------------------------------------------------------------
# 10.10  Lowest‑scoring knowledge questions (knowledge gaps)
# -----------------------------------------------------------------
question_stats = (
    big_df[big_df["is_correct"].notna()]                 # knowledge only
    .groupby("Question")
    .apply(lambda d: (d["is_correct"].sum() / d["is_correct"].count()) * 100)
    .reset_index(name="pct_correct")
    .sort_values("pct_correct")
)

bottom_n = 5
for _, row in question_stats.head(bottom_n).iterrows():
    metric_rows.append({
        "Metric": f"Lowest‑scoring question – '{row['Question']}'",
        "Value": round(row["pct_correct"], 2)
    })

# -----------------------------------------------------------------
# 10.11  Domain difficulty – which topics are easy / hard?
# -----------------------------------------------------------------
if "topic" in big_df.columns:
    domain_stats = (
        big_df[big_df["is_correct"].notna()]               # knowledge only
        .groupby("topic")["is_correct"]
        .mean()
        .multiply(100)
        .round(2)
        .reset_index(name="pct_correct")
        .sort_values("pct_correct")
    )

    for _, row in domain_stats.iterrows():
        metric_rows.append({"Metric": f"Domain difficulty – {row['topic']}", "Value": row["pct_correct"]})

    hardest = domain_stats.head(3)
    easiest = domain_stats.tail(3)

    for _, row in hardest.iterrows():
        metric_rows.append({"Metric": f"Hardest domain – {row['topic']}", "Value": row["pct_correct"]})
    for _, row in easiest.iterrows():
        metric_rows.append({"Metric": f"Easiest domain – {row['topic']}", "Value": row["pct_correct"]})

# -----------------------------------------------------------------
# 10.12  Hardest / easiest knowledge question per test (session + before/after)
# -----------------------------------------------------------------
test_question_stats = (
    big_df[big_df["is_correct"].notna()]                 # knowledge only
    .groupby(["session", "timepoint", "Question"])
    .apply(lambda d: (d["is_correct"].sum() / d["is_correct"].count()) * 100)
    .reset_index(name="pct_correct")
)

for (sess, tp), sub in test_question_stats.groupby(["session", "timepoint"]):
    hardest_row = sub.loc[sub["pct_correct"].idxmin()]
    metric_rows.append({
        "Metric": f"Hardest knowledge question – session {sess} ({tp})",
        "Value": f"\"{hardest_row['Question']}\" ({hardest_row['pct_correct']:.2f}%)"
    })
    easiest_row = sub.loc[sub["pct_correct"].idxmax()]
    metric_rows.append({
        "Metric": f"Easiest knowledge question – session {sess} ({tp})",
        "Value": f"\"{easiest_row['Question']}\" ({easiest_row['pct_correct']:.2f}%)"
    })

# -----------------------------------------------------------------
# 10.13  Percent of participants whose role‑clarity rating improved
# -----------------------------------------------------------------
if role_clarity_keywords:
    role_df = big_df[
        big_df["Question"].str.contains("role", case=False, na=False)
    ].copy()
    role_df = role_df.dropna(subset=["likert_score"])

    role_before = role_df[role_df["timepoint"] == "pre"]
    role_after  = role_df[role_df["timepoint"] == "post"]
    role_merge = pd.merge(
        role_before[["Player", "likert_score"]].rename(columns={"likert_score": "before"}),
        role_after[["Player", "likert_score"]].rename(columns={"likert_score": "after"}),
        on="Player",
        how="inner"
    )
    role_merge["improved"] = role_merge["after"] > role_merge["before"]
    pct_role_improved = role_merge["improved"].mean() * 100
    metric_rows.append({
        "Metric": "Percent of participants whose role clarity rating improved",
        "Value": round(pct_role_improved, 2)
    })

# -----------------------------------------------------------------
# 10.14  Average confidence change per participant
# -----------------------------------------------------------------
if conf_post_scores:
    conf_long = big_df[
        big_df["Question"].apply(
            lambda q: any(
                any(kw.lower() in (q or "").lower() for kw in snippet)
                for snippet in confidence_questions
            )
        )
    ].copy()
    conf_long = conf_long.dropna(subset=["likert_score"])

    conf_avg = (
        conf_long
        .groupby(["Player", "timepoint"])["likert_score"]
        .mean()
        .reset_index()
    )
    conf_wide = conf_avg.pivot(index="Player", columns="timepoint", values="likert_score")
    conf_wide["change"] = conf_wide["post"] - conf_wide["pre"]
    avg_conf_change = conf_wide["change"].mean()
    metric_rows.append({
        "Metric": "Average confidence change (after – before)",
        "Value": round(avg_conf_change, 2)
    })

# -----------------------------------------------------------------
# 10.15  Practical‑application free‑text metric (optional)
# -----------------------------------------------------------------
if "Planned Change" in big_df.columns:
    planned = big_df[
        (big_df["timepoint"] == "post") &
        big_df["Question"].str.contains("improve", case=False, na=False)
    ].copy()
    planned = planned.dropna(subset=["Planned Change"])
    pct_planned = (planned["Planned Change"].str.strip().astype(bool).mean()) * 100
    metric_rows.append({
        "Metric": "Percent of participants who identified at least one planned change",
        "Value": round(pct_planned, 2)
    })

# -----------------------------------------------------------------
# 10.16  New‑learning Likert item (if present)
# -----------------------------------------------------------------
new_learning_keywords = ["learned something new", "new information"]
likert_metrics(big_df, new_learning_keywords, "new learning")

# -----------------------------------------------------------------
# 10.17  Training‑recommendation Likert item (if present)
# -----------------------------------------------------------------
recommend_keywords = ["recommend", "would recommend"]
likert_metrics(big_df, recommend_keywords, "training recommendation")

# -----------------------------------------------------------------
# 10.18  Percent of participants meeting target proficiency per domain
# -----------------------------------------------------------------
target_proficiency = 80   # you can change this if you wish
if "topic" in big_df.columns:
    domain_gap = (
        big_df[big_df["is_correct"].notna()]
        .groupby(["topic", "Player"])["is_correct"]
        .mean()
        .reset_index(name="pct_correct_per_part")
        .groupby("topic")
        .agg(
            avg_pct_correct=("pct_correct_per_part", "mean"),
            participants_meeting_target=("pct_correct_per_part",
                                         lambda s: (s * 100 >= target_proficiency).mean() * 100)
        )
        .reset_index()
    )
    for _, row in domain_gap.iterrows():
        metric_rows.append({
            "Metric": f"Percent of participants meeting target ({target_proficiency} %) in domain {row['topic']}",
            "Value": round(row["participants_meeting_target"], 2)
        })

# -----------------------------------------------------------------
# 10.19  Participant‑level summary CSV (handles optional free‑text columns)
# -----------------------------------------------------------------
optional_text_cols = [c for c in ["Intended Application", "Open Text Feedback"] if c in big_df.columns]

agg_dict = {
    "final_score": ("Current Total Score (points)", "max"),
    "role_clarity": ("likert_score", lambda s: s.mean() if s.name.lower().find("role") != -1 else None),
    "confidence": ("likert_score", lambda s: s.mean()
                   if any(k in s.name.lower() for k in ["confidence", "map security", "rmf roles", "system description"])
                   else None),
    "practical": ("likert_score", lambda s: s.mean() if "improve" in s.name.lower() else None),
    "usefulness": ("likert_score", lambda s: s.mean() if "useful" in s.name.lower() else None),
    "new_learning": ("likert_score", lambda s: s.mean()
                     if "learned something new" in s.name.lower() else None),
    "recommendation": ("likert_score", lambda s: s.mean()
                       if "recommend" in s.name.lower() else None),
}
# add any optional free‑text columns
for txt_col in optional_text_cols:
    agg_dict[txt_col] = (txt_col, lambda s: s.dropna().iloc[0] if not s.dropna().empty else None)

participant_summary = (
    big_df
    .groupby(["Player", "session", "session_name", "timepoint"])
    .agg(**agg_dict)          # unpack the dict
    .reset_index()
)

# Pivot so before/after appear as separate columns
part_sum_wide = participant_summary.pivot(
    index=["Player", "session", "session_name"],
    columns="timepoint"
)

# Flatten MultiIndex column names
part_sum_wide.columns = [
    f"{metric}_{tp}" if tp else metric
    for metric, tp in part_sum_wide.columns
]

participant_path = OUT_DIR / "RMF_participant_summary.csv"
part_sum_wide.to_csv(participant_path, index=True)
print(f"📊  Participant‑level summary written to → {participant_path}")

# -----------------------------------------------------------------
# 10.20  Open‑text word‑frequency (optional)
# -----------------------------------------------------------------
open_text_col = "Open Text Feedback"   # change if you have a different column name
if open_text_col in big_df.columns:
    words = (
        big_df[open_text_col]
        .dropna()
        .str.lower()
        .str.replace(r"[^\w\s]+", "", regex=True)
        .str.split()
        .explode()
    )
    stop_words = set("""the and a an of to in is was were be been for on with at by from
                     will would could should it its that this these those as per""".split())
    freq = (
        words[~words.isin(stop_words)]
        .value_counts()
        .rename_axis("word")
        .reset_index(name="count")
        .head(20)       # top‑20 most common terms
    )
    open_text_path = OUT_DIR / "open_text_top_words.csv"
    freq.to_csv(open_text_path, index=False)
    print(f"🗒️  Top open‑text terms written to → {open_text_path}")

# -----------------------------------------------------------------
# 1️⃣1️⃣  Assemble the metric summary DataFrame & write to CSV
# -----------------------------------------------------------------
metrics_df = pd.DataFrame(metric_rows)

metrics_path = OUT_DIR / "RMF_metrics_summary.csv"
metrics_df.to_csv(metrics_path, index=False)

# -----------------------------------------------------------------
# 1️⃣2️⃣  Write the three standard CSV files (before, after, wide)
# -----------------------------------------------------------------
pre_path  = OUT_DIR / "RMF_pre_summary.csv"
post_path = OUT_DIR / "RMF_post_summary.csv"
wide_path = OUT_DIR / "RMF_wide_summary.csv"

pre_df  = knowledge_summary[knowledge_summary["timepoint"] == "pre"]
post_df = knowledge_summary[knowledge_summary["timepoint"] == "post"]

pre_df.to_csv(pre_path, index=False)
post_df.to_csv(post_path, index=False)
wide.to_csv(wide_path, index=False)

# -----------------------------------------------------------------
# 🎉  Done
# -----------------------------------------------------------------
print("\n✅  Finished! Files written:")
print(f"   Before‑training summary   → {pre_path}")
print(f"   After‑training summary    → {post_path}")
print(f"   Wide (before vs after)   → {wide_path}")
print(f"   Metric summary           → {metrics_path}")
print(f"   Participant summary      → {participant_path}")
if open_text_col in big_df.columns:
    print(f"   Open‑text top words      → {open_text_path}")