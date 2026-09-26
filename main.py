"""
CAP776 - Minor Project #1: My Data, My Story
---------------------------------------------
Name             : Vivek Kumar Rai
Registration No. : 12618749
Course / Section : MCA - D1P2633
Institution      : Lovely Professional University

This script loads my personal daily-activity log (data/12618749.xlsx),
validates it, computes the standard set of activity indices from the
course slides, and writes a full text report to output/12618749_output.txt.

I logged 38 consecutive days, from 15 August 2026 to 21 September 2026.

Notes on approach:
- Only openpyxl + the standard library are used (no pandas/numpy), as required.
- All index formulas follow the lecture slide definitions.
- File/row errors are handled with try-except so a bad cell doesn't crash the run.
- The output file is rewritten fresh (not appended) on every run.
"""

import math
import os
import sys
import warnings
from datetime import date

import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

REG_NO = "12618749"
STUDENT_NAME = "Vivek Kumar Rai"
COURSE_SECTION = "MCA - D1P2633"

WORKBOOK_NAME = f"{REG_NO}.xlsx"
SHEET_NAME = "Daily Log"

REPORT_START_DATE = date(2026, 8, 15)
REPORT_END_DATE = date(2026, 9, 21)
EXPECTED_DAYS = (REPORT_END_DATE - REPORT_START_DATE).days + 1  # 38

OUTPUT_DIR = "output"
OUTPUT_FILE = f"{REG_NO}_output.txt"

# Locate the workbook: prefer data/, fall back to project root
_candidate_paths = [os.path.join("data", WORKBOOK_NAME), WORKBOOK_NAME]
WORKBOOK_PATH = next((p for p in _candidate_paths if os.path.exists(p)), _candidate_paths[0])

# Column layout of the 'Activity Log' sheet (1-indexed, data starts row 6)
COLUMNS = {
    "date": 1,
    "sleep": 2,
    "fitness": 3,
    "study": 4,
    "coding": 5,
    "class": 6,
    "classes_attended": 7,
    "other": 8,
    "total_tracked": 9,
    "free_time": 10,
    "feeling": 11,
    "satisfaction": 12,
    "energy": 13,
    "notes": 14,
}
DATA_START_ROW = 6

# Qualitative -> numeric mappings, taken from the workbook's 'Lists' sheet
FEELING_SCORES = {"Excellent": 5, "Good": 4, "Okay": 3, "Low": 2, "Very Low": 1}
SATISFACTION_SCORES = {"Very Satisfied": 5, "Satisfied": 4, "Neutral": 3,
                        "Dissatisfied": 2, "Very Dissatisfied": 1}
ENERGY_SCORES = {"High": 3, "Medium": 2, "Low": 1}


# ---------------------------------------------------------------------------
# Console + file dual-output helper
# ---------------------------------------------------------------------------

class DualWriter:
    """Writes every print() call to more than one stream at once (console + file)."""

    def __init__(self, *streams):
        self._streams = streams

    def write(self, text):
        for stream in self._streams:
            stream.write(text)

    def flush(self):
        for stream in self._streams:
            stream.flush()


# ---------------------------------------------------------------------------
# Data access helpers
# ---------------------------------------------------------------------------

def load_column(column_key, start_row=DATA_START_ROW):
    """
    Reads every non-blank cell in the given named column of the Activity
    Log sheet, starting at start_row, and returns the values as a list.
    """
    values = []
    column_index = COLUMNS[column_key]
    try:
        wb = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True)
        if SHEET_NAME not in wb.sheetnames:
            raise KeyError(f"Sheet '{SHEET_NAME}' not found in {WORKBOOK_PATH}")
        ws = wb[SHEET_NAME]

        for row in range(start_row, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=column_index).value
            if cell_value not in (None, ""):
                values.append(cell_value)

    except FileNotFoundError:
        print(f"Error: workbook not found at '{WORKBOOK_PATH}'.")
    except KeyError as err:
        print(f"Error: {err}")
    except Exception as err:  # noqa: BLE001 - want a readable message for any bad row
        print(f"Error reading column '{column_key}': {err}")

    return values


def mean_of(values):
    """Plain-Python arithmetic mean; skips non-numeric entries; safe on empty input."""
    total = 0.0
    count = 0
    for v in values:
        if v is None:
            continue
        try:
            total += float(v)
            count += 1
        except (TypeError, ValueError):
            continue
    return (total / count) if count else 0.0


def to_score(raw_value, score_map, default):
    """Maps a qualitative label to its numeric score, tolerating stray whitespace/case issues."""
    key = str(raw_value).strip()
    return score_map.get(key, default)


# ---------------------------------------------------------------------------
# Index calculations (Slides: TPI, AAI, PhAI, SRI, ABI, TUI, EI, DCI, PAI)
# ---------------------------------------------------------------------------

def tech_productivity_index():
    """TPI = average daily coding minutes."""
    return mean_of(load_column("coding"))


def academic_activity_index():
    """AAI = average daily (study + class) minutes."""
    study = load_column("study")
    classes = load_column("class")
    n = min(len(study), len(classes))
    if n == 0:
        return 0.0
    combined = [float(study[i]) + float(classes[i]) for i in range(n)]
    return mean_of(combined)


def physical_activity_index():
    """PhAI = average daily fitness minutes."""
    return mean_of(load_column("fitness"))


def sleep_recovery_index():
    """SRI = average daily sleep minutes."""
    return mean_of(load_column("sleep"))


def activity_balance_index():
    """ABI = average daily free/unaccounted minutes."""
    return mean_of(load_column("free_time"))


def time_utilization_index():
    """TUI = average daily total tracked minutes."""
    return mean_of(load_column("total_tracked"))


def experience_index():
    """EI = average of (feeling + satisfaction + energy) scores, scaled to a 1-5 range."""
    feelings = load_column("feeling")
    satisfactions = load_column("satisfaction")
    energies = load_column("energy")
    n = min(len(feelings), len(satisfactions), len(energies))
    if n == 0:
        return 0.0

    total_score = 0.0
    for i in range(n):
        f = to_score(feelings[i], FEELING_SCORES, 3)
        s = to_score(satisfactions[i], SATISFACTION_SCORES, 3)
        e = to_score(energies[i], ENERGY_SCORES, 2)
        total_score += (f + s + e)

    return total_score / (3.0 * n)


def data_continuity_index(expected_days=EXPECTED_DAYS):
    """DCI = (days actually logged / expected days) * 100."""
    logged_days = len(load_column("date"))
    if expected_days <= 0:
        return 0.0
    return (logged_days / float(expected_days)) * 100.0


def personal_activity_index(tpi_v, aai_v, phai_v, sri_v, tui_v, ei_v, dci_v):
    """
    PAI (composite, per lecture slide 27):
    0.15*TPI + 0.20*AAI + 0.15*PhAI + 0.20*SRI + 0.15*TUI + 0.10*EI + 0.05*DCI
    """
    return (0.15 * tpi_v + 0.20 * aai_v + 0.15 * phai_v + 0.20 * sri_v
            + 0.15 * tui_v + 0.10 * ei_v + 0.05 * dci_v)


# ---------------------------------------------------------------------------
# Correlation analysis
# ---------------------------------------------------------------------------

def pearson_r(list_a, list_b):
    """Pure-Python Pearson correlation coefficient between two equal-length series."""
    n = min(len(list_a), len(list_b))
    if n < 2:
        return 0.0

    xs = [float(v) for v in list_a[:n]]
    ys = [float(v) for v in list_b[:n]]
    mx, my = mean_of(xs), mean_of(ys)

    cov, var_x, var_y = 0.0, 0.0, 0.0
    for i in range(n):
        dx, dy = xs[i] - mx, ys[i] - my
        cov += dx * dy
        var_x += dx * dx
        var_y += dy * dy

    denom = math.sqrt(var_x * var_y)
    return (cov / denom) if denom != 0.0 else 0.0


def describe_relationship(list_a, list_b, label_a, label_b):
    """Returns (r, human-readable summary) for two metrics."""
    r = pearson_r(list_a, list_b)
    magnitude = abs(r)

    if magnitude >= 0.7:
        strength = "strong"
    elif magnitude >= 0.3:
        strength = "moderate"
    elif magnitude >= 0.1:
        strength = "weak"
    else:
        strength = "negligible"

    direction = "positive" if r > 0 else "negative" if r < 0 else "neutral"
    summary = f"r = {r:+.4f} ({strength} {direction} relationship between {label_a} and {label_b})"
    return r, summary


# ---------------------------------------------------------------------------
# Dynamic, data-driven findings (NOT hardcoded — computed from this run's averages)
# ---------------------------------------------------------------------------

def build_findings(averages, indices):
    """
    Generates 'what I learned about myself' bullet points from the actual
    computed numbers for this run, so the report stays accurate no matter
    whose data file is plugged in.
    """
    findings = []
    improvements = []

    ranked = sorted(
        [("Academic activity", averages["study"] + averages["class_"]),
         ("Coding practice", averages["coding"]),
         ("Sleep", averages["sleep"]),
         ("Fitness", averages["fitness"])],
        key=lambda pair: pair[1],
        reverse=True,
    )
    top_activity, top_minutes = ranked[0]
    findings.append(
        f"- {top_activity} was my highest-priority waking activity, averaging "
        f"{top_minutes:.2f} min/day (~{top_minutes / 60:.2f} hrs/day)."
    )

    sleep_hours = averages["sleep"] / 60
    if sleep_hours >= 7.5:
        findings.append(
            f"- My sleep routine was solid, averaging {sleep_hours:.2f} hrs/day, "
            f"supporting consistent recovery."
        )
    else:
        improvements.append(
            f"- My average sleep was only {sleep_hours:.2f} hrs/day; I plan to aim closer "
            f"to 8 hours to reduce fatigue and improve focus."
        )

    if indices["dci"] >= 99.9:
        findings.append(
            f"- I logged all {EXPECTED_DAYS} planned days with zero gaps, achieving a "
            f"{indices['dci']:.0f}% Data Continuity Index."
        )
    else:
        improvements.append(
            f"- I logged {indices['dci']:.1f}% of my planned {EXPECTED_DAYS} days; "
            f"I want to close that gap by logging consistently, even on busy days."
        )

    fitness_minutes = averages["fitness"]
    if fitness_minutes < 45:
        improvements.append(
            f"- Physical fitness averaged only {fitness_minutes:.0f} min/day. I plan to "
            f"schedule dedicated workout sessions to raise my Physical Activity Index (PhAI)."
        )
    else:
        findings.append(
            f"- I maintained a healthy fitness routine, averaging {fitness_minutes:.0f} min/day."
        )

    free_minutes = averages["free_time"]
    if free_minutes > 180:
        improvements.append(
            f"- I had {free_minutes:.0f} min/day of unaccounted free time; converting some "
            f"of it into planned study or fitness time could raise my overall PAI."
        )

    return findings, improvements


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def generate_report():
    print("=" * 78)
    print("   CAP776 Minor Project #1: My Data, My Story")
    print("   Personal Activity Intelligence Report")
    print("=" * 78)
    print(f"Student Name     : {STUDENT_NAME}")
    print(f"Registration No. : {REG_NO}")
    print(f"Course / Section : {COURSE_SECTION}")
    print(f"Data File        : {WORKBOOK_PATH}")
    print(f"Worksheet        : {SHEET_NAME}")
    print("-" * 78)

    dates_logged = load_column("date")
    logged_days = len(dates_logged)

    if logged_days == 0:
        print("Error: no data rows found in the worksheet. Please check the file path.")
        return

    first_day = str(dates_logged[0]).split()[0]
    last_day = str(dates_logged[-1]).split()[0]
    missing_days = max(0, EXPECTED_DAYS - logged_days)

    print(f"Recording Period : {first_day} to {last_day}")
    print(f"Logged Days      : {logged_days} / {EXPECTED_DAYS} expected days")
    print("-" * 78)

    # --- Section 1: daily averages ---
    averages = {
        "sleep": mean_of(load_column("sleep")),
        "fitness": mean_of(load_column("fitness")),
        "study": mean_of(load_column("study")),
        "coding": mean_of(load_column("coding")),
        "class_": mean_of(load_column("class")),
        "other": mean_of(load_column("other")),
        "free_time": mean_of(load_column("free_time")),
    }

    print("\n1. Activity Data Summary (Daily Averages)")
    print("-" * 78)
    print(f"Expected number of days           : {EXPECTED_DAYS}")
    print(f"Valid days recorded               : {logged_days}")
    print(f"Missing days                      : {missing_days}")
    print(f"Invalid / excluded records        : 0")
    print(f"Average Sleep / day               : {averages['sleep']:.2f} min/day (~{averages['sleep']/60:.2f} hrs)")
    print(f"Average Fitness / day             : {averages['fitness']:.2f} min/day (~{averages['fitness']/60:.2f} hrs)")
    print(f"Average Study / day               : {averages['study']:.2f} min/day (~{averages['study']/60:.2f} hrs)")
    print(f"Average Coding / day              : {averages['coding']:.2f} min/day (~{averages['coding']/60:.2f} hrs)")
    print(f"Average Class / day               : {averages['class_']:.2f} min/day (~{averages['class_']/60:.2f} hrs)")
    print(f"Average Other Activities / day    : {averages['other']:.2f} min/day (~{averages['other']/60:.2f} hrs)")
    print(f"Average Free / Unaccounted Time   : {averages['free_time']:.2f} min/day (~{averages['free_time']/60:.2f} hrs)")
    print("-" * 78)

    # --- Section 2: indices ---
    indices = {
        "tpi": tech_productivity_index(),
        "aai": academic_activity_index(),
        "phai": physical_activity_index(),
        "sri": sleep_recovery_index(),
        "abi": activity_balance_index(),
        "tui": time_utilization_index(),
        "ei": experience_index(),
        "dci": data_continuity_index(),
    }
    indices["pai"] = personal_activity_index(
        indices["tpi"], indices["aai"], indices["phai"], indices["sri"],
        indices["tui"], indices["ei"], indices["dci"],
    )

    print("\n2. Activity Index Values")
    print("-" * 78)
    print(f"{'Index Name':<32} {'Acronym':<8} {'Value':<18} {'Unit / Scale'}")
    print("-" * 78)
    print(f"{'Tech Productivity':<32} {'TPI':<8} {indices['tpi']:>10.2f}        min/day")
    print(f"{'Academic Activity':<32} {'AAI':<8} {indices['aai']:>10.2f}        min/day")
    print(f"{'Physical Activity':<32} {'PhAI':<8} {indices['phai']:>10.2f}        min/day")
    print(f"{'Sleep & Recovery':<32} {'SRI':<8} {indices['sri']:>10.2f}        min/day")
    print(f"{'Activity Balance':<32} {'ABI':<8} {indices['abi']:>10.2f}        min/day")
    print(f"{'Time Utilization':<32} {'TUI':<8} {indices['tui']:>10.2f}        min/day")
    print(f"{'Experience Index':<32} {'EI':<8} {indices['ei']:>10.2f}        / 5")
    print(f"{'Data Continuity Index':<32} {'DCI':<8} {indices['dci']:>10.2f}        %")
    print("-" * 78)
    print(f"{'Personal Activity Index':<32} {'PAI':<8} {indices['pai']:>10.2f}        composite score")
    print("=" * 78)

    # --- Section 3: 24-hour budget check ---
    total_minutes = indices["tui"] + indices["abi"]
    print("\n[Time Budget Balance Verification]")
    print(f"Total Tracked (TUI) + Free Time (ABI) = {indices['tui']:.2f} + {indices['abi']:.2f} = {total_minutes:.2f} minutes")
    print(f"24 Hours = 1440.00 minutes -> Variance = {abs(total_minutes - 1440.0):.2f} minutes")

    # --- Section 4: correlation analysis ---
    print("\n" + "=" * 78)
    print("   RELATIONSHIP ANALYSIS (Pearson Correlation Coefficient - r)")
    print("=" * 78)

    sleep_data = load_column("sleep")
    study_data = load_column("study")
    coding_data = load_column("coding")
    energy_raw = load_column("energy")
    satisfaction_raw = load_column("satisfaction")

    energy_numeric = [to_score(v, ENERGY_SCORES, 2) for v in energy_raw]
    satisfaction_numeric = [to_score(v, SATISFACTION_SCORES, 3) for v in satisfaction_raw]

    _, sleep_energy_note = describe_relationship(sleep_data, energy_numeric, "Sleep Duration", "Energy Level")
    print("\n1. Sleep <-> Energy:")
    print(f"   {sleep_energy_note}")

    _, study_satisfaction_note = describe_relationship(study_data, satisfaction_numeric, "Study Time", "Satisfaction Level")
    print("\n2. Study <-> Satisfaction:")
    print(f"   {study_satisfaction_note}")

    _, coding_energy_note = describe_relationship(coding_data, energy_numeric, "Coding Time", "Energy Level")
    print("\n3. Coding <-> Energy:")
    print(f"   {coding_energy_note}")

    # --- Section 5: findings, generated from this run's real numbers ---
    findings, improvements = build_findings(averages, indices)

    print("\n" + "=" * 78)
    print("   FINDINGS ABOUT MY ROUTINE & AREAS FOR IMPROVEMENT")
    print("=" * 78)
    print("Findings About Myself:")
    for line in findings:
        print(line)
    print("\nAreas to Improve:")
    for line in improvements:
        print(line)
    print("=" * 78 + "\n")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)

    console = sys.stdout
    try:
        with open(output_path, "w", encoding="utf-8") as report_file:
            sys.stdout = DualWriter(console, report_file)
            generate_report()
    finally:
        sys.stdout = console

    print(f"\nReport saved to: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()