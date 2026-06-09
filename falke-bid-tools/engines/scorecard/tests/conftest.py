"""Shared fixtures + SYNTHETIC sample-card ground truth.

All firms and figures here are fictional. The end-to-end validation matrix is a
client binary that is gitignored and absent from the shipped bundle, so the
integration tests that need it skip; the constants below describe the synthetic
sample card the suite validates against."""
import json
import os

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))            # .../tests
SKILL_ROOT = os.path.dirname(HERE)                            # .../scorecard
# The client Inputs/ folder lives under the scorecard PROJECT root (00_Scorecard),
# which is the PARENT of the skill package — NOT two levels up (that over-climbed
# to .../FALKE, so the matrix xlsx never resolved and every gold test silently
# SKIPPED — false green). Resolve absolutely against the project root.
PROJECT_ROOT = os.path.dirname(SKILL_ROOT)                    # .../00_Scorecard

SAMPLE_XLSX = os.path.join(
    PROJECT_ROOT, "Inputs",
    "Sample Condominium - Bid Comparison Matrix.xlsx")
BASELINE_JSON = os.path.join(SKILL_ROOT, "examples", "sample_baseline.json")
GOLD_OVERRIDES_JSON = os.path.join(
    SKILL_ROOT, "examples", "sample_gold_overrides.json")

# bidders excluded by a §1.4 set-aside ruling (applied via --exclude)
GOLD_EXCLUSIONS = ["Harbor Builders Inc.", "Borealis Builders Solutions"]

# display-name aliases (Marvin §1.5): the matrix carries full legal firm names
# but the sample board card uses short names. Keys are matched on the normalized
# raw/display name; the raw matrix name is retained in the run log for audit.
# (A dotted acronym is already handled by display_name normalization, so it is
# NOT listed here.)
GOLD_ALIASES = {
    "Acme Restoration": "Acme",
    "Granite Remodel Group": "Granite",
    "Harbor Builders Inc.": "Harbor",
    "Borealis Builders Solutions": "Borealis",
}

# sample-card published Overall (curve ON, 100% coverage; Cascade via +5 bonus)
GOLD_OVERALL = {
    "Acme": 84, "Borealis": 82, "Cascade": 75, "Dorne": 69,
    "Crest": 65, "Fjord": 56, "Granite": 51,
}

# ---- sample-card run parameters ----
SF_BASIS = 16000
BAND_LOW = 3.35
BAND_HIGH = 3.55
MID = 3.40
VARIANCE_MID = 3.45

# ---- sample-card grand-total row totals for the 7 KEPT bidders ----
# (drops Harbor, Borealis, duplicate Dorne; keeps Dorne J)
GOLD_TOTALS = {
    "Crest": 4400000,
    "Dorne": 3680000,
    "Fjord": 2080000,
    "Granite": 1950000,
    "Acme": 3360000,
    "Borealis": 3370000,
    "Cascade": 3050000,
}

# ---- sample-card $/SF (Marvin §3) ----
GOLD_PER_SF = {
    "Crest": 275,
    "Dorne": 230,
    "Borealis": 211,
    "Acme": 210,
    "Cascade": 191,
    "Fjord": 130,
    "Granite": 122,
}

# ---- sample-card tiers (Marvin §4.1) ----
GOLD_TIERS = {
    "Borealis": "TOP",
    "Acme": "TOP",
    "Cascade": "MID",
    "Fjord": "RISK",
    "Granite": "RISK",
    "Dorne": "DEFENSIVE",
    "Crest": "PREMIUM",
}

# ---- dropped bidders (must NOT appear in included field) ----
DROPPED = {"Harbor", "Mc Bride Builders", "Borealis Builders Solutions"}

# ---- sample card Overall ranking order (Marvin §9, curve ON) ----
GOLD_RANK_ORDER = ["Acme", "Borealis", "Cascade", "Dorne", "Crest", "Fjord", "Granite"]


@pytest.fixture(scope="session")
def sample_xlsx_available():
    return os.path.exists(SAMPLE_XLSX)


@pytest.fixture
def baseline_lines():
    with open(BASELINE_JSON, "r", encoding="utf-8") as fh:
        return json.load(fh)


@pytest.fixture
def run_overrides():
    """sf_basis + band overrides for the synthetic sample gold run."""
    return {
        "sf_basis": SF_BASIS,
        "band_low": BAND_LOW,
        "band_high": BAND_HIGH,
        "modeled_mid_takeoff": MID,
        "variance_mid": VARIANCE_MID,
    }
