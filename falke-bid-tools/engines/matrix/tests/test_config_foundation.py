"""Matrix config-foundation unit tests (Boris).

Covers the loader/validation/matcher/gate built for the matrix generalization
sprint: known_firms.yaml loading + schema validation (C8), the firm matcher +
collision detection (C3), and the RunInputs identity validation + SF-basis gate
(M2). Mirrors the scorecard engine's test idioms (test_config.py /
test_cli_sf_gate.py): real YAML written to tmp_path, typed-error assertions, no
mocking.

Run from the engine root (engines/matrix/):
    python3 -m pytest tests/test_config_foundation.py -v
"""
from __future__ import annotations

import textwrap

import pytest

from src.config_errors import KnownFirmsConfigError, MissingParameterError, MatrixConfigError
from src.firm_config import (
    DEFAULT_KNOWN_FIRMS_PATH,
    load_known_firms,
)
from src.run_config import (
    SF_GATE_STOP,
    RunInputs,
    load_run_config,
    resolve_sf_basis,
)


def _write(tmp_path, text):
    p = tmp_path / "known_firms.yaml"
    p.write_text(textwrap.dedent(text), encoding="utf-8")
    return str(p)


# ---------------------------------------------------------------------------
# Shipped file loads + validates (sanity that §4 content is schema-valid)
# ---------------------------------------------------------------------------

def test_shipped_known_firms_loads_and_validates():
    cfg = load_known_firms(DEFAULT_KNOWN_FIRMS_PATH)
    ids = {f.firm_id for f in cfg.firms}
    assert ids == {"robmar", "pbs"}
    pbs = next(f for f in cfg.firms if f.firm_id == "pbs")
    assert pbs.match == ["principal builders"]  # NOT bare "pbs" (collision-safe)
    assert pbs.code_format_profile == "csi_1995_2digit"
    robmar = next(f for f in cfg.firms if f.firm_id == "robmar")
    assert len(robmar.reclassifications) == 2


# ---------------------------------------------------------------------------
# Loader / schema validation (C8)
# ---------------------------------------------------------------------------

def test_malformed_yaml_hard_stops(tmp_path):
    path = _write(tmp_path, "firms: [oops: : :\n")
    with pytest.raises(KnownFirmsConfigError, match="not valid YAML"):
        load_known_firms(path)


def test_missing_file_hard_stops():
    with pytest.raises(KnownFirmsConfigError, match="not found"):
        load_known_firms("/no/such/known_firms.yaml")


def test_bad_from_division_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: ["x"]
            reclassifications:
              - rule_id: R
                from: "DIV 99 00 00"
                to:   "DIV 09 00 00"
                when_description_contains_all: ["flooring"]
    """)
    with pytest.raises(KnownFirmsConfigError, match="`from`.*not a canonical"):
        load_known_firms(path)


def test_bad_to_division_fabrication_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: ["x"]
            reclassifications:
              - rule_id: R
                from: "DIV 13 00 00"
                to:   "DIV 99 00 00"
                when_description_contains_all: ["flooring"]
    """)
    with pytest.raises(KnownFirmsConfigError, match="fabricate a non-canonical"):
        load_known_firms(path)


def test_from_equals_to_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: ["x"]
            reclassifications:
              - rule_id: R
                from: "DIV 09 00 00"
                to:   "DIV 09 00 00"
                when_description_contains_all: ["flooring"]
    """)
    with pytest.raises(KnownFirmsConfigError, match="no-op reclass"):
        load_known_firms(path)


def test_empty_match_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: []
    """)
    with pytest.raises(KnownFirmsConfigError, match="`match` must be a non-empty"):
        load_known_firms(path)


def test_empty_keywords_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: ["x"]
            reclassifications:
              - rule_id: R
                from: "DIV 13 00 00"
                to:   "DIV 09 00 00"
                when_description_contains_all: []
    """)
    with pytest.raises(KnownFirmsConfigError, match="keyword guard"):
        load_known_firms(path)


def test_two_rule_cycle_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: ["x"]
            reclassifications:
              - rule_id: A
                from: "DIV 13 00 00"
                to:   "DIV 09 00 00"
                when_description_contains_all: ["flooring"]
              - rule_id: B
                from: "DIV 09 00 00"
                to:   "DIV 13 00 00"
                when_description_contains_all: ["tile"]
    """)
    with pytest.raises(KnownFirmsConfigError, match="cycle"):
        load_known_firms(path)


def test_duplicate_firm_id_hard_stops(tmp_path):
    path = _write(tmp_path, """
        firms:
          - firm_id: dup
            match: ["a"]
          - firm_id: dup
            match: ["b"]
    """)
    with pytest.raises(KnownFirmsConfigError, match="duplicate firm_id"):
        load_known_firms(path)


def test_validate_false_skips_deep_checks(tmp_path):
    # bad division code passes when validate=False (structure-only parse).
    path = _write(tmp_path, """
        firms:
          - firm_id: x
            match: ["x"]
            reclassifications:
              - rule_id: R
                from: "DIV 99 00 00"
                to:   "DIV 09 00 00"
                when_description_contains_all: ["flooring"]
    """)
    cfg = load_known_firms(path, validate=False)
    assert cfg.firms[0].firm_id == "x"


# ---------------------------------------------------------------------------
# Firm matcher + collision (C3)
# ---------------------------------------------------------------------------

def test_match_single_robmar():
    cfg = load_known_firms(DEFAULT_KNOWN_FIRMS_PATH)
    res = cfg.match("Robmar Construction LLC")
    assert not res.ambiguous
    assert res.firm is not None and res.firm.firm_id == "robmar"


def test_match_pbs_by_principal_builders():
    cfg = load_known_firms(DEFAULT_KNOWN_FIRMS_PATH)
    assert cfg.match("Principal Builders Solutions").firm.firm_id == "pbs"
    # bare "pbs"-substring names must NOT match (collision-safe term choice).
    assert cfg.match("Upbstate Builders Inc.").firm is None


def test_match_none_for_unknown_firm():
    cfg = load_known_firms(DEFAULT_KNOWN_FIRMS_PATH)
    res = cfg.match("Coastal Concrete Restoration LLC")
    assert not res.ambiguous and res.firm is None and res.matched_firm_ids == []


def test_collision_is_ambiguous_no_first_wins(tmp_path):
    # constructed collision fixture (GS-7): two entries both hit one name.
    path = _write(tmp_path, """
        firms:
          - firm_id: robmar
            match: ["robmar"]
          - firm_id: robmar_restoration
            match: ["robmar restoration"]
    """)
    cfg = load_known_firms(path)
    res = cfg.match("Robmar Restoration LLC")
    assert res.ambiguous
    assert res.firm is None  # no first-wins
    assert set(res.matched_firm_ids) == {"robmar", "robmar_restoration"}


# ---------------------------------------------------------------------------
# RunInputs identity validation (M1) + SF-basis gate (M2)
# ---------------------------------------------------------------------------

def test_run_inputs_valid():
    ri = RunInputs(project_name="P", project_address="A", gross_sf=10000.0)
    ri.validate()  # no raise


def test_missing_project_name_hard_stops():
    ri = RunInputs(project_name="", project_address="A", gross_sf=10000.0)
    with pytest.raises(MissingParameterError, match="project_name"):
        ri.validate()


def test_missing_gross_sf_hard_stops():
    ri = RunInputs(project_name="P", project_address="A", gross_sf=None)
    with pytest.raises(MissingParameterError, match="gross_sf"):
        ri.validate()


def test_nonpositive_gross_sf_hard_stops():
    ri = RunInputs(project_name="P", project_address="A", gross_sf=0.0)
    with pytest.raises(MatrixConfigError, match="must be positive"):
        ri.validate()


def test_load_run_config_overrides_win(tmp_path):
    p = tmp_path / "project.yaml"
    p.write_text("project_name: FromFile\nproject_address: A\ngross_sf: 5000\n")
    ri = load_run_config(str(p), overrides={"gross_sf": 9999, "sf_source": "explicit"})
    assert ri.project_name == "FromFile"
    assert ri.gross_sf == 9999
    assert ri.sf_source == "explicit"


def test_sf_gate_explicit():
    assert resolve_sf_basis(12345.0, False, None) == (12345.0, "explicit")


def test_sf_gate_confirmed_uses_extracted():
    assert resolve_sf_basis(None, True, 8000.0) == (8000.0, "matrix-confirmed")


def test_sf_gate_neither_stops_with_suggestion():
    val, msg = resolve_sf_basis(None, False, 8000.0)
    assert val == SF_GATE_STOP and "8,000 SF" in msg


def test_sf_gate_confirmed_but_no_gsf_stops():
    val, msg = resolve_sf_basis(None, True, None)
    assert val == SF_GATE_STOP and "explicitly" in msg
