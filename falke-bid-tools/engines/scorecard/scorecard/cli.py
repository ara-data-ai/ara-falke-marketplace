"""Command-line entry point for the scorecard skill.

Usage:
  # 1. PREVIEW the cost baseline (echoes it + runs the bid-anchoring check,
  #    renders NOTHING). Review with the owner, then re-run to render.
  python -m scorecard.cli --preview-baseline \
      --matrix "/path/<project> bid matrix.xlsx" \
      --project-name "<Project · Scope>" \
      --sf-basis <SF> --band-low <low $M> --band-high <high $M> --mid <mid $M> \
      --baseline path/to/baseline.json

  # 2. RENDER the scorecard (only after the baseline is confirmed).
  python -m scorecard.cli --baseline-confirmed \
      --matrix "/path/<project> bid matrix.xlsx" \
      --project-name "<Project · Scope>" \
      --sf-basis <SF> --band-low <low $M> --band-high <high $M> --mid <mid $M> \
      --baseline path/to/baseline.json \
      --out-dir Outputs --refit

REQUIRED each run: --matrix, --project-name, the band (--band-low/--band-high/
--mid), and a CONFIRMED SF basis (see the SF gate below). The band hard-stops
with MissingParameterError if omitted; --project-name has no default so a new
project can never silently inherit another project's name on a board deliverable.
See examples/sample_run.yaml for the synthetic validation values.

SF-BASIS SUGGEST-AND-CONFIRM GATE (relaxed from the old hard refusal): the skill
now READS the matrix's own Row-10 'TOTAL GSF' and offers it as a SUGGESTED
default — but it NEVER silently renders with it. A render REQUIRES one of:
  * --sf-basis <value>  — explicit override; use this value, no prompt; or
  * --sf-confirmed       — accept the matrix Row-10 GSF as the SF basis.
A render with NEITHER hard-stops (exit 2) with a message naming the matrix SF:
  "[STOP] SF basis not confirmed — the matrix reports <N> SF; re-run with
   --sf-basis <value> to override, or --sf-confirmed to accept the matrix SF."
$/SF is always computed against whichever SF is confirmed. --preview-baseline
echoes the matrix-detected SF (the suggested default) and renders nothing, so the
owner can confirm or override before any card is built.

BASELINE-CONFIRMATION GATE (REQUIRED, mirrors the SF gate): the modeled cost
baseline is the yardstick the whole scorecard measures against, and it can show
signs of being bid-derived. A render run therefore HARD-STOPS (exit 2) unless
--baseline-confirmed is passed. The intended flow is: run --preview-baseline,
review the echo + fingerprint check with the owner, then re-run with
--baseline-confirmed. --preview-baseline renders nothing and ignores
--baseline-confirmed.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from .config import load_config
from .errors import ScorecardError
from .modeling import refit_all
from .pipeline import audit_run, preview_baseline, render_summary, run_scorecard
from .render import build_context, render_html, render_pdf, write_html


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Falke bid-comparison scorecard skill")
    ap.add_argument("--matrix", required=True, help="path to bid-comparison xlsx")
    ap.add_argument("--config", default=None, help="path to scorecard_config.yaml")
    # REQUIRED parameters (also can live in config run_inputs)
    ap.add_argument("--sf-basis", type=float, default=None,
                    help="$/SF area basis — EXPLICIT OVERRIDE. When supplied it "
                         "is used as-is (no prompt). When omitted, the matrix's "
                         "Row-10 GSF is the suggested default and a render "
                         "requires --sf-confirmed (see the SF gate).")
    ap.add_argument("--sf-confirmed", action="store_true",
                    help="accept the matrix's Row-10 'TOTAL GSF' as the SF basis. "
                         "Required to render when --sf-basis is NOT supplied; "
                         "ignored when --sf-basis IS supplied (the explicit value "
                         "wins) and by --preview-baseline.")
    ap.add_argument("--band-low", type=float, default=None, help="band low $M")
    ap.add_argument("--band-high", type=float, default=None, help="band high $M")
    ap.add_argument("--mid", type=float, default=None,
                    help="modeled mid (takeoff) $M")
    ap.add_argument("--variance-mid", type=float, default=None,
                    help="Section C variance reference $M (default=band center)")
    # PRESENTATION labels (NOT modeled). Defaults to generic region labels.
    ap.add_argument("--region", default=None,
                    help='short region label for the cost-band chip + Section A '
                         'title, e.g. "South FL" (default "South FL")')
    ap.add_argument("--region-full", default=None,
                    help='long region label for the Section A title, e.g. '
                         '"South Florida" (default: the --region value, or '
                         '"South Florida" when region is "South FL")')
    ap.add_argument("--baseline-year", type=int, default=None,
                    help="baseline/pricing year shown in the Section A title "
                         "(default: the current year)")
    ap.add_argument("--baseline", default=None,
                    help="JSON file of Section A baseline lines (PARAMETER)")
    ap.add_argument("--qual-notes", default=None,
                    help="JSON file of per-bidder qualitative notes")
    ap.add_argument("--overrides", default=None,
                    help="JSON file of per-bidder category overrides")
    ap.add_argument("--exclude", default=None,
                    help="comma-separated bidder names to EXCLUDE from the scored "
                         "field per a human ruling (matched on normalized name), "
                         'e.g. --exclude "Harbor Builders Inc.,Borealis Builders '
                         'Solutions". Default = include all & flag (Marvin §1.4).')
    ap.add_argument("--exclusions", default=None,
                    help='JSON file with an exclusions list, either ["Name", ...] '
                         'or {"exclude": ["Name", ...]}. Merged with --exclude.')
    ap.add_argument("--aliases", default=None,
                    help='JSON file mapping raw/normalized firm name -> short '
                         'display name, e.g. {"Acme Restoration": "Acme"}. '
                         'Applied to the DISPLAYED bidder name; the raw matrix '
                         'name is retained in the run log for audit (Marvin §1.5).'
                         ' Merged over config["aliases"] (this file wins). '
                         'Default = no rename.')
    ap.add_argument("--project-name", required=True,
                    help="project title shown on the board scorecard (REQUIRED — "
                         "no default, so a new project can never silently inherit "
                         "another project's name on a board deliverable). "
                         'e.g. --project-name "Sample Condominium · Lobby Renovation"')
    ap.add_argument("--out-dir", default=".", help="output directory")
    ap.add_argument("--engine", default="chromium",
                    choices=["chromium", "auto", "weasyprint"],
                    help="PDF engine (default chromium — installed in Falke env; "
                         "weasyprint is an optional alternative)")
    ap.add_argument("--refit", action="store_true",
                    help="re-fit Section C + Overall curve with scipy and print "
                         "vs Darvish ranges (FIRST build step)")
    ap.add_argument("--html-only", action="store_true",
                    help="emit HTML, skip PDF (useful when no PDF engine)")
    ap.add_argument("--audit", dest="audit", action="store_true", default=True,
                    help="run the deterministic self-audit after artifacts "
                         "(DEFAULT ON); writes audit_report.md + audit.json")
    ap.add_argument("--no-audit", dest="audit", action="store_false",
                    help="skip the self-audit (not recommended for board runs)")
    ap.add_argument("--preview-baseline", action="store_true",
                    help="ECHO the supplied cost baseline (trade lines, "
                         "subtotals, OH&P, band in $M AND $/SF) + run the "
                         "bid-anchoring fingerprint check, then EXIT 0 WITHOUT "
                         "rendering a scorecard. Review with the owner first.")
    ap.add_argument("--baseline-confirmed", action="store_true",
                    help="REQUIRED to render (mirrors --sf-basis): confirms the "
                         "owner reviewed the baseline via --preview-baseline. "
                         "Without it a render run HARD-STOPS (exit 2). Ignored "
                         "by --preview-baseline.")
    args = ap.parse_args(argv)

    # ---- SF-BASIS SUGGEST-AND-CONFIRM GATE. Resolve the SF basis BEFORE config
    # so $/SF is always computed against a CONFIRMED value. Reads the matrix's
    # own Row-10 GSF as the suggested default; the gate then decides:
    #   * --sf-basis supplied  -> explicit override (used as-is);
    #   * --sf-confirmed (no --sf-basis) -> accept the matrix GSF;
    #   * preview mode          -> use explicit if given else the matrix GSF, and
    #                              surface the suggestion (renders nothing);
    #   * a RENDER with neither -> hard-stop (exit 2) naming the matrix SF.
    # The matrix GSF is detected with the SAME detector the full parse uses
    # (MatrixParser.detect_sf), so the suggested value matches what the audit
    # later sees.
    try:
        # validate=False: we only need the static matrix block to detect the
        # Row-10 GSF; sf_basis/band are not yet resolved here.
        cfg_probe = load_config(args.config, validate=False)
    except ScorecardError:
        cfg_probe = None
    matrix_gsf = None
    if args.sf_basis is None and cfg_probe is not None:
        try:
            from .matrix import MatrixParser
            _, matrix_gsf = MatrixParser(cfg_probe.block("matrix")).detect_sf(
                args.matrix)
        except ScorecardError as e:
            # a missing/unreadable matrix is reported the same way the parse path
            # would report it; the gate below still asks the user to act.
            print(f"[STOP] {e}", file=sys.stderr)
            return 2

    if args.sf_basis is not None:
        sf_basis, sf_source = args.sf_basis, "explicit"
    elif args.preview_baseline:
        # preview never blocks; show the matrix GSF as the suggested basis.
        sf_basis, sf_source = matrix_gsf, "matrix-confirmed"
    elif args.sf_confirmed:
        sf_basis, sf_source = matrix_gsf, "matrix-confirmed"
    else:
        # RENDER with neither an explicit basis nor confirmation -> suggest+stop.
        if matrix_gsf is None:
            print("[STOP] SF basis not confirmed and the matrix reports no "
                  "Row-10 'TOTAL GSF' to suggest — re-run with --sf-basis "
                  "<value> to set it explicitly.", file=sys.stderr)
        else:
            print(f"[STOP] SF basis not confirmed — the matrix reports "
                  f"{matrix_gsf:,.0f} SF; re-run with --sf-basis <value> to "
                  f"override, or --sf-confirmed to accept the matrix SF.",
                  file=sys.stderr)
        return 2

    # When confirming/previewing the matrix GSF but none was detected, there is
    # nothing to confirm — STOP rather than fall through to a None-basis config.
    if sf_basis is None:
        print("[STOP] --sf-confirmed given but the matrix reports no Row-10 "
              "'TOTAL GSF' to confirm — supply --sf-basis <value> explicitly.",
              file=sys.stderr)
        return 2

    overrides_inputs = {
        "sf_basis": sf_basis,
        "sf_source": sf_source,
        "band_low": args.band_low,
        "band_high": args.band_high,
        "modeled_mid_takeoff": args.mid,
        "variance_mid": args.variance_mid,
        "region": args.region,
        "region_full": args.region_full,
        "pricing_year": args.baseline_year,
    }

    try:
        cfg = load_config(args.config, overrides=overrides_inputs)
    except ScorecardError as e:
        print(f"[STOP] {e}", file=sys.stderr)
        return 2

    # ---- PREVIEW MODE: echo the baseline + run the fingerprint check, then
    # EXIT 0 without rendering. The owner SEES the yardstick (incl. the
    # matrix-suggested SF basis) before any card is built. (--baseline-confirmed
    # and --sf-confirmed are both ignored here.)
    if args.preview_baseline:
        try:
            preview = preview_baseline(
                args.matrix, cfg, baseline_lines=_load_json(args.baseline))
        except ScorecardError as e:
            print(f"[STOP] {e}", file=sys.stderr)
            return 2
        for line in preview["echo"]:
            print(line)
        print("\n(No scorecard rendered — review the baseline AND the SF basis "
              "with the owner, then re-run with --baseline-confirmed plus either "
              "--sf-basis <value> or --sf-confirmed to build the card.)")
        return 0

    # ---- BASELINE-CONFIRMATION GATE (REQUIRED to render; mirrors the SF gate).
    # The cost baseline is the yardstick the scorecard measures against and can
    # be bid-derived; it must be confirmed each run.
    if not args.baseline_confirmed:
        print("[STOP] Baseline not confirmed — run with --preview-baseline, "
              "review it with the owner, then re-run with --baseline-confirmed. "
              "(The cost baseline is the yardstick the scorecard measures "
              "against; it must be confirmed each run.)", file=sys.stderr)
        return 2

    if args.refit:
        print("=== CURVE RE-FIT (scipy.optimize.least_squares) ===")
        for rr in refit_all(cfg.run.variance_mid):
            print(f"\n[{rr.name}] params={rr.params}")
            print(f"  max|resid|={rr.max_abs_residual} mean|resid|={rr.mean_abs_residual}")
            print(f"  in_range={rr.in_range}")
            print(f"  note: {rr.notes}")
        print()

    baseline_lines = _load_json(args.baseline)
    qual_notes = _load_json(args.qual_notes)
    overrides = _load_json(args.overrides)
    exclude = _parse_exclusions(args.exclude, _load_json(args.exclusions))
    aliases = _load_json(args.aliases)

    try:
        result = run_scorecard(
            args.matrix, cfg,
            baseline_lines=baseline_lines,
            qualitative_notes=qual_notes,
            overrides=overrides,
            exclude=exclude,
            aliases=aliases,
            project_name=args.project_name,
        )
    except ScorecardError as e:
        print(f"[STOP] {e}", file=sys.stderr)
        return 2

    os.makedirs(args.out_dir, exist_ok=True)
    print("=== RUN LOG ===")
    for line in result["log"]:
        print("  " + line)

    ctx = build_context(result, cfg)
    html = render_html(ctx)
    base = os.path.join(args.out_dir, "scorecard")
    write_html(html, base + ".html")
    print(f"\nHTML -> {base}.html")
    if not args.html_only:
        try:
            render_pdf(html, base + ".pdf", engine=args.engine)
            print(f"PDF  -> {base}.pdf")
        except ScorecardError as e:
            print(f"[WARN] PDF render skipped: {e}", file=sys.stderr)

    # provenance JSON for the board / audit trail
    with open(base + "_run.json", "w", encoding="utf-8") as fh:
        json.dump({
            "run_id": result["meta"]["run_id"],
            "full_coverage": result["full_coverage"],
            "overall_label": result["overall_label"],
            "log": result["log"],
            "bidders": [{
                "name": b["name"], "rank": b["rank"], "total": b["total"],
                "per_sf": b["per_sf"], "tier": b["tier"],
                "overall": b["overall"],
            } for b in result["bidders"]],
        }, fh, indent=2, default=str)
    print(f"JSON -> {base}_run.json")

    # ---- Scorecard Summary companion (plain-English; matched-set, every run) ----
    try:
        summary_paths = render_summary(
            result, args.out_dir, engine=args.engine, html_only=args.html_only)
        print(f"SUMMARY -> {summary_paths['summary_html']}")
        if "summary_pdf" in summary_paths:
            print(f"SUMMARY -> {summary_paths['summary_pdf']}")
    except ScorecardError as e:
        print(f"[WARN] summary render skipped: {e}", file=sys.stderr)

    # ---- self-audit (Floyd-lite, every run; default ON) ----
    if args.audit:
        ar, paths = audit_run(result, cfg, args.out_dir, aliases=aliases)
        print(f"AUDIT -> {paths['report_md']}")
        print(f"AUDIT -> {paths['audit_json']}")
        print(f"\n=== SELF-AUDIT VERDICT: {ar.verdict} "
              f"({ar.counts['blocker']} blocker(s), {ar.counts['warn']} "
              f"warning(s), {ar.counts['info']} info) ===")
        if ar.verdict == "FAIL":
            print("[FAIL] Self-audit found a BLOCKER — do NOT deliver this "
                  "scorecard as final; remediate and re-run.", file=sys.stderr)
            return 1
    return 0


def _load_json(path):
    if not path:
        return None
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _parse_exclusions(exclude_csv, exclusions_json):
    """Merge --exclude (comma-separated) and --exclusions (JSON list or
    {"exclude": [...]}) into a de-duplicated list of names. Returns None when no
    exclusions are supplied (preserving the include-all default)."""
    names = []
    if exclude_csv:
        names.extend(n.strip() for n in str(exclude_csv).split(",") if n.strip())
    if exclusions_json:
        items = (exclusions_json.get("exclude")
                 if isinstance(exclusions_json, dict) else exclusions_json)
        if isinstance(items, list):
            names.extend(str(n).strip() for n in items if str(n).strip())
    # de-dup preserving order
    seen, out = set(), []
    for n in names:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out or None


if __name__ == "__main__":
    raise SystemExit(main())
