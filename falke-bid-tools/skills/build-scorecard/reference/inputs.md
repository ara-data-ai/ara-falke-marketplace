# Per-project inputs

Per-project inputs live in JSON/CLI, never in the shipped config. Copy the
blanks, fill them for the project, pass them on the CLI.

## Blank templates

In the bundled `engines/scorecard/examples/_templates/`:

- `baseline.template.json` → Section A modeled baseline lines (`--baseline`).
  The estimator's takeoff; the engine does NOT derive it from bidder numbers
  (circularity). `value` keys feed the QA fingerprint test.
- `qual_scores.template.json` → per-bidder category overrides (`--overrides`);
  the human-input checkpoint, needed at 100% coverage for the Overall curve.
- `aliases.template.json` → optional short display names for the board card
  (`--aliases`).
- `exclusions.template.json` → optional human set-aside ruling (`--exclusions`,
  or `--exclude "Name,Name"`).

## Filled validation examples (synthetic sample)

In the bundled `engines/scorecard/examples/`: `sample_baseline.json`,
`sample_qual_notes.json`, `sample_gold_overrides.json`, `sample_aliases.json`.
Use these as worked examples of a filled set (all firms/figures are fictional).

## Matrix format the engine expects

Counts/widths are DETECTED, not assumed, but the matrix must follow the
structural format: a bidder-name row; a `COST / COST SUBTOTALS / $/SF / $/SF
SUBTOTALS` sub-header quartet below it; per-division `... SUBTOTAL` rows; and a
`GRAND TOTAL CONSTRUCTION COST` row (the compared total — never the pre-markup
`CONSTRUCTION COST SUBTOTAL`).
