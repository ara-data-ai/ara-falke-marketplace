---
name: scorecard
description: >
  Generate a Falke-branded board bid-comparison scorecard (PDF) from a
  bid-comparison matrix (.xlsx). TRIGGER on phrases like "create the Scorecard",
  "create the matrix scorecard", "build the scorecard", "regenerate the
  scorecard", or "refresh the scorecard" (case-insensitive) for a condo/HOA
  construction project. The matrix is normally a session-uploaded .xlsx (Claude
  Code: @path token from drag/drop or @-reference; Cowork:
  /sessions/<id>/mnt/uploads/<file>) — resolve the upload path per the rule in
  reference/runbook.md and ASK to confirm if ambiguous (do not guess
  most-recent). The SF basis is read PER RUN from THIS matrix and the user is
  asked to confirm it (--sf-confirmed) or override it (--sf-basis); the modeled
  baseline band (low/high/mid $M) is also required and confirmed. The render
  hard-stops (exit 2) without an SF decision and without --baseline-confirmed;
  this skill PROMPTS for what's missing. The skill also asks WHERE to save, and
  after the run it produces a plain-English Scorecard Summary alongside the
  scorecard, runs an audit (FAIL stops delivery), and offers to DRAFT (never
  auto-send) a submission email. Output is Falke-branded per Anna's template.
argument-hint: "[matrix.xlsx] [--project-name ...] [--sf-basis ...] [--band-low/high/mid ...]"
allowed-tools: Read, Bash(* -m scorecard.cli *), Bash(python3 -m scorecard.cli *), Bash(python3 -m pytest *)
---

# Falke Bid-Comparison Scorecard

Generate the Falke-branded board scorecard PDF by running the **scorecard
engine** (a Python package) over a bid matrix. Your job is to resolve the
uploaded matrix, gather the required parameters, run the engine, run the audit
step, and confirm the artifacts — not to recompute anything by hand.

## Engine location

The engine is the Python package bundled at:

```
${CLAUDE_PLUGIN_ROOT}/engines/scorecard
```

Run it as a module with the bootstrapped venv interpreter — see the exact
command in `reference/runbook.md`. This skill references that package as-is; it
does not duplicate it.

## Resolving the uploaded matrix

The matrix almost always arrives as a session upload. Resolve the exact path
per the **Upload Detection** rule in `reference/runbook.md` (Claude Code uses
the `@path` token from drag/drop or an `@`-reference; Cowork uses
`/sessions/<id>/mnt/uploads/<file>`). If the path is ambiguous — multiple
`.xlsx` files in the upload area, or no clear path — **stop and ask the user
to confirm the exact path**. Do not guess "most recent."

## When to STOP and ask

This is a board deliverable, so do **not** guess inputs. Stop and ask the user
if any of these is missing — the engine hard-stops anyway and this skill
**prompts for missing parameters** before invoking the engine:

- `--matrix` — path to the bid-comparison `.xlsx` (resolve per Upload Detection)
- `--project-name` — board title (never inherit another project's)
- **SF basis** — read PER RUN from THIS matrix, then verified with the user (see
  Step 3). Every matrix carries its OWN square-footage and it differs every job;
  never reuse a remembered/fixed value. Surface the detected number and confirm
  it before using it for $/SF. The render hard-stops (exit 2) without either
  `--sf-confirmed` (accept the matrix SF) or `--sf-basis <value>` (override).
- `--band-low` / `--band-high` / `--mid` — the modeled baseline band ($M)
- **baseline confirmation** — the cost baseline must be previewed and explicitly
  confirmed by the owner before any render (see Step 4); the final run hard-stops
  (exit 2) without `--baseline-confirmed`.
- **save location** — ask the user where to save the outputs (`--out-dir`) before
  rendering (see Step 5); do not assume a path.

If you have the matrix but not the parameters, ask for them; do not substitute
the matrix GSF or invent a band.

## Workflow

1. **Resolve the matrix.** Detect the uploaded file per the Upload Detection
   section in `reference/runbook.md`. If ambiguous, confirm with the user.
2. **Collect parameters.** Get the four required parameters from the user
   (prompt if missing). For per-project JSON inputs (baseline, qual scores,
   aliases), copy the blanks in `reference/inputs.md` and have the user fill
   them.
3. **Verify the SF basis from THIS matrix (REQUIRED).** Run the engine with
   `--preview-baseline` (it renders nothing and echoes the matrix-detected
   square-footage, labeled "SUGGESTED from matrix Row-10"). Read THAT number off
   THIS run — it is a per-run read of the submitted file, never a remembered or
   default value (every matrix has its own SF). Present it plainly and ask:
   *"The matrix lists **{N} SF**. I'll use that for the $/SF figures — is that
   correct, or should I use a different square-footage?"* Then **STOP for the
   answer.** If the user confirms → the final render uses `--sf-confirmed`. If the
   user gives a different number → the final render uses `--sf-basis <that>`. A
   render with neither hard-stops (exit 2). See `reference/runbook.md`.
4. **Confirm the baseline (REQUIRED).** The same `--preview-baseline` run also
   echoes the baseline — the trade-scope lines, the subtotal + OH&P, and the
   modeled band in both $ and $/SF — AND surfaces any baseline-anchoring
   fingerprint hits. Show the owner, with the honest framing: *"This baseline is
   the yardstick every bid is measured against. The tool detected that line X
   matches bidder Y's number within Z% — meaning the baseline may be partly
   derived from the bids rather than independent. Please confirm this baseline is
   correct, or tell me what to change."* Then **STOP and wait for the owner's
   explicit answer.** Do not proceed on silence or assumption.
   - **Confirmed** → the final render uses `--baseline-confirmed`.
   - **Changes** → the owner edits the baseline input (baseline JSON / trade lines
     / band), then re-run `--preview-baseline` and re-confirm. Loop until confirmed.
   The final run MUST include `--baseline-confirmed`; the engine hard-stops
   (exit 2) without it.
5. **Ask where to save (REQUIRED).** Before rendering, ASK the user where the
   outputs should go (the `--out-dir`) — do not assume a path.
6. **Run the engine.** Use the command in `reference/runbook.md` (include the
   SF decision from Step 3 — `--sf-confirmed` or `--sf-basis <value>` — and
   `--baseline-confirmed` from Step 4). For a first build/QA run, add `--refit`
   to re-fit the curves and print them vs the published modeling ranges.
7. **Read the run log.** The engine prints a RUN LOG; surface any QA-fingerprint
   hits, duplicate drops, or completeness flags to the user (these are board
   disclosure items, not auto-fails).
8. **Run the audit step.** Run the post-engine audit per the **Audit Step**
   section in `reference/runbook.md`. It returns PASS / PASS-WITH-WARNINGS /
   FAIL. **A FAIL stops the ship** — do not present the deliverable. Surface
   PASS-WITH-WARNINGS items to the user before handing off.
9. **Confirm the artifacts.** Check `--out-dir` for `scorecard.pdf` (Falke-
   branded per Anna's template), `scorecard.html`, the auto-produced
   `scorecard_summary.html` + `scorecard_summary.pdf` (the plain-English winner +
   why + caveats companion the Falke reviewer reads — **surface it to the user**),
   `scorecard_run.json`, and `audit_report.md`. Report the coverage state from
   `scorecard_run.json` (`full_coverage`, `overall_label`) — see the coverage
   rule in `reference/runbook.md`.
10. **Offer to draft the submission email (ALWAYS).** After the scorecard +
    summary are produced, ALWAYS offer to DRAFT an email for submitting the
    scorecard (e.g. to the board / client) — subject + body, pulling the winner
    and key points from `scorecard_run.json` / the Scorecard Summary — for the
    user to review and send. Do NOT auto-send: sending on the user's behalf needs
    explicit per-instance permission; drafting is fine.
11. **Hand to Floyd.** This is a solution deliverable; it goes through Floyd's
    gate before it ships. The `scorecard_run.json` plus `audit_report.md` are
    the audit trail.

## Reference (load only when needed)

- `reference/runbook.md` — the exact run command, the inputs, the outputs, the
  Upload Detection rule, the Audit Step, and the coverage/curve rule (the one
  source of truth for how to invoke).
- `reference/inputs.md` — the per-project JSON input templates and where to
  find the blanks and the synthetic sample validation examples.
- `reference/config.md` — the tunable config reference (what lives in
  `config/scorecard_config.yaml` and what is a CLI parameter).
- For the engine's own deep docs (modeling specs, package internals), see the
  bundled `engines/scorecard/` package and its docstrings.

## Eval

Before trusting a change to this skill, run the skill-behavior eval:
`eval/run_eval.sh` (see `eval/README.md`). It is separate from the 40 modeling
tests and checks that the skill triggers and produces the artifacts.
