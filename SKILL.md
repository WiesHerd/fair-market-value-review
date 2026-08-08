---
name: fair-market-value-review
description: Guided fair market value (FMV) review for physician and APP compensation. Walks the user step-by-step — survey data, provider details, compensation components, productivity, proposal — then generates a filled review workbook, a DOCX/PDF packet, an email-ready HTML report, or a new-hire market range. Tracks every request in a register. Use when the user mentions FMV review, market review, compensation benchmarking, a salary adjustment request, a new-hire offer range, blended survey percentiles, or wants to know where a provider sits versus market.
---

# Fair Market Value Review — Guided Workflow

## What this is
A step-through workflow for a provider compensation manager or analyst. The agent runs
the intake like a valuation analyst would: asks for what it needs one item at a time,
validates each input as it arrives, shows every calculation before using it, and logs
the session so the review is reproducible.

**Two situations this covers:**
- **New hire** — a candidate with no offer yet. Output: a defensible market range.
- **Existing provider** — an incumbent up for a salary adjustment. Output: current vs
  proposed positioning, cost impact, and committee-ready deliverables.

---

## Start here: zero-state behavior

**If the user's first message is vague** — "fmv", "help", "how do I use this", "what does
this do", "/fmv", or they clearly just installed the skill — **do not ask them a question
yet. Teach them first.** Print this orientation (adjust lightly for tone), then stop and wait:

> **Fair Market Value Review — here's how this works.**
>
> I'll walk you through a market review of a provider's compensation, one question at a
> time. You don't need to prepare anything in advance — I'll tell you what I need when I
> need it, and you can paste values or hand me a file at every step.
>
> **Three things I can do:**
> 1. **New hire** — you have a candidate, no offer yet → I produce a defensible range.
> 2. **Salary adjustment** — existing provider up for a change → current-vs-proposed
>    positioning, cost impact, and a committee-ready packet.
> 3. **Quick benchmark** — "where does Dr. X sit right now?" → answer on screen, no files.
>
> **What I'll ask you for, in order:**
> `survey data → provider details → compensation components → productivity (wRVUs) → the proposal`
>
> **What you get:** a filled review workbook, a Word/PDF packet, an email-ready HTML
> summary, and a market-band chart — plus a logged entry in your request register.
>
> **Two things up front:** this is a market-approach analysis to inform a decision, not a
> certified FMV opinion. And I'll show you every calculation before I generate anything.
>
> **Which of the three do you want to start with?** (Or say "set up my files" if this is
> your first time and I should learn your spreadsheet layout.)

**Do not** dump the phase list, methodology, or command table on them. The orientation
above is the whole first response; everything else is revealed as they progress.

**If their first message already tells you the situation** ("I need a range for a PICU
candidate"), skip the orientation and go straight to the matching flow — but still say in
one line what you'll need and that you'll show your math.

---

## Commands (how a user starts)

No native slash commands — the user says these in plain language and you run the flow.

| User says | Flow | Produces |
|---|---|---|
| "run an FMV review" / "/fmv" | Full guided intake | Whatever they pick at the end |
| "new hire range" / "what should we offer" | Phases 1, 2, 6–8 (skip current comp) | Market range pack |
| "adjustment for [provider]" | Full intake | Workbook fill + review packet |
| "just the benchmarks" / "where does X sit" | Phases 1, 2, 4–6 | On-screen positioning, no files |
| "set up my files" / "first time" | **Setup flow** | Confirmed column mapping |
| "what's outstanding" / "show my queue" | Register lookup | List of open requests |
| "log this request" | Register entry | Tracked request ID |

Ask which situation applies if it isn't already obvious.

---

## Setup flow (first time with a new organization's files)

The most common failure is assuming someone else's spreadsheet layout. Never guess.

1. Ask the user to point at their salary roster workbook and their survey workbook.
2. Run the inspector on each — read-only, modifies nothing:
   ```bash
   python3 scripts/inspect_workbook.py path/to/roster.xlsx --guess-roster
   python3 scripts/inspect_workbook.py path/to/survey.xlsx --guess-survey
   ```
3. Show the detected sheet names, columns and sample rows. Ask them to confirm or correct
   — especially name, specialty, FTE, current salary and the percentile columns.
4. Record the confirmed mapping (sheet-name constants at the top of
   `scripts/committee_template_generator.py`, or carry it in the session).
5. Confirm the **benefits/fringe rate** by reading their template's own formula. Never assume.

Once mapped, the user never repeats this step.

---

## The guided intake

**Interaction rules — every time:**
- **One question at a time.** Never dump a list of required fields.
- **Echo back** what you understood after each answer, then move on.
- Accept **a file or a typed answer** at every step.
- If ambiguous, ask. If a sensible default exists (FTE 1.00), propose it explicitly and
  get a yes or no. Never silently guess.
- Show a **progress line**: `[✓ survey] [✓ provider] [→ compensation] [ ] productivity [ ] proposal`

### Phase 0 — Orient
Confirm new hire vs existing provider, then preview the shopping list so nothing surprises them.

### Phase 1 — Survey data
Ask for their survey workbook (any Excel with specialty benchmark rows), pasted benchmark
text, or a simple anchors JSON (p25/p50/p75/p90).

Validate immediately and report back: which sheet and row matched the specialty (watch the
dash pitfall), which sources carry data and their n-counts, and any suppressed sources
(n=0) excluded from the blend. If the specialty isn't present, say so and list the closest
matches rather than failing.

### Phase 2 — Provider
Single provider or a cohort?
- **Single:** name, specialty/subspecialty, years of experience, FTE, clinical FTE.
- **Cohort:** a TSV/Excel, or collect them one at a time.
- **If they have a roster file:** look the provider up, show what was found, and ask them
  to confirm it's the right person before continuing.

### Phase 3 — Compensation components *(skip for a new hire with no offer)*
Build TCC piece by piece, confirming each:
1. Current base salary
2. Planned base (post annual increase), if an increase cycle applies
3. Additional cash components — stipends (admin, call, program director), productivity
   incentive, quality/retention, other cash. **Ask what the organization calls these**;
   the workbook labels them Component 1–3.
4. Benefits/fringe rate — **read it from their template's formula.** Never assume.

Echo the TCC build-up as a small table before moving on.

### Phase 4 — Productivity
Ask for **total annual wRVUs** — the sum of the wRVU values across all services performed
during the measurement year (a production total, not a count of encounters) — plus any
governance-file wRVU percentile they already have. If the role isn't wRVU-measured, note
it and skip. Don't block.

### Phase 5 — The proposal
- **Existing provider:** proposed base, proposed components, review type, tracker number.
- **New hire:** no proposal needed — the output *is* the recommended range.

### Phase 6 — Confirm, with the math shown
Before generating anything, present in one place:
- every input, labeled **from file** / **user-provided** / **derived**
- the blended benchmark written out with real numbers:
  `blended p50 = (325,000×89 + 315,000×34 + 320,000×142) ÷ 265 = $321,038`
- current and proposed percentile positioning, labeled **directional** if interpolated
- the **compensation-to-productivity alignment check** and any flag it raises, with the
  documented business reason if compensation sits materially above production
- both cost bases: Δ vs current **and** Δ vs planned
- which deliverables are about to be produced

Get an explicit go-ahead.

### Phase 7 — Generate and verify
Run the matching command. Afterwards: the workbook filler self-verifies on re-open (6
structural checks) — report the result; present each file with a one-line summary; offer
the HTML email version if a packet was produced.

### Phase 8 — Log the session
Append to **`logs/fmv_review_log.md`** and write **`logs/fmv_review_<date>_<subject>.json`**.
Capture: timestamp, review type, subject, input files and sheets used, typed inputs,
blended values and the n behind them, positioning results and whether interpolated,
deliverables produced, verification result, and every caveat issued. Someone reading the
log later should be able to reproduce the review exactly. If the user asks to exclude
something, exclude it.

### Phase 9 — Register the request
Every review gets an entry so nothing lives only in an inbox:

```bash
# when the request arrives
python3 scripts/fmv_register.py add --provider "NAME" --specialty "SPECIALTY" \
  --type adjustment --requester "WHO ASKED" --due 2026-09-01

# after the analysis
python3 scripts/fmv_register.py update FMV-0001 --status in_review \
  --proposed-base 325000 --tcc-percentile 75.3 --wrvu-percentile 50.9 \
  --alignment comp_above_production --deliverables "review.xlsx; packet.pdf"

# when the committee decides
python3 scripts/fmv_register.py update FMV-0001 --status approved \
  --decision "Approved at $325,000 with documented call-burden rationale"
```

Ask "what's outstanding?" any time:
```bash
python3 scripts/fmv_register.py list --open
python3 scripts/fmv_register.py summary     # counts, overdue, alignment flags to document
```

Plain CSV at `register/fmv_requests.csv` — opens in Excel, diffs in git, no database.
Statuses: intake → awaiting_data → in_review → pending_committee → approved / declined /
withdrawn. **Offer to log the request at the start of any review**, and update it at the
end without being asked.

---

## Commands the phases run

**Fill the review workbook:**
```bash
python3 scripts/committee_template_generator.py \
  --name "Provider Name" --salary-file roster.xlsx --survey-file survey.xlsx \
  --template fmv_review_template.xlsx \
  --proposed-base 325000 --stipend 42000 --wrvu 4228 --track-num 100026 \
  --output review.xlsx
```
Looks up the provider, blends survey sources by sample size, writes the blended Survey row
plus interpolated-percentile **formulas**, preserves every template formula, and verifies
the saved file on re-open. Use `--benchmarks file.txt` instead of `--survey-file` for
pasted text. Handles `.dec` portal-encrypted files. `--wrvu` and `--track-num` are
required — there is no safe default for either.

**Review packet / HTML email / new-hire range / workbook inspection:**
```bash
python3 scripts/build_adjustment_report.py --config review_config.json
python3 scripts/build_html_email.py --config review_config.json
python3 scripts/build_cv_only_market_anchor.py --name "Candidate" \
  --specialty "Specialty" --yoe 8 --survey-file survey.xlsx --output-dir ./range_out
python3 scripts/inspect_workbook.py file.xlsx --guess-roster
```

---

## Methodology grounding (read before advising anyone)

This tool performs the **market approach** — positioning compensation against published
survey benchmarks. A defensible FMV determination weighs the market, income and cost
approaches together with the facts of the arrangement. Say this plainly when a user treats
a percentile as a conclusion.

1. **Alignment beats level.** No percentile is automatically FMV and none is automatically
   not. CMS has never set a threshold. The question is whether compensation aligns with
   productivity and scope. `scripts/fmv_analysis.py` computes this and flags gaps.
2. **The p75 myth.** "Under the 75th is safe" is wrong in both directions. 25% of providers
   sit at or above p75 by definition; paying p75 to someone producing at p25 is a real
   misalignment regardless of the label.
3. **Above p90 is not automatically impermissible.** 10% of providers are there. Stacked
   arrangements can support it — document each component separately *and* in aggregate.
4. **Surveys compute each metric independently.** A p90 wRVU producer is not entitled to
   p90 comp-per-wRVU. Market comp-per-wRVU is relatively flat across quartiles; applying a
   high-quartile ratio to high production double-counts productivity.
5. **Blending has a cost.** Multiple independent surveys is prudent practice, but providers
   responding to more than one survey get double-counted. Disclose it.
6. **Survey definitions differ** and surveys publish one to two years in arrears.
7. **Prefer national data.** Survey "regions" mix dissimilar markets and differ by publisher.
8. **Apples-to-apples productivity.** APP wRVUs billed under a physician's identifier
   inflate apparent production.
9. **Comp-to-collections over 100% is not a red flag by itself** for coverage-based or
   high-Medicaid specialties.
10. **Small samples move fast.** One respondent can shift a percentile materially.

See `references/fmv-methodology.md` for the full write-up and sources.

## Transparency rules (non-negotiable)
- Show the blended weighted-average math with real numbers before using it.
- Label interpolated percentiles **directional estimates**, never exact survey percentiles.
- Label every figure from-file / user-provided / derived.
- Always present both cost bases (vs current, vs planned).
- State explicitly whether benefits/fringe are included.
- Never fabricate, extrapolate, or fill in benchmark data absent from the user's sources.

## Out of scope (for now)
- **Call coverage pay, medical directorships, and other non-clinical arrangements.** These
  are valued differently (hourly rates, burden/intensity analysis, market rate per day of
  coverage). If a user asks, say so plainly and offer the clinical review instead — do not
  stretch clinical TCC benchmarks onto a call-pay question.

## Hard boundaries
- This produces an **informational market-approach analysis, not a certified FMV opinion.**
  It does not perform the income or cost approaches and does not evaluate commercial
  reasonableness (a separate test — an arrangement can be at FMV and still not be
  commercially reasonable). Where regulatory reliance is required (Stark/Anti-Kickback), an
  independent qualified valuation professional should perform the determination.
- Never invent survey numbers or n-counts. Suppressed data stays suppressed.
- Provider names and compensation are confidential: keep them in the user's own files and
  logs, never in examples or anything committed to a repository.

## Deliverables
1. **Review workbook** — their template, both sheets, formulas preserved and verified
2. **DOCX + PDF packet** — styled, 0.6" margins, KPI cards, callouts, market-band exhibit
3. **HTML email report** — self-contained, embedded exhibit, PDF export
4. **Exhibit PNG** — market band p25–p75, p50 line, p90 marker, current/planned/proposed
5. **CSV** audit extract (cohort) / **summary JSON** (new hire)

## Style
Write like a compensation analyst: **empirical, concise, neutral.** Label what is from file
versus derived. Do **not** FTE-adjust governance percentiles. Prefer tables and one strong
exhibit over long narrative.

## Pitfalls / QA checks
- **Benefits rate:** read from the template formula. A fully-loaded physician fringe
  (employer FICA ~7.65% + retirement + health + malpractice + CME) commonly runs 15–30%;
  a narrow internal allocation can be low single digits. The example workbook's rate is an
  arbitrary placeholder, not a recommendation.
- **wRVUs are a sum, not a count** — the annual total of wRVU values across all services.
- **Specialty dash mismatch:** roster files often use ASCII `-` where survey files use a
  Unicode en-dash `–` (U+2013). Normalize with `unicodedata.normalize('NFKD')` + ASCII
  encoding, then fuzzy-match checking **all** dash-separated parts — "Pediatrics" alone
  matches "Pediatrics – General" and picks the wrong row.
- **Name matching:** normalize "Last, First" vs "First Last", punctuation, suffixes.
- **Percentile math:** never present interpolated estimates as exact survey percentiles.
- **Cost basis:** always show vs-current *and* vs-planned.
- **Equity check:** confirm proposals don't create senior-junior inversions.
- **Prefer formulas over hardcoded percentiles** so positioning updates if benchmarks change.
- **Do not FTE-adjust** governance percentiles — lower positioning for part-time is correct.
- **Suppressed data** (`*`, blank, `0`) → n=0, excluded from blending automatically.
- **openpyxl cannot evaluate formulas** — verify via LibreOffice headless recalculation or
  recompute in Python. Never trust an unevaluated cell.

## Files
- `scripts/committee_template_generator.py` — workbook fill (lookup, blending, verification)
- `scripts/fmv_analysis.py` — market-approach analytics: positioning, comp-per-wRVU,
  alignment flags, caveat set
- `scripts/fmv_register.py` — request register (add / update / list / summary)
- `scripts/inspect_workbook.py` — read-only setup helper for unfamiliar workbooks
- `scripts/build_adjustment_report.py` — DOCX/PDF packet + exhibit + CSV
- `scripts/build_html_email.py` — self-contained HTML email report
- `scripts/build_cv_only_market_anchor.py` — new-hire range pack
- `scripts/generate_example_xlsx_fixtures.py` — regenerate the synthetic example files
- `references/fmv-methodology.md` — how FMV reviews are actually done, with sources
- `references/` — cell maps, survey file structure, worked example, new-hire notes
- `WALKTHROUGH.md` — full analyst scenario on the bundled example data
