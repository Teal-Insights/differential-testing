# Technical Standard: Differential Testing

This standard defines how an Excel-to-Python extraction is validated by differential
testing: the rule for comparing values, the report format, the acceptance bar, and the
coverage a conforming harness must exercise.

---

## 0. What this standard governs

**Differential testing** is a long-established software validation technique (McKeeman,
*Differential Testing for Software*, 1998). Feed **identical inputs** to two independent
implementations of the same specification and compare their outputs. If they disagree,
one of them has a bug, and the disagreement shows exactly where to look.

Here the two implementations are called **oracles**. The original Excel workbook is the
trusted reference, the **golden master**, so golden-master validation and differential
testing mean the same thing in this standard. When the oracles disagree, the golden
master wins and the disagreement localizes a bug in the system under test. The workbook
is its own specification: no one has to write down what "correct" means for each of
thousands of formulas, because the new implementation is held to the workbook's answers.

| Term | Definition in this standard |
|---|---|
| **Golden master** (reference oracle) | The trusted Excel workbook, run live in Microsoft Excel through `xlwings` (COM automation). Source of truth. |
| **System under test (SUT)** | The generated Python code, tested at one of two layers. The *extracted computation* (the in-memory dependency graph, evaluated directly) answers "is the extraction faithful?". The *shipped package* (the standalone library, called only through its public API) answers "is what users receive faithful?". |
| **Harness** | The code that writes each scenario to both oracles, reads their outputs, compares them, and writes the report. |
| **Scenario / test suite** | A scenario is one complete set of input values. The test suite is the set of scenarios a run executes (§6). |
| **Comparison** | One output cell in one scenario: the *golden value* against the *SUT value*. The unit that §1 judges and §2 records. |
| **Parity report** | The CSV and TXT record of every comparison (§2). |
| **Acceptance bar** | 100% of comparisons pass under the precision policy of §1. One mismatch fails the run. |

The sections follow the logic of the test. §1 defines when two values are equal. §2–§3
define how every comparison is reported. §4–§5 check that the harness is wired
correctly, and §6 that the scenarios exercise the model, so that a pass means what it
says. §7–§8 apply the standard to Q-CRAFT v2 and DDT. §9 lists what a pass still does
not prove. A conforming validation MUST satisfy §1–§6; the conformance checklist at the
end turns them into review items.

---

## 1. Precision policy (the comparison contract)

This section is the core of the standard: the exact, ordered rule that decides whether a
golden value and a SUT value are equal. Every tool uses the same rule, so parity reports
are comparable across tools.

### 1.1 Hybrid numeric tolerance: `atol = 1e-6`, `rtol = 1e-12`

A numeric comparison passes when **either** the absolute or the relative difference is
within tolerance:

```
abs_diff = abs(golden - sut)
rel_diff = abs_diff / abs(golden)        (infinite when golden == 0)

passed   ⇔  abs_diff <= 1e-6   or   rel_diff <= 1e-12
```

- `rel_diff` is measured against the golden value only. When the golden value is `0`,
  `rel_diff` is infinite, so only the absolute bound can pass the comparison.
- `atol` and `rtol` MUST each be defined once, as named constants passed through the run
  as configuration, and never inlined at a comparison site. A tool's tolerances are then
  auditable in one place.
- The verdict MUST be computed from the same `rel_diff` value written to the report, so a
  reader can reproduce every verdict from the `abs_diff` and `rel_diff` columns alone.

**Why `atol = 1e-6`.** Spreadsheets typically show values to about 6 decimals, and Excel
and Python can differ in the last few bits from floating-point rounding without any real
defect. `1e-6` absorbs that rounding without hiding a genuine divergence. On faithful
extractions, comparisons decided by the absolute bound are observed to be overwhelmingly
`0.0` (bit-exact): the tolerance is a margin, not a substitute for agreement. Passes
decided by the relative bound have `abs_diff > atol` by definition.

**Why `rtol = 1e-12`.** An absolute bound stops measuring parity once values grow large.
At magnitude `1e16`, one unit in the last place (ULP) of an IEEE 754 double is about 2,
so `atol = 1e-6` is finer than the number format can resolve; an absolute-only gate would
demand identical operation order rather than computational agreement. `rtol = 1e-12`
requires agreement to at least 12 significant digits. Doubles carry about 15.95, so this
allows about 4 digits of rounding accumulated through deep recursions while staying
orders of magnitude stricter than any genuine modeling divergence.

**Why both bounds.** Neither bound works alone. The absolute bound cannot be met once
`|golden|` exceeds about `8.6e9`, where one ULP already exceeds `1e-6`; somewhat below
that, a few ULPs of accumulated rounding exceed it while staying near machine epsilon in
relative terms. The relative bound cannot judge a zero golden value and becomes
arbitrarily strict as values approach zero. The hybrid rule applies whichever bound is
meaningful at the magnitude of the value being compared. The conventional
`rel_tol = 1e-9` of `math.isclose` is too loose for this contract and MUST NOT be
substituted.

**What the relative bound does not excuse.** The hybrid rule stops the comparison from
failing agreement that the number format cannot express. It is not a license for loose
parity. A divergence such as golden `51.14` against SUT `1075.14` (typical of ULP-level
differences in very large intermediates surviving catastrophic cancellation at degenerate
inputs) still fails, and MUST fail. Ill-conditioned outputs at degenerate inputs are a
question of scenario design and model validity, not of tolerance. Widening either
constant, or changing the `rel_diff` denominator, to absorb such rows is prohibited: it
changes what "parity" means and breaks comparability of reports across tools.

### 1.2 The ordered comparison ladder

Each comparison runs down a fixed list of rules, and the first rule that matches decides.
Order matters: blanks, errors and non-finite values are settled before the tolerance
rule, so they never reach the arithmetic.

| # | Condition | Verdict | `abs_diff` / `rel_diff` |
|---|---|---|---|
| 1 | both blank (`None`) | **pass** | `0.0` / `0.0` |
| 2 | exactly one side blank | **fail** | none |
| 3 | value not numeric-coercible (e.g. a spreadsheet error) | pass **iff** both sides are equal by typed identity | none |
| 4 | numeric but non-finite (`NaN`, `±inf`) | pass **iff** exactly equal, **or** both `NaN` | none |
| 5 | finite numbers | pass **iff** `abs_diff <= atol` **or** `rel_diff <= rtol` | computed |

Rules 1–2 are the blank rules and rule 4 is the `NaN` rule (§1.4). Rule 3 is error-class
equality (§1.3). Rule 5 is the tolerance rule (§1.1).

### 1.3 Error-class equality

Spreadsheet errors (`#DIV/0!`, `#VALUE!`, `#N/A`, …) are comparison values, not noise.
Two errors match **iff they are the same error class**: `#DIV/0!` against `#N/A` fails.

**Compare typed errors, not strings.** The equality MUST be a typed, class-level
equality: both sides are recognized as errors and agree on *which* error. Normalize the
spreadsheet driver's raw error representation to a typed error value, then compare the
types. Do **not** rely on the driver's error text happening to equal a language-level
sentinel. That can pass by accident, for example when the sentinel is a string subclass
whose members spell the Excel errors exactly, and it stays unverified until an error path
actually runs (§6, §9).

**Matched errors fail the run unless declared.** A matched error passes its comparison
but proves nothing: both oracles failed the same way, which is no evidence of
computational parity. A common cause is a dropdown label mismatch (logical `"High"`
against workbook `"High "`) that silently produces errors on long-horizon outputs only
(§5.2). A matched error in a scenario that does not declare `expects_error_values=True`
MUST fail the run (exit code `1`) and be listed in the report. A harness MAY offer an
override flag (`--allow-matched-errors`) that downgrades these to warnings for triage; it
MUST NOT be the default, and CI MUST NOT pass it.

### 1.4 Blank and `NaN` rules

- **Blanks (`None`):** both blank passes; one blank against a value fails. A blank
  against a number is a real divergence (the SUT left a cell empty that the golden master
  filled, or the reverse) and MUST fail, never be silently tolerated.
- **`NaN` and infinities:** `NaN` compares false against everything, so tolerance cannot
  judge it, and it MUST be handled before the arithmetic. Two `NaN`s are equal; `±inf`
  matches only the same infinity.

### 1.5 Crash containment

An exception in one scenario MUST NOT abort the test run. It is recorded as one
**failing** comparison per output cell, with both oracle values replaced by the
exception's class and message and the diff fields left empty. The crash then appears in
the report as a failure instead of losing the whole run.

---

## 2. Parity-report convention

Every conforming run writes **two files** to a report directory. Both are required,
because they serve different readers: a CSV with one row per comparison, for diffing and
pivoting, and a TXT summary, for a human to read first.

### 2.1 CSV: one row per comparison

Recommended schema:

```
scenario_id, cell_address, cell_label, golden_value, sut_value, abs_diff, rel_diff, passed
```

The CSV MUST contain exactly one row per comparison, with a stable scenario identifier,
the cell address **and** a human-readable label, both oracle values, `abs_diff` and
`rel_diff`, and a boolean pass column. Single-axis scenario identifiers MUST let a reader
recover per-axis pass rates by grouping the CSV (e.g. `single_axis:<axis>=<value>`).

### 2.2 TXT: the summary a human reads first

Required content, in order:

1. **Header:** title, generation timestamp (UTC, ISO 8601, to the second), workbook
   identity, SUT identity, and both tolerances (`atol = 1e-06`, `rtol = 1e-12`).
2. **Counts:** total comparisons, passed, failed, and pass rate (%).
3. **The acceptance bar, printed literally:** `Acceptance bar: 100.00%`.
4. **Result:** `PASS` or `FAIL`.
5. **First divergence** (§3), when any comparison failed.
6. **Every failing comparison.**

The TXT SHOULD also include per-axis and per-point pass-rate tables, so the axis that
broke is visible at a glance. When the SUT is the extracted graph, it SHOULD also list
**absent inputs**: input cells the test suite tried to set that the SUT does not expose.
An absent input is itself a differential signal about the extraction.

### 2.3 The 100% pass-rate acceptance bar

The bar is **100%**: one mismatch fails the run. The bar MUST be asserted in three
places, and all three MUST agree:

- the TXT prints `Acceptance bar: 100.00%`;
- the process **exit code** is `0` iff every comparison passes and no undeclared matched
  errors remain (§1.3), `1` if any comparison fails, and `2` if a prerequisite is
  missing (no workbook, no Excel, package not generated);
- CI (the pull-request test workflow) and reviewers treat any non-`PASS` report as a
  blocking failure.

### 2.4 Reports are versioned artifacts, not scratch output

A parity report is valid only for the exact workbook, extraction configuration and test
suite that produced it. When any of them changes, the affected test suite MUST be re-run
and the new report committed with the tool; a stale report describes a model that no
longer exists. Ship a reference report with the tool and treat any diff against it as a
regression signal.

---

## 3. First-divergence reporting convention

When a run fails, the TXT points to a single starting place: the *first* failing
comparison in a deterministic order (scenario order, then cell-label order). The block
MUST name the scenario and the exact cell, with enough detail to reproduce and localize
the divergence:

```
First divergence:
  scenario:  <scenario_id>
  cell:      <cell_address>  (<cell_label>)
  golden:    <golden_value>
  sut:       <sut_value>
  abs_diff:  <abs_diff>
  rel_diff:  <rel_diff>
```

**Why first, and why deterministic.** The scenario order and cell-label order are fixed,
so the same bug surfaces at the same scenario and cell on every run with the same inputs.
An engineer starts from one scenario and one cell instead of triaging thousands of rows.
This is the localization that differential testing promises (§0). A conforming tool
provides at least this pointer, and MAY also list all failures ranked by descending
`abs_diff` to surface the largest divergence.

---

## 4. Structural pre-flight checks (fail fast, before any scenario)

These checks run before any scenario. Each rules out a different cause of failure, so a
passing run means what it says. All three are required.

1. **Path and package check.** The workbook and the generated package must exist. If
   either is missing, stop immediately with a hint for regenerating it (exit code `2`,
   §2.3); no scenario runs against missing inputs.
2. **Staleness check.** Compare the workbook with the SUT's embedded snapshot of the
   model's constants. A stale snapshot makes a later diff look like a code-generation
   bug. The check MUST compare content hashes and **fail hard** on a mismatch; a
   warning, such as one based on file timestamps, can be ignored and lets real drift
   through (§9).
3. **Binding-cell check.** Before comparing any values, assert that the SUT's input
   setters and output readers target exactly the cell addresses the test suite declares,
   after address normalization (§5.1). This separates binding correctness from
   calculation correctness: without it, a setter bound to the wrong cell can go
   unnoticed whenever that cell's default happens to equal the test value. A conforming
   tool MUST verify its input and output bindings before trusting a passing calculation.

---

## 5. Harness boundary rules

The pre-flight checks (§4) verify the harness before it runs. These rules govern the
values it handles while it runs: cell addresses used for lookups, and input values
written to both oracles. Both must be normalized or resolved the same way for both
oracles; otherwise the comparison tests the harness, not the model.

### 5.1 Address-key normalization

`excel-grapher` stores sheet-qualified addresses in canonical form (e.g.
`'Discrete Risks'!H2`). Human-authored binding keys and graph `leaf_keys()` may quote
sheet names differently but normalize to the same form: `normalize_cell_type_env_key()`
for constraint matching and `normalize_key()` for graph lookup. Harnesses and audits must
normalize before comparing configured addresses to graph keys, using `normalize_key` /
`parse_address` before graph lookups and `xlwings` writes, never a naive
`split("!", 1)`. See [issue #51](https://github.com/Teal-Insights/extraction-pipeline-template/issues/51).

### 5.2 Workbook-exact label resolution

Scenario matrices and exported APIs use clean logical values (`"High"`,
`"Real interest rate"`). Many workbooks branch on **exact string equality** with
reference label cells (e.g. `IF($B$2=$B$9, …)`). If a harness writes the logical string
instead of the workbook's exact label, formulas can silently fall through to `""` and
produce `#VALUE!` in later projection years. Parity still passes, because both oracles
fail the same way: the matched-error case of §1.3.

**Workbook reference label cells are the source of truth for values written to public
input cells.** Scenario ids and API parameter names may stay human-readable, but every
Excel-facing write must match the reference literal character for character, including
trailing spaces and suffixes such as `(a)`.

Workbook-neutral procedure:

1. **Find the reference labels.** Identify the cells the workbook uses in
   IF/MATCH/CHOOSE guards for each public enum input. Sheet and range names are
   project-specific.
2. **Record exact literals.** Capture each value with `repr()`; do not assume trimmed or
   canonical spellings.
3. **Align enum domains.** `domain.enum.values` on each public input must use those exact
   strings, not cleaned-up scenario names.
4. **Resolve at the harness boundary.** Scenarios keep logical values. A small resolver
   maps each logical value to its workbook-exact label immediately before the golden and
   graph `set_inputs` calls, matching with `.strip()` plus prefix rules for suffix
   variants (`"Real interest rate"` → `"Real interest rate (a)"`). **Write the raw
   workbook string.**
5. **Normalize only internally.** Block lookups, dict keys and test ids may strip
   whitespace; public input writes may not, until the source workbook is corrected.
6. **Test label resolution.** Assert that every scenario value for each enum public input
   resolves to a loaded workbook label, and that every binding enum value appears among
   the reference labels. Example:
   [`tests/differential/workbook_labels.py`](https://github.com/Teal-Insights/qcraft-extraction-pipeline/blob/main/tests/differential/workbook_labels.py)
   in the Q-CRAFT extraction pipeline.

---

## 6. Coverage design (what the test suite must exercise)

The comparison rule (§1) proves parity only on the inputs the test suite actually runs.
The standard test suite has four scenario groups:

- **Canonical scenarios:** the tool's headline configurations (e.g. baseline plus each
  named shock for each modeled entity). They anchor the test suite to real cases.
- **Single-axis isolation:** change exactly one parameter from a canonical scenario, so a
  break shows up cleanly in the per-axis pass-rate table (§2.2).
- **Full categorical factorial:** every combination of the categorical axes, to catch 2-
  and 3-way interactions that single-axis scenarios cannot reach.
- **Boundary coverage:** the first and last periods of any recursion or time dimension,
  not just interior ones, where off-by-one and boundary bugs live.

**Mandatory: the error and boundary domain.** A test suite of well-formed inputs only is
the *weakest* differential test: divergences cluster at the edges, and well-formed inputs
exclude the edges. A conforming test suite MUST include an **error/boundary scenario
group** that drives every reachable error path and domain edge, for example:

- out-of-range dispatch indices;
- unknown lookup keys (`#N/A`);
- inputs that zero a denominator (`#DIV/0!`);
- triggers that never fire and triggers that always fire;
- non-integer inputs that hit truncation.

This group is what actually *exercises* error-class equality (§1.3), rather than merely
defining it.

---

## 7. Applying the standard to Q-CRAFT v2

Q-CRAFT (Quantitative Climate Risk Assessment Fiscal Tool; Batini et al., 2024) is an IMF
Fiscal Affairs Department workbook with 100+ country datasets and rich scenario logic.
The **contract** of this standard applies unchanged; only the **tool-specific map** is
new.

**Held constant across tools (do not re-derive):**

- the precision policy of §1, verbatim: the same hybrid gate (`atol = 1e-6`,
  `rtol = 1e-12`), comparison ladder, and error-class, blank and `NaN` rules. This is
  what makes a Q-CRAFT parity report comparable to any other;
- the two-file report (§2), the 100% bar and exit codes (§2.3), the first-divergence
  block (§3), and the three pre-flight checks (§4).

**Specified per tool:**

- **Binding map:** the named ranges and cells for Q-CRAFT's inputs and outputs, and each
  input's typed shape. Verify them with the binding-cell check (§4.3) before trusting
  results.
- **Scenario axes:** the categorical axes (country, climate-shock type and severity,
  horizon), the continuous baselines, and the mandatory error/boundary group (§6).
  Country lookup is central to Q-CRAFT, so that group includes unknown-country `#N/A`
  scenarios.
- **Output cell set:** the indicator, threshold and ratio-path cells Q-CRAFT reports.

With 100+ countries, expect a large scenario count. The report format (per-axis tables,
first divergence, groupable CSV) is built to stay readable at that size and needs no
structural change.

## 8. Applying the standard to DDT

DDT, the second target tool, uses the same contract and the same split. Hold constant
the precision policy, report format, acceptance bar, first-divergence rule and pre-flight
checks. Specify only DDT's binding map, scenario axes (including the error/boundary
group) and output cell set.

If DDT outputs categorical or classification cells, ladder rule 3 (not numeric-coercible,
so typed equality) already covers them, and no new comparison rule is needed. DDT's
scenario design only has to enumerate the valid categorical values.

For both tools, the conformance checklist at the end of this standard is the review.

---

## 9. What a green "100% PASS" does NOT prove

A pass proves parity only within the scope of the run. Design against these gaps from
the start:

- **Untested inputs.** Parity holds only for the scenarios that ran. A 100% pass on
  well-formed inputs proves "matches Excel on well-formed inputs", not "matches Excel".
  The mandatory error/boundary group (§6) is the mitigation.
- **Incidental error equality.** If errors are compared by string identity rather than
  typed class (§1.3), the comparison is unverified until an error path runs. Make it
  explicit and exercise it.
- **Warn-only staleness.** A staleness warning can be ignored, so real workbook/snapshot
  drift passes silently. This is why §4.2 requires a content-hash check that fails hard.
- **Environment drift.** Parity can depend on the environment in the last bits. Record
  the Python and Excel environment of the reference report and, ideally, reproduce it.

---

## Conformance checklist: is a harness "to standard"?

A reviewer checks a finished harness (for Q-CRAFT v2, DDT or any other Excel-to-Python
target) against every item below in one pass. Each item is a property of the harness,
not a build step, and names the section it enforces. Any unchecked item means the
harness does not conform.

### Precision contract (§1)

- [ ] `atol = 1e-6` and `rtol = 1e-12`, each defined **once** as a named constant and
  passed as configuration, never inlined at a comparison site (§1.1).
- [ ] The gate is the **hybrid** rule (pass iff `abs_diff <= atol` or
  `rel_diff <= rtol`), computed from the same `rel_diff` written to the report; a zero
  golden value is decided by the absolute bound alone (§1.1).
- [ ] The comparison ladder matches §1.2 **in order**: blank → error class →
  `NaN`/non-finite → tolerance, first match wins (§1.2).
- [ ] Errors compare by **typed class** (`#DIV/0!` ≠ `#N/A`) after normalization to typed
  error values, not by incidental string identity (§1.3).
- [ ] A matched error in a scenario without `expects_error_values=True` fails the run and
  is listed in the report; any `--allow-matched-errors` flag is opt-in and absent from
  CI (§1.3).
- [ ] Blanks: both blank → pass; exactly one side blank → fail (§1.4).
- [ ] `NaN` and non-finite values are handled before the arithmetic: two `NaN`s are
  equal; `±inf` matches only the same infinity (§1.4).
- [ ] An exception in one scenario becomes a **failing** comparison per output cell, with
  the exception class and message as oracle values, and does not abort the run (§1.5).

### Report and result (§2–§3)

- [ ] **Both** report files exist: the CSV has the §2.1 columns, one row per comparison
  and groupable single-axis ids; the TXT has the §2.2 sections in order (§2.1–§2.2).
- [ ] The TXT prints `Acceptance bar: 100.00%`, the exit code enforces it (`0` pass,
  `1` fail, `2` prerequisite missing), and CI blocks on any non-`PASS` (§2.3).
- [ ] A forced failure produces a **well-formed, deterministic first-divergence block**:
  the first failure in scenario order, then cell-label order, with every field needed to
  reproduce and localize it (§3).
- [ ] The shipped **reference report** matches the current workbook, extraction
  configuration and test suite (§2.4).

### Harness wiring and coverage (§4–§6)

- [ ] All **three pre-flight checks** run before any scenario: path and package,
  staleness (content-hash **hard-fail**), and binding cells (§4).
- [ ] Sheet-qualified addresses are normalized with `normalize_key` / `parse_address`
  before graph lookups and `xlwings` writes, never compared or split with a naive
  `split("!", 1)` (§5.1).
- [ ] Enum public input writes use **workbook-exact reference labels**; logical values
  are resolved at the harness boundary, not trimmed or canonicalized on write (§5.2).
- [ ] The test suite includes the canonical, single-axis, full-factorial and boundary
  groups (§6).
- [ ] The test suite includes an **error/boundary group** that actually *exercises*
  error-class equality through reachable error paths and domain edges (§6).

### Per-tool map (§7–§8)

- [ ] The binding map, scenario axes and output cell set are the **tool's own**; for
  DDT, the valid categorical values are enumerated (ladder rule 3 covers them, so no new
  comparison rule is needed) (§7–§8).

### Limitations designed against (§9)

- [ ] The Python and Excel **environment is recorded**, so parity is reproducible in the
  last bits (§9).
