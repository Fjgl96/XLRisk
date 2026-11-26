# Changelog

All notable changes to this fork of XLRisk are documented here.
Versions follow [Semantic Versioning](https://semver.org). The original XLRisk by
pyscripter has no version numbers; its last change (19 March 2025, "Fix #29") is
treated as version 1.0.0.

## [1.2.0] - 2026-11-25

### Added
- **Output Mean Tornado** (new section of every Risk Results sheet). For each output and input the
  iterations are sorted by the input and split in 10 groups of the same size; the bars go from the
  lowest to the highest mean of the output among the groups, measured from the overall mean, so
  the tornado reads in the units of the output (soles, dollars), like the "change in output mean"
  tornado of @RISK. A table next to each chart lists both means, the swing and an approximate share
  of the variance (correlation squared, rescaled to 100%). Inputs that do not change the output
  are left out.
- **Compare Runs** button: overlays the cumulative distribution of one output from every
  "Risk Results N" sheet on a new "Risk Compare" sheet, with mean, standard deviation, minimum,
  5th percentile, median, 95th percentile, maximum and probability below 0 of each run.
- **Swap Functions** button: replaces every Risk distribution function of the workbook by its
  expected value (only the function call, the rest of the formula stays), so the file can be sent
  to people without XLRisk. The original formulas are kept in a hidden "XLRisk Swap" sheet; the
  second press puts them back. Run explains what to do while the functions are swapped out.
- `RiskTriGen(bottom, mostLikely, top, [bottomPerc], [topPerc])`: triangular distribution given
  two percentiles (10% and 90% by default) and the most likely value, as in @RISK.
- Histograms in Excel 2013 and older: a column chart of the counts in 20 equal bins (also used
  if the Excel 2016 histogram chart fails).
- 27 new self tests (58 in total).

### Fixed
- A simulation that fails now says why. The original stopped without any message on errors, for
  example when an output of the XLRisk sheet showed `#NAME?`; that case now names the cell and
  how to fix it.
- `RiskDiscrete` accepts the values in any order: it sorts them together with their
  probabilities. Before, unordered values gave `#VALUE!`. It also rejects a negative first
  probability, which was not checked.
- `RiskBernoulli` and `RiskBinomial` show their exact expected value outside a simulation
  (`RiskBernoulli(0.15)` shows 0.15, not a rounded 0; `RiskBinomial(10, 0.15)` shows 1.5, not 2).

### Changed
- "Calculate data tables during simulation" is off in new XLRisk sheets. Data tables recalculate in
  every iteration and made simulations of models with sensitivity tables much slower. Existing
  workbooks keep their setting.
- The Excel 2016 histogram chart type is written as a number, so the add-in also compiles in
  older Excel.

## Also in 1.2.0 (developed as 1.2.0-beta.2)

### Changed
- **Faster simulations.** `Application.Calculate` recalculates every open workbook in every
  iteration. During a simulation XLRisk now suspends the recalculation of the sheets of the other
  open workbooks (`Worksheet.EnableCalculation`) and restores exactly those sheets at the end, also
  when the run fails or is stopped. Add a cell named `RecalcOtherWorkbooks` with the value `TRUE`
  in the XLRisk sheet to keep the previous behaviour.
- The status bar is refreshed, and Excel processes the Stop button, about four times per second
  instead of in every iteration.

### Added
- **Functions with the same names and arguments as in @RISK:**
  - `RiskOutput("Name")`: marks a cell as an output from its formula, as in
    `=RiskOutput("Profit") + C17 - C18`. Returns 0. The Add Output button keeps working; a cell
    registered both ways counts once.
  - `RiskMean`, `RiskStdDev`, `RiskMin`, `RiskMax`, `RiskPercentile` (and `RiskPtoX`),
    `RiskTarget` (and `RiskXtoP`): statistics of an output or input cell in the latest simulation,
    read from the most recent "Risk Results N" sheet, so they work after reopening the workbook.
    They return `#N/A` while a simulation runs and for cells without results.
  - `RiskPoisson(lambda)` and `RiskLognorm2(mu, sigma)` (same as XLRisk's `RiskLogNorm`).
- Self tests for all of the above (31 checks).
- `ClsSimulation.LoopSeconds`: seconds taken by the iterations of the last run. The self tests
  report it (INFO row) and check that the other workbooks recalculate again after a run.

## [1.1.0] - 2026-11-25

### Changed
- **Random number generator.** VBA `Rnd()`, a 24-bit linear congruential generator
  with a period of about 16.7 million, is replaced by MRG32k3a (L'Ecuyer, 1999):
  period of about 2^191, passes the TestU01 BigCrush battery, and gives identical
  results on 32-bit Excel, 64-bit Excel and Excel for the Mac. The new `Rng` module
  holds the generator.
- Because the generator changed, a workbook with a fixed Seed gives different
  numbers than in XLRisk 1.0. To reproduce 1.0 results, add a cell named
  `LegacyRnd` with the value `TRUE` in the XLRisk sheet.

### Fixed
- **Seed did not reproduce simulations.** The random samples were drawn before the
  generator was seeded, so the first run of a session and any run after a
  recalculation with Show Samples gave different numbers for the same Seed. The
  generator is now seeded before any sample is drawn.
- **Latin Hypercube shuffle was not uniform.** `Shuffle` rounded the random index
  with `CLng`, which gave the first and last positions half the probability of the
  others. It is now a standard Fisher-Yates shuffle. In tests the old shuffle did not
  create measurable correlation between independent inputs, but the permutation
  was biased.

### Added
- `XLRiskVersion()` worksheet function and `XLRISK_VERSION` constant.
- Self tests (`SelfTests` module, macro `XLRiskSelfTests`): reference values of the
  generator, uniformity, shuffle, and a full simulation of a model with known exact
  results, including reproducibility with a fixed Seed. See `tests/README.md`.
- `tests/reference/mrg32k3a_reference.py`: regenerates every reference value of the
  self tests outside Excel.
- `build/build.ps1` builds `dist/XLRisk.xlam` from `Source/` and writes its SHA-256
  checksum; `build/export.ps1` exports code edited in the VBA editor back to `Source/`.
- `.gitattributes` keeps the VBA files with Windows line endings.
- Spanish user manual, as an Excel workbook with a runnable example model
  (`docs/XLRisk_Manual_ES.xlsx`) and as Markdown (`docs/manual-es.md`), both generated by
  `docs/build_manual.py`.
- README rewritten: download link, install steps, quick start and function reference.
- `XLRisk.xlam` at the repository root is now the 1.1.0 build (not digitally signed).

### Documented
- `RiskLogNorm(mu, sigma)` takes the mean and standard deviation of ln(X), like `RiskLognorm2`
  in @RISK, not those of X. The manual shows how to convert them.
