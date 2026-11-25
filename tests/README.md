# XLRisk self tests

The self tests live inside the add-in (module `Source/SelfTests.bas`), so they run
in the same Excel where XLRisk is used.

## Run

1. Load the add-in you want to test (for example `dist\XLRisk.xlam`).
2. Press **Alt+F8**, type `XLRisk.xlam!XLRiskSelfTests` in "Macro name" and press **Run**.
3. A new workbook lists every check with PASS or FAIL, and a message gives the total.
   Nothing is saved. A full run takes about a minute.

## What is checked

| Group | Check | Why |
| --- | --- | --- |
| Generator | Three published MRG32k3a values (six seeds 12345) | The generator is implemented exactly |
| Generator | Three values after `RngSeed 12345`; same seed gives the same 1,000 numbers | Seeding is deterministic |
| Generator | Mean, variance, lag-1 autocorrelation and chi-square of 200,000 uniforms | Basic statistical quality |
| Latin Hypercube | Shuffle keeps every value; chi-square of the final position of one value | The shuffle is an unbiased permutation |
| Simulation | Price x quantity model: mean profit 28,850 and probability of loss 35% (exact values) | A full simulation gives correct results |
| Simulation | `RiskNormal(100, 15)`: mean and standard deviation | Continuous distributions |
| Simulation | Correlation induced by `RiskCorrmat`, target 0.7 | Iman-Conover correlation |
| Simulation | Two runs with the same Seed give identical results | Reproducibility |

Simulation checks use 5,000 iterations and a tolerance of about 3 standard errors.

## Reference values outside Excel

```bash
python tests/reference/mrg32k3a_reference.py
```

It re-implements the exact arithmetic of `Rng.bas` and the shuffle of `Math.bas`,
prints every reference value used by the self tests and runs the same statistical
checks. It also shows the bias of the XLRisk 1.0 shuffle.
