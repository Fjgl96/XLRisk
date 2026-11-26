# XLRisk
XLRisk is an Excel add-in for [Monte Carlo simulation](https://en.wikipedia.org/wiki/Monte_Carlo_method).
It is **free and open source** and **compatible with Excel for the Mac**. It works in a similar fashion to
commercial add-ins such as @RISK.

**[⬇ Download XLRisk 1.2.0](https://github.com/Fjgl96/XLRisk/raw/master/XLRisk.xlam)** ·
**[Manual en español](docs/manual-es.md)** ([Excel](docs/XLRisk_Manual_ES.xlsx)) ·
**[Test workbooks](examples/)** ([basic](examples/XLRisk_Prueba_Basica.xlsx), [multi-variable](examples/XLRisk_Prueba_Multivariable.xlsx)) ·
[Changelog](CHANGELOG.md)

*Requires Excel 2013 or later. Excel 2016 or newer is recommended.*

![The XLRisk Ribbon Tab](https://github.com/pyscripter/XLRisk/wiki/Images/XLRisk_Ribbon.jpg)

## About this fork
This fork, maintained by [Fjgl96](https://github.com/Fjgl96), extends the original
[XLRisk by pyscripter](https://github.com/pyscripter/XLRisk) towards the practical feature set of
commercial Monte Carlo add-ins, while keeping it free, open source and compatible with existing XLRisk
workbooks.

**What is new in 1.2.0**
* **Tornado in the units of the output** (Output Mean Tornado), with an approximate share of variance per input.
* **Compare Runs**: overlays the cumulative distributions of several simulations with a table of statistics.
* **Swap Functions**: replaces the Risk functions by their expected values to share the model, and back.
* Functions with the same names as in @RISK: `RiskOutput`, `RiskMean`, `RiskStdDev`, `RiskMin`, `RiskMax`,
  `RiskPercentile`, `RiskTarget`, `RiskTriGen`, `RiskPoisson`, `RiskLognorm2`.
* About 3 times faster runs when other workbooks are open.
* `RiskDiscrete` accepts values in any order; `RiskBernoulli`/`RiskBinomial` show exact expected values;
  data tables are off by default; histograms also in Excel 2013.

**1.1.0**: professional random number generator (MRG32k3a), a fixed Seed reproduces the whole simulation,
unbiased Latin Hypercube shuffle, self tests (58 in 1.2.0), build script and a Spanish manual.

## Features
* Free and open source (MIT license)
* Excel for Windows and Excel for Mac
* 15 distributions, plus correlated inputs with the
  [Iman-Conover method](https://www.uio.no/studier/emner/matnat/math/STK4400/v05/undervisningsmateriale/A%20distribution-free%20approach%20to%20rank%20correlation.pdf)
* [Latin Hypercube sampling](https://en.wikipedia.org/wiki/Latin_hypercube_sampling) for faster convergence
* Statistics, percentiles, histograms, cumulative distributions and tornado diagrams

## Install
1. Download [`XLRisk.xlam`](https://github.com/Fjgl96/XLRisk/raw/master/XLRisk.xlam) (SHA-256 `092b23c72f04c3315b5d85854cb495775ca23213ffa207aada1fe8138d1df7a2`).
2. Windows: right click the file > Properties > tick **Unblock** > OK. Otherwise Excel opens it in
   Protected View and it does not load.
3. Copy it to `%APPDATA%\Microsoft\AddIns` (a trusted location for Excel).
4. In Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse > select `XLRisk.xlam` > OK.
   Untick any older XLRisk first and keep a single copy of the file on your computer.
5. Check: `=XLRiskVersion()` in any cell returns `1.2.0`. Optionally run the self tests
   (Alt+F8, `XLRisk.xlam!XLRiskSelfTests`).

## Quick start
1. Build your model as usual.
2. Replace each uncertain assumption with a risk function, one per cell, for example `=RiskTriang(45, 50, 56)`.
   XLRisk finds the inputs by scanning the formulas of the workbook.
3. Select the result cell and press **Add Output**.
4. Press **Setup** to choose the Seed and options, set the iterations and press **Run**.
5. Read the new results sheet: statistics, percentiles, histogram, cumulative distribution, tornado and
   output mean tornado. Run again with other assumptions and press **Compare Runs** to overlay the runs.

Outside a simulation every risk function returns its expected value, so the model keeps showing a base case.

## Risk functions
| Function | Parameters | Typical use |
| --- | --- | --- |
| `RiskTriang(min, mode, max)` | min < mode < max | Three-point expert estimate |
| `RiskPert(min, mode, max)` | min < mode < max | Like triangular, less weight on the extremes |
| `RiskNormal(mean, stdev)` | | Symmetric variables |
| `RiskLogNorm(mu, sigma)` | **mean and stdev of ln(X)**, like @RISK's `RiskLognorm2` | Positive, right-skewed variables |
| `RiskUniform(min, max)` | | Only the range is known |
| `RiskTriGen(bottom, mode, top, [p_bottom], [p_top])` | values at two percentiles (10% and 90% by default) | Expert estimates of P10, most likely and P90 |
| `RiskDiscrete(values, probabilities)` | values in any order, probabilities sum to 1 | Scenarios |
| `RiskDUniform(values)` | | Equally likely options |
| `RiskBernoulli(p)`, `RiskBinomial(n, p)`, `RiskPoisson(lambda)` | | Events that happen or not; counts |
| `RiskExpon(mean)`, `RiskGamma(alpha, beta)`, `RiskErlang(k, lambda)`, `RiskWeibull(alpha, beta)` | beta and lambda are scale parameters | Times, sizes, lifetimes |
| `RiskBeta(alpha, beta, [A], [B])` | A, B default to 0, 1 | Rates and proportions |
| `RiskCumul(min, max, xValues, cumProbs)` | | Expert percentiles |
| `RiskLognorm2(mu, sigma)` | same as `RiskLogNorm` | Name used by @RISK |
| `RiskCorrmat(matrix, position)` | last argument of any distribution | Correlated inputs |

Outputs and results: `RiskOutput("Name")` marks an output from its formula (`=RiskOutput("NPV") + ...`);
`RiskMean`, `RiskStdDev`, `RiskMin`, `RiskMax`, `RiskPercentile` (`RiskPtoX`) and `RiskTarget` (`RiskXtoP`)
read the latest simulation.

Helpers: `RiskIsValidCorrmat`, `RiskCorrectCorrmat`, `RiskSCorrel`, `XLRiskVersion`.
Full detail, settings and examples in the [Spanish manual](docs/manual-es.md) and in the
[original wiki](https://github.com/pyscripter/XLRisk/wiki).

## Original documentation (wiki of pyscripter/XLRisk)
* [Ribbon tab](https://github.com/pyscripter/XLRisk/wiki/Ribbon_Tab) ·
  [Options dialog](https://github.com/pyscripter/XLRisk/wiki/OptionsDialog) ·
  [Usage](https://github.com/pyscripter/XLRisk/wiki/Usage) ·
  [Risk functions](https://github.com/pyscripter/XLRisk/wiki/RiskFunctions)
* [Worked example](https://github.com/pyscripter/XLRisk/wiki/WorkedExample) ·
  [Simulation results](https://github.com/pyscripter/XLRisk/wiki/SimResults) ·
  [Correlated inputs example](https://github.com/pyscripter/XLRisk/wiki/CorrelatedInputsExample)
* Tutorials: [article by Zhijing Eu](https://medium.com/analytics-vidhya/building-a-probabilistic-risk-estimate-using-monte-carlo-simulations-cf904b1ab503) ·
  [video](https://www.youtube.com/watch?v=0RiEquwDjNg)

## Developers
The VBA code lives as text in `Source/`. Build the add-in with [`build/build.ps1`](build/README.md) and
check it with the self tests in [`tests/README.md`](tests/README.md). The Spanish manual is generated by
`docs/build_manual.py`.

## En español
XLRisk es un complemento gratuito y open source para simulación Monte Carlo en Excel, compatible con los
libros del XLRisk original. **[Descarga la versión 1.2.0](https://github.com/Fjgl96/XLRisk/raw/master/XLRisk.xlam)**
y sigue el **[manual en español](docs/manual-es.md)**, que también está como
**[libro de Excel](docs/XLRisk_Manual_ES.xlsx)** con un modelo de ejemplo listo para simular.
Para probarlo hay dos libros listos para correr (solo abrir y pulsar Run): uno
**[básico](examples/XLRisk_Prueba_Basica.xlsx)** y uno **[multivariable](examples/XLRisk_Prueba_Multivariable.xlsx)**.

## Credits
XLRisk was created by [pyscripter](https://github.com/pyscripter) (Kiriakos Vlahos). He thanks the
[ESMT](https://www.esmt.org/) students for providing an incentive for building this add-in and helping to
improve it, and Howard Rudd for his
[series of posts](https://www.howardrudd.net/how-tos/vba-monte-carlo-risk-analysis-spreadsheet-with-correlation-part-1/)
on Monte Carlo simulation; XLRisk uses some of his code. @RISK is a trademark of Lumivero; this project is
independent.
