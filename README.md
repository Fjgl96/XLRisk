# XLRisk
XLRisk is an Excel addin for performing [Monte Carlo simulation](https://en.wikipedia.org/wiki/Monte_Carlo_method).  It is **free and open source** and it is **compatible with Excel for the Mac**.  It works in a similar fashion to commercial addins such as [@RISK by Palisade](http://www.palisade.com).

*Note:*  
Excel 2013 or later required.  Excel 16 or newer is recommended.

## About this fork
This fork, maintained by [Fjgl96](https://github.com/Fjgl96), extends the original
[XLRisk by pyscripter](https://github.com/pyscripter/XLRisk) towards the practical feature set of
commercial Monte Carlo add-ins, keeping it free, open source and compatible with existing XLRisk
workbooks. Changes are listed in the [CHANGELOG](CHANGELOG.md).

* **1.1.0** replaces VBA `Rnd()` with the MRG32k3a generator, makes a fixed Seed reproduce the whole
  simulation, fixes the bias of the Latin Hypercube shuffle and adds self tests and a build script.
* Developers: the VBA code lives as text in `Source/`. Build the add-in with
  [`build/build.ps1`](build/README.md) and check it with the self tests described in
  [`tests/README.md`](tests/README.md).

**En español:** este fork busca que XLRisk sea una alternativa open source práctica a los complementos
comerciales de simulación Monte Carlo, sin romper los libros que ya funcionan con XLRisk. La versión 1.1.0
cambia el generador de números aleatorios por MRG32k3a, hace que la semilla reproduzca toda la simulación
y agrega pruebas automáticas. El detalle está en el [CHANGELOG](CHANGELOG.md).

![The XLRisk Ribbon Tab](https://github.com/pyscripter/XLRisk/wiki/Images/XLRisk_Ribbon.jpg)

## Features
* Free and open source
* Can be used with both Excel for Windows and Excel for Mac
* Very easy to use
* Extensive simulation statistics and graphical output
* Implements the [Latin Hypercube sampling](https://en.wikipedia.org/wiki/Latin_hypercube_sampling) method for faster convergence and more accurate results
* Implements the [Iman-Conover method](https://www.uio.no/studier/emner/matnat/math/STK4400/v05/undervisningsmateriale/A%20distribution-free%20approach%20to%20rank%20correlation.pdf) for generating samples from correlated distributions

## Installation instructions and documentation
* [Download and Install](https://github.com/pyscripter/XLRisk/wiki/Installation)
* [The XLRisk Ribbon tab](https://github.com/pyscripter/XLRisk/wiki/Ribbon_Tab)
* [The XLRisk Options dialog](https://github.com/pyscripter/XLRisk/wiki/OptionsDialog)
* [Usage Instructions](https://github.com/pyscripter/XLRisk/wiki/Usage)
* [Risk/distribution functions](https://github.com/pyscripter/XLRisk/wiki/RiskFunctions)
* [Worked Example](https://github.com/pyscripter/XLRisk/wiki/WorkedExample)
* [Simulation Results](https://github.com/pyscripter/XLRisk/wiki/SimResults)
* [Correlated Inputs Example](https://github.com/pyscripter/XLRisk/wiki/CorrelatedInputsExample)

## Tutorials
* [Article on Risk Analysis using XLRisk](https://medium.com/analytics-vidhya/building-a-probabilistic-risk-estimate-using-monte-carlo-simulations-cf904b1ab503) by Zhijing Eu
* [Video tutorial](https://www.youtube.com/watch?v=0RiEquwDjNg)

**Credits:**  
I would like to thank the [ESMT](https://www.esmt.org/) students for providing an incentive for building this addin and helping to improve it.
Thanks are due also to Howard Rudd for his excellent [series of posts](https://www.howardrudd.net/how-tos/vba-monte-carlo-risk-analysis-spreadsheet-with-correlation-part-1/) on Monte Carlo simulation.
XLRisk makes use of some of his code.

