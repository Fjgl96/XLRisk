Attribute VB_Name = "SelfTests"
Option Explicit

' Self tests for XLRisk
'
' How to run them: in Excel press Alt+F8, type XLRisk.xlam!XLRiskSelfTests
' in the "Macro name" box and press Run. A new workbook lists every check
' with PASS or FAIL. Nothing is saved.
'
' Reference values
' - The MRG32k3a values for six seeds equal to 12345 are the published values
'   of L'Ecuyer's reference implementation (first output 0.1270111220...).
' - tests/reference/mrg32k3a_reference.py regenerates every reference value
'   used here and runs the same statistical checks outside Excel.
' - The simulation checks compare simulated statistics with exact values of
'   the model, with a tolerance of about 3 standard errors.

Private Const SIM_ITERATIONS As Long = 5000
Private Const SIM_SEED As Double = 20261125#

Private mReport As Worksheet
Private mRow As Long
Private mChecks As Long
Private mFailures As Long

Public Sub XLRiskSelfTests()
    Dim StartTime As Double
    Dim OldLegacy As Boolean

    StartTime = Timer
    OldLegacy = UseLegacyRnd
    UseLegacyRnd = False
    mChecks = 0
    mFailures = 0

    CreateReport
    TestRngReferenceValues
    TestRngSeeding
    TestRngUniformity
    TestShuffle
    TestSimulation
    TestVersion12

    UseLegacyRnd = OldLegacy
    RngRandomize
    FinishReport Timer - StartTime
End Sub

'---------------------------------------------------------------------------
' Random number generator
'---------------------------------------------------------------------------

Private Sub TestRngReferenceValues()
    Const G As String = "Generator"
    RngSetState 12345#, 12345#, 12345#, 12345#, 12345#, 12345#
    CheckNear G, "MRG32k3a published value #1", 0.127011122046577, RngUniform(), 1E-12
    CheckNear G, "MRG32k3a published value #2", 0.318527565396795, RngUniform(), 1E-12
    CheckNear G, "MRG32k3a published value #3", 0.30918601558327, RngUniform(), 1E-12
End Sub

Private Sub TestRngSeeding()
    Const G As String = "Generator"
    Dim First(1 To 1000) As Double
    Dim I As Long
    Dim Same As Boolean

    RngSeed 12345#
    CheckNear G, "RngSeed 12345, value #1", 0.19447201221487, RngUniform(), 1E-12
    CheckNear G, "RngSeed 12345, value #2", 1.20970596364207E-02, RngUniform(), 1E-12
    CheckNear G, "RngSeed 12345, value #3", 0.188748790477344, RngUniform(), 1E-12

    RngSeed 777#
    For I = 1 To 1000
        First(I) = RngUniform()
    Next I
    RngSeed 777#
    Same = True
    For I = 1 To 1000
        If RngUniform() <> First(I) Then Same = False
    Next I
    CheckTrue G, "The same seed gives the same 1,000 numbers", Same, IIf(Same, "identical", "different")
End Sub

Private Sub TestRngUniformity()
    Const G As String = "Generator"
    Const N As Long = 200000
    Const BINS As Long = 100
    Dim Counts(1 To 100) As Long
    Dim I As Long
    Dim Bin As Long
    Dim U As Double
    Dim Prev As Double
    Dim SumU As Double
    Dim SumU2 As Double
    Dim SumLag As Double
    Dim MeanU As Double
    Dim VarU As Double
    Dim Lag1 As Double
    Dim ChiSq As Double
    Dim Expected As Double

    RngSeed 2026#
    Prev = 0#
    For I = 1 To N
        U = RngUniform()
        SumU = SumU + U
        SumU2 = SumU2 + U * U
        If I > 1 Then SumLag = SumLag + U * Prev
        Bin = Int(U * BINS) + 1
        Counts(Bin) = Counts(Bin) + 1
        Prev = U
    Next I

    MeanU = SumU / N
    VarU = SumU2 / N - MeanU * MeanU
    Lag1 = (SumLag / (N - 1) - 0.25) * 12#
    Expected = N / BINS
    For I = 1 To BINS
        ChiSq = ChiSq + (Counts(I) - Expected) ^ 2 / Expected
    Next I

    CheckNear G, "Mean of 200,000 uniforms", 0.5, MeanU, 4# * Sqr(1# / 12# / N)
    CheckNear G, "Variance of 200,000 uniforms", 1# / 12#, VarU, 0.0007
    CheckNear G, "Lag-1 autocorrelation", 0#, Lag1, 4# / Sqr(N)
    CheckBelow G, "Chi-square of 100 bins (99.9% critical value)", 148.23, ChiSq
End Sub

'---------------------------------------------------------------------------
' Latin Hypercube shuffle
'---------------------------------------------------------------------------

Private Sub TestShuffle()
    Const G As String = "Latin Hypercube"
    Const REPS As Long = 20000
    Dim V(1 To 10) As Double
    Dim Counts(1 To 10) As Long
    Dim R As Long
    Dim I As Long
    Dim Total As Double
    Dim ChiSq As Double
    Dim Expected As Double
    Dim IntegrityOK As Boolean

    RngSeed 99#
    IntegrityOK = True
    For R = 1 To REPS
        For I = 1 To 10
            V(I) = I
        Next I
        Shuffle V
        Total = 0#
        For I = 1 To 10
            Total = Total + V(I)
            If V(I) = 1# Then Counts(I) = Counts(I) + 1
        Next I
        If Total <> 55# Then IntegrityOK = False
    Next R

    Expected = REPS / 10
    For I = 1 To 10
        ChiSq = ChiSq + (Counts(I) - Expected) ^ 2 / Expected
    Next I

    CheckTrue G, "Shuffle keeps every value", IntegrityOK, IIf(IntegrityOK, "sum = 55 in every run", "values lost")
    CheckBelow G, "Shuffle is uniform: chi-square of 10 positions (99.9%)", 27.88, ChiSq
End Sub

'---------------------------------------------------------------------------
' Full simulation of a small model with known exact results
'---------------------------------------------------------------------------

Private Sub TestSimulation()
    Const G As String = "Simulation"
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim XS As Worksheet
    Dim Sim As ClsSimulation
    Dim Res1 As Variant
    Dim Res2 As Variant
    Dim C As Long
    Dim Same As Boolean

    On Error GoTo Failed
    Set WB = Workbooks.Add(xlWBATWorksheet)
    Set WS = WB.Worksheets(1)
    WS.Name = "Model"
    BuildModel WS

    WB.Activate
    Set XS = SetUpXLRisk
    XS.Range("Seed").Value = SIM_SEED
    XS.Range("Iterations").Value = SIM_ITERATIONS
    XS.Range("ScreenUpdate").Value = False
    XS.Range("LatinHypercube").Value = True
    AddTestOutput XS, WS.Range("F3"), "Profit"
    AddTestOutput XS, WS.Range("F4"), "Normal"
    AddTestOutput XS, WS.Range("F5"), "CorrA"
    AddTestOutput XS, WS.Range("F6"), "CorrB"
    AddTestOutput XS, WS.Range("F7"), "Loss"
    XS.Calculate

    ' First run
    Set Sim = RunSimulation(WB)
    If Sim.OutputCount <> 8 Then
        WriteRow G, "Simulation found the 8 test outputs", 8, Sim.OutputCount, 0, False
        WB.Close SaveChanges:=False
        Exit Sub
    End If
    Res1 = Sim.LastResults
    C = 1 + Sim.InputCount  ' last column before the outputs

    CheckNear G, "Iterations completed", SIM_ITERATIONS, UBound(Res1, 1), 0
    ' Price x quantity: price 45/50/56 (25%/50%/25%), quantity 8,000/10,000/12,000
    ' (30%/50%/20%), unit cost 32, fixed cost 150,000.
    ' Exact mean profit 28,850 and exact probability of loss 35%.
    CheckNear G, "Mean profit of the price x quantity model", 28850#, ColMean(Res1, C + 1), 2000#
    CheckNear G, "Probability of loss of the price x quantity model", 0.35, ColMean(Res1, C + 5), 0.02
    CheckNear G, "RiskNormal(100, 15): mean", 100#, ColMean(Res1, C + 2), 0.64
    CheckNear G, "RiskNormal(100, 15): standard deviation", 15#, ColStd(Res1, C + 2), 0.45
    CheckNear G, "Correlation induced by RiskCorrmat (target 0.7)", 0.7, ColCorrel(Res1, C + 3, C + 4), 0.03
    WriteInfo G, "Seconds for " & Format(SIM_ITERATIONS, "#,##0") & " iterations", Sim.LoopSeconds

    ' Functions with the same names as in @RISK
    Const F As String = "@RISK functions"
    CheckNear F, "RiskOutput(""Name"") in a formula adds the output", ColMean(Res1, C + 1), ColMean(Res1, C + 6), 0.000001
    CheckNear F, "RiskPoisson(4): mean", 4#, ColMean(Res1, C + 7), 0.09
    CheckNear F, "RiskLognorm2(0, 0.25): mean", 1.0317434074991, ColMean(Res1, C + 8), 0.012
    CheckNear F, "RiskMean(cell)", ColMean(Res1, C + 1), CellNumber(WS.Range("K1")), 0.001
    CheckNear F, "RiskStdDev(cell)", ColStd(Res1, C + 1), CellNumber(WS.Range("K2")), 0.001
    CheckNear F, "RiskMin(cell)", WorksheetFunction.Min(ColValues(Res1, C + 1)), CellNumber(WS.Range("K3")), 0.001
    CheckNear F, "RiskMax(cell)", WorksheetFunction.Max(ColValues(Res1, C + 1)), CellNumber(WS.Range("K4")), 0.001
    CheckNear F, "RiskPercentile(cell, 0.5)", _
              WorksheetFunction.Percentile_Inc(ColValues(Res1, C + 1), 0.5), CellNumber(WS.Range("K5")), 0.001
    CheckNear F, "RiskTarget(cell, 0) = probability of loss", ColMean(Res1, C + 5), CellNumber(WS.Range("K6")), 0.000001
    CheckTrue F, "Statistics of a cell without results return #N/A", IsError(WS.Range("K7").Value), _
              IIf(IsError(WS.Range("K7").Value), "#N/A", "a value")
    CheckTrue G, "Other open workbooks recalculate again after the run", _
              mReport.EnableCalculation, IIf(mReport.EnableCalculation, "restored", "still suspended")

    ' Second run with the same seed must reproduce every number
    Set Sim = RunSimulation(WB)
    Res2 = Sim.LastResults
    Same = SameResults(Res1, Res2)
    CheckTrue G, "A fixed Seed reproduces the whole simulation", Same, IIf(Same, "identical", "different")

    WB.Close SaveChanges:=False
    Exit Sub

Failed:
    WriteRow G, "Simulation ran without errors", "no error", Err.Description, "", False
    On Error Resume Next
    Set gSimulation = Nothing
    If Not WB Is Nothing Then WB.Close SaveChanges:=False
End Sub

Private Sub TestVersion12()
' Version 1.2: RiskDiscrete in any order, Bernoulli and Binomial without rounding, RiskTriGen,
' data tables off by default, histogram for older Excel, tornado of the change in output mean,
' Compare Runs and Swap Functions.
    Const G As String = "Version 1.2"
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim XS As Worksheet
    Dim RR As Worksheet
    Dim CS As Worksheet
    Dim Sim As ClsSimulation
    Dim Res As Variant
    Dim V As Variant
    Dim T As Range
    Dim C As Long
    Dim I As Long
    Dim N As Long
    Dim N45 As Long
    Dim N56 As Long
    Dim OnlyListed As Boolean
    Dim OldSample As Boolean
    Dim Original As String
    Dim Detail As String

    OldSample = ProduceRandomSample
    On Error GoTo Failed
    Set WB = Workbooks.Add(xlWBATWorksheet)
    Set WS = WB.Worksheets(1)
    WS.Name = "Model12"
    With WS
        .Range("A1").Formula = "=RiskDiscrete({56,45,50},{0.25,0.25,0.5})"
        .Range("A2").Formula = "=RiskBernoulli(0.15)"
        .Range("A3").Formula = "=RiskBinomial(10,0.15)"
        .Range("A4").Formula = "=RiskTriGen(80,100,130,0.1,0.9)"
        .Range("A5").Formula = "=2*A1+A4"
        .Range("A6").Formula = "=A1"
        .Range("A7").Formula = "=A4"
        .Range("A8").Formula = "=A2"
    End With

    ' Expected values, with Show Samples off
    ProduceRandomSample = False
    WB.Activate
    Application.CalculateFull
    CheckNear G, "RiskDiscrete with values in any order: expected value", 50.25, CellNumber(WS.Range("A1")), 0.000000001
    CheckNear G, "RiskBernoulli(0.15) shows 0.15, not a rounded 0", 0.15, CellNumber(WS.Range("A2")), 0.000000001
    CheckNear G, "RiskBinomial(10, 0.15) shows 1.5, not a rounded 2", 1.5, CellNumber(WS.Range("A3")), 0.000000001
    CheckNear G, "RiskTriGen(80, 100, 130, 10%, 90%): expected value", 104.296539858006, CellNumber(WS.Range("A4")), 0.0001

    ' A new XLRisk sheet has data tables off
    Set XS = SetUpXLRisk
    CheckTrue G, "New XLRisk sheet: Calculate data tables is off", (XS.Range("CalcDataTables").Value = False), _
              CStr(XS.Range("CalcDataTables").Value)
    XS.Range("Seed").Value = SIM_SEED
    XS.Range("Iterations").Value = SIM_ITERATIONS
    XS.Range("ScreenUpdate").Value = False
    XS.Range("LatinHypercube").Value = True
    AddTestOutput XS, WS.Range("A5"), "Sum"
    AddTestOutput XS, WS.Range("A6"), "Disc"
    AddTestOutput XS, WS.Range("A7"), "TriGen"
    AddTestOutput XS, WS.Range("A8"), "Bern"
    XS.Calculate

    gForceHistogramFallback = True
    Set Sim = RunSimulation(WB)
    gForceHistogramFallback = False
    Res = Sim.LastResults
    C = 1 + Sim.InputCount

    ' RiskDiscrete with unordered values
    V = ColValues(Res, C + 2)
    N = UBound(V)
    OnlyListed = True
    For I = 1 To N
        If V(I) = 45 Then
            N45 = N45 + 1
        ElseIf V(I) = 56 Then
            N56 = N56 + 1
        ElseIf V(I) <> 50 Then
            OnlyListed = False
        End If
    Next I
    CheckTrue G, "RiskDiscrete in any order: only the listed values", OnlyListed, IIf(OnlyListed, "45, 50, 56", "other values")
    CheckNear G, "RiskDiscrete in any order: share of 45", 0.25, N45 / N, 0.02
    CheckNear G, "RiskDiscrete in any order: share of 56", 0.25, N56 / N, 0.02
    CheckNear G, "RiskBernoulli(0.15): mean in the simulation", 0.15, ColMean(Res, C + 4), 0.015
    CheckNear G, "RiskTriGen: 10th percentile is the bottom value 80", 80#, _
              WorksheetFunction.Percentile_Inc(ColValues(Res, C + 3), 0.1), 1.5
    CheckNear G, "RiskTriGen: 90th percentile is the top value 130", 130#, _
              WorksheetFunction.Percentile_Inc(ColValues(Res, C + 3), 0.9), 1.5

    Set RR = WB.Worksheets("Risk Results 1")

    ' Histogram for Excel 2013 and older (forced)
    Set T = Nothing
    On Error Resume Next
    Set T = RR.Range("OutDistTable1")
    On Error GoTo Failed
    CheckTrue G, "Histogram for older Excel: table of counts", Not T Is Nothing, IIf(T Is Nothing, "missing", "found")
    If Not T Is Nothing Then
        CheckNear G, "Histogram for older Excel: counts add up to the iterations", SIM_ITERATIONS, _
                  WorksheetFunction.Sum(T.Offset(1, 1).Resize(20, 1)), 0
    End If

    ' Tornado of the change in output mean, output Sum = 2 x Discrete + TriGen
    Set T = Nothing
    On Error Resume Next
    Set T = RR.Range("MeanTornadoTable1")
    On Error GoTo Failed
    CheckTrue G, "Output mean tornado: table for the first output", Not T Is Nothing, IIf(T Is Nothing, "missing", "found")
    If Not T Is Nothing Then
        Detail = CStr(T.Offset(1, 0).Value) & " / " & CStr(T.Offset(2, 0).Value) & " / " & CStr(T.Offset(3, 0).Value)
        CheckTrue G, "Output mean tornado: only the 2 inputs that change the output, largest swing first", _
                  (InStr(1, T.Offset(1, 0).Value, "$A$4") > 0) And (InStr(1, T.Offset(2, 0).Value, "$A$1") > 0) _
                  And Len(CStr(T.Offset(3, 0).Value)) = 0, Detail
        CheckNear G, "Output mean tornado: swing of RiskTriGen (exact 63.4)", 63.42, CellNumber(T.Offset(1, 5)), 5
        CheckNear G, "Output mean tornado: swing of 2 x RiskDiscrete (exact 22)", 22#, CellNumber(T.Offset(2, 5)), 4
        CheckNear G, "Output mean tornado: share of variance of RiskTriGen (exact 84.9%)", 0.849, CellNumber(T.Offset(1, 6)), 0.03
        CheckNear G, "Output mean tornado: shares add up to 100%", 1#, _
                  CellNumber(T.Offset(1, 6)) + CellNumber(T.Offset(2, 6)), 0.000001
    End If

    ' Compare Runs: a second run with the same seed must give the same column
    Set Sim = RunSimulation(WB)
    Set CS = CompareRunsFor(WB, AddressWithSheet(WS.Range("A5")))
    CheckTrue G, "Compare Runs: sheet with the two runs", Not CS Is Nothing, IIf(CS Is Nothing, "missing", "found")
    If Not CS Is Nothing Then
        CheckTrue G, "Compare Runs: one column per run", _
                  (CS.Range("B3").Value = "Risk Results 1") And (CS.Range("C3").Value = "Risk Results 2"), _
                  CStr(CS.Range("B3").Value) & " / " & CStr(CS.Range("C3").Value)
        CheckNear G, "Compare Runs: mean of the first run", ColMean(Res, C + 1), CellNumber(CS.Range("B4")), 0.000001
        CheckNear G, "Compare Runs: same seed gives the same mean", CellNumber(CS.Range("B4")), CellNumber(CS.Range("C4")), 0.000001
        CheckNear G, "Compare Runs: chart with one curve per run", 2, CS.ChartObjects(1).Chart.SeriesCollection.Count, 0
    End If

    ' Swap Functions
    Original = WS.Range("A4").Formula
    ProduceRandomSample = False
    N = SwapOut(WB)
    CheckNear G, "Swap Functions: the 4 Risk functions are replaced", 4, N, 0
    CheckTrue G, "Swap Functions: the cell holds the expected value", _
              (InStr(1, WS.Range("A1").Formula, "Risk", vbTextCompare) = 0) And _
              (Abs(CellNumber(WS.Range("A1")) - 50.25) < 0.000000001), WS.Range("A1").Formula
    N = SwapIn(WB)
    CheckTrue G, "Swap Functions: the second press restores the formulas", _
              (WS.Range("A4").Formula = Original) And Not HasSwapSheet(WB), WS.Range("A4").Formula

    ProduceRandomSample = OldSample
    WB.Close SaveChanges:=False
    Exit Sub

Failed:
    WriteRow G, "Version 1.2 checks ran without errors", "no error", Err.Description, "", False
    On Error Resume Next
    gForceHistogramFallback = False
    ProduceRandomSample = OldSample
    Set gSimulation = Nothing
    If Not WB Is Nothing Then WB.Close SaveChanges:=False
End Sub

Private Sub BuildModel(WS As Worksheet)
    With WS
        .Range("A1").Value = 45
        .Range("A2").Value = 50
        .Range("A3").Value = 56
        .Range("B1").Value = 0.25
        .Range("B2").Value = 0.5
        .Range("B3").Value = 0.25
        .Range("C1").Value = 8000
        .Range("C2").Value = 10000
        .Range("C3").Value = 12000
        .Range("D1").Value = 0.3
        .Range("D2").Value = 0.5
        .Range("D3").Value = 0.2
        .Range("H1").Value = 1
        .Range("I1").Value = 0.7
        .Range("H2").Value = 0.7
        .Range("I2").Value = 1
        .Range("F1").Formula = "=RiskDiscrete(A1:A3,B1:B3)"
        .Range("F2").Formula = "=RiskDiscrete(C1:C3,D1:D3)"
        .Range("F3").Formula = "=(F1-32)*F2-150000"
        .Range("F4").Formula = "=RiskNormal(100,15)"
        .Range("F5").Formula = "=RiskNormal(0,1,RiskCorrmat(H1:I2,1))"
        .Range("F6").Formula = "=RiskNormal(0,1,RiskCorrmat(H1:I2,2))"
        .Range("F7").Formula = "=IF(F3<0,1,0)"
        ' Outputs marked in the formula, as in @RISK
        .Range("F8").Formula = "=RiskOutput(""ProfitCopy"")+F3"
        .Range("F9").Formula = "=RiskOutput(""Poisson"")+RiskPoisson(4)"
        .Range("F10").Formula = "=RiskOutput(""LN2"")+RiskLognorm2(0,0.25)"
        ' Statistics of the latest simulation
        .Range("K1").Formula = "=RiskMean(F3)"
        .Range("K2").Formula = "=RiskStdDev(F3)"
        .Range("K3").Formula = "=RiskMin(F3)"
        .Range("K4").Formula = "=RiskMax(F3)"
        .Range("K5").Formula = "=RiskPercentile(F3,0.5)"
        .Range("K6").Formula = "=RiskTarget(F3,0)"
        .Range("K7").Formula = "=RiskMean(A1)"
    End With
End Sub

Private Sub AddTestOutput(XS As Worksheet, Cell As Range, ByVal OutputName As String)
' Registers an output the same way as the Add Output button
    Dim R As Range
    Set R = XS.Range("RiskOutputs").CurrentRegion
    Set R = R.Rows(R.Rows.Count).Offset(1)
    R.Cells(1, 1).Formula = "=AddressWithSheet(" & AddressWithSheet(Cell) & ")"
    R.Cells(1, 2).Value = OutputName
End Sub

Private Function RunSimulation(WB As Workbook) As ClsSimulation
' Runs a simulation the same way as the Run button
    Dim Sim As ClsSimulation
    WB.Activate
    Set Sim = New ClsSimulation
    Set gSimulation = Sim
    Sim.Run
    Set gSimulation = Nothing
    Set RunSimulation = Sim
End Function

'---------------------------------------------------------------------------
' Statistics helpers
'---------------------------------------------------------------------------

Private Function ColMean(R As Variant, ByVal Col As Long) As Double
    Dim I As Long
    Dim S As Double
    For I = LBound(R, 1) To UBound(R, 1)
        S = S + R(I, Col)
    Next I
    ColMean = S / (UBound(R, 1) - LBound(R, 1) + 1)
End Function

Private Function ColStd(R As Variant, ByVal Col As Long) As Double
    Dim I As Long
    Dim M As Double
    Dim S As Double
    M = ColMean(R, Col)
    For I = LBound(R, 1) To UBound(R, 1)
        S = S + (R(I, Col) - M) ^ 2
    Next I
    ColStd = Sqr(S / (UBound(R, 1) - LBound(R, 1)))
End Function

Private Function ColCorrel(R As Variant, ByVal ColX As Long, ByVal ColY As Long) As Double
    Dim I As Long
    Dim MX As Double
    Dim MY As Double
    Dim SXY As Double
    Dim SXX As Double
    Dim SYY As Double
    MX = ColMean(R, ColX)
    MY = ColMean(R, ColY)
    For I = LBound(R, 1) To UBound(R, 1)
        SXY = SXY + (R(I, ColX) - MX) * (R(I, ColY) - MY)
        SXX = SXX + (R(I, ColX) - MX) ^ 2
        SYY = SYY + (R(I, ColY) - MY) ^ 2
    Next I
    ColCorrel = SXY / Sqr(SXX * SYY)
End Function

Private Function ColValues(R As Variant, ByVal Col As Long) As Variant
' One column of the results as a 1-based array of Doubles
    Dim I As Long
    Dim V() As Double
    ReDim V(1 To UBound(R, 1) - LBound(R, 1) + 1)
    For I = LBound(R, 1) To UBound(R, 1)
        V(I - LBound(R, 1) + 1) = R(I, Col)
    Next I
    ColValues = V
End Function

Private Function CellNumber(Cell As Range) As Double
' Value of a cell, or a value that fails every check when the cell holds an error
    If IsError(Cell.Value) Or Not IsNumeric(Cell.Value) Then
        CellNumber = -1E+300
    Else
        CellNumber = Cell.Value
    End If
End Function

Private Function SameResults(A As Variant, B As Variant) As Boolean
    Dim I As Long
    Dim J As Long
    SameResults = False
    If UBound(A, 1) <> UBound(B, 1) Or UBound(A, 2) <> UBound(B, 2) Then Exit Function
    For I = LBound(A, 1) To UBound(A, 1)
        For J = LBound(A, 2) To UBound(A, 2)
            If A(I, J) <> B(I, J) Then Exit Function
        Next J
    Next I
    SameResults = True
End Function

'---------------------------------------------------------------------------
' Report
'---------------------------------------------------------------------------

Private Sub CreateReport()
    Dim WB As Workbook
    Set WB = Workbooks.Add(xlWBATWorksheet)
    Set mReport = WB.Worksheets(1)
    mReport.Name = "XLRisk Self Tests"
    With mReport
        .Range("A1").Value = "XLRisk " & XLRISK_VERSION & " self tests"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A3").Value = "Group"
        .Range("B3").Value = "Check"
        .Range("C3").Value = "Expected"
        .Range("D3").Value = "Obtained"
        .Range("E3").Value = "Tolerance"
        .Range("F3").Value = "Result"
        .Range("A3:F3").Font.Bold = True
    End With
    mRow = 4
End Sub

Private Sub WriteRow(ByVal Group As String, ByVal CheckName As String, ByVal Expected As Variant, _
                     ByVal Obtained As Variant, ByVal Tolerance As Variant, ByVal Passed As Boolean)
    With mReport
        .Cells(mRow, 1).Value = Group
        .Cells(mRow, 2).Value = CheckName
        .Cells(mRow, 3).Value = AsText(Expected)
        .Cells(mRow, 4).Value = AsText(Obtained)
        .Cells(mRow, 5).Value = Tolerance
        With .Cells(mRow, 6)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            If Passed Then
                .Value = "PASS"
                .Interior.Color = RGB(0, 140, 0)
            Else
                .Value = "FAIL"
                .Interior.Color = RGB(200, 0, 0)
            End If
        End With
    End With
    mChecks = mChecks + 1
    If Not Passed Then mFailures = mFailures + 1
    mRow = mRow + 1
End Sub

Private Function AsText(ByVal V As Variant) As Variant
' A text that looks like a formula (a formula shown as detail) is written as text
    If VarType(V) = vbString Then
        If Left(V, 1) = "=" Then V = "'" & V
    End If
    AsText = V
End Function

Private Sub WriteInfo(ByVal Group As String, ByVal Label As String, ByVal Value As Variant)
' Informative row: not counted as a check
    With mReport
        .Cells(mRow, 1).Value = Group
        .Cells(mRow, 2).Value = Label
        .Cells(mRow, 4).Value = Value
        With .Cells(mRow, 6)
            .Value = "INFO"
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 56, 100)
        End With
    End With
    mRow = mRow + 1
End Sub

Private Sub CheckNear(ByVal Group As String, ByVal CheckName As String, ByVal Expected As Double, _
                      ByVal Obtained As Double, ByVal Tolerance As Double)
    WriteRow Group, CheckName, Expected, Obtained, Tolerance, (Abs(Obtained - Expected) <= Tolerance)
End Sub

Private Sub CheckBelow(ByVal Group As String, ByVal CheckName As String, ByVal Limit As Double, ByVal Obtained As Double)
    WriteRow Group, CheckName, "< " & CStr(Limit), Obtained, "", (Obtained < Limit)
End Sub

Private Sub CheckTrue(ByVal Group As String, ByVal CheckName As String, ByVal Condition As Boolean, ByVal Detail As String)
    WriteRow Group, CheckName, "TRUE", Detail, "", Condition
End Sub

Private Sub FinishReport(ByVal Seconds As Double)
    With mReport
        .Range("A2").Value = mChecks & " checks, " & mFailures & " failed, " & Format(Seconds, "0") & " seconds"
        .Columns("A:F").AutoFit
        .Parent.Activate
        .Activate
    End With
    If mFailures = 0 Then
        MsgBox "All " & mChecks & " XLRisk self tests passed.", vbInformation, "XLRisk " & XLRISK_VERSION
    Else
        MsgBox mFailures & " of " & mChecks & " XLRisk self tests failed. See the report.", _
               vbExclamation, "XLRisk " & XLRISK_VERSION
    End If
End Sub
