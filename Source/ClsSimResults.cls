VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSimResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOutSheet As Worksheet
Private mRiskInputs As Collection
Private mRiskOutputs As Collection
Private mIterations As Long
Private mSimResults() As Variant
Private mDependencyMatrix() As Boolean

Private Const ChartHeight As Integer = 22 'rows
Private Const ChartWidth As Integer = 10 'columns


Private Sub Class_Initialize()
    Set mOutSheet = CreateOutputSheet
End Sub

Public Sub Init(RiskInputs As Collection, RiskOutputs As Collection, Iterations As Long, SimResults() As Variant)
    Set mRiskInputs = RiskInputs
    Set mRiskOutputs = RiskOutputs
    mIterations = Iterations
    mSimResults = SimResults
End Sub

Public Sub ProduceResults()
    mOutSheet.Activate
    InitialiseResults
    'Produce Output
    mOutSheet.Cells(4, 2).Resize(mIterations, 1 + mRiskInputs.Count + mRiskOutputs.Count) = mSimResults
    ' Produce Statistics
    ProduceStatistics
    ' Calculate before producing the graphs
    Application.Calculate
    ' Produce Cumulative Distributions
    ProduceCumulativeDistributions
    ' Produce histograms (Excel 2016+ histogram chart, or a column chart of counts in older Excel)
    ProduceHistograms
    ' Produce Tornado Diagram
    If mRiskInputs.Count > 1 Then
        ProduceTornadoDiagram
        ProduceMeanTornado
    End If
End Sub

Private Sub InitialiseResults()
    Dim RiskInput As ClsRiskInput
    Dim Curr As Range
    Dim RiskOutput As ClsRiskOutput
    
    With mOutSheet
        .Range("C1").Value = "Inputs"
        .Range("B3").Value = "Iterations"
    End With
      
    '  Setup risk inputs
    Set Curr = mOutSheet.Range("C2")
    mOutSheet.Names.Add Name:="InputResults", RefersTo:=Curr.Offset(2)
    For Each RiskInput In mRiskInputs
        Curr.Value = NameOrAddress(RiskInput.Cell)
        If Curr.PrefixCharacter = "'" Then Curr.Value = "''" & Curr.Value
        Curr.Columns.AutoFit
        Curr.ColumnWidth = WorksheetFunction.Max(Curr.ColumnWidth, 10)
        Curr.HorizontalAlignment = xlCenter
        Curr.Offset(1, 0) = Right(RiskInput.Cell.Formula, Len(RiskInput.Cell.Formula) - 1)
        Set Curr = Curr.Offset(0, 1)
    Next RiskInput
    'Format Input
    With mOutSheet.Range("C1").Resize(1, mRiskInputs.Count)
        If mRiskInputs.Count > 1 Then .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With mOutSheet.Range("InputResults")
        ThickBorders .Offset(0, -1).Resize(mIterations, 1)
        ThickBorders .Offset(-2, -1).Resize(2, 1)
        ThickBorders .Offset(-3).Resize(1, mRiskInputs.Count)
        ThickBorders .Resize(mIterations, mRiskInputs.Count)
        ThickBorders .Offset(-2).Resize(2, mRiskInputs.Count)
    End With
    
    '   Setup risk outputs
    Curr.Offset(-1, 0) = "Outputs"
    mOutSheet.Names.Add Name:="OutputResults", RefersTo:=Curr.Offset(2)
    For Each RiskOutput In mRiskOutputs
        Curr = AddressWithSheet(RiskOutput.Cell)
        Curr.Offset(1, 0) = RiskOutput.Name
        Set Curr = Curr.Offset(0, 1)
    Next RiskOutput
    'Format Output
    With mOutSheet.Range("C1").Offset(0, 1).Resize(1, mRiskOutputs.Count)
        If mRiskOutputs.Count > 1 Then .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    '   AutoFit
    With mOutSheet.Range("OutputResults").Offset(-2).Resize(2, mRiskOutputs.Count)
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    With mOutSheet.Range("OutputResults")
        ThickBorders .Offset(-3).Resize(1, mRiskOutputs.Count)
        ThickBorders .Resize(mIterations, mRiskOutputs.Count)
        ThickBorders .Offset(-2).Resize(2, mRiskOutputs.Count)
    End With
    
    '   Setup Simulation Statistics
    Set Curr = Curr.Offset(0, 2)
    Curr.Offset(-1, 0) = "Simulation Statistics"
    mOutSheet.Names.Add Name:="SimStats", RefersTo:=Curr.Offset(2, -1)
    Curr.Resize(2, mRiskOutputs.Count).Value = _
        mOutSheet.Range("OutputResults").Offset(-2).Resize(2, mRiskOutputs.Count).Value
    'Format Simulation Results
    With mOutSheet.Range("SimStats").Offset(-3).Resize(1, mRiskOutputs.Count + 1)
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    'AutoFit
    With mOutSheet.Range("SimStats").Offset(-2, 1).Resize(2, mRiskOutputs.Count)
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    With mOutSheet.Range("SimStats")
        ThickBorders .Offset(-3).Resize(1, mRiskOutputs.Count + 1)
        ThickBorders .Resize(14, 1)
        ThickBorders .Offset(0, 1).Resize(14, mRiskOutputs.Count)
        ThickBorders .Offset(-2, 1).Resize(2, mRiskOutputs.Count)
        ThickBorders .Offset(-2).Resize(16, mRiskOutputs.Count + 1)
    End With
    
    '  Setup Output Cumulative Distributions
    Set Curr = Curr.Offset(0, 1 + mRiskOutputs.Count)
    Curr.Offset(-1, 0) = "Output Cumulative Distributions"
    mOutSheet.Names.Add Name:="OutCumDist", RefersTo:=Curr.Offset(1, 0)
    With Curr.Offset(-1, 0).Resize(1, ChartWidth)
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    '  Setup Output Histograms
    Set Curr = Curr.Offset(0, ChartWidth + 1)
    Curr.Offset(-1, 0) = "Output Distributions"
    mOutSheet.Names.Add Name:="OutDist", RefersTo:=Curr.Offset(1, 0)
    With Curr.Offset(-1, 0).Resize(1, ChartWidth)
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    If mRiskInputs.Count > 1 Then
        '  Setup Tornado Diagram
        Set Curr = Curr.Offset(0, ChartWidth + 1)
        Curr.Offset(-1, 0) = "Risk Tornado Diagram"
        mOutSheet.Names.Add Name:="Tornado", RefersTo:=Curr.Offset(1, 0)
        With Curr.Offset(-1, 0).Resize(1, ChartWidth)
            .Merge
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        '  Setup the tornado of the change in output mean, after the hidden sorted table of the tornado
        Set Curr = Curr.Offset(0, ChartWidth + mRiskOutputs.Count + 3)
        Curr.Offset(-1, 0) = "Output Mean Tornado"
        mOutSheet.Names.Add Name:="MeanTornado", RefersTo:=Curr.Offset(1, 0)
        With Curr.Offset(-1, 0).Resize(1, ChartWidth)
            .Merge
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If
    
    ' Add Hyperlinks
    Set Curr = mOutSheet.Range("A1")
    With Curr
        .Value = "Navigation"
        .Font.Bold = True
    End With
    
    With mOutSheet.Hyperlinks
        .Add Anchor:=Curr.Offset(1), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("OutputResults").Offset(-3).Address, TextToDisplay:="Risk Outputs"
        .Add Anchor:=Curr.Offset(2), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("SimStats").Offset(-3).Address, TextToDisplay:="Simulation Statistics"
        .Add Anchor:=Curr.Offset(3), Address:=vbNullString, _
            SubAddress:="Percentiles", TextToDisplay:="Risk Output Percentiles"
        .Add Anchor:=Curr.Offset(4), Address:=vbNullString, _
            SubAddress:="PercentRank", TextToDisplay:="Calculate Percent Rank"
        .Add Anchor:=Curr.Offset(5), Address:=vbNullString, _
            SubAddress:="Correlations", TextToDisplay:="Correlations"
        .Add Anchor:=Curr.Offset(6), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("OutCumDist").Offset(-2).Address, TextToDisplay:="Output Cumulative Distributions"
        .Add Anchor:=Curr.Offset(7), Address:=vbNullString, _
            SubAddress:=mOutSheet.Range("OutDist").Offset(-2).Address, TextToDisplay:="Output Distributions"
        If mRiskInputs.Count > 1 Then
            .Add Anchor:=Curr.Offset(8), Address:=vbNullString, _
                SubAddress:=mOutSheet.Range("Tornado").Offset(-2).Address, TextToDisplay:="Risk Tornado Diagram"
            .Add Anchor:=Curr.Offset(9), Address:=vbNullString, _
                SubAddress:=mOutSheet.Range("MeanTornado").Offset(-2).Address, TextToDisplay:="Output Mean Tornado"
        End If
    End With
    mOutSheet.Columns(1).AutoFit
End Sub

Private Sub StatHelper(Cell As Range, StatName As String, StatFormula As String, Address As String)
    Cell = StatName
    Cell.Offset(0, 1).Formula = "=" & StatFormula & "(" & Address & ")"
End Sub

Private Sub StatAggregateHelper(Cell As Range, StatName As String, FormulaCode As Integer, Address As String)
    Cell = StatName
    Cell.Offset(0, 1).Formula = "=AGGREGATE(" & CStr(FormulaCode) & ",6," & Address & ")"
End Sub

Private Sub ProduceStatistics()
'  Use the Aggregate Excel function so that stats can be calculated even if the outputs contain a few errors
    Dim FirstOutput As Range
    Dim Cell As Range
    Dim Address As String
    Dim I As Long
    
    Set FirstOutput = mOutSheet.Range("OutputResults").Resize(mIterations, 1)
    Set Cell = mOutSheet.Range("SimStats")
    
    Address = FirstOutput.Address(True, False)
    StatAggregateHelper Cell, "Mean", 1, Address
    StatAggregateHelper Cell.Offset(1), "Median", 12, Address
    StatAggregateHelper Cell.Offset(2), "Mode", 13, Address
    StatAggregateHelper Cell.Offset(3), "Std. Deviation", 7, Address
    StatAggregateHelper Cell.Offset(4), "Variance", 10, Address
    StatHelper Cell.Offset(5), "Kurtosis", "KURT", Address
    StatHelper Cell.Offset(6), "Skewness", "Skew", Address
    StatAggregateHelper Cell.Offset(7), "Minimum", 5, Address
    StatAggregateHelper Cell.Offset(8), "Maximum", 4, Address
    Cell.Offset(9) = "Range"
    Cell.Offset(9, 1).Formula = "=" & Cell.Offset(8, 1).Address(False, False) & "-" & Cell.Offset(7, 1).Address(False, False)
    StatHelper Cell.Offset(10), "Count", "Count", Address
    Cell.Offset(11) = "Error Count"
    Cell.Offset(11, 1).FormulaArray = "=COUNT(IF(ISERROR(" & Address & "), 1, """"))"
    Cell.Offset(12) = "Std. Error"
    Cell.Offset(12, 1).Formula = "=" & Cell.Offset(3, 1).Address(False, False) & "/SQRT(" & Cell.Offset(10, 1).Address(False, False) & ")"
    Cell.Offset(13) = "Confidence Level (95%)"
    Cell.Offset(13, 1).Formula = "=CONFIDENCE.T(5%," & Cell.Offset(3, 1).Address(False, False) & "," & Cell.Offset(10, 1).Address(False, False) & ")"
    If mRiskOutputs.Count > 1 Then
        For I = 2 To mRiskOutputs.Count
            Cell.Offset(0, I).Resize(14, 1).FormulaR1C1 = Cell.Offset(0, 1).Resize(14, 1).FormulaR1C1
        Next I
    End If
    
    'Percentiles
    ProducePercentiles Cell.Offset(15)
    
    'Percent Rank
    ProducePercentRank Cell.Offset(60)
    
    'Correlations
    ProduceCorrelations Cell.Offset(66)
    
    'AutoFit
    mOutSheet.Columns(Cell.Column).AutoFit
End Sub

Private Sub ProducePercentiles(Cell As Range)
    Dim Perc As Integer
    Dim PCount As Integer
    Dim Address As String
    
    mOutSheet.Names.Add Name:="Percentiles", RefersTo:=Cell.Offset(3).Resize(41, mRiskOutputs.Count + 1)
    Perc = 0
    For PCount = 1 To 41
        With Cell.Offset(PCount + 2)
            .Value = Perc / 1000
            .NumberFormat = "0.0%"
            .HorizontalAlignment = xlRight
        End With
        Perc = Perc + 25
    Next PCount
    'Address of First Output with absolute rows
    Address = mOutSheet.Range("OutputResults").Resize(mIterations, 1).Address(True, False)
    Cell.Offset(3, 1).Formula = "=AGGREGATE(16,6," & Address & "," & Cell.Offset(3).Address(False, True) & ")"
    Cell.Offset(3, 1).Copy Cell.Offset(3, 1).Resize(41, mRiskOutputs.Count)
    ThickBorders Cell.Offset(3).Resize(41, 1)
    ThickBorders Cell.Offset(3, 1).Resize(41, mRiskOutputs.Count)
    mOutSheet.Range("SimStats").Offset(-3).Resize(3, 1 + mRiskOutputs.Count).Copy Cell
    Cell.Value = "Percentiles"
End Sub

Private Sub ProducePercentRank(Cell As Range)
    Dim Address As String
    
    mOutSheet.Names.Add Name:="PercentRank", RefersTo:=Cell
    Cell.Offset(3) = "Enter Value:"
    mOutSheet.Range("Percentiles").Cells(3, 2).Resize(1, mRiskOutputs.Count).Calculate
    Cell.Offset(3, 1).Resize(1, mRiskOutputs.Count).Value = _
        mOutSheet.Range("Percentiles").Cells(3, 2).Resize(1, mRiskOutputs.Count).Value '5th Percentile
    Cell.Offset(4) = "% rank:"
    Address = mOutSheet.Range("OutputResults").Resize(mIterations, 1).Address(True, False)
    Cell.Offset(4, 1).Formula = "=PERCENTRANK.INC(" & Address & "," & Cell.Offset(3, 1).Address(True, False) & ")"
    If mRiskOutputs.Count > 1 Then Cell.Offset(4, 1).Copy Cell.Offset(4, 2).Resize(1, mRiskOutputs.Count - 1)
    Cell.Offset(4, 1).Resize(1, mRiskOutputs.Count).NumberFormat = "0.00%"
    Cell.Offset(3).Resize(2, 1).HorizontalAlignment = xlRight
    ThickBorders Cell.Offset(3).Resize(2, 1)
    ThickBorders Cell.Offset(3, 1).Resize(2, mRiskOutputs.Count)
    mOutSheet.Range("SimStats").Offset(-3).Resize(3, 1 + mRiskOutputs.Count).Copy Cell
    Cell = "Percent Rank"
End Sub

Sub ProduceCorrelations(Cell As Range)
    Dim FirstInput As Range
    Dim InputNames As Range
    Dim OutAddress As String
    Dim I As Long
    Dim J As Long
    
    Cell.Font.Bold = True
    mOutSheet.Names.Add Name:="Correlations", RefersTo:=Cell
    
    Set InputNames = mOutSheet.Range("InputResults").Offset(-2)
    Set FirstInput = mOutSheet.Range("InputResults").Resize(mIterations, 1)
    OutAddress = mOutSheet.Range("OutputResults").Resize(mIterations, 1).Address(True, False)
    For I = 1 To mRiskInputs.Count
        Cell.Offset(2 + I).Value = InputNames.Offset(0, I - 1).Value
        If Cell.Offset(2 + I).PrefixCharacter = "'" Then Cell.Offset(2 + I).Value = "''" & Cell.Offset(2 + I).Value

        Cell.Offset(2 + I, 1).Formula = _
        "=CORREL(" & FirstInput.Offset(0, I - 1).Address & "," & OutAddress & ")"
    Next I
    If mRiskOutputs.Count > 1 Then _
        Cell.Offset(3, 1).Resize(mRiskInputs.Count, 1).Copy _
        Cell.Offset(3, 2).Resize(mRiskInputs.Count, mRiskOutputs.Count - 1)
    ThickBorders Cell.Offset(3).Resize(mRiskInputs.Count, 1)
    ThickBorders Cell.Offset(3, 1).Resize(mRiskInputs.Count, mRiskOutputs.Count)
    
    CalculateDependencyMatrix
    For I = 1 To mRiskInputs.Count
        For J = 1 To mRiskOutputs.Count
            If Not mDependencyMatrix(I, J) Or (VarType(Cell.Offset(2 + I, J).Value) = vbError) Then _
              Cell.Offset(2 + I, J).ClearContents
        Next J
    Next I
    
    mOutSheet.Range("SimStats").Offset(-3).Resize(3, 1 + mRiskOutputs.Count).Copy Cell
    Cell = "Correlations"
    Cell.Offset(2) = "Risk Inputs"
End Sub

Private Sub ProduceCumulativeDistributions()
    Dim I As Integer
    Dim R As Range
    Dim Percentiles As Range
    Dim NewChart As Chart
    
    On Error GoTo Finish
      
    Set Percentiles = mOutSheet.Range("Percentiles")
    Set R = mOutSheet.Range("OutCumDist").Resize(ChartHeight, ChartWidth)
        
    Percentiles.Columns(1).Select
    For I = 1 To mRiskOutputs.Count
        Set NewChart = mOutSheet.Shapes.AddChart2(-1, xlXYScatterLines, R.Left, R.Top, R.Width, R.Height).Chart
        
        With NewChart
            .HasLegend = False
            .SeriesCollection(1).XValues = Percentiles.Columns(I + 1)
            .Axes(xlValue).MaximumScale = 1
            If Percentiles.Cells(1, I + 1) > 0 Then
                .Axes(xlCategory).MinimumScale = WorksheetFunction.RoundDown(Percentiles.Cells(1, I + 1), 0)
            End If
            '.Axes(xlCategory).MaximumScale = WorksheetFunction.RoundUp(Percentiles.Cells(21, I + 1), 0)
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.text = "Cum. Distribution of " & mRiskOutputs(I).Name & _
                " (" & AddressWithSheet(mRiskOutputs(I).Cell) & ")"
            .ChartTitle.Font.Size = 14
            .ChartTitle.Font.Bold = False
        End With
        Set R = R.Offset(ChartHeight + 1)
    Next I
    mOutSheet.Range("A1").Select
Finish:
End Sub

Private Sub ProduceHistograms()
' One histogram per output. Excel 2016 and later draw it with their histogram chart; older
' versions, or a failure of that chart, get a column chart of the counts in 20 equal bins.
    Dim SimOutput As Range
    Dim R As Range
    Dim RiskOutput As ClsRiskOutput
    Dim J As Long
    Dim Done As Boolean

    Set SimOutput = mOutSheet.Range("OutputResults").Resize(mIterations, 1)
    Set R = mOutSheet.Range("OutDist").Resize(ChartHeight, ChartWidth)
    J = 0
    For Each RiskOutput In mRiskOutputs
        J = J + 1
        Done = False
        If Val(Application.Version) >= 16 And Not gForceHistogramFallback Then
            Done = NativeHistogram(SimOutput, R, RiskOutput)
        End If
        If Not Done Then HistogramFallback R, RiskOutput, J
        Set SimOutput = SimOutput.Offset(0, 1)
        Set R = R.Offset(ChartHeight + 1)
    Next RiskOutput
    On Error Resume Next
    mOutSheet.Range("A1").Select
End Sub

Private Function NativeHistogram(SimOutput As Range, R As Range, RiskOutput As ClsRiskOutput) As Boolean
    Const HISTOGRAM_CHART As Long = 118   ' xlHistogram, written as a number so older Excel can compile
    Dim ChartShape As Shape

    On Error GoTo Failed
    ' Range needs to be selected
    'https://stackoverflow.com/questions/37912746/vba-why-xlhistogram-does-not-work
    SimOutput.Select
    Set ChartShape = mOutSheet.Shapes.AddChart2(-1, HISTOGRAM_CHART, R.Left, R.Top, R.Width, R.Height)
    With ChartShape.Chart
       .SetElement (msoElementChartTitleAboveChart)
       .ChartTitle.text = "Distribution of " & RiskOutput.Name & " (" & AddressWithSheet(RiskOutput.Cell) & ")"
    End With
    NativeHistogram = True
    Exit Function
Failed:
    NativeHistogram = False
End Function

Private Sub HistogramFallback(R As Range, RiskOutput As ClsRiskOutput, ByVal J As Long)
' Column chart of the counts of output J in 20 equal bins. The table of counts is written in the
' cells under the chart (first two columns of R).
    Const BINS As Long = 20
    Dim Col As Long
    Dim I As Long
    Dim K As Long
    Dim N As Long
    Dim X As Double
    Dim MinV As Double
    Dim MaxV As Double
    Dim W As Double
    Dim Counts(1 To 20) As Long
    Dim Table As Range
    Dim ChartObj As ChartObject

    On Error GoTo Finish
    Col = 1 + mRiskInputs.Count + J
    N = 0
    For I = 1 To mIterations
        If IsNumber(mSimResults(I, Col)) Then
            X = CDbl(mSimResults(I, Col))
            If N = 0 Then
                MinV = X
                MaxV = X
            Else
                If X < MinV Then MinV = X
                If X > MaxV Then MaxV = X
            End If
            N = N + 1
        End If
    Next I
    If N = 0 Then Exit Sub
    W = (MaxV - MinV) / BINS
    For I = 1 To mIterations
        If IsNumber(mSimResults(I, Col)) Then
            X = CDbl(mSimResults(I, Col))
            If W > 0 Then K = Int((X - MinV) / W) + 1 Else K = 1
            If K > BINS Then K = BINS
            If K < 1 Then K = 1
            Counts(K) = Counts(K) + 1
        End If
    Next I

    Set Table = R.Cells(1, 1).Resize(BINS + 1, 2)
    Table.Cells(1, 1).Value = "Bin upper limit"
    Table.Cells(1, 2).Value = "Count"
    For K = 1 To BINS
        Table.Cells(K + 1, 1).Value = MinV + W * K
        Table.Cells(K + 1, 2).Value = Counts(K)
    Next K
    Table.Columns(1).Offset(1).Resize(BINS).NumberFormat = "#,##0.00"
    mOutSheet.Names.Add Name:="OutDistTable" & CStr(J), RefersTo:=Table.Cells(1, 1)

    Set ChartObj = mOutSheet.ChartObjects.Add(R.Left, R.Top, R.Width, R.Height)
    With ChartObj.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Values = Table.Cells(2, 2).Resize(BINS, 1)
            .XValues = Table.Cells(2, 1).Resize(BINS, 1)
            .Name = "Count"
        End With
        .HasLegend = False
        .ChartGroups(1).GapWidth = 5
        .HasTitle = True
        .ChartTitle.text = "Distribution of " & RiskOutput.Name & " (" & AddressWithSheet(RiskOutput.Cell) & ")"
    End With
Finish:
End Sub

Sub ProduceTornadoDiagram()
    Dim Cell As Range
    Dim R As Range
    Dim SortedTable As Range
    Dim NewChart As Chart
    Dim source As Series
    Dim I As Integer

    Set Cell = mOutSheet.Range("Tornado")
    Cell.Value = "Select Output:"
    Cell.Offset(1).Value = "Name:"

    Set SortedTable = Cell.Offset(0, ChartWidth + 1)

    mOutSheet.Range("Correlations").Resize(mRiskInputs.Count + 3, mRiskOutputs.Count + 1).Copy SortedTable.Offset(-1)
    SortedTable.Resize(mRiskInputs.Count + 1, mRiskOutputs.Count + 1).Columns.AutoFit
    SortedTable.Resize(mRiskInputs.Count + 1, mRiskOutputs.Count + 1).Columns.Hidden = True

    Cell.Offset(0, 1).Value = SortedTable.Offset(0, 1).Value 'First output cell address
    ThickBorders Cell.Offset(0, 1)
    With Cell.Offset(0, 1).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=" & SortedTable.Offset(0, 1).Resize(1, mRiskOutputs.Count).Address
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Select Output Cell"
        .ShowInput = True
        .ShowError = True
    End With

    'Formulas
    Cell.Offset(1, 1).Formula = "=HLOOKUP(" & Cell.Offset(0, 1).Address & "," & _
        SortedTable.Offset(0, 1).Resize(2, mRiskOutputs.Count).Address & ",2,0)" 'Risk output name
    Cell.Offset(2).Formula = "=MATCH(" & Cell.Offset(0, 1).Address & "," & _
        SortedTable.Offset(0, 1).Resize(1, mRiskOutputs.Count).Address & ",0)" 'Column Index used in the formula below
    Cell.Offset(3).Formula = "=""Tornado Diagram for ""&" & Cell.Offset(1, 1).Address & "&"" (""&" & _
        Cell.Offset(0, 1).Address & "&"")""" ' Chart Title
    SortedTable.Offset(2).Resize(mRiskInputs.Count, mRiskOutputs.Count + 1).FormulaArray = _
        "=SortedTable(" & mOutSheet.Range("Correlations").Offset(3).Resize(mRiskInputs.Count, mRiskOutputs.Count + 1).Address _
        & "," & Cell.Offset(2).Address & "+1,1,1)" 'Sorted Table
    SortedTable.Offset(-1).Formula = "=""Correlations - Column ""&" & Cell.Offset(2).Address & "&"" sorted""" ' Table title

    Cell.Offset(2).NumberFormat = ";;;" 'Hide
    Cell.Resize(2, 2).Columns.AutoFit
    
    Cell.Offset(4).Formula = "=CELL(""filename"", A1)"
    If Cell.Offset(4).Value = vbNullString Then 'filename not saved
      Cell.Offset(5).Formula = "=""'" & mOutSheet.Name & "'!"""
    Else
      ' This is so that the chart is shown correctly when the sheet is renamed
      Cell.Offset(5).Formula = "=""'""&" & "MID(" & Cell.Offset(4).Address & ",FIND(""]""," & Cell.Offset(4).Address & ")+1,31)" & "&""'!"""
    End If
    Cell.Offset(6).Formula = "=" & Cell.Offset(5).Address & "&" & "ADDRESS(" & SortedTable.Row + 2 & "+" & "COUNTBLANK(" & _
        "OFFSET(" & SortedTable.Offset(2).Resize(mRiskInputs.Count, 1).Address & ",0," & _
        Cell.Offset(2).Address & "))" & "," & SortedTable.Column & ")&"":" & _
        SortedTable.Offset(mRiskInputs.Count + 1).Address & """"

    '!!!! https://github.com/pyscripter/XLRisk/issues/11
    Cell.Offset(7).Formula = "=INDIRECT(" & Cell.Offset(5).Value & Cell.Offset(6).Address() & ")"
    Cell.Offset(8).Formula = "=OFFSET(INDIRECT(" & Cell.Offset(5).Value & Cell.Offset(6).Address() & "),0," & Cell.Offset(5).Value & Cell.Offset(2).Address & ")"
    mOutSheet.Names.Add Name:="TornadoX", RefersTo:=Cell.Offset(7).FormulaLocal
    mOutSheet.Names.Add Name:="TornadoY", RefersTo:=Cell.Offset(8).FormulaLocal
    Cell.Offset(7).Clear
    Cell.Offset(8).Clear
    
    mOutSheet.Range("TornadoY").Select
    'Create Chart
    Set R = Cell.Offset(3).Resize(WorksheetFunction.Max(ChartHeight, mRiskInputs.Count), ChartWidth)
    'Set NewChart = mOutSheet.ChartObjects.Add(Left:=R.Left, Top:=R.Top, Width:=R.Width, Height:=R.Height).Chart
    Set NewChart = mOutSheet.Shapes.AddChart2(-1, xlBarClustered, R.Left, R.Top, R.Width, R.Height).Chart


    With NewChart
        '.ChartType = xlBarClustered
        .HasLegend = False
        .PlotVisibleOnly = False
        Set source = .SeriesCollection(1)
        source.Values = "='" & mOutSheet.Name & "'!TornadoY"
        source.XValues = "='" & mOutSheet.Name & "'!TornadoX"
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).MinimumScale = -1
        .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
        .ApplyDataLabels xlDataLabelsShowValue
        .SetElement (msoElementDataLabelOutSideEnd)
        source.DataLabels.NumberFormat = "0.00"
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Caption = "='" & mOutSheet.Name & "'!" & Cell.Offset(3).Address
        .ChartTitle.Font.Size = 14
        .ChartTitle.Font.Bold = False
        .ChartGroups(1).GapWidth = 50
    End With
End Sub

Private Sub ProduceMeanTornado()
' Tornado of the change in output mean, in the units of the output (as the "change in output mean"
' tornado of @RISK). For each input the iterations are sorted by the value of the input and split
' in 10 groups of the same size (deciles). The bar goes from the lowest to the highest mean of the
' output among those groups, measured from the overall mean of the output. The table also gives an
' approximate share of the variance of the output: correlation squared, rescaled to add up to 100%.
    Const GROUPS As Long = 10
    Dim Start As Range
    Dim TopCell As Range
    Dim Table As Range
    Dim ChartObj As ChartObject
    Dim RiskOutput As ClsRiskOutput
    Dim I As Long
    Dim J As Long
    Dim K As Long
    Dim N As Long
    Dim G As Long
    Dim NRows As Long
    Dim InCol As Long
    Dim OutCol As Long
    Dim Keys() As Double
    Dim Idx() As Long
    Dim GroupSum(1 To 10) As Double
    Dim GroupCount(1 To 10) As Long
    Dim GroupMean As Double
    Dim OverallMean As Double
    Dim LowMean() As Double
    Dim HighMean() As Double
    Dim Corr() As Double
    Dim Used() As Boolean
    Dim Order() As Long
    Dim SumR2 As Double
    Dim Tmp As Long
    Dim Offs As Long

    If mIterations < GROUPS Then Exit Sub
    On Error GoTo Finish
    Set Start = mOutSheet.Range("MeanTornado")
    Offs = 0
    J = 0
    For Each RiskOutput In mRiskOutputs
        J = J + 1
        OutCol = 1 + mRiskInputs.Count + J
        ReDim LowMean(1 To mRiskInputs.Count)
        ReDim HighMean(1 To mRiskInputs.Count)
        ReDim Corr(1 To mRiskInputs.Count)
        ReDim Used(1 To mRiskInputs.Count)
        OverallMean = ColumnMean(OutCol)
        NRows = 0
        SumR2 = 0
        For I = 1 To mRiskInputs.Count
            InCol = 1 + I
            If mDependencyMatrix(I, J) Then
                ' Iterations where both the input and the output are numbers
                N = 0
                ReDim Keys(1 To mIterations)
                ReDim Idx(1 To mIterations)
                For K = 1 To mIterations
                    If IsNumber(mSimResults(K, InCol)) And IsNumber(mSimResults(K, OutCol)) Then
                        N = N + 1
                        Keys(N) = CDbl(mSimResults(K, InCol))
                        Idx(N) = K
                    End If
                Next K
                If N >= GROUPS Then
                    SortByKey Keys, Idx, 1, N
                    For K = 1 To GROUPS
                        GroupSum(K) = 0
                        GroupCount(K) = 0
                    Next K
                    For K = 1 To N
                        G = Int((K - 1) * GROUPS / N) + 1
                        GroupSum(G) = GroupSum(G) + CDbl(mSimResults(Idx(K), OutCol))
                        GroupCount(G) = GroupCount(G) + 1
                    Next K
                    For K = 1 To GROUPS
                        If GroupCount(K) > 0 Then
                            GroupMean = GroupSum(K) / GroupCount(K)
                            If K = 1 Then
                                LowMean(I) = GroupMean
                                HighMean(I) = GroupMean
                            Else
                                If GroupMean < LowMean(I) Then LowMean(I) = GroupMean
                                If GroupMean > HighMean(I) Then HighMean(I) = GroupMean
                            End If
                        End If
                    Next K
                    Corr(I) = PairCorrelation(InCol, OutCol)
                    SumR2 = SumR2 + Corr(I) ^ 2
                    Used(I) = True
                    NRows = NRows + 1
                End If
            End If
        Next I

        Set TopCell = Start.Offset(Offs, 0)
        TopCell.Value = "Output: " & RiskOutput.Name & " (" & AddressWithSheet(RiskOutput.Cell) & ")"
        TopCell.Font.Bold = True
        If NRows > 0 Then
            ' Order the inputs by swing, largest first
            ReDim Order(1 To NRows)
            K = 0
            For I = 1 To mRiskInputs.Count
                If Used(I) Then
                    K = K + 1
                    Order(K) = I
                End If
            Next I
            For I = 2 To NRows
                Tmp = Order(I)
                K = I - 1
                Do While K >= 1
                    If HighMean(Order(K)) - LowMean(Order(K)) >= HighMean(Tmp) - LowMean(Tmp) Then Exit Do
                    Order(K + 1) = Order(K)
                    K = K - 1
                Loop
                Order(K + 1) = Tmp
            Next I

            Set Table = TopCell.Offset(1, ChartWidth + 1).Resize(NRows + 1, 7)
            Table.Rows(1).Value = Array("Input", "Lowest decile mean", "Highest decile mean", _
                "Low vs mean", "High vs mean", "Swing", "% of variance (approx.)")
            Table.Rows(1).Font.Bold = True
            For K = 1 To NRows
                I = Order(K)
                Table.Cells(K + 1, 1).Value = NameOrAddress(mRiskInputs(I).Cell)
                Table.Cells(K + 1, 2).Value = LowMean(I)
                Table.Cells(K + 1, 3).Value = HighMean(I)
                Table.Cells(K + 1, 4).Value = LowMean(I) - OverallMean
                Table.Cells(K + 1, 5).Value = HighMean(I) - OverallMean
                Table.Cells(K + 1, 6).Value = HighMean(I) - LowMean(I)
                If SumR2 > 0 Then Table.Cells(K + 1, 7).Value = Corr(I) ^ 2 / SumR2 Else Table.Cells(K + 1, 7).Value = 0
            Next K
            Table.Columns(2).Resize(, 5).Offset(1).Resize(NRows).NumberFormat = "#,##0.00"
            Table.Columns(7).Offset(1).Resize(NRows).NumberFormat = "0.0%"
            ThickBorders Table
            Table.Columns.AutoFit
            mOutSheet.Names.Add Name:="MeanTornadoTable" & CStr(J), RefersTo:=Table.Cells(1, 1)

            Set ChartObj = mOutSheet.ChartObjects.Add(TopCell.Offset(1, 0).Left, TopCell.Offset(1, 0).Top, _
                TopCell.Offset(1, 0).Resize(1, ChartWidth).Width, TopCell.Offset(1, 0).Resize(ChartHeight, 1).Height)
            With ChartObj.Chart
                .ChartType = xlBarClustered
                Do While .SeriesCollection.Count > 0
                    .SeriesCollection(1).Delete
                Loop
                With .SeriesCollection.NewSeries
                    .Name = "Lowest decile"
                    .Values = Table.Cells(2, 4).Resize(NRows, 1)
                    .XValues = Table.Cells(2, 1).Resize(NRows, 1)
                    .Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
                End With
                With .SeriesCollection.NewSeries
                    .Name = "Highest decile"
                    .Values = Table.Cells(2, 5).Resize(NRows, 1)
                    .XValues = Table.Cells(2, 1).Resize(NRows, 1)
                    .Format.Fill.ForeColor.RGB = RGB(0, 140, 0)
                End With
                .ChartGroups(1).Overlap = 100
                .ChartGroups(1).GapWidth = 50
                .Axes(xlCategory).ReversePlotOrder = True
                .Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
                .HasLegend = True
                .Legend.Position = xlLegendPositionBottom
                .HasTitle = True
                .ChartTitle.text = "Change in mean of " & RiskOutput.Name & " (mean " & Format(OverallMean, "#,##0.00") & ")"
                .ChartTitle.Font.Size = 12
            End With
        Else
            TopCell.Offset(1, 0).Value = "No input changes this output."
        End If
        If NRows + 3 > ChartHeight Then Offs = Offs + NRows + 5 Else Offs = Offs + ChartHeight + 3
    Next RiskOutput
Finish:
End Sub

Private Function IsNumber(X As Variant) As Boolean
    If IsError(X) Then
        IsNumber = False
    Else
        IsNumber = (VarType(X) = vbDouble Or VarType(X) = vbLong Or VarType(X) = vbInteger _
            Or VarType(X) = vbSingle Or VarType(X) = vbCurrency Or VarType(X) = vbDecimal)
    End If
End Function

Private Function ColumnMean(ByVal Col As Long) As Double
    Dim K As Long
    Dim N As Long
    Dim S As Double
    For K = 1 To mIterations
        If IsNumber(mSimResults(K, Col)) Then
            S = S + CDbl(mSimResults(K, Col))
            N = N + 1
        End If
    Next K
    If N > 0 Then ColumnMean = S / N
End Function

Private Function PairCorrelation(ByVal ColX As Long, ByVal ColY As Long) As Double
' Pearson correlation over the iterations where both columns hold numbers
    Dim K As Long
    Dim N As Long
    Dim SX As Double, SY As Double, SXX As Double, SYY As Double, SXY As Double
    Dim X As Double, Y As Double
    Dim Den As Double
    For K = 1 To mIterations
        If IsNumber(mSimResults(K, ColX)) And IsNumber(mSimResults(K, ColY)) Then
            X = CDbl(mSimResults(K, ColX))
            Y = CDbl(mSimResults(K, ColY))
            N = N + 1
            SX = SX + X
            SY = SY + Y
            SXX = SXX + X * X
            SYY = SYY + Y * Y
            SXY = SXY + X * Y
        End If
    Next K
    If N < 2 Then Exit Function
    Den = Sqr((N * SXX - SX * SX) * (N * SYY - SY * SY))
    If Den > 0 Then PairCorrelation = (N * SXY - SX * SY) / Den
End Function

Private Sub SortByKey(Keys() As Double, Idx() As Long, ByVal Lo As Long, ByVal Hi As Long)
' Quicksort of Keys(Lo..Hi), moving Idx along. Ties keep no particular order.
    Dim I As Long
    Dim J As Long
    Dim Pivot As Double
    Dim TK As Double
    Dim TI As Long
    Do While Lo < Hi
        I = Lo
        J = Hi
        Pivot = Keys((Lo + Hi) \ 2)
        Do While I <= J
            Do While Keys(I) < Pivot
                I = I + 1
            Loop
            Do While Keys(J) > Pivot
                J = J - 1
            Loop
            If I <= J Then
                TK = Keys(I): Keys(I) = Keys(J): Keys(J) = TK
                TI = Idx(I): Idx(I) = Idx(J): Idx(J) = TI
                I = I + 1
                J = J - 1
            End If
        Loop
        ' Recurse into the smaller part, loop on the larger one
        If J - Lo < Hi - I Then
            If Lo < J Then SortByKey Keys, Idx, Lo, J
            Lo = I
        Else
            If I < Hi Then SortByKey Keys, Idx, I, Hi
            Hi = J
        End If
    Loop
End Sub

Sub CalculateDependencyMatrix()
    Dim AppCalculation As XlCalculation
    Dim SavedOutputs() As Variant
    Dim SavedInputs() As String
    Dim OldProduceRandomSample As Boolean
    Dim I As Long
    Dim J As Long
    
    ReDim mDependencyMatrix(1 To mRiskInputs.Count, 1 To mRiskOutputs.Count)
    ReDim SavedInputs(1 To mRiskInputs.Count)
    ReDim SavedOutputs(1 To mRiskOutputs.Count)
    
    ' Save ProduceRandomSample
    OldProduceRandomSample = ProduceRandomSample
    ProduceRandomSample = False
    ' Save Calculation Mode
    AppCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    On Error GoTo Restore
    
    ' Save values
    Application.CalculateFull
    For I = 1 To mRiskInputs.Count
        SavedInputs(I) = mRiskInputs(I).Cell.Formula
    Next I
    For J = 1 To mRiskOutputs.Count
        SavedOutputs(J) = mRiskOutputs(J).Cell.Value
    Next J
    For I = 1 To mRiskInputs.Count
        mRiskInputs(I).Cell.Value = -100 ' Start with a negative value
        Application.Calculate
        For J = 1 To mRiskOutputs.Count
            If IsError(SavedOutputs(J)) Then
                mDependencyMatrix(I, J) = False
            ElseIf IsError(mRiskOutputs(J).Cell.Value) Then
                mDependencyMatrix(I, J) = True
            Else
                mDependencyMatrix(I, J) = mRiskOutputs(J).Cell.Value <> SavedOutputs(J)
            End If
        Next J
        mRiskInputs(I).Cell.Value = 100 ' Now try a positive value
        Application.Calculate
        For J = 1 To mRiskOutputs.Count
            If Not mDependencyMatrix(I, J) Then
                If IsError(SavedOutputs(J)) Then
                    mDependencyMatrix(I, J) = False
                ElseIf IsError(mRiskOutputs(J).Cell.Value) Then
                    mDependencyMatrix(I, J) = True
                Else
                    mDependencyMatrix(I, J) = mRiskOutputs(J).Cell.Value <> SavedOutputs(J)
                End If
            End If
        Next J
        
        mRiskInputs(I).Cell.Formula = SavedInputs(I)
    Next I
Restore:
    On Error GoTo 0
    ' Save again just in case
    For I = 1 To mRiskInputs.Count
        mRiskInputs(I).Cell.Formula = SavedInputs(I)
    Next I
    'Restore ProduceRandomSample
    ProduceRandomSample = OldProduceRandomSample
    'Restore Calculation Mode
    Application.Calculation = AppCalculation
    Application.CalculateFull
End Sub

