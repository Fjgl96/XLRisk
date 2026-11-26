Attribute VB_Name = "Compare"
Option Explicit

' Compares the runs of a workbook: overlays the cumulative distribution of one output from every
' "Risk Results N" sheet on a new sheet named "Risk Compare", with a table of statistics.

Public Const COMPARE_SHEET As String = "Risk Compare"

Public Sub CompareRuns()
' Action for the Compare Runs button
    Dim WB As Workbook
    Dim Runs As Collection
    Dim Addr() As String
    Dim Nm() As String
    Dim N As Long
    Dim I As Long
    Dim Choice As Variant
    Dim Prompt As String
    Dim WS As Worksheet

    Set WB = ActiveWorkBook
    Set Runs = ResultSheets(WB)
    If Runs.Count < 2 Then
        MsgBox "Compare Runs needs at least two simulations (two Risk Results sheets) in this workbook.", _
            Title:="XLRisk"
        Exit Sub
    End If
    N = OutputHeaders(Runs(Runs.Count), Addr, Nm)
    If N = 0 Then
        MsgBox "The latest Risk Results sheet has no outputs.", Title:="XLRisk"
        Exit Sub
    End If
    If N = 1 Then
        Choice = 1
    Else
        Prompt = "Which output do you want to compare?" & vbCrLf
        For I = 1 To N
            Prompt = Prompt & vbCrLf & CStr(I) & ". " & Nm(I) & " (" & Addr(I) & ")"
        Next I
        Choice = Application.InputBox(Prompt, "Compare Runs", 1, Type:=1)
        If VarType(Choice) = vbBoolean Then Exit Sub   ' Cancel
        If Choice < 1 Or Choice > N Or Choice <> Int(Choice) Then
            MsgBox "Type a number between 1 and " & CStr(N) & ".", Title:="XLRisk"
            Exit Sub
        End If
    End If
    Set WS = CompareRunsFor(WB, Addr(CLng(Choice)))
    If WS Is Nothing Then
        MsgBox "No run has the output " & Addr(CLng(Choice)) & ".", Title:="XLRisk"
    Else
        WS.Activate
    End If
End Sub

Public Function ResultSheets(WB As Workbook) As Collection
' The sheets named "Risk Results N", in increasing order of N
    Dim Result As New Collection
    Dim WS As Worksheet
    Dim Numbers() As Long
    Dim Shts() As Worksheet
    Dim N As Long
    Dim I As Long
    Dim J As Long
    Dim Num As Long
    Dim Tmp As Worksheet

    N = 0
    For Each WS In WB.Worksheets
        If Left(WS.Name, 13) = "Risk Results " And IsNumeric(Mid(WS.Name, 14)) Then
            N = N + 1
            ReDim Preserve Numbers(1 To N)
            ReDim Preserve Shts(1 To N)
            Numbers(N) = CLng(Mid(WS.Name, 14))
            Set Shts(N) = WS
        End If
    Next WS
    For I = 2 To N
        Num = Numbers(I)
        Set Tmp = Shts(I)
        J = I - 1
        Do While J >= 1
            If Numbers(J) <= Num Then Exit Do
            Numbers(J + 1) = Numbers(J)
            Set Shts(J + 1) = Shts(J)
            J = J - 1
        Loop
        Numbers(J + 1) = Num
        Set Shts(J + 1) = Tmp
    Next I
    For I = 1 To N
        Result.Add Shts(I)
    Next I
    Set ResultSheets = Result
End Function

Public Function OutputHeaders(WS As Worksheet, Addr() As String, Nm() As String) As Long
' Addresses (row 2) and names (row 3) of the outputs of a Risk Results sheet
    Dim C As Range
    Dim N As Long

    On Error GoTo NoOutputs
    Set C = WS.Range("OutputResults").Offset(-2)
    N = 0
    Do While Len(CStr(C.Offset(0, N).Value)) > 0
        N = N + 1
        ReDim Preserve Addr(1 To N)
        ReDim Preserve Nm(1 To N)
        Addr(N) = CStr(C.Offset(0, N - 1).Value)
        Nm(N) = CStr(C.Offset(1, N - 1).Value)
    Loop
    OutputHeaders = N
    Exit Function
NoOutputs:
    OutputHeaders = 0
End Function

Private Function SameAddress(ByVal A As String, ByVal B As String) As Boolean
    SameAddress = (UCase(Replace(A, "'", "")) = UCase(Replace(B, "'", "")))
End Function

Public Function CompareRunsFor(WB As Workbook, ByVal OutputAddress As String) As Worksheet
' Builds the sheet "Risk Compare" for the output at OutputAddress and returns it
' (Nothing if no run has that output)
    Const STAT_ROWS As Long = 8
    Dim Runs As Collection
    Dim RunWS As Worksheet
    Dim WS As Worksheet
    Dim Addr() As String
    Dim Nm() As String
    Dim OutName As String
    Dim N As Long
    Dim K As Long
    Dim Col As Long
    Dim Found As Long
    Dim Iters As Long
    Dim Data As Range
    Dim Perc As Range
    Dim StatTop As Range
    Dim CumTop As Range
    Dim ChartObj As ChartObject
    Dim I As Long
    Dim OldAlerts As Boolean

    Set Runs = ResultSheets(WB)
    Set CompareRunsFor = Nothing

    ' New comparison sheet, replacing the previous one
    OldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    WB.Worksheets(COMPARE_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = OldAlerts
    Set WS = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
    WS.Name = COMPARE_SHEET

    Set StatTop = WS.Range("A3")
    StatTop.Value = "Statistic"
    StatTop.Offset(1, 0).Resize(STAT_ROWS, 1).Value = Application.WorksheetFunction.Transpose( _
        Array("Mean", "Std. Deviation", "Minimum", "5th percentile", "Median", "95th percentile", "Maximum", _
              "Probability < 0"))
    Set CumTop = WS.Range("A14")
    CumTop.Offset(-1, 0).Value = "Cumulative distributions"
    CumTop.Value = "Percentile"
    For I = 0 To 40
        CumTop.Offset(1 + I, 0).Value = I * 0.025
    Next I
    CumTop.Offset(1, 0).Resize(41, 1).NumberFormat = "0.0%"

    Found = 0
    For Each RunWS In Runs
        N = OutputHeaders(RunWS, Addr, Nm)
        Col = 0
        For K = 1 To N
            If SameAddress(Addr(K), OutputAddress) Then
                Col = K
                Exit For
            End If
        Next K
        If Col > 0 Then
            Found = Found + 1
            If Len(OutName) = 0 Then OutName = Nm(Col)
            Iters = Application.WorksheetFunction.Count(RunWS.Range("B4", RunWS.Cells(RunWS.Rows.Count, 2)))
            Set Data = RunWS.Range("OutputResults").Offset(0, Col - 1).Resize(Iters, 1)
            Set Perc = RunWS.Range("Percentiles")
            StatTop.Offset(0, Found).Value = RunWS.Name
            CumTop.Offset(0, Found).Value = RunWS.Name
            WriteStats StatTop.Offset(1, Found), Data
            For I = 1 To 41
                CumTop.Offset(I, Found).Value = Perc.Cells(I, Col + 1).Value
            Next I
        End If
    Next RunWS

    If Found = 0 Then
        Application.DisplayAlerts = False
        WS.Delete
        Application.DisplayAlerts = OldAlerts
        Exit Function
    End If

    WS.Range("A1").Value = "Comparison of runs: " & OutName & " (" & OutputAddress & ")"
    WS.Range("A1").Font.Bold = True
    WS.Range("A1").Font.Size = 13
    CumTop.Offset(-1, 0).Font.Bold = True
    StatTop.Resize(1, Found + 1).Font.Bold = True
    CumTop.Resize(1, Found + 1).Font.Bold = True
    StatTop.Offset(1, 1).Resize(STAT_ROWS - 1, Found).NumberFormat = "#,##0.00"
    StatTop.Offset(STAT_ROWS, 1).Resize(1, Found).NumberFormat = "0.0%"
    CumTop.Offset(1, 1).Resize(41, Found).NumberFormat = "#,##0.00"
    ThickBorders StatTop.Resize(STAT_ROWS + 1, Found + 1)
    ThickBorders CumTop.Resize(42, Found + 1)
    WS.Columns(1).Resize(, Found + 1).AutoFit

    Set ChartObj = WS.ChartObjects.Add(WS.Cells(3, Found + 3).Left, WS.Range("A3").Top, 520, 330)
    With ChartObj.Chart
        .ChartType = xlXYScatterLines
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For I = 1 To Found
            With .SeriesCollection.NewSeries
                .Name = "='" & COMPARE_SHEET & "'!" & CumTop.Offset(0, I).Address
                .XValues = CumTop.Offset(1, I).Resize(41, 1)
                .Values = CumTop.Offset(1, 0).Resize(41, 1)
                .MarkerStyle = xlMarkerStyleNone
            End With
        Next I
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 1
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .HasTitle = True
        .ChartTitle.text = "Cumulative distribution of " & OutName
    End With
    WS.Range("A1").Select
    Set CompareRunsFor = WS
End Function

Private Sub WriteStats(Top As Range, Data As Range)
    Dim WF As WorksheetFunction
    Dim Numbers As Double
    Set WF = Application.WorksheetFunction
    On Error Resume Next
    Numbers = WF.Count(Data)
    Top.Offset(0, 0).Value = WF.Average(Data)
    Top.Offset(1, 0).Value = WF.StDev_S(Data)
    Top.Offset(2, 0).Value = WF.Min(Data)
    Top.Offset(3, 0).Value = WF.Percentile_Inc(Data, 0.05)
    Top.Offset(4, 0).Value = WF.Median(Data)
    Top.Offset(5, 0).Value = WF.Percentile_Inc(Data, 0.95)
    Top.Offset(6, 0).Value = WF.Max(Data)
    If Numbers > 0 Then Top.Offset(7, 0).Value = WF.CountIf(Data, "<0") / Numbers
End Sub
