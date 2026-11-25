Attribute VB_Name = "RiskStatistics"
Option Explicit

' RiskOutput and the simulation statistics functions, with the same names and
' arguments as in @RISK.
'
' RiskOutput marks a cell as a simulation output from its own formula:
'     =RiskOutput("Profit") + C17 - C18
' It returns 0, so it does not change the value of the cell.
'
' The statistics functions read the results of the latest simulation of a cell
' (an output, or a risk input) from the most recent "Risk Results N" sheet of its
' workbook, so they keep working after the workbook is saved and reopened:
'     =RiskMean(C19)   =RiskPercentile(C19, 0.05)   =RiskTarget(C19, 0)
' They return #N/A while a simulation runs and when the cell has no results.

' Incremented after every simulation; invalidates the cache of results
Public gResultsRevision As Long

Private mCache As Collection
Private mCacheRevision As Long

Private Const RESULTS_PREFIX As String = "Risk Results "

'---------------------------------------------------------------------------
' RiskOutput
'---------------------------------------------------------------------------

Public Function RiskOutput(Optional Name As Variant, Optional RangeName As Variant, _
                           Optional Position As Variant) As Double
Attribute RiskOutput.VB_Description = "Marks the cell as a simulation output. Use it as =RiskOutput(""Name"") + formula. Returns 0."
Attribute RiskOutput.VB_ProcData.VB_Invoke_Func = " \n20"
    RiskOutput = 0
End Function

'---------------------------------------------------------------------------
' Statistics of the latest simulation
'---------------------------------------------------------------------------

Public Function RiskMean(Cell As Range) As Variant
Attribute RiskMean.VB_Description = "Mean of the cell in the latest simulation"
Attribute RiskMean.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Data As Variant
    If GetResults(Cell, Data) Then
        RiskMean = WorksheetFunction.Average(Data)
    Else
        RiskMean = CVErr(xlErrNA)
    End If
End Function

Public Function RiskStdDev(Cell As Range) As Variant
Attribute RiskStdDev.VB_Description = "Standard deviation of the cell in the latest simulation"
Attribute RiskStdDev.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Data As Variant
    If GetResults(Cell, Data) Then
        If UBound(Data) > 1 Then
            RiskStdDev = WorksheetFunction.StDev_S(Data)
        Else
            RiskStdDev = 0
        End If
    Else
        RiskStdDev = CVErr(xlErrNA)
    End If
End Function

Public Function RiskMin(Cell As Range) As Variant
Attribute RiskMin.VB_Description = "Minimum of the cell in the latest simulation"
Attribute RiskMin.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Data As Variant
    If GetResults(Cell, Data) Then
        RiskMin = WorksheetFunction.Min(Data)
    Else
        RiskMin = CVErr(xlErrNA)
    End If
End Function

Public Function RiskMax(Cell As Range) As Variant
Attribute RiskMax.VB_Description = "Maximum of the cell in the latest simulation"
Attribute RiskMax.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Data As Variant
    If GetResults(Cell, Data) Then
        RiskMax = WorksheetFunction.Max(Data)
    Else
        RiskMax = CVErr(xlErrNA)
    End If
End Function

Public Function RiskPercentile(Cell As Range, Percentile As Double) As Variant
Attribute RiskPercentile.VB_Description = "Value of the cell at a percentile (between 0 and 1) in the latest simulation"
Attribute RiskPercentile.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Data As Variant
    If Percentile < 0 Or Percentile > 1 Then
        RiskPercentile = CVErr(xlErrValue)
    ElseIf GetResults(Cell, Data) Then
        RiskPercentile = WorksheetFunction.Percentile_Inc(Data, Percentile)
    Else
        RiskPercentile = CVErr(xlErrNA)
    End If
End Function

Public Function RiskPtoX(Cell As Range, Percentile As Double) As Variant
Attribute RiskPtoX.VB_Description = "Same as RiskPercentile: value of the cell at a percentile (between 0 and 1)"
Attribute RiskPtoX.VB_ProcData.VB_Invoke_Func = " \n20"
    RiskPtoX = RiskPercentile(Cell, Percentile)
End Function

Public Function RiskTarget(Cell As Range, TargetValue As Double) As Variant
Attribute RiskTarget.VB_Description = "Probability that the cell is less than or equal to a value in the latest simulation"
Attribute RiskTarget.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim Data As Variant
    Dim I As Long
    Dim Below As Long
    If GetResults(Cell, Data) Then
        For I = LBound(Data) To UBound(Data)
            If Data(I) <= TargetValue Then Below = Below + 1
        Next I
        RiskTarget = Below / (UBound(Data) - LBound(Data) + 1)
    Else
        RiskTarget = CVErr(xlErrNA)
    End If
End Function

Public Function RiskXtoP(Cell As Range, TargetValue As Double) As Variant
Attribute RiskXtoP.VB_Description = "Same as RiskTarget: probability that the cell is less than or equal to a value"
Attribute RiskXtoP.VB_ProcData.VB_Invoke_Func = " \n20"
    RiskXtoP = RiskTarget(Cell, TargetValue)
End Function

'---------------------------------------------------------------------------
' Reading the results
'---------------------------------------------------------------------------

Private Function GetResults(Cell As Range, Data As Variant) As Boolean
' Numeric results of Cell in the latest simulation, as a 1-based array of Doubles
    Dim Target As Range
    Dim Key As String

    On Error GoTo Failed
    If Not gSimulation Is Nothing Then
        If gSimulation.IsRunning Then Exit Function
    End If
    Set Target = Cell.Cells(1, 1)
    Key = Target.Parent.Parent.Name & "|" & AddressWithSheet(Target)

    If mCache Is Nothing Or mCacheRevision <> gResultsRevision Then
        Set mCache = New Collection
        mCacheRevision = gResultsRevision
    End If
    If CacheLookup(Key, Data) Then
        GetResults = True
        Exit Function
    End If
    If Not ReadResults(Target, Data) Then Exit Function
    mCache.Add Data, Key
    GetResults = True
    Exit Function
Failed:
    GetResults = False
End Function

Private Function CacheLookup(Key As String, Data As Variant) As Boolean
    On Error GoTo NotFound
    Data = mCache(Key)
    CacheLookup = True
    Exit Function
NotFound:
    CacheLookup = False
End Function

Private Function ReadResults(Target As Range, Data As Variant) As Boolean
' Finds the column of Target in the most recent "Risk Results N" sheet
    Dim WS As Worksheet
    Dim Latest As Worksheet
    Dim Best As Long
    Dim N As Long
    Dim Header As Range
    Dim Col As Long
    Dim Found As Boolean
    Dim ByAddress As String
    Dim ByName As String
    Dim Label As String
    Dim Iterations As Long
    Dim Raw As Variant
    Dim Values() As Double
    Dim I As Long
    Dim K As Long

    For Each WS In Target.Parent.Parent.Worksheets
        If Left(WS.Name, Len(RESULTS_PREFIX)) = RESULTS_PREFIX Then
            N = Val(Mid(WS.Name, Len(RESULTS_PREFIX) + 1))
            If N > Best Then
                Best = N
                Set Latest = WS
            End If
        End If
    Next WS
    If Latest Is Nothing Then Exit Function

    ' Row 2 holds the address of every output and the name or address of every input
    ByAddress = NormalizeRef(AddressWithSheet(Target))
    ByName = NormalizeRef(NameOrAddress(Target))
    Set Header = Latest.Range("C2")
    Do While Not IsEmpty(Header.Offset(0, Col).Value) And Col < 16000
        Label = NormalizeRef(CStr(Header.Offset(0, Col).Value))
        If Label = ByAddress Or Label = ByName Then
            Found = True
            Exit Do
        End If
        Col = Col + 1
    Loop
    If Not Found Then Exit Function

    ' Iteration numbers run down column B from row 4
    If IsEmpty(Latest.Range("B4").Value) Then Exit Function
    If IsEmpty(Latest.Range("B5").Value) Then
        Iterations = 1
    Else
        Iterations = Latest.Range("B4").End(xlDown).Row - 3
    End If

    If Iterations = 1 Then
        ReDim Raw(1 To 1, 1 To 1)
        Raw(1, 1) = Header.Offset(2, Col).Value
    Else
        Raw = Header.Offset(2, Col).Resize(Iterations, 1).Value
    End If

    ReDim Values(1 To Iterations)
    For I = 1 To Iterations
        If Not IsError(Raw(I, 1)) Then
            If Not IsEmpty(Raw(I, 1)) And IsNumeric(Raw(I, 1)) Then
                K = K + 1
                Values(K) = Raw(I, 1)
            End If
        End If
    Next I
    If K = 0 Then Exit Function
    ReDim Preserve Values(1 To K)
    Data = Values
    ReadResults = True
End Function

Private Function NormalizeRef(ByVal S As String) As String
' Compares references without quotes around sheet names and without case
    NormalizeRef = UCase(Replace(S, "'", ""))
End Function
