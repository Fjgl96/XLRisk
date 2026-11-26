Attribute VB_Name = "FunctionSwap"
Option Explicit

' Swap Functions: replaces every Risk distribution function of the workbook by its expected value,
' so the file can be sent to people without XLRisk, and puts the functions back on the second press.
' The original formulas are kept in a hidden sheet named "XLRisk Swap".

Public Const SWAP_SHEET As String = "XLRisk Swap"

Public Sub SwapFunctions()
' Action for the Swap Functions button
    Dim WB As Workbook
    Dim N As Long

    Set WB = ActiveWorkBook
    If HasSwapSheet(WB) Then
        N = SwapIn(WB)
        MsgBox CStr(N) & " cells got their Risk functions back.", Title:="XLRisk Swap Functions"
    Else
        If MsgBox("Replace every Risk function of this workbook by its expected value?" & vbCrLf & vbCrLf & _
            "The original formulas are saved in a hidden sheet. Press Swap Functions again to restore them.", _
            vbOKCancel + vbQuestion, "XLRisk Swap Functions") = vbCancel Then Exit Sub
        N = SwapOut(WB)
        If N = 0 Then
            MsgBox "This workbook has no Risk functions to swap.", Title:="XLRisk Swap Functions"
        Else
            MsgBox CStr(N) & " cells now hold expected values instead of Risk functions." & vbCrLf & _
                "Press Swap Functions again to restore them.", Title:="XLRisk Swap Functions"
        End If
    End If
End Sub

Public Function HasSwapSheet(WB As Workbook) As Boolean
    Dim WS As Worksheet
    On Error Resume Next
    Set WS = WB.Worksheets(SWAP_SHEET)
    HasSwapSheet = Not WS Is Nothing
End Function

Private Function IsModelSheet(WS As Worksheet) As Boolean
    IsModelSheet = WS.Name <> SWAP_SHEET And WS.Name <> "XLRisk" And WS.Name <> COMPARE_SHEET _
        And Left(WS.Name, 13) <> "Risk Results "
End Function

Public Function SwapOut(WB As Workbook) As Long
' Replaces the Risk functions by their expected values. Returns the number of cells changed.
    Dim OldSample As Boolean
    Dim OldAlerts As Boolean
    Dim Store As Worksheet
    Dim Sht As Worksheet
    Dim Formulas As Range
    Dim Cell As Range
    Dim NewFormula As String
    Dim N As Long

    OldSample = ProduceRandomSample
    ProduceRandomSample = False
    Set Store = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
    Store.Name = SWAP_SHEET
    Store.Range("A1:C1").Value = Array("Sheet", "Cell", "Original formula")
    N = 0
    For Each Sht In WB.Worksheets
        If IsModelSheet(Sht) Then
            Set Formulas = Nothing
            On Error Resume Next
            Set Formulas = Sht.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            If Not Formulas Is Nothing Then
                For Each Cell In Formulas
                    If Not Cell.HasArray Then
                        NewFormula = ReplaceRiskCalls(Sht, Cell.Formula)
                        If NewFormula <> Cell.Formula Then
                            N = N + 1
                            Store.Cells(N + 1, 1).Value = Sht.Name
                            Store.Cells(N + 1, 2).Value = Cell.Address
                            Store.Cells(N + 1, 3).Value = "'" & Cell.Formula
                            Cell.Formula = NewFormula
                        End If
                    End If
                Next Cell
            End If
        End If
    Next Sht
    ProduceRandomSample = OldSample
    If N = 0 Then
        OldAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        Store.Delete
        Application.DisplayAlerts = OldAlerts
    Else
        Store.Visible = xlSheetHidden
    End If
    Application.CalculateFull
    SwapOut = N
End Function

Public Function SwapIn(WB As Workbook) As Long
' Puts back the formulas saved by SwapOut. Returns the number of cells restored.
    Dim Store As Worksheet
    Dim Row As Long
    Dim N As Long
    Dim OldAlerts As Boolean

    Set Store = WB.Worksheets(SWAP_SHEET)
    Row = 2
    N = 0
    Do While Len(CStr(Store.Cells(Row, 1).Value)) > 0
        On Error Resume Next
        WB.Worksheets(CStr(Store.Cells(Row, 1).Value)).Range(CStr(Store.Cells(Row, 2).Value)).Formula = _
            CStr(Store.Cells(Row, 3).Value)
        If Err.Number = 0 Then N = N + 1
        Err.Clear
        On Error GoTo 0
        Row = Row + 1
    Loop
    OldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Store.Visible = xlSheetVisible
    Store.Delete
    Application.DisplayAlerts = OldAlerts
    Application.CalculateFull
    SwapIn = N
End Function

Public Function ReplaceRiskCalls(Sht As Worksheet, ByVal Formula As String) As String
' Replaces each call to a Risk distribution function inside Formula by its expected value
    Dim FunctionList As Variant
    Dim Fn As Variant
    Dim Pos As Long
    Dim Best As Long
    Dim BestLen As Long
    Dim P As Long
    Dim Q As Long
    Dim CallText As String
    Dim V As Variant
    Dim Lit As String

    FunctionList = RiskFunctionList()
    Pos = 1
    Do
        ' Earliest call of any Risk distribution function from Pos
        Best = 0
        For Each Fn In FunctionList
            P = FindCall(Formula, CStr(Fn), Pos)
            If P > 0 And (Best = 0 Or P < Best) Then
                Best = P
                BestLen = Len(Fn)
            End If
        Next Fn
        If Best = 0 Then Exit Do
        Q = MatchingParen(Formula, Best + BestLen)
        If Q = 0 Then Exit Do
        CallText = Mid(Formula, Best, Q - Best + 1)
        V = Sht.Evaluate(CallText)
        If IsError(V) Or Not IsNumeric(V) Or VarType(V) = vbString Or VarType(V) = vbBoolean Then
            Pos = Q + 1
        Else
            Lit = NumberLiteral(CDbl(V))
            Formula = Left(Formula, Best - 1) & Lit & Mid(Formula, Q + 1)
            Pos = Best + Len(Lit)
        End If
    Loop
    ReplaceRiskCalls = Formula
End Function

Private Function FindCall(ByVal Formula As String, ByVal Fn As String, ByVal Start As Long) As Long
' Position of "Fn(" in Formula from Start, as a whole function name and outside text in quotes
    Dim P As Long
    Dim Before As String
    P = InStr(Start, Formula, Fn & "(", vbTextCompare)
    Do While P > 0
        If P > 1 Then Before = Mid(Formula, P - 1, 1) Else Before = ""
        If Not (Before Like "[A-Za-z0-9_.]") And Not InsideQuotes(Formula, P) Then
            FindCall = P
            Exit Function
        End If
        P = InStr(P + 1, Formula, Fn & "(", vbTextCompare)
    Loop
    FindCall = 0
End Function

Private Function InsideQuotes(ByVal Formula As String, ByVal P As Long) As Boolean
    Dim I As Long
    Dim Inside As Boolean
    For I = 1 To P - 1
        If Mid(Formula, I, 1) = """" Then Inside = Not Inside
    Next I
    InsideQuotes = Inside
End Function

Private Function MatchingParen(ByVal Formula As String, ByVal OpenPos As Long) As Long
' Position of the parenthesis that closes the one at OpenPos; 0 if there is none
    Dim I As Long
    Dim Depth As Long
    Dim Ch As String
    Dim InText As Boolean
    For I = OpenPos To Len(Formula)
        Ch = Mid(Formula, I, 1)
        If Ch = """" Then
            InText = Not InText
        ElseIf Not InText Then
            If Ch = "(" Then
                Depth = Depth + 1
            ElseIf Ch = ")" Then
                Depth = Depth - 1
                If Depth = 0 Then
                    MatchingParen = I
                    Exit Function
                End If
            End If
        End If
    Next I
    MatchingParen = 0
End Function

Private Function NumberLiteral(ByVal X As Double) As String
' A number as formula text, with a dot as decimal separator and negatives in parentheses
    Dim S As String
    S = Trim(Str(X))
    If Left(S, 1) = "." Then S = "0" & S
    If Left(S, 2) = "-." Then S = "-0" & Mid(S, 2)
    If X < 0 Then S = "(" & S & ")"
    NumberLiteral = S
End Function
