VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWB As Workbook
Private mProcessingCorrmatInfo As Boolean
Private mSeed As Double
Private mIterations As Long
Private mScreenUpdate As Boolean
Private mLatinHypercube As Boolean
Private mCalculateDataTables As Boolean
Private mMacroBefore As String
Private mMacroAfter As String
Private mMacroAfterSimulation As String
Private mSimError As Boolean
Private mSimErrorMsg As String
' The following two-dimensional array will contain the values of Iter, Risk Input Samples, Risk Output results
Private mSimResults() As Variant
Private mRiskInputs As Collection
Private mRiskOutputs As Collection
Private mInputGroups As Collection
Private mCurrentIteration As Long
' Sheets of other open workbooks whose recalculation is suspended during the simulation
Private mSuspendedSheets As Collection
Private mRecalcOtherWorkbooks As Boolean
Private mLoopSeconds As Double
Private mRunning As Boolean

Private Const NoCorrelationGroupName As String = "No Correlation"
' How often the status bar is refreshed and Excel processes the Stop button
Private Const REFRESH_SECONDS As Double = 0.25

Property Get ActiveWorkBook() As Workbook
    Set ActiveWorkBook = mWB
End Property

Property Get LastResults() As Variant
' Iteration number, risk input samples and risk output values of the last run
    LastResults = mSimResults
End Property

Property Get InputCount() As Long
    If Not mRiskInputs Is Nothing Then InputCount = mRiskInputs.Count
End Property

Property Get OutputCount() As Long
    If Not mRiskOutputs Is Nothing Then OutputCount = mRiskOutputs.Count
End Property

Property Get IsRunning() As Boolean
    IsRunning = mRunning
End Property

Property Get LoopSeconds() As Double
' Seconds taken by the iterations of the last run (without producing the results sheet)
    LoopSeconds = mLoopSeconds
End Property

Property Get ProcessingCorrmatInfo() As Boolean
    ProcessingCorrmatInfo = mProcessingCorrmatInfo
End Property

Private Sub Class_Initialize()
    Dim XLRisk As Worksheet
    Set mWB = Application.ActiveWorkBook
    
    Set XLRisk = SetUpXLRisk
    mSeed = XLRisk.Range("Seed").Value
    mIterations = XLRisk.Range("Iterations").Value
    mScreenUpdate = XLRisk.Range("ScreenUpdate").Value
    mLatinHypercube = XLRisk.Range("LatinHypercube").Value
    mCalculateDataTables = XLRisk.Range("CalcDataTables").Value
    mMacroBefore = XLRisk.Range("MacroBefore").Value
    mMacroAfter = XLRisk.Range("MacroAfter").Value
    On Error Resume Next
    mMacroAfterSimulation = XLRisk.Range("MacroAfterSimulation").Value
    ' Optional setting: TRUE goes back to VBA Rnd() (results of XLRisk 1.0)
    UseLegacyRnd = False
    UseLegacyRnd = XLRisk.Range("LegacyRnd").Value
    ' Optional setting: TRUE keeps recalculating the other open workbooks in every iteration
    mRecalcOtherWorkbooks = False
    mRecalcOtherWorkbooks = XLRisk.Range("RecalcOtherWorkbooks").Value
    mCurrentIteration = 0
    mProcessingCorrmatInfo = False
End Sub

Private Sub SuspendOtherWorkbooks()
' Stops the recalculation of the sheets of every other open workbook.
' Only the sheets that were calculating are recorded, so that exactly those are restored.
    Dim WB As Workbook
    Dim WS As Worksheet

    Set mSuspendedSheets = New Collection
    If mRecalcOtherWorkbooks Then Exit Sub
    On Error Resume Next
    For Each WB In Application.Workbooks
        If Not WB Is mWB Then
            For Each WS In WB.Worksheets
                If WS.EnableCalculation Then
                    WS.EnableCalculation = False
                    If Not WS.EnableCalculation Then mSuspendedSheets.Add WS
                End If
            Next WS
        End If
    Next WB
End Sub

Private Sub ResumeOtherWorkbooks()
' Restores the recalculation of the sheets suspended by SuspendOtherWorkbooks
    Dim Item As Variant

    If mSuspendedSheets Is Nothing Then Exit Sub
    On Error Resume Next
    For Each Item In mSuspendedSheets
        Item.EnableCalculation = True
    Next Item
    Set mSuspendedSheets = Nothing
End Sub

Private Sub SimIteration(Iter As Long)
    Dim RiskInput As ClsRiskInput
    Dim RiskOutput As ClsRiskOutput
    Dim I As Integer
      
    On Error GoTo SSError
    
    'Run Macro Before
    If mMacroBefore <> vbNullString Then Application.Run "'" & mWB.Name & "'!" & mMacroBefore
        
    'Recalculate
    If mCalculateDataTables Then
        Application.Calculate
    Else
        Application.Calculation = xlCalculationSemiautomatic
        Application.Calculation = xlCalculationManual
    End If
    
    'Run Macro After
    If mMacroAfter <> vbNullString Then Application.Run "'" & mWB.Name & "'!" & mMacroAfter
      
    ' Inputs
    mSimResults(Iter, 1) = Iter
    I = 2
    For Each RiskInput In mRiskInputs
        mSimResults(Iter, I) = RiskInput.Cell.Value2
        I = I + 1
    Next RiskInput
    
    'Outputs
    For Each RiskOutput In mRiskOutputs
        mSimResults(Iter, I) = RiskOutput.Cell.Value
        I = I + 1
    Next RiskOutput
    
    Exit Sub
SSError:
    mSimError = True
    mSimErrorMsg = "Error in simulation iteration"
End Sub

Public Sub Run()
    Dim AppCalculation As XlCalculation
    Dim Iter As Long
    Dim OldProduceRandomSample As Boolean
    Dim SimResults As ClsSimResults
    Dim InputGroup As ClsInputGroup
    Dim RiskInput As ClsRiskInput
    Dim StartTime As Double
    Dim LastRefresh As Double
    Dim CurrentTime As Double
    Dim FailMessage As String
    
    ' Statistics functions (RiskMean...) return #N/A while the simulation runs
    mRunning = True

    ' Save ProduceRandomSample
    OldProduceRandomSample = ProduceRandomSample
    
    ' Save Calculation Mode
    AppCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    On Error GoTo Failed
        
    '  Stop Screen Updating?
    Application.ScreenUpdating = mScreenUpdate
    Application.Cursor = xlWait
    
    Set mRiskInputs = New Collection
    CollectRiskInputs mRiskInputs
    If mRiskInputs.Count = 0 Then
        If HasSwapSheet(ActiveWorkBook) Then
            MsgBox "The Risk functions are swapped out for their expected values. " & _
                "Press Swap Functions to restore them and run again.", Title:="XLRisk simulation error"
        Else
            MsgBox "No risk inputs defined", Title:="XLRisk simulation error"
        End If
        GoTo RestoreExcel
    End If
    
    If Not OneRiskFunctionPerCell(mRiskInputs) Then GoTo RestoreExcel
    
    Set mRiskOutputs = New Collection
    CollectRiskOutputs mRiskOutputs
    If mRiskOutputs.Count = 0 Then
        MsgBox "No risk outputs defined", Title:="XLRisk simulation error"
        GoTo RestoreExcel
    End If
        
    'Perform simulation
    UserStopped = False
    ProduceRandomSample = True
    
    'Setup Risk Input Groups
    Set mInputGroups = New Collection
    
    ' Application.Calculate recalculates every open workbook, but the model lives only in mWB
    SuspendOtherWorkbooks

    ' Do a full calculation and process correlation information
    mProcessingCorrmatInfo = True
    Application.CalculateFull
    mProcessingCorrmatInfo = False
        
    ' Create an InputGroup with uncorrelated inputs
    For Each RiskInput In mRiskInputs
        If RiskInput.CorrmatAddress = vbNullString Then
            If InputGroup Is Nothing Then
                Set InputGroup = New ClsInputGroup
                InputGroup.Init mIterations, mLatinHypercube
                mInputGroups.Add InputGroup, NoCorrelationGroupName
            End If
            RiskInput.CorrmatAddress = NoCorrelationGroupName
            RiskInput.Index = InputGroup.AddRiskInput(RiskInput)
        End If
    Next RiskInput
    
    'Seed the random number generator before any sample is drawn,
    'so that a fixed Seed reproduces the whole simulation
    If mSeed <> 0 Then
        RngSeed mSeed
    Else
        RngRandomize
    End If

    'Check Input Groups
    For Each InputGroup In mInputGroups
        If Not InputGroup.IsValid Then GoTo RestoreExcel
        InputGroup.GenerateRndSamples
    Next InputGroup
    
    ' Allocate storage space
    ReDim mSimResults(1 To mIterations, 1 To 1 + mRiskInputs.Count + mRiskOutputs.Count)
    ' Set up Risk InputGroups
      
    StartTime = Timer
    LastRefresh = StartTime
    For Iter = 1 To mIterations
        mCurrentIteration = Iter
        If mSimError Then
            mSimError = False
            MsgBox mSimErrorMsg
            Exit For
        End If
        SimIteration Iter
        ' Refresh the status bar and let Excel process the Stop button a few times
        ' per second, instead of in every iteration
        CurrentTime = Timer
        If CurrentTime - LastRefresh >= REFRESH_SECONDS Or CurrentTime < LastRefresh Or Iter = mIterations Then
            Application.StatusBar = "Iteration: " & CStr(Iter) & "/" & CStr(mIterations)
            DoEvents
            LastRefresh = CurrentTime
        End If
        'Check whether to Stop
        If UserStopped Then
            UserStopped = False
            MsgBox "The simulation was interrupted"
            GoTo RestoreExcel
        End If
    Next Iter
    mCurrentIteration = 0
    mLoopSeconds = Timer - StartTime
    If mLoopSeconds < 0 Then mLoopSeconds = mLoopSeconds + 86400#  ' run crossed midnight
    
    Set SimResults = New ClsSimResults
    SimResults.Init mRiskInputs, mRiskOutputs, mIterations, mSimResults
    SimResults.ProduceResults

    'Run Macro After Simulation
    If mMacroAfterSimulation <> vbNullString Then Application.Run "'" & mWB.Name & "'!" & mMacroAfterSimulation
    GoTo RestoreExcel

Failed:
    ' The original XLRisk stopped without a word on any error; now the user learns why
    FailMessage = Err.Description
    Resume RestoreExcel

RestoreExcel:
    ' The cleanup must finish even if one of its steps fails
    On Error Resume Next
    'Restore ProduceRandomSample
    ProduceRandomSample = OldProduceRandomSample
    
    'Let the other open workbooks recalculate again
    ResumeOtherWorkbooks

    'Statistics functions read the new results in the final recalculation
    mRunning = False
    gResultsRevision = gResultsRevision + 1
    
    'Restore Calculation Mode
    Application.Calculation = AppCalculation
    Application.CalculateFull
    
    ' Restore Status Bar
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

    If Len(FailMessage) > 0 Then
        MsgBox "The simulation stopped because of an error:" & vbCrLf & vbCrLf & FailMessage, _
            vbExclamation, "XLRisk simulation error"
    End If
End Sub

Public Function GetRndSample(Cell As Range) As Double
    Dim RiskInput As ClsRiskInput
    Dim InputGroup As ClsInputGroup
    If (mCurrentIteration < 1) Or Not (Cell.Worksheet.Parent Is mWB) Then
        GetRndSample = RngUniform()
    Else
        Set RiskInput = mRiskInputs.Item(AddressWithSheet(Cell))
        Set InputGroup = mInputGroups.Item(RiskInput.CorrmatAddress)
        GetRndSample = InputGroup.GetRndSample(mCurrentIteration, RiskInput.Index)
    End If
End Function

Public Sub ProcessCorrmatInfo(Cell As Range, CorrmatRng As Range, Index As Long)
    Dim InputGroup As ClsInputGroup
    Dim RiskInput As ClsRiskInput
    Dim Address As String
    
    Address = AddressWithSheet(CorrmatRng)
    On Error GoTo InputNotFound
    Set RiskInput = mRiskInputs.Item(AddressWithSheet(Cell))
    On Error GoTo 0
    RiskInput.CorrmatAddress = Address
    RiskInput.Index = Index

    On Error GoTo InputGroupNotFound
    Set InputGroup = mInputGroups.Item(Address)
    On Error GoTo 0
    InputGroup.AddRiskInput RiskInput
    Exit Sub
InputGroupNotFound:
    On Error GoTo 0
    Set InputGroup = New ClsInputGroup
    InputGroup.Init mIterations, mLatinHypercube
    Set InputGroup.CorrmatRng = CorrmatRng
    mInputGroups.Add InputGroup, Address
    InputGroup.AddRiskInput RiskInput
    Exit Sub
InputNotFound:
    Err.Raise Number:=vbObjectError + 15, _
              source:="ProcessCorrmatInfo", _
              Description:="Risk Input Not Found"
End Sub


