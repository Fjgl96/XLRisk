Attribute VB_Name = "RiskFunctions"
Option Explicit
Option Base 1

Public ProduceRandomSample As Boolean

' Random numbers come from RngUniform (MRG32k3a, see the Rng module)

Public Function RiskFunctionList() As Variant
' Returns a list of risk functions
' Needs to be updated as new risk functions are added
    RiskFunctionList = Array("RiskBernoulli", "RiskBeta", "RiskBinomial", _
    "RiskCumul", "RiskDiscrete", "RiskDUniform", "RiskErlang", "RiskExpon", _
    "RiskGamma", "RiskLogNorm", "RiskNormal", "RiskTriang", "RiskPert", _
    "RiskUniform", "RiskWeibull", "RiskLognorm2", "RiskPoisson", "RiskTriGen")
End Function

Public Function RiskBinomial(N As Long, P As Double, Optional Corrmat As Long = 0) As Variant
Attribute RiskBinomial.VB_Description = "Generate random sample from a Binomial destribution"
Attribute RiskBinomial.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (N < 0) Or (P < 0) Or (P > 1) Then
        RiskBinomial = CVErr(xlErrValue)
        Exit Function
    End If
    
    If (N = 0) Or (P = 0) Then
        RiskBinomial = 0
    ElseIf ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskBinomial = WorksheetFunction.Binom_Inv(N, P, RndValue)
    Else
        ' Expected value, not rounded (a Long result used to round 40 x 0.15 or 0.15 itself)
        RiskBinomial = CDbl(N) * P
    End If
End Function

Public Function RiskBernoulli(P As Double, Optional Corrmat As Long = 0) As Variant
Attribute RiskBernoulli.VB_Description = "Generate random sample from a Bernoulli destribution"
Attribute RiskBernoulli.VB_ProcData.VB_Invoke_Func = " \n20"
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (P < 0) Or (P > 1) Then
        RiskBernoulli = CVErr(xlErrValue)
        Exit Function
    End If
    
    RiskBernoulli = RiskBinomial(1, P)
End Function

Public Function RiskExpon(Mean As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskExpon.VB_Description = "Generate random sample from an exponential destribution"
Attribute RiskExpon.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If Mean <= 0 Then
        RiskExpon = CVErr(xlErrValue)
        Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskExpon = -Mean * Log(1 - RndValue)
    Else
        RiskExpon = Mean
    End If
End Function

Function RiskErlang(kappa As Long, lambda As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskErlang.VB_Description = "Generate random sample from an Erlang destribution"
Attribute RiskErlang.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random Sample from an Erlang distribution (same as Gamma with an integer alpha)

    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (kappa <= 0) Or (lambda <= 0) Then
      RiskErlang = CVErr(xlErrValue)
      Exit Function
    End If
    
    RiskErlang = RiskGamma(CDbl(kappa), lambda)
End Function

Function RiskGamma(alpha As Double, beta As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskGamma.VB_Description = "Generate random sample from a Gamma destribution"
Attribute RiskGamma.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Gamma distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (alpha <= 0) Or (beta <= 0) Then
      RiskGamma = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskGamma = WorksheetFunction.Gamma_Inv(RndValue, alpha, beta)
    Else
        RiskGamma = alpha * beta
    End If
End Function

Public Function RiskUniform(Min As Double, Max As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskUniform.VB_Description = "Generate random sample from a uniform destribution"
Attribute RiskUniform.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Uniform distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If Max < Min Then
      RiskUniform = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskUniform = Min + RndValue * (Max - Min)
    Else
        RiskUniform = (Min + Max) / 2
    End If
End Function

Public Function RiskDiscrete(Values As Variant, Probabilities As Variant, Optional Corrmat As Long = 0) As Variant
Attribute RiskDiscrete.VB_Description = "Generate random sample from a discrete probability destribution"
Attribute RiskDiscrete.VB_ProcData.VB_Invoke_Func = " \n20"
'   Generates samples for a discrete distribution specified by the values and their probabilities.
'   The values may come in any order: they are sorted together with their probabilities, so the
'   inverse of the cumulative distribution stays increasing (needed by Latin Hypercube and RiskCorrmat).

    Dim V() As Double
    Dim Pr() As Double
    Dim Count As Long
    Dim I As Long
    Dim J As Long
    Dim RndValue As Double
    Dim Cum As Double
    Dim Total As Double
    Dim TmpV As Double
    Dim TmpP As Double

    Application.Volatile (ProduceRandomSample)

    If Not ReadNumbers(Values, V) Then
        RiskDiscrete = CVErr(xlErrValue)
        Exit Function
    End If
    If Not ReadNumbers(Probabilities, Pr) Then
        RiskDiscrete = CVErr(xlErrValue)
        Exit Function
    End If
    Count = UBound(V)
    If UBound(Pr) <> Count Then
        RiskDiscrete = CVErr(xlErrValue)
        Exit Function
    End If

    Total = 0
    For I = 1 To Count
        If Pr(I) < 0 Then
            RiskDiscrete = CVErr(xlErrValue)
            Exit Function
        End If
        Total = Total + Pr(I)
    Next I
    If Round(Total, 10) <> 1 Then
        RiskDiscrete = CVErr(xlErrValue)
        Exit Function
    End If

    If ProduceRandomSample Then
        ' Insertion sort of the pairs (value, probability) by value
        For I = 2 To Count
            TmpV = V(I)
            TmpP = Pr(I)
            J = I - 1
            Do While J >= 1
                If V(J) <= TmpV Then Exit Do
                V(J + 1) = V(J)
                Pr(J + 1) = Pr(J)
                J = J - 1
            Loop
            V(J + 1) = TmpV
            Pr(J + 1) = TmpP
        Next I
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        Cum = 0
        For I = 1 To Count
            Cum = Cum + Pr(I)
            If RndValue <= Cum Then
                RiskDiscrete = V(I)
                Exit Function
            End If
        Next I
        RiskDiscrete = V(Count)
    Else
        Cum = 0
        For I = 1 To Count
            Cum = Cum + V(I) * Pr(I)
        Next I
        RiskDiscrete = Cum
    End If
End Function

Private Function ReadNumbers(Source As Variant, Result() As Double) As Boolean
' Copies the numbers of a range, an array or a single value into Result(1 To n).
' Returns False if Source holds no numbers or holds text, blanks or errors.
    Dim Data As Variant
    Dim Item As Variant
    Dim N As Long

    On Error GoTo Invalid
    If TypeName(Source) = "Range" Then
        Data = Source.Value
    Else
        Data = Source
    End If
    N = 0
    If IsArray(Data) Then
        For Each Item In Data
            If Not IsPlainNumber(Item) Then GoTo Invalid
            N = N + 1
            ReDim Preserve Result(1 To N)
            Result(N) = CDbl(Item)
        Next Item
    Else
        If Not IsPlainNumber(Data) Then GoTo Invalid
        N = 1
        ReDim Result(1 To 1)
        Result(1) = CDbl(Data)
    End If
    ReadNumbers = (N > 0)
    Exit Function
Invalid:
    ReadNumbers = False
End Function

Private Function IsPlainNumber(X As Variant) As Boolean
' True for numbers; False for text, blanks, booleans and errors
    Select Case VarType(X)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            IsPlainNumber = True
        Case Else
            IsPlainNumber = False
    End Select
End Function

Public Function RiskDUniform(Values As Variant, Optional Corrmat As Long = 0) As Variant
Attribute RiskDUniform.VB_Description = "Generate random sample from a uniform discrete destribution"
Attribute RiskDUniform.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Discrete Uniform distribution
'   Values can be a range or an array of values
    
    Dim Count As Integer
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    Count = WorksheetFunction.Count(Values)

    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskDUniform = Values(Int(RndValue * Count) + 1)
    Else
        RiskDUniform = WorksheetFunction.Sum(Values) / Count
    End If
End Function

Public Function RiskNormal(Mean As Double, StDev As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskNormal.VB_Description = "Generate random sample from a normal destribution"
Attribute RiskNormal.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Normal distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskNormal = WorksheetFunction.Norm_Inv(RndValue, Mean, StDev)
    Else
        RiskNormal = Mean
    End If
End Function

Public Function RiskLogNorm(Mean As Double, StDev As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskLogNorm.VB_Description = "Generate random sample from a lognormal destribution"
Attribute RiskLogNorm.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Log Normal distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskLogNorm = WorksheetFunction.LogNorm_Inv(RndValue, Mean, StDev)
    Else
        RiskLogNorm = Exp(Mean + 0.5 * StDev ^ 2)
    End If
End Function

Public Function RiskLognorm2(Mu As Double, Sigma As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskLognorm2.VB_Description = "Generate random sample from a lognormal distribution given the mean and standard deviation of Ln(X). Same as RiskLogNorm"
Attribute RiskLognorm2.VB_ProcData.VB_Invoke_Func = " \n20"
'   Same parameters as RiskLognorm2 in @RISK: mean and standard deviation of Ln(X).
'   XLRisk's RiskLogNorm uses these same parameters.
    Application.Volatile (ProduceRandomSample)
    RiskLognorm2 = RiskLogNorm(Mu, Sigma)
End Function

Public Function RiskPoisson(Lambda As Double, Optional Corrmat As Long = 0) As Variant
Attribute RiskPoisson.VB_Description = "Generate random sample from a Poisson distribution"
Attribute RiskPoisson.VB_ProcData.VB_Invoke_Func = " \n20"
'   Number of events in a period when on average Lambda events occur
    Dim RndValue As Double
    Dim K As Long
    Dim Cum As Double

    Application.Volatile (ProduceRandomSample)
    'Error checking
    If Lambda <= 0 Then
        RiskPoisson = CVErr(xlErrValue)
        Exit Function
    End If
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        ' Inverse of the cumulative distribution
        K = 0
        Cum = WorksheetFunction.Poisson_Dist(0, Lambda, True)
        Do While Cum < RndValue And K < 1000000
            K = K + 1
            Cum = WorksheetFunction.Poisson_Dist(K, Lambda, True)
        Loop
        RiskPoisson = K
    Else
        RiskPoisson = Lambda
    End If
End Function

Function RiskTriang(Min As Double, Mode As Double, Max As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskTriang.VB_Description = "Generate random sample from a triangular destribution"
Attribute RiskTriang.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Triangular distribution
'   See https://en.wikipedia.org/wiki/Triangular_distribution
    
    Dim LowerRange As Double
    Dim HigherRange As Double
    Dim TotalRange As Double
    Dim CumulativeProb As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (Mode <= Min) Or (Max <= Mode) Then
      RiskTriang = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        LowerRange = Mode - Min
        HigherRange = Max - Mode
        TotalRange = Max - Min
        If gSimulation Is Nothing Then
            CumulativeProb = RngUniform()
        Else
            CumulativeProb = gSimulation.GetRndSample(Application.Caller)
        End If
        If CumulativeProb < (LowerRange / TotalRange) Then
            RiskTriang = Min + Sqr(CumulativeProb * LowerRange * TotalRange)
        Else
            RiskTriang = Max - Sqr((1 - CumulativeProb) * HigherRange * TotalRange)
        End If
    Else
        RiskTriang = (Min + Mode + Max) / 3
    End If
End Function

Public Function RiskTriGen(Bottom As Double, MostLikely As Double, Top As Double, _
    Optional BottomPerc As Double = 0.1, Optional TopPerc As Double = 0.9, Optional Corrmat As Long = 0) As Variant
Attribute RiskTriGen.VB_Description = "Generate random sample from a triangular distribution given two percentiles and the most likely value"
Attribute RiskTriGen.VB_ProcData.VB_Invoke_Func = " \n20"
'   Triangular distribution defined like @RISK's RiskTrigen: Bottom is the value at the
'   BottomPerc percentile, Top the value at the TopPerc percentile (10% and 90% by default).
'   The minimum and maximum of the triangle are found by fixed-point iteration.
    Dim A As Double
    Dim B As Double
    Application.Volatile (ProduceRandomSample)

    If Not TriGenLimits(Bottom, MostLikely, Top, BottomPerc, TopPerc, A, B) Then
        RiskTriGen = CVErr(xlErrValue)
        Exit Function
    End If
    RiskTriGen = TriangSample(A, MostLikely, B)
End Function

Public Function TriGenLimits(ByVal Bottom As Double, ByVal MostLikely As Double, ByVal Top As Double, _
    ByVal BottomPerc As Double, ByVal TopPerc As Double, ByRef A As Double, ByRef B As Double) As Boolean
' Finds the minimum A and maximum B of the triangular distribution whose cumulative probability
' is BottomPerc at Bottom and TopPerc at Top. Returns False if there is no such triangle.
    Dim I As Long
    Dim NewA As Double
    Dim NewB As Double

    TriGenLimits = False
    If Not (Bottom <= MostLikely And MostLikely <= Top And Bottom < Top) Then Exit Function
    If Not (BottomPerc > 0 And BottomPerc < TopPerc And TopPerc < 1) Then Exit Function
    A = Bottom
    B = Top
    For I = 1 To 1000
        NewA = Bottom - Sqr(BottomPerc * (B - A) * (MostLikely - A))
        NewB = Top + Sqr((1 - TopPerc) * (B - A) * (B - MostLikely))
        If Abs(NewA - A) <= 0.000000000001 * (1 + Abs(A)) And Abs(NewB - B) <= 0.000000000001 * (1 + Abs(B)) Then
            A = NewA
            B = NewB
            TriGenLimits = (A < MostLikely And MostLikely < B)
            Exit Function
        End If
        A = NewA
        B = NewB
    Next I
End Function

Private Function TriangSample(ByVal Min As Double, ByVal Mode As Double, ByVal Max As Double) As Double
' Sample (or expected value when samples are off) of a triangular distribution, for RiskTriGen
    Dim P As Double
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            P = RngUniform()
        Else
            P = gSimulation.GetRndSample(Application.Caller)
        End If
        If P < (Mode - Min) / (Max - Min) Then
            TriangSample = Min + Sqr(P * (Mode - Min) * (Max - Min))
        Else
            TriangSample = Max - Sqr((1 - P) * (Max - Mode) * (Max - Min))
        End If
    Else
        TriangSample = (Min + Mode + Max) / 3
    End If
End Function

Function RiskBeta(alpha As Double, beta As Double, Optional A As Double = 0, Optional B As Double = 1, Optional Corrmat As Long = 0) As Double
Attribute RiskBeta.VB_Description = "Generate random sample from a beta destribution"
Attribute RiskBeta.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random Sample from a Beta distribution
    
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (B <= A) Then
      RiskBeta = CVErr(xlErrValue)
      Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskBeta = WorksheetFunction.Beta_Inv(RndValue, alpha, beta, A, B)
    Else
        RiskBeta = A + (alpha / (alpha + beta)) * (B - A)
    End If
End Function

Function RiskPert(Min As Double, Mode As Double, Max As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskPert.VB_Description = "Generate random sample from a PERT destribution"
Attribute RiskPert.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a Pert distribution a special case of the Beta distribution
'   A smoother version of the triangular distribution
'   See https://www.coursera.org/lecture/excel-vba-for-creative-problem-solving-part-3-projects/the-beta-pert-distribution-GJVsK

    Dim alpha As Double
    Dim beta As Double
    
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (Mode <= Min) Or (Max <= Mode) Then
      RiskPert = CVErr(xlErrValue)
      Exit Function
    End If
    
    ' Reparameterize the RiskBeta distribution as per above link
    alpha = (4 * Mode + Max - 5 * Min) / (Max - Min)
    beta = (5 * Max - Min - 4 * Mode) / (Max - Min)
    RiskPert = RiskBeta(alpha, beta, Min, Max)
End Function

Public Function RiskCumul(MinValue As Double, MaxValue As Double, _
                          XValues As Variant, YValues As Variant, Optional Corrmat As Long = 0) As Double
Attribute RiskCumul.VB_Description = "Generate random sample from a cumulative destribution"
Attribute RiskCumul.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random sample from a distribution specifed by data points of its cumulative function
'   Values can be a range or an array of values
    
    Dim I As Integer
    Dim Count As Integer
    Dim ParamError As Boolean
    Dim RndValue As Double
    Dim Slope As Double
    
    Application.Volatile (ProduceRandomSample)
    
    Count = WorksheetFunction.Count(XValues)

    'Error checking
    ParamError = False
    If MinValue > MaxValue Then
        ParamError = True
    ElseIf Count < 1 Then
      ParamError = True
    ElseIf WorksheetFunction.Count(XValues) <> Count Then
        ParamError = True
    ElseIf (XValues(1) <= MinValue) Or (XValues(Count) >= MaxValue) Then
        ParamError = True
    Else
        'Check that XValues and YValues are in strict increasing order
        For I = 2 To Count
            If XValues(I) <= XValues(I - 1) Then
                ParamError = True
                Exit For
            End If
        Next I
        If Not ParamError Then
            For I = 2 To Count
                If YValues(I) <= YValues(I - 1) Then
                    ParamError = True
                    Exit For
                End If
            Next I
        End If
    End If
    
    If ParamError Then
        RiskCumul = CVErr(xlErrValue)
        Exit Function
    End If
    
    If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        If RndValue <= YValues(1) Then
            RiskCumul = MinValue + (XValues(1) - MinValue) * RndValue / YValues(1)
        ElseIf RndValue > YValues(Count) Then
            RiskCumul = XValues(Count) + (MaxValue - XValues(Count)) * _
                (RndValue - YValues(Count)) / (1 - YValues(Count))
        Else
          For I = 2 To Count
            If RndValue <= YValues(I) Then
                RiskCumul = XValues(I - 1) + (XValues(I) - XValues(I - 1)) * _
                    (RndValue - YValues(I - 1)) / (YValues(I) - YValues(I - 1))
                Exit For
            End If
          Next I
        End If
    Else
        Slope = YValues(1) / (XValues(1) - MinValue)
        RiskCumul = 0.5 * Slope * (XValues(1) ^ 2 - MinValue ^ 2)
        For I = 2 To Count
            Slope = (YValues(I) - YValues(I - 1)) / (XValues(I) - XValues(I - 1))
            RiskCumul = RiskCumul + 0.5 * Slope * (XValues(I) ^ 2 - XValues(I - 1) ^ 2)
        Next I
        Slope = (1 - YValues(Count)) / (MaxValue - XValues(Count))
        RiskCumul = RiskCumul + 0.5 * Slope * (MaxValue ^ 2 - XValues(Count) ^ 2)
    End If
End Function

Function RiskWeibull(alpha As Double, beta As Double, Optional Corrmat As Long = 0) As Double
Attribute RiskWeibull.VB_Description = "Generate random sample from a Weibull destribution"
Attribute RiskWeibull.VB_ProcData.VB_Invoke_Func = " \n20"
'   Random Sample from a Weibull distribution (same as Gamma with an integer alpha)
    Dim RndValue As Double
    Application.Volatile (ProduceRandomSample)
    
    'Error checking
    If (alpha <= 0) Or (beta <= 0) Then
      RiskWeibull = CVErr(xlErrValue)
      Exit Function
    End If
    
   If ProduceRandomSample Then
        If gSimulation Is Nothing Then
            RndValue = RngUniform()
        Else
            RndValue = gSimulation.GetRndSample(Application.Caller)
        End If
        RiskWeibull = beta * (-Log(1 - RndValue)) ^ (1 / alpha)
    Else
        RiskWeibull = beta * WorksheetFunction.Gamma(1 + 1 / alpha)
    End If
End Function

Private Sub CreateFunctionDescription(FuncName As String, FuncDesc As String, ArgDesc As Variant)
'   Creates a description for a function and its arguments
'   They are used by the Excel function wizard
    On Error Resume Next
    Call Application.MacroOptions( _
      Macro:=FuncName, _
      Description:=FuncDesc, _
      Category:="XLRisk", _
      ArgumentDescriptions:=ArgDesc)
End Sub

Public Sub FunctionDescriptions()
    Const CorrMatDesc As String = "Optional. Placeholder for linking to a correlation matrix using the RiskCorrmat function)"
    ' Distribution functions
    Call CreateFunctionDescription("RiskBernoulli", "Generate random sample from a Bernoulli destribution", _
        Array("Probability of success", CorrMatDesc))
    Call CreateFunctionDescription("RiskBeta", "Generate random sample from a beta destribution", _
        Array("Shape parameter", "Shape parameter", "Optional minimum - 0 if omitted", "Optional maximum - 1 if omitted", CorrMatDesc))
    Call CreateFunctionDescription("RiskBinomial", "Generate random sample from a Binomial destribution", _
        Array("Number of Trials", "Probability of success", CorrMatDesc))
    Call CreateFunctionDescription("RiskCumul", "Generate random sample from a cumulative destribution", _
        Array("Minimum Value", "Maximum Value", "Range or array of X coordinates", _
        "Range or Array of Y coordinates (cumulative probabilities)", CorrMatDesc))
    Call CreateFunctionDescription("RiskDiscrete", "Generate random sample from a discrete probability destribution", _
        Array("Range or array of values (any order)", "Range or array of probabilities", CorrMatDesc))
    Call CreateFunctionDescription("RiskDUniform", "Generate random sample from a uniform discrete destribution", _
        Array("Range or array of values", CorrMatDesc))
    Call CreateFunctionDescription("RiskGamma", "Generate random sample from a Gamma destribution", _
        Array("Shape parameter", "Scale parameter", CorrMatDesc))
    Call CreateFunctionDescription("RiskErlang", "Generate random sample from an Erlang destribution", _
        Array("Integer shape parameter", "Scale parameter", CorrMatDesc))
    Call CreateFunctionDescription("RiskExpon", "Generate random sample from an exponential destribution", _
        Array("Mean", CorrMatDesc))
    Call CreateFunctionDescription("RiskNormal", "Generate random sample from a normal destribution", _
        Array("Mean", "Standard Deviation", CorrMatDesc))
    Call CreateFunctionDescription("RiskLogNorm", "Generate random sample from a lognormal destribution", _
        Array("Mean of Ln(X)", "Standard Deviation of Ln(X)", CorrMatDesc))
    Call CreateFunctionDescription("RiskPert", "Generate random sample from a PERT destribution", _
        Array("Minimum value", "Mode", "Maximum value", CorrMatDesc))
    Call CreateFunctionDescription("RiskTriang", "Generate random sample from a triangular destribution", _
        Array("Minimum value", "Mode", "Maximum value", CorrMatDesc))
    Call CreateFunctionDescription("RiskTriGen", "Generate random sample from a triangular distribution given two percentiles and the most likely value", _
        Array("Value at the bottom percentile", "Most likely value", "Value at the top percentile", _
        "Optional bottom percentile - 0.1 if omitted", "Optional top percentile - 0.9 if omitted", CorrMatDesc))
    Call CreateFunctionDescription("RiskUniform", "Generate random sample from a uniform destribution", _
        Array("Minimum value", "Maximum Value", CorrMatDesc))
    Call CreateFunctionDescription("RiskWeibull", "Generate random sample from a Weibull destribution", _
        Array("Shape parameter", "Scale parameter", CorrMatDesc))
    Call CreateFunctionDescription("RiskLognorm2", "Generate random sample from a lognormal destribution (same as RiskLogNorm)", _
        Array("Mean of Ln(X)", "Standard Deviation of Ln(X)", CorrMatDesc))
    Call CreateFunctionDescription("RiskPoisson", "Generate random sample from a Poisson destribution", _
        Array("Mean number of events", CorrMatDesc))
    ' Outputs and statistics of the latest simulation
    Call CreateFunctionDescription("RiskOutput", "Marks the cell as a simulation output. Use it as =RiskOutput(""Name"") + formula. Returns 0", _
        Array("Optional name of the output", "Optional. Ignored (kept for compatibility)", "Optional. Ignored (kept for compatibility)"))
    Call CreateFunctionDescription("RiskMean", "Mean of a cell in the latest simulation", Array("Output or input cell"))
    Call CreateFunctionDescription("RiskStdDev", "Standard deviation of a cell in the latest simulation", Array("Output or input cell"))
    Call CreateFunctionDescription("RiskMin", "Minimum of a cell in the latest simulation", Array("Output or input cell"))
    Call CreateFunctionDescription("RiskMax", "Maximum of a cell in the latest simulation", Array("Output or input cell"))
    Call CreateFunctionDescription("RiskPercentile", "Value of a cell at a percentile in the latest simulation", _
        Array("Output or input cell", "Percentile between 0 and 1, for example 0.05"))
    Call CreateFunctionDescription("RiskPtoX", "Same as RiskPercentile", _
        Array("Output or input cell", "Percentile between 0 and 1, for example 0.05"))
    Call CreateFunctionDescription("RiskTarget", "Probability that a cell is less than or equal to a value in the latest simulation", _
        Array("Output or input cell", "Target value"))
    Call CreateFunctionDescription("RiskXtoP", "Same as RiskTarget", Array("Output or input cell", "Target value"))
    ' Other functions
    Call CreateFunctionDescription("RiskCorrmat", "When added as an optional last argument to a Risk function it creates a link to a correlation matrix", _
        Array("Range containing a lower triangular correlation matrix", "Position of this risk input function in the correlation matrix"))
    Call CreateFunctionDescription("RiskCorrectCorrmat", "Fixes an invalid correlation matrix and returns the corrected matrix", _
        Array("Range containing a lower triangular correlation matrix", "Optional tolerance with default value 1.0E-10"))
    Call CreateFunctionDescription("RiskIsValidCorrmat", "Checks whether a correlation matrix is valid.  Returns a Boolean value", _
        Array("Range containing a lower triangular correlation matrix", "Optional tolerance with default value 1.0E-10"))
    Call CreateFunctionDescription("RiskSCorrel", "Returns the Spearman's rank correlation of two arrays", _
        Array("First Range or array", "Second Range or array"))
End Sub

Public Function RiskCorrmat(CorrmatRng As Range, Position As Long) As Long
Attribute RiskCorrmat.VB_Description = "When added as an optional last argument to a Risk function it creates a link to a correlation matrix"
Attribute RiskCorrmat.VB_ProcData.VB_Invoke_Func = " \n20"
    If Not (gSimulation Is Nothing) And (TypeOf Application.Caller Is Range) Then
        If gSimulation.ProcessingCorrmatInfo And (gSimulation.ActiveWorkBook Is Application.Caller.Parent.Parent) Then
            gSimulation.ProcessCorrmatInfo Application.Caller, CorrmatRng, Position
        End If
    End If
    RiskCorrmat = 0
End Function

Public Function RndM(Optional ByVal Number As Long) As Double
' Wichman-Hill Pseudo Random Number Generator: an alternative for VB Rnd() function
' http://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd%28%29-function
' See also https://www.random.org/analysis/#visual
' Not currently used
    Static lngX As Long
    Static lngY As Long
    Static lngZ As Long
    Static blnInit As Boolean
    Dim dblRnd As Double
    ' if initialized and no input number given
    If blnInit And Number = 0 Then
        ' lngX, lngY and lngZ will never be 0
        lngX = (171 * lngX) Mod 30269
        lngY = (172 * lngY) Mod 30307
        lngZ = (170 * lngZ) Mod 30323
    Else
        ' if no initialization, use Timer, otherwise ensure positive Number
        If Number = 0 Then Number = Timer * 60 Else Number = Number And &H7FFFFFFF
        lngX = (Number Mod 30269)
        lngY = (Number Mod 30307)
        lngZ = (Number Mod 30323)
        ' lngX, lngY and lngZ must be bigger than 0
        If lngX <= 0 Then lngX = 171
        If lngY <= 0 Then lngY = 172
        If lngZ <= 0 Then lngZ = 170
        ' mark initialization state
        blnInit = True
    End If
    ' generate a random number
    dblRnd = CDbl(lngX) / 30269# + CDbl(lngY) / 30307# + CDbl(lngZ) / 30323#
    ' return a value between 0 and 1
    RndM = dblRnd - Int(dblRnd)
End Function
