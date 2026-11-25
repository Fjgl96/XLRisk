Attribute VB_Name = "Rng"
Option Explicit

' Random number generator for XLRisk
'
' MRG32k3a combined multiple recursive generator
' P. L'Ecuyer, "Good Parameters and Implementations for Combined Multiple
' Recursive Random Number Generators", Operations Research 47(1), 1999.
'
' - Period of about 2^191 and passes the TestU01 BigCrush battery.
' - Every intermediate product stays below 2^53, so the Double arithmetic
'   used here is exact: the same seed gives the same sequence on 32-bit
'   Excel, 64-bit Excel and Excel for the Mac.
' - Replaces VBA Rnd(), a 24-bit linear congruential generator whose period
'   (about 16.7 million) is too short for large simulations.
'
' Set the named cell "LegacyRnd" to TRUE in the XLRisk sheet to go back to
' VBA Rnd() and reproduce results of XLRisk 1.0.

Private Const M1 As Double = 4294967087#
Private Const M2 As Double = 4294944443#
Private Const A12 As Double = 1403580#
Private Const A13N As Double = 810728#
Private Const A21 As Double = 527612#
Private Const A23N As Double = 1370589#
Private Const NORM_DIVISOR As Double = 4294967088#
Private Const TWO_POW_32 As Double = 4294967296#

Private S10 As Double
Private S11 As Double
Private S12 As Double
Private S20 As Double
Private S21 As Double
Private S22 As Double
Private mSeeded As Boolean

' When True, RngUniform falls back to VBA Rnd()
Public UseLegacyRnd As Boolean

Public Function RngUniform() As Double
' Returns a uniform random sample in the open interval (0, 1)
    Dim P1 As Double
    Dim P2 As Double
    Dim K As Double

    If UseLegacyRnd Then
        RngUniform = Rnd()
        Exit Function
    End If
    If Not mSeeded Then RngRandomize

    ' First component
    P1 = A12 * S11 - A13N * S10
    K = Int(P1 / M1)
    P1 = P1 - K * M1
    If P1 < 0# Then P1 = P1 + M1
    If P1 >= M1 Then P1 = P1 - M1
    S10 = S11
    S11 = S12
    S12 = P1

    ' Second component
    P2 = A21 * S22 - A23N * S20
    K = Int(P2 / M2)
    P2 = P2 - K * M2
    If P2 < 0# Then P2 = P2 + M2
    If P2 >= M2 Then P2 = P2 - M2
    S20 = S21
    S21 = S22
    S22 = P2

    ' Combination
    If P1 > P2 Then
        RngUniform = (P1 - P2) / NORM_DIVISOR
    Else
        RngUniform = (P1 - P2 + M1) / NORM_DIVISOR
    End If
End Function

Public Sub RngSeed(ByVal Seed As Double)
' Seeds the generator from any number. The same seed always gives the same sequence.
    Dim X As Double

    X = DMod(Abs(Int(Seed)), TWO_POW_32)
    X = NextLcg(X)
    S10 = DMod(X, M1)
    X = NextLcg(X)
    S11 = DMod(X, M1)
    X = NextLcg(X)
    S12 = DMod(X, M1)
    X = NextLcg(X)
    S20 = DMod(X, M2)
    X = NextLcg(X)
    S21 = DMod(X, M2)
    X = NextLcg(X)
    S22 = DMod(X, M2)

    ' Each component needs at least one non-zero seed
    If S10 = 0# And S11 = 0# And S12 = 0# Then S10 = 12345#
    If S20 = 0# And S21 = 0# And S22 = 0# Then S20 = 12345#
    mSeeded = True

    ' Keep the legacy generator reproducible as well
    If UseLegacyRnd Then
        Rnd (-1)
        Randomize (Seed)
    End If
End Sub

Public Sub RngRandomize()
' Seeds the generator from the clock (used when the Seed setting is 0)
    Randomize
    RngSeed Int(Rnd() * 2147483647#) + Int(Timer * 1000#)
End Sub

Public Sub RngSetState(ByVal A0 As Double, ByVal A1 As Double, ByVal A2 As Double, _
                       ByVal B0 As Double, ByVal B1 As Double, ByVal B2 As Double)
' Sets the six seeds directly. Used by the self tests to check published reference values.
    S10 = A0
    S11 = A1
    S12 = A2
    S20 = B0
    S21 = B1
    S22 = B2
    mSeeded = True
End Sub

Private Function NextLcg(ByVal X As Double) As Double
' 32-bit linear congruential step used only to spread a seed into six values
    NextLcg = DMod(69069# * X + 1#, TWO_POW_32)
End Function

Private Function DMod(ByVal A As Double, ByVal B As Double) As Double
' Modulo for Doubles (the VBA Mod operator converts to Long and overflows above 2^31)
    Dim R As Double
    R = A - Int(A / B) * B
    If R < 0# Then R = R + B
    If R >= B Then R = R - B
    DMod = R
End Function
