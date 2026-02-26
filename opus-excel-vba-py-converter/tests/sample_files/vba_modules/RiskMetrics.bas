Attribute VB_Name = "RiskMetrics"
' ============================================================
' Module: RiskMetrics
' Purpose: Risk and capital adequacy calculations including
'          Value at Risk (VaR), Expected Shortfall, and
'          Solvency ratios.
' ============================================================
Option Explicit

' ============================================================
' Function: NormalInverse
' Approximation of the inverse standard normal CDF
' using the Abramowitz & Stegun rational approximation.
' ============================================================
Public Function NormalInverse(ByVal p As Double) As Double
    ' Coefficients
    Const a1 As Double = -3.969683028665376E+01
    Const a2 As Double = 2.209460984245205E+02
    Const a3 As Double = -2.759285104469687E+02
    Const a4 As Double = 1.383577518672690E+02
    Const a5 As Double = -3.066479806614716E+01
    Const a6 As Double = 2.506628277459239E+00
    
    Const b1 As Double = -5.447609879822406E+01
    Const b2 As Double = 1.615858368580409E+02
    Const b3 As Double = -1.556989798598866E+02
    Const b4 As Double = 6.680131188771972E+01
    Const b5 As Double = -1.328068155288572E+01
    
    Const c1 As Double = -7.784894002430293E-03
    Const c2 As Double = -3.223964580411365E-01
    Const c3 As Double = -2.400758277161838E+00
    Const c4 As Double = -2.549732539343734E+00
    Const c5 As Double = 4.374664141464968E+00
    Const c6 As Double = 2.938163982698783E+00
    
    Const d1 As Double = 7.784695709041462E-03
    Const d2 As Double = 3.224671290700398E-01
    Const d3 As Double = 2.445134137142996E+00
    Const d4 As Double = 3.754408661907416E+00
    
    Dim q As Double, r As Double
    
    If p < 0.02425 Then
        ' Lower tail
        q = Sqr(-2# * Log(p))
        NormalInverse = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                        ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#)
    ElseIf p <= 0.97575 Then
        ' Central region
        q = p - 0.5
        r = q * q
        NormalInverse = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
                        (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1#)
    Else
        ' Upper tail
        q = Sqr(-2# * Log(1# - p))
        NormalInverse = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
                         ((((d1 * q + d2) * q + d3) * q + d4) * q + 1#))
    End If
End Function

' ============================================================
' Function: ValueAtRisk
' Parametric VaR assuming normal distribution.
'   VaR = mu + sigma * z_alpha
' ============================================================
Public Function ValueAtRisk(ByVal meanLoss As Double, ByVal stdDev As Double, _
                             ByVal confidenceLevel As Double) As Double
    ValueAtRisk = meanLoss + stdDev * NormalInverse(confidenceLevel)
End Function

' ============================================================
' Function: ExpectedShortfall
' Conditional Tail Expectation (CTE) / Expected Shortfall
' for a normal distribution.
'   ES = mu + sigma * phi(z_alpha) / (1 - alpha)
' ============================================================
Public Function ExpectedShortfall(ByVal meanLoss As Double, ByVal stdDev As Double, _
                                   ByVal confidenceLevel As Double) As Double
    Dim z As Double
    z = NormalInverse(confidenceLevel)
    
    ' Standard normal PDF at z
    Dim phi As Double
    phi = Exp(-0.5 * z * z) / Sqr(2# * 3.14159265358979)
    
    ExpectedShortfall = meanLoss + stdDev * phi / (1# - confidenceLevel)
End Function

' ============================================================
' Function: SolvencyRatio
' Calculates the solvency ratio = Available Capital / Required Capital.
' ============================================================
Public Function SolvencyRatio(ByVal availableCapital As Double, _
                               ByVal requiredCapital As Double) As Double
    If requiredCapital > 0 Then
        SolvencyRatio = availableCapital / requiredCapital
    Else
        SolvencyRatio = 0
    End If
End Function

' ============================================================
' Function: RiskBasedCapital
' Simplified Risk-Based Capital calculation.
'   RBC = Sqrt(C1^2 + C2^2 + C3^2 + C4^2)
'   where C1-C4 are risk charges for different categories.
' ============================================================
Public Function RiskBasedCapital(ByVal assetRisk As Double, _
                                  ByVal insuranceRisk As Double, _
                                  ByVal interestRateRisk As Double, _
                                  ByVal businessRisk As Double) As Double
    RiskBasedCapital = Sqr(assetRisk ^ 2 + insuranceRisk ^ 2 + _
                           interestRateRisk ^ 2 + businessRisk ^ 2)
End Function

' ============================================================
' Function: LossRatio
' Calculates the loss ratio = Incurred Losses / Earned Premiums
' ============================================================
Public Function LossRatio(ByVal incurredLosses As Double, _
                           ByVal earnedPremiums As Double) As Double
    If earnedPremiums > 0 Then
        LossRatio = incurredLosses / earnedPremiums
    Else
        LossRatio = 0
    End If
End Function

' ============================================================
' Function: CombinedRatio
' Combined Ratio = Loss Ratio + Expense Ratio
' ============================================================
Public Function CombinedRatio(ByVal incurredLosses As Double, _
                               ByVal expenses As Double, _
                               ByVal earnedPremiums As Double) As Double
    If earnedPremiums > 0 Then
        CombinedRatio = (incurredLosses + expenses) / earnedPremiums
    Else
        CombinedRatio = 0
    End If
End Function

' ============================================================
' Sub: CalculateRiskMetrics
' Populates the RiskAnalysis sheet.
' ============================================================
Public Sub CalculateRiskMetrics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RiskAnalysis")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim r As Long
    For r = 3 To lastRow
        Dim meanL As Double: meanL = CDbl(ws.Cells(r, 2).Value)
        Dim stdL As Double: stdL = CDbl(ws.Cells(r, 3).Value)
        
        ' VaR at 95% and 99.5%
        ws.Cells(r, 4).Value = ValueAtRisk(meanL, stdL, 0.95)
        ws.Cells(r, 5).Value = ValueAtRisk(meanL, stdL, 0.995)
        
        ' Expected Shortfall at 95% and 99.5%  
        ws.Cells(r, 6).Value = ExpectedShortfall(meanL, stdL, 0.95)
        ws.Cells(r, 7).Value = ExpectedShortfall(meanL, stdL, 0.995)
    Next r
    
    MsgBox "Risk metrics calculated!", vbInformation
End Sub