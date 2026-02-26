Attribute VB_Name = "PremiumCalculations"
' ============================================================
' Module: PremiumCalculations
' Purpose: Insurance premium computation for term life,
'          whole life, and endowment products.
' ============================================================
Option Explicit

' ============================================================
' Function: AnnuityDue
' Calculates the present value of a life annuity-due.
'   ax_due = SUM(t=0..n-1) v^t * tPx
' ============================================================
Public Function AnnuityDue(ByVal age As Integer, ByVal term As Integer, _
                            ByVal interest As Double, ByVal gender As String) As Double
    Dim v As Double
    v = 1# / (1# + interest)
    
    Dim total As Double
    total = 0#
    Dim t As Integer
    For t = 0 To term - 1
        total = total + (v ^ t) * SurvivalProbability(age, t, gender)
    Next t
    
    AnnuityDue = total
End Function

' ============================================================
' Function: PureEndowment
' Calculates nEx = v^n * nPx
' ============================================================
Public Function PureEndowment(ByVal age As Integer, ByVal n As Integer, _
                               ByVal interest As Double, ByVal gender As String) As Double
    Dim v As Double
    v = 1# / (1# + interest)
    PureEndowment = (v ^ n) * SurvivalProbability(age, n, gender)
End Function

' ============================================================
' Function: TermLifeNetPremium
' Net annual premium for an n-year term life insurance.
'   P = A(1)x:n / ax_due:n
' where A(1)x:n = sum(t=0..n-1) v^(t+1) * tPx * q(x+t)
' ============================================================
Public Function TermLifeNetPremium(ByVal age As Integer, ByVal term As Integer, _
                                    ByVal sumAssured As Double, ByVal interest As Double, _
                                    ByVal gender As String) As Double
    Dim v As Double
    v = 1# / (1# + interest)
    
    ' Net single premium (term insurance)
    Dim nsp As Double
    nsp = 0#
    Dim t As Integer
    For t = 0 To term - 1
        Dim tPx As Double
        tPx = SurvivalProbability(age, t, gender)
        Dim qxt As Double
        qxt = GetMortalityRate(age + t, gender)
        nsp = nsp + (v ^ (t + 1)) * tPx * qxt
    Next t
    nsp = nsp * sumAssured
    
    ' Annuity due
    Dim annuity As Double
    annuity = AnnuityDue(age, term, interest, gender)
    
    If annuity > 0 Then
        TermLifeNetPremium = nsp / annuity
    Else
        TermLifeNetPremium = 0
    End If
End Function

' ============================================================
' Function: WholeLifeNetPremium
' Net annual premium for whole life insurance.
' ============================================================
Public Function WholeLifeNetPremium(ByVal age As Integer, ByVal sumAssured As Double, _
                                     ByVal interest As Double, ByVal gender As String) As Double
    WholeLifeNetPremium = TermLifeNetPremium(age, 120 - age, sumAssured, interest, gender)
End Function

' ============================================================
' Function: EndowmentNetPremium
' Net annual premium for an n-year endowment.
'   P = (A(1)x:n + nEx) * SA / ax_due:n
' ============================================================
Public Function EndowmentNetPremium(ByVal age As Integer, ByVal term As Integer, _
                                     ByVal sumAssured As Double, ByVal interest As Double, _
                                     ByVal gender As String) As Double
    Dim v As Double
    v = 1# / (1# + interest)
    
    ' Term insurance component
    Dim termIns As Double
    termIns = 0#
    Dim t As Integer
    For t = 0 To term - 1
        termIns = termIns + (v ^ (t + 1)) * SurvivalProbability(age, t, gender) * _
                  GetMortalityRate(age + t, gender)
    Next t
    
    ' Pure endowment component
    Dim pe As Double
    pe = PureEndowment(age, term, interest, gender)
    
    ' Total NSP
    Dim nsp As Double
    nsp = (termIns + pe) * sumAssured
    
    ' Annuity due
    Dim annuity As Double
    annuity = AnnuityDue(age, term, interest, gender)
    
    If annuity > 0 Then
        EndowmentNetPremium = nsp / annuity
    Else
        EndowmentNetPremium = 0
    End If
End Function

' ============================================================
' Function: GrossPremium
' Adds expense and profit loading to net premium.
' ============================================================
Public Function GrossPremium(ByVal netPremium As Double, _
                              ByVal expenseRatio As Double, _
                              ByVal profitMargin As Double) As Double
    GrossPremium = netPremium / (1 - expenseRatio - profitMargin)
End Function

' ============================================================
' Sub: CalculateAllPremiums
' Populates the PremiumSummary sheet with premium calculations
' for a range of policy configurations.
' ============================================================
Public Sub CalculateAllPremiums()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("PremiumCalculations")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim r As Long
    For r = 3 To lastRow
        Dim age As Integer: age = CInt(ws.Cells(r, 1).Value)
        Dim gender As String: gender = CStr(ws.Cells(r, 2).Value)
        Dim term As Integer: term = CInt(ws.Cells(r, 3).Value)
        Dim sa As Double: sa = CDbl(ws.Cells(r, 4).Value)
        Dim rate As Double: rate = CDbl(ws.Cells(r, 5).Value)
        Dim prodType As String: prodType = CStr(ws.Cells(r, 6).Value)
        
        Dim netP As Double
        Select Case UCase(prodType)
            Case "TERM"
                netP = TermLifeNetPremium(age, term, sa, rate, gender)
            Case "WHOLE LIFE"
                netP = WholeLifeNetPremium(age, sa, rate, gender)
            Case "ENDOWMENT"
                netP = EndowmentNetPremium(age, term, sa, rate, gender)
            Case Else
                netP = 0
        End Select
        
        ws.Cells(r, 7).Value = netP
        ws.Cells(r, 8).Value = GrossPremium(netP, 0.15, 0.05)
    Next r
    
    MsgBox "Premium calculations complete!", vbInformation
End Sub