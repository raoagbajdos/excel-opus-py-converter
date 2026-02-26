Attribute VB_Name = "clsPolicy"
' ============================================================
' Class Module: clsPolicy
' Represents an individual insurance policy with properties
' and methods for valuation.
' ============================================================
Option Explicit

Private pPolicyID As String
Private pInsuredName As String
Private pAge As Integer
Private pGender As String
Private pProductType As String
Private pSumAssured As Double
Private pTerm As Integer
Private pInterestRate As Double
Private pIssueDate As Date

' --- Properties ---
Public Property Get PolicyID() As String
    PolicyID = pPolicyID
End Property
Public Property Let PolicyID(ByVal value As String)
    pPolicyID = value
End Property

Public Property Get InsuredName() As String
    InsuredName = pInsuredName
End Property
Public Property Let InsuredName(ByVal value As String)
    pInsuredName = value
End Property

Public Property Get Age() As Integer
    Age = pAge
End Property
Public Property Let Age(ByVal value As Integer)
    pAge = value
End Property

Public Property Get Gender() As String
    Gender = pGender
End Property
Public Property Let Gender(ByVal value As String)
    pGender = value
End Property

Public Property Get ProductType() As String
    ProductType = pProductType
End Property
Public Property Let ProductType(ByVal value As String)
    pProductType = value
End Property

Public Property Get SumAssured() As Double
    SumAssured = pSumAssured
End Property
Public Property Let SumAssured(ByVal value As Double)
    pSumAssured = value
End Property

Public Property Get Term() As Integer
    Term = pTerm
End Property
Public Property Let Term(ByVal value As Integer)
    pTerm = value
End Property

Public Property Get InterestRate() As Double
    InterestRate = pInterestRate
End Property
Public Property Let InterestRate(ByVal value As Double)
    pInterestRate = value
End Property

Public Property Get IssueDate() As Date
    IssueDate = pIssueDate
End Property
Public Property Let IssueDate(ByVal value As Date)
    pIssueDate = value
End Property

' ============================================================
' Function: CalculateNetPremium
' Returns the net premium for this policy.
' ============================================================
Public Function CalculateNetPremium() As Double
    Select Case UCase(pProductType)
        Case "TERM"
            CalculateNetPremium = TermLifeNetPremium(pAge, pTerm, pSumAssured, pInterestRate, pGender)
        Case "WHOLE LIFE"
            CalculateNetPremium = WholeLifeNetPremium(pAge, pSumAssured, pInterestRate, pGender)
        Case "ENDOWMENT"
            CalculateNetPremium = EndowmentNetPremium(pAge, pTerm, pSumAssured, pInterestRate, pGender)
        Case Else
            CalculateNetPremium = 0
    End Select
End Function

' ============================================================
' Function: CalculateReserve
' Returns the policy reserve at a given duration t.
'   tVx = A(x+t) - P * ax_due(x+t)  (prospective method)
' ============================================================
Public Function CalculateReserve(ByVal duration As Integer) As Double
    If duration >= pTerm Then
        CalculateReserve = 0
        Exit Function
    End If
    
    Dim futureAge As Integer
    futureAge = pAge + duration
    Dim remainingTerm As Integer
    remainingTerm = pTerm - duration
    
    ' Future benefit (NSP of remaining coverage)
    Dim v As Double
    v = 1# / (1# + pInterestRate)
    
    Dim futureBenefit As Double
    futureBenefit = 0#
    Dim t As Integer
    For t = 0 To remainingTerm - 1
        futureBenefit = futureBenefit + (v ^ (t + 1)) * _
                        SurvivalProbability(futureAge, t, pGender) * _
                        GetMortalityRate(futureAge + t, pGender)
    Next t
    futureBenefit = futureBenefit * pSumAssured
    
    ' Future premiums
    Dim futurePremiums As Double
    futurePremiums = CalculateNetPremium() * AnnuityDue(futureAge, remainingTerm, pInterestRate, pGender)
    
    CalculateReserve = futureBenefit - futurePremiums
End Function