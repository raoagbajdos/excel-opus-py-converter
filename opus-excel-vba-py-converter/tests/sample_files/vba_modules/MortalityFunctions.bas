Attribute VB_Name = "MortalityFunctions"
' ============================================================
' Module: MortalityFunctions
' Purpose: Life table and mortality calculations for actuarial
'          pricing of life insurance products.
' ============================================================
Option Explicit

' --- Constants ---
Private Const MAX_AGE As Integer = 120
Private Const BASE_MORTALITY_RATE As Double = 0.001

' ============================================================
' Function: GetMortalityRate
' Returns the annual probability of death (qx) for a given
' age and gender using a simplified Makeham-Gompertz model.
'
' Parameters:
'   age    - Integer, the attained age
'   gender - String, "M" or "F"
' Returns:
'   Double - the qx value
' ============================================================
Public Function GetMortalityRate(ByVal age As Integer, ByVal gender As String) As Double
    Dim A As Double, B As Double, c As Double
    
    ' Makeham-Gompertz parameters
    If UCase(gender) = "M" Then
        A = 0.0005
        B = 0.00004
        c = 1.1
    Else
        A = 0.0003
        B = 0.000025
        c = 1.095
    End If
    
    ' qx = A + B * c^x
    Dim qx As Double
    qx = A + B * (c ^ age)
    
    ' Cap at 1.0
    If qx > 1# Then qx = 1#
    
    GetMortalityRate = qx
End Function

' ============================================================
' Function: SurvivalProbability
' Returns tPx - the probability of surviving t years from age x.
' ============================================================
Public Function SurvivalProbability(ByVal age As Integer, ByVal t As Integer, _
                                     ByVal gender As String) As Double
    Dim px As Double
    px = 1#
    Dim i As Integer
    For i = 0 To t - 1
        px = px * (1# - GetMortalityRate(age + i, gender))
        If px < 0.000001 Then
            px = 0#
            Exit For
        End If
    Next i
    SurvivalProbability = px
End Function

' ============================================================
' Function: LifeExpectancy
' Calculates curtate life expectancy ex for a given age.
' ============================================================
Public Function LifeExpectancy(ByVal age As Integer, ByVal gender As String) As Double
    Dim ex As Double
    ex = 0#
    Dim t As Integer
    For t = 1 To MAX_AGE - age
        ex = ex + SurvivalProbability(age, t, gender)
    Next t
    LifeExpectancy = ex
End Function

' ============================================================
' Sub: BuildMortalityTable
' Populates a worksheet with a complete mortality table.
' ============================================================
Public Sub BuildMortalityTable()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MortalityTable")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "MortalityTable"
    End If
    
    ' Headers
    ws.Cells(1, 1).Value = "Age"
    ws.Cells(1, 2).Value = "qx (Male)"
    ws.Cells(1, 3).Value = "qx (Female)"
    ws.Cells(1, 4).Value = "lx (Male)"
    ws.Cells(1, 5).Value = "lx (Female)"
    ws.Cells(1, 6).Value = "ex (Male)"
    ws.Cells(1, 7).Value = "ex (Female)"
    
    Dim lxM As Double, lxF As Double
    lxM = 100000
    lxF = 100000
    
    Dim row As Integer
    For row = 0 To MAX_AGE
        ws.Cells(row + 2, 1).Value = row
        
        Dim qxM As Double, qxF As Double
        qxM = GetMortalityRate(row, "M")
        qxF = GetMortalityRate(row, "F")
        
        ws.Cells(row + 2, 2).Value = qxM
        ws.Cells(row + 2, 3).Value = qxF
        ws.Cells(row + 2, 4).Value = lxM
        ws.Cells(row + 2, 5).Value = lxF
        ws.Cells(row + 2, 6).Value = LifeExpectancy(row, "M")
        ws.Cells(row + 2, 7).Value = LifeExpectancy(row, "F")
        
        lxM = lxM * (1 - qxM)
        lxF = lxF * (1 - qxF)
    Next row
    
    MsgBox "Mortality table generated with " & MAX_AGE + 1 & " rows.", vbInformation
End Sub