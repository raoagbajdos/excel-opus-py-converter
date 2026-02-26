Attribute VB_Name = "LossReserving"
' ============================================================
' Module: LossReserving
' Purpose: Property & Casualty loss reserving using the
'          Chain-Ladder (Development Factor) method.
' ============================================================
Option Explicit

' ============================================================
' Function: ChainLadderFactor
' Calculates the age-to-age development factor between two
' consecutive development periods.
' ============================================================
Public Function ChainLadderFactor(ByVal triangle As Range, _
                                   ByVal devPeriod As Integer) As Double
    Dim sumCurrent As Double, sumPrior As Double
    sumCurrent = 0#
    sumPrior = 0#
    
    Dim r As Integer
    For r = 1 To triangle.Rows.Count - devPeriod
        Dim priorVal As Variant, currentVal As Variant
        priorVal = triangle.Cells(r, devPeriod).Value
        currentVal = triangle.Cells(r, devPeriod + 1).Value
        
        If IsNumeric(priorVal) And IsNumeric(currentVal) Then
            If CDbl(priorVal) > 0 Then
                sumPrior = sumPrior + CDbl(priorVal)
                sumCurrent = sumCurrent + CDbl(currentVal)
            End If
        End If
    Next r
    
    If sumPrior > 0 Then
        ChainLadderFactor = sumCurrent / sumPrior
    Else
        ChainLadderFactor = 1#
    End If
End Function

' ============================================================
' Function: CumulativeFactor
' Product of all remaining development factors from a given
' period to ultimate.
' ============================================================
Public Function CumulativeFactor(ByVal triangle As Range, _
                                  ByVal fromPeriod As Integer) As Double
    Dim prod As Double
    prod = 1#
    Dim maxDev As Integer
    maxDev = triangle.Columns.Count - 1
    
    Dim d As Integer
    For d = fromPeriod To maxDev
        prod = prod * ChainLadderFactor(triangle, d)
    Next d
    
    CumulativeFactor = prod
End Function

' ============================================================
' Function: UltimateLoss
' Projects the ultimate loss for a given accident year.
' ============================================================
Public Function UltimateLoss(ByVal triangle As Range, _
                              ByVal accidentYear As Integer) As Double
    ' Find the latest diagonal value for this accident year
    Dim latestDev As Integer
    latestDev = triangle.Columns.Count - accidentYear + 1
    If latestDev > triangle.Columns.Count Then latestDev = triangle.Columns.Count
    
    Dim latestValue As Double
    latestValue = CDbl(triangle.Cells(accidentYear, latestDev).Value)
    
    ' Multiply by cumulative development factor
    If latestDev < triangle.Columns.Count Then
        UltimateLoss = latestValue * CumulativeFactor(triangle, latestDev)
    Else
        UltimateLoss = latestValue
    End If
End Function

' ============================================================
' Function: IBNRReserve
' Calculates IBNR (Incurred But Not Reported) reserve.
'   IBNR = Ultimate - Paid-to-date
' ============================================================
Public Function IBNRReserve(ByVal triangle As Range, _
                             ByVal accidentYear As Integer) As Double
    Dim latestDev As Integer
    latestDev = triangle.Columns.Count - accidentYear + 1
    If latestDev > triangle.Columns.Count Then latestDev = triangle.Columns.Count
    
    Dim paidToDate As Double
    paidToDate = CDbl(triangle.Cells(accidentYear, latestDev).Value)
    
    IBNRReserve = UltimateLoss(triangle, accidentYear) - paidToDate
End Function

' ============================================================
' Sub: RunChainLadder
' Runs the chain-ladder analysis on the LossTriangle sheet.
' ============================================================
Public Sub RunChainLadder()
    Dim wsTriangle As Worksheet
    Set wsTriangle = ThisWorkbook.Worksheets("LossTriangle")
    
    ' Find the triangle range (assumes it starts at B3)
    Dim lastRow As Long, lastCol As Long
    lastRow = wsTriangle.Cells(wsTriangle.Rows.Count, 2).End(xlUp).row
    lastCol = wsTriangle.Cells(3, wsTriangle.Columns.Count).End(xlToLeft).Column
    
    Dim triangleRange As Range
    Set triangleRange = wsTriangle.Range(wsTriangle.Cells(3, 2), wsTriangle.Cells(lastRow, lastCol))
    
    ' Output development factors
    Dim outputRow As Long
    outputRow = lastRow + 3
    wsTriangle.Cells(outputRow, 1).Value = "Development Factors"
    wsTriangle.Cells(outputRow, 1).Font.Bold = True
    
    Dim d As Integer
    For d = 1 To lastCol - 2
        wsTriangle.Cells(outputRow + 1, d + 1).Value = d & " to " & d + 1
        wsTriangle.Cells(outputRow + 2, d + 1).Value = ChainLadderFactor(triangleRange, d)
        wsTriangle.Cells(outputRow + 2, d + 1).NumberFormat = "0.0000"
    Next d
    
    ' Output IBNR reserves
    outputRow = outputRow + 5
    wsTriangle.Cells(outputRow, 1).Value = "IBNR Reserves"
    wsTriangle.Cells(outputRow, 1).Font.Bold = True
    wsTriangle.Cells(outputRow + 1, 1).Value = "Acc. Year"
    wsTriangle.Cells(outputRow + 1, 2).Value = "Paid to Date"
    wsTriangle.Cells(outputRow + 1, 3).Value = "Ultimate"
    wsTriangle.Cells(outputRow + 1, 4).Value = "IBNR"
    
    Dim numYears As Integer
    numYears = lastRow - 2
    
    Dim ay As Integer
    For ay = 1 To numYears
        Dim latestDev As Integer
        latestDev = lastCol - 1 - ay + 1
        If latestDev > lastCol - 1 Then latestDev = lastCol - 1
        
        wsTriangle.Cells(outputRow + 1 + ay, 1).Value = wsTriangle.Cells(2 + ay, 1).Value
        wsTriangle.Cells(outputRow + 1 + ay, 2).Value = triangleRange.Cells(ay, latestDev).Value
        wsTriangle.Cells(outputRow + 1 + ay, 3).Value = UltimateLoss(triangleRange, ay)
        wsTriangle.Cells(outputRow + 1 + ay, 4).Value = IBNRReserve(triangleRange, ay)
        
        wsTriangle.Cells(outputRow + 1 + ay, 2).NumberFormat = "#,##0"
        wsTriangle.Cells(outputRow + 1 + ay, 3).NumberFormat = "#,##0"
        wsTriangle.Cells(outputRow + 1 + ay, 4).NumberFormat = "#,##0"
    Next ay
    
    MsgBox "Chain-Ladder analysis complete! " & numYears & " accident years processed.", vbInformation
End Sub