Attribute VB_Name = "ThisWorkbook"
' ============================================================
' ThisWorkbook Module
' Handles workbook-level events and initialisation.
' ============================================================
Option Explicit

Private Sub Workbook_Open()
    ' Display welcome message
    MsgBox "Welcome to the Actuarial Insurance Model." & vbCrLf & _
           "Use the macros in the Developer tab to run analyses:" & vbCrLf & vbCrLf & _
           "  - BuildMortalityTable: Generate full life table" & vbCrLf & _
           "  - CalculateAllPremiums: Compute premiums for all policies" & vbCrLf & _
           "  - RunChainLadder: Run loss reserving analysis" & vbCrLf & _
           "  - CalculateRiskMetrics: Compute VaR and risk measures", _
           vbInformation, "Actuarial Insurance Model v2.0"
End Sub