VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    If Not (Intersect(Target, Range("A4:C1048576")) Is Nothing) Then
        Dim timesheetPivotEndRow As Integer, poTemplateEndRow As Integer, autoFillSourceRow As Integer
        Dim autoFillTargetRow As Integer
        Dim originalRange As Range
        timesheetPivotEndRow = Range("C4").End(xlDown).Row - 1
        poTemplateEndRow = Range("V4").End(xlDown).Row
        If timesheetPivotEndRow <> poTemplateEndRow Then
            Application.ScreenUpdating = False
            
            Set originalRange = ActiveCell
            If timesheetPivotEndRow > poTemplateEndRow Then
                 Range(Cells(poTemplateEndRow, "T"), Cells(poTemplateEndRow, "AE")).AutoFill _
                    Destination:=Range(Cells(poTemplateEndRow, "T"), Cells(timesheetPivotEndRow, "AE")), _
                    Type:=xlFillDefault
            ElseIf timesheetPivotEndRow < poTemplateEndRow And poTemplateEndRow > 5 Then
                Range(Cells(timesheetPivotEndRow + 1, "T"), Cells(poTemplateEndRow, "AE")).ClearContents
            End If
            originalRange.Activate
            
            Application.ScreenUpdating = True
        End If
    End If
End Sub

