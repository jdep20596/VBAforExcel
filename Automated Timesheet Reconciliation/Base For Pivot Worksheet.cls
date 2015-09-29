VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim daysColumnEnd As Range, formulaColumnsEnd As Range

Private Sub Worksheet_Change(ByVal Target1 As Range)
    If Not (Intersect(Target1, Range("A4:E1048576")) Is Nothing) Then
        Dim autoFillSources As Range
        Dim autoFillTargets As Range, originalRange As Range
        Application.ScreenUpdating = False
        
        Set originalRange = ActiveCell
        
        Range("E3").End(xlDown).Activate
        If ActiveCell.Row > 10000 Then
            MsgBox ("Please make sure that Column E is not empty")
        Else
            Set daysColumnEnd = ActiveCell
            
            Set formulaColumnsEnd = Range("F3").End(xlDown)
            
            If formulaColumnsEnd.Row < daysColumnEnd.Row Then
                Set autoFillSources = Range(Cells(formulaColumnsEnd.Row, "F"), Cells(formulaColumnsEnd.Row, "K"))
                Set autoFillTargets = Range(Cells(formulaColumnsEnd.Row, "F"), Cells(daysColumnEnd.Row, "K"))
                autoFillSources.AutoFill Destination:=Range(autoFillSources, autoFillTargets), _
                    Type:=xlFillDefault
                    
                Range(Cells(formulaColumnsEnd.Row + 1, "H"), Cells(daysColumnEnd.Row, "H")).ClearContents
            
            ElseIf formulaColumnsEnd.Row > daysColumnEnd.Row Then
            
                If daysColumnEnd.Row < 10000 Then
                    Range(Cells(daysColumnEnd.Row + 1, "F"), Cells(formulaColumnsEnd.Row, "K")).ClearContents
                End If
            End If
            UpdateCapPivot
            originalRange.Activate
            Application.ScreenUpdating = True
        End If
    End If
End Sub

Private Sub UpdateCapPivot()
    Worksheets("Reconciliation").PivotTables("CAPTimesheetPivotTable").ChangePivotCache _
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets _
        ("Base for Pivot").Range("A3", Cells(daysColumnEnd.Row, "G")))
    Worksheets("Reconciliation").PivotTables("CAPTimesheetPivotTable").PivotCache.Refresh
    
    Worksheets("PO Template").PivotTables("PivotTable1").ChangePivotCache _
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Worksheets _
        ("Base for Pivot").Range("A3", Cells(daysColumnEnd.Row, "G")))
    Worksheets("PO Template").PivotTables("PivotTable1").PivotCache.Refresh
End Sub
