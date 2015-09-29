Attribute VB_Name = "Module1"
Sub LoadChronos()
    Dim fileToOpen As Variant
    Dim workbookName As String, month As String
    Dim originalSheet As Worksheet
    Dim a As Range, b As Range, c As Range
    Dim i As Long
    
    Set originalSheet = ActiveSheet
    
    Worksheets("PO Template").Activate
    month = Application.WorksheetFunction.Text(Worksheets("PO Template") _
        .Range("V2").Value, "mmm")
    
    fileToOpen = Application.GetOpenFilename("Excel Workbooks , *.xls*", False)
    If fileToOpen <> False Then
        Application.ScreenUpdating = False
        'clear contents on reconciliation spreadsheet
        originalSheet.Activate
        originalSheet.Range("A5:L10000").ClearContents
    
        'open Chronos workbook and register filename
        Workbooks.Open (fileToOpen)
        workbookName = GetFilenameFromPath(fileToOpen)
        Workbooks(workbookName).Worksheets(1).Activate
        
        'find the month to be copied over and adjust Chronos extract layout
        Set a = Range("A1:BT1").Find(month, LookIn:=xlValues)
        a.Activate
        ActiveCell.Offset(1, 0).Activate
        ActiveCell.Offset(0, 2).Activate
        Set a = ActiveCell
        i = ActiveCell.Column
        Chronos_Layout_Setup (i)
        
        'copy details over
        Set b = Range("A1:Z100").Find("Project Code", LookIn:=xlValues)
        Set c = Range("A1:Z100").Find("Charge Rate", LookIn:=xlValues)
        Range(b, c).Select
        Range(Selection, Selection.End(xlDown)).Copy
        originalSheet.Activate
        Range("A3").Activate
        ActiveCell.PasteSpecial (xlPasteValues)
        Workbooks(workbookName).Worksheets(1).Activate
        a.Activate
        Set b = a.End(xlDown)
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Cells(b.Row, a.Column - 2)).Copy
        originalSheet.Activate
        Range("J3").Activate
        ActiveCell.PasteSpecial (xlPasteValues)
        
        Application.ScreenUpdating = True
    Else
        MsgBox ("New Extract Cancelled")
    End If
End Sub

Function GetFilenameFromPath(ByVal strPath As String) As String
' Function taken from http://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If

End Function
