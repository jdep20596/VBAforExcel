Attribute VB_Name = "Module11"
'Find Cell Content Macro; (For questions/suggestions, please contact CBRASIL);
'
'Find the contents of one cell on another specific document:
'Make sure the cells where you are getting the values from are allingned vertically (in the same column);
'Make sure that the cells to the right of that column are blank;
'Make sure that the sheet where the search will be performed does not have any active filters;
'Then, select the cell containing the first value you want to find and run the Macro.
'If there are no errors, the macro should write the addresses of the cells containing the values you searched for.
'They will be written on the cells next to the source cells.
'Shortcut: Ctrl + Shift + F
Option Explicit

Sub FindCellContent()
Attribute FindCellContent.VB_ProcData.VB_Invoke_Func = "F\n14"
    Application.ScreenUpdating = False
    Dim poToFind As String, tracker As String, whereToFind As String, sourceSheet As String
    Dim constantCells As Range, cell As Range
    
    'prompts the user for the name of the target sheet.
    whereToFind = Application.InputBox(prompt:="Please type in the name of the sheet where you want to perform the search:")
    
    'selects the whole column of values to search.
    Range(Selection, Selection.End(xlDown)).Select
    Set constantCells = Selection.SpecialCells(xlConstants)
    
    'performs the search
    For Each cell In constantCells
        poToFind = cell
        On Error Resume Next
        tracker = Worksheets(whereToFind).Cells.Find(cell).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        cell.Offset(0, 1) = tracker
        tracker = ""
    Next cell
    
    'returns to source sheet before ending the sub.
    cell.Activate
    Application.ScreenUpdating = True
End Sub


