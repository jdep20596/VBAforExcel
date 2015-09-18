Attribute VB_Name = "Module2"
' Chronos_Layout_Setup Macro; (For questions/suggestions, please contact CBRASIL)
'
' Sets up the layout for Chronos to facilitate viewing for GR.
'
' Keyboard Shortcut: Ctrl+Shift+N

Sub Chronos_Layout_Setup()
Attribute Chronos_Layout_Setup.VB_Description = "Sets up the layout for Chronos"
Attribute Chronos_Layout_Setup.VB_ProcData.VB_Invoke_Func = "N\n14"
    Columns("A:A").ColumnWidth = 7.71
    Columns("B:B").ColumnWidth = 12.29
    Columns("C:C").ColumnWidth = 15
    Columns("D:D").ColumnWidth = 13.14
    Columns("E:E").ColumnWidth = 12
    Columns("F:F").ColumnWidth = 4.14
    Columns("G:G").ColumnWidth = 4.14
    Columns("H:H").ColumnWidth = 3.71
    Columns("I:I").ColumnWidth = 6.57
    Rows("2:2").EntireRow.AutoFit
    Columns("J:AA").ColumnWidth = 7.14
    Range("J3").Activate
    ActiveWindow.FreezePanes = True
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    'Next step: Find out how to exclude DPR's and Holidays from filter
    If AutoFilterMode = False Then
        Selection.AutoFilter _
        field:=1, _
        Criteria1:="IT*"
    End If
    Range("A2").Select
    Columns("J:X").Group
    Range("Y1:AA1").Interior.Color = 65535
    Range("E2").Interior.Color = 65535
    Range("C2").Interior.Color = 65535
    Range("A2").Interior.Color = 65535
    
End Sub
