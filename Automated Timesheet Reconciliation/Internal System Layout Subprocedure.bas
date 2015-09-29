Attribute VB_Name = "Module2"
' Chronos_Layout_Setup Macro;
'
' Sets up the layout for Chronos to facilitate viewing for GR.
'
' Keyboard Shortcut: Ctrl+Shift+N

Sub Chronos_Layout_Setup(ByVal monthColumn As Long)
    Columns("A:A").ColumnWidth = 7.71
    Columns("B:B").ColumnWidth = 12.29
    Columns("C:C").ColumnWidth = 15
    Columns("D:D").ColumnWidth = 6.43
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
    
    'exclude DPR's and Holidays from filter
    If AutoFilterMode = False Then
        With Selection
            .AutoFilter field:=1, Criteria1:="IT*"
            .AutoFilter field:=5, Criteria1:="Capgemini"
            .AutoFilter field:=monthColumn, Criteria1:="<>"
        End With
    End If
    
    Range("A2").Select
    Columns("J:X").Group
    Range("Y1:AA1").Interior.Color = 65535
    Range("E2").Interior.Color = 65535
    Range("C2").Interior.Color = 65535
    Range("A2").Interior.Color = 65535
    
    'sort cell values
    With ActiveSheet.AutoFilter.Sort.SortFields
        .Clear
        .Add Key:=Range("E3:E2104"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Add Key:=Range("C3:C2104"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End With
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

