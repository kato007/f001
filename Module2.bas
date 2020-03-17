Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Add Key:=Range("E2:E182"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("データ").Sort.SortFields.Add Key:=Range("U2:U182"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("データ").Sort
        .SetRange Range("A1:BC182")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
