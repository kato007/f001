Attribute VB_Name = "Module6"
Sub Macro8()
Attribute Macro8.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro8 Macro
'

'
    Range("K24:N25").Select
    ActiveWorkbook.Worksheets("0348M970表紙").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("0348M970表紙").Sort.SortFields.Add Key:=Range("K24") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("0348M970表紙").Sort
        .SetRange Range("K24:N25")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
