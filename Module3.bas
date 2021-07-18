Attribute VB_Name = "Module3"
Sub green()
Attribute green.VB_ProcData.VB_Invoke_Func = " \n14"
'
' green Makro
'

'
    ActiveSheet.ListObjects("Tablo4").Range.AutoFilter Field:=3, Criteria1:=RGB _
        (110, 168, 70), Operator:=xlFilterCellColor
End Sub
Sub red()
Attribute red.VB_ProcData.VB_Invoke_Func = " \n14"
'
' red Makro
'

'
    ActiveSheet.ListObjects("Tablo4").Range.AutoFilter Field:=3, Criteria1:=RGB _
        (255, 129, 129), Operator:=xlFilterCellColor
End Sub
Sub all()
Attribute all.VB_ProcData.VB_Invoke_Func = " \n14"
'
' all Makro
'

'
    ActiveSheet.ListObjects("Tablo4").Range.AutoFilter Field:=3
End Sub
