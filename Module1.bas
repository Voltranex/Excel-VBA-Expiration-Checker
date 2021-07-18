Attribute VB_Name = "Module1"
Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=now()"
    x = 0
    y = 0
Dim cell As Range
For Each cell In Range("Tablo4[Expiration Date]")
If cell.Value < Range("B1").Value Then
cell.Interior.Color = 8487423
End If
If cell.Value > Range("B1").Value Then
cell.Interior.Color = 4630638
x = x + 1
End If
If cell.Value = "" Then
cell.Interior.Color = VBA.ColorConstants.vbWhite
y = y + 1
End If
Next cell
Range("B2") = x
Range("B3") = 96 - y - x
End Sub

