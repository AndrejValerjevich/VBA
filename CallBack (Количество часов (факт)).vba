Public v_fc As Integer, v_lc As Integer, v_fr As Integer, v_lr As Integer
Const input_sheet = "Лист3"

Sub CallBack(ParamArray varname())
Dim r_dp As Range

With Worksheets(input_sheet)

If varname(0) = "DP_7" Then  'определение требуемого инфопровайдера
   Set r_dp = varname(1)  'получаение данных заданного инфопровайдера
   Call RangeAdress(r_dp)
End If

End With

End Sub

Sub RangeAdress(r_dp As Range)  'функция определения диапазона заданного инфопровайдера
 v_fr = r_dp.Row
 v_lr = r_dp.Rows.Count + r_dp.Row - 1
 v_fc = r_dp.Column
 v_lc = r_dp.Columns.Count + r_dp.Column - 1
End Sub

Sub FillCells()  'функция переноса выбранных данных в ячейки последней строки инфопровайдера

ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 1) = ThisWorkbook.Sheets(input_sheet).Cells(1, 3)
ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 2) = ThisWorkbook.Sheets(input_sheet).Cells(2, 3)
ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 5) = ThisWorkbook.Sheets(input_sheet).Cells(3, 3)
ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 3) = ThisWorkbook.Sheets(input_sheet).Cells(4, 2)
ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 4) = ThisWorkbook.Sheets(input_sheet).Cells(5, 2)
ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 6) = ThisWorkbook.Sheets(input_sheet).Cells(6, 2)

End Sub
