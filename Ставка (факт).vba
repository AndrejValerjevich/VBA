Public v_fc As Integer, v_lc As Integer, v_fr As Integer, v_lr As Integer
Const input_sheet = "Лист3"

Sub CallBack(ParamArray varname())
Dim r_dp As Range

With Worksheets(input_sheet)

If varname(0) = "DP_5" Then  'определение требуемого инфопровайдера
   Set r_dp = varname(1)  'получаение данных заданного инфопровайдера
   Call RangeAdress(r_dp)
End If

End With

End Sub

Sub RangeAdress(r_dp As Range) 'функция определения диапазона заданного инфопровайдера
 v_fr = r_dp.Row
 v_lr = r_dp.Rows.Count + r_dp.Row - 1
 v_fc = r_dp.Column
 v_lc = r_dp.Columns.Count + r_dp.Column - 1
End Sub


Sub FillCells() 'функция переноса выбранных данных в ячейки последней строки инфопровайдера
Dim dt As Date

With ThisWorkbook.Sheets(input_sheet)
    dt = ThisWorkbook.Sheets(input_sheet).Cells(3, 3)
    days_amount = ThisWorkbook.Sheets(input_sheet).Cells(6, 3)
        For counter = 0 To days_amount
            ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 1) = ThisWorkbook.Sheets(input_sheet).Cells(1, 3)  'проект
            ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 2) = ThisWorkbook.Sheets(input_sheet).Cells(2, 3)  'квалификация
            ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 3).NumberFormat = "dd.mm.yyyy"
            ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 3) = dt                                            'дата
            ThisWorkbook.Sheets(input_sheet).Cells(v_lr, 4) = ThisWorkbook.Sheets(input_sheet).Cells(5, 3)  'ставка
            dt = dt + 1
            Call ThisWorkbook.Sheets(input_sheet).BUTTON_11_Click  'вызов кнопки "Сохранение"
            Call ThisWorkbook.Sheets(input_sheet).BUTTON_10_Click  'вызов кнопки "Акутализация"
        Next counter
End With
