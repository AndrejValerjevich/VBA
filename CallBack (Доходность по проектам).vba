Public v_fc As Integer, v_lc As Integer, v_fr As Integer, v_lr As Integer, selected_dp As String, selected_sheet As String
Const sheet1 = "Лист1", sheet_psd = "Лист2", sheet_pmd = "Лист3", sheet_pd = "Лист4"

Sub CallBack(ParamArray varname())

Dim r_dp As Range

With Worksheets(sheet_psd)  'считывание значения ячейки с названием выбранного инфопровайдера
    .Activate
    selected_dp = Range("I1").Value

    If selected_dp = "" Then
        selected_dp = ""
    End If
    
End With

If selected_dp = "DP_9" Then  'определение выбранного пользователем типа диаграммы
    selected_sheet = sheet_psd
    Else
        If selected_dp = "DP_10" Then
        selected_sheet = sheet_pmd
        Else
            If selected_dp = "DP_11" Then
            selected_sheet = sheet_pd
        End If
    End If
End If
    
With Worksheets(selected_sheet)
    If varname(0) = selected_dp Then  'определение требуемого инфопровайдера
        Set r_dp = varname(1)  'получаение данных заданного инфопровайдера
        Call RangeAdress(r_dp)
    End If
End With

Worksheets(sheet1).Activate

End Sub

Sub CreateGraph()  'функция построения графика (требуемого пользователем типа)

Dim oChart As ChartObject
Dim range1 As Range


With Worksheets(sheet1).ChartObjects  'очистка листа, если график был уже построен
If .Count > 0 Then .Delete
End With

With ThisWorkbook.Sheets(selected_sheet)  'определение диапазона данных выбранного инфопровайдера
Set range1 = .Range(.Cells(1, 1), .Cells(v_lr - 1, v_lc))
End With

a = range1.Address

Set oChart = ThisWorkbook.Sheets(sheet1).ChartObjects.Add(440, 10, 1100, 270)  'создание объекта диаграммы

oChart.Chart.SetSourceData (Sheets(selected_sheet).Range(a))  'заполнение диаграммы данными из выбранного диапазона

End Sub

Sub RangeAdress(r_dp As Range)  'функция определения диапазона заданного инфопровайдера
 v_fr = r_dp.Row
 v_lr = r_dp.Rows.Count + r_dp.Row - 1
 v_fc = r_dp.Column
 v_lc = r_dp.Columns.Count + r_dp.Column - 1
End Sub
