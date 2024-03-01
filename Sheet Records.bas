Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Rows.Count > 1 Then Exit Sub
    If Target.Columns.Count > 1 Then Exit Sub
    
    Set sh = ThisWorkbook.Sheets("Учет")
    Set dateEndColumn = sh.ListObjects("УчетРемонта").ListColumns("Дата окончания")
    Set operationsColumn = sh.ListObjects("УчетРемонта").ListColumns("Работы")
    Set oilChangeColumn = sh.ListObjects("УчетРемонта").ListColumns("Замена масла")
    Set kmColumn = sh.ListObjects("УчетРемонта").ListColumns("Пробег")
    Set oilChangeKmColumn = sh.ListObjects("УчетРемонта").ListColumns("Следующая замена масла")
    Set stasusColumn = sh.ListObjects("УчетРемонта").ListColumns("В работе")
    
    If Not Intersect(Target, dateEndColumn.DataBodyRange) Is Nothing Then
        If Not Cells(Target.Row, Target.Column).Value = "" Then stasusColumn.DataBodyRange(Target.Row - 1, 1).Value = "В работе" Else stasusColumn.DataBodyRange(Target.Row - 1, 1).Value = "В ремонте"
    End If
    
    If Not Intersect(Target, operationsColumn.DataBodyRange) Is Nothing Then
        If Not Cells(Target.Row, Target.Column).Value = "" And dateEndColumn.DataBodyRange(Target.Row - 1, 1) = "" Then stasusColumn.DataBodyRange(Target.Row - 1, 1).Value = "В ремонте"
    End If
    
    If Not Intersect(Target, kmColumn.DataBodyRange) Is Nothing Then
        If IsNumeric(Target.Value) And oilChangeColumn.DataBodyRange(Target.Row - 1, 1) = "Да" Then oilChangeKmColumn.DataBodyRange(Target.Row - 1, 1).Value = Target.Value + 10000
        If Target.Value = "" Then oilChangeKmColumn.DataBodyRange(Target.Row - 1, 1).Value = ""
    End If
End Sub