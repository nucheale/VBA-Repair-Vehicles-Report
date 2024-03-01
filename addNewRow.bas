Function twoDimArrayToOneDim(oldArr)
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For i = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(i) = oldArr(i, 1)
    Next i
    twoDimArrayToOneDim = newArr
End Function

Private Sub addNewRow()
    Set sh = ThisWorkbook.Sheets("Учет")
    Set dictSh = ThisWorkbook.Sheets("Справочник")
    dropdownCars = dictSh.ListObjects("Авто").ListColumns("Именование").DataBodyRange
    dropdownCars = twoDimArrayToOneDim(dropdownCars)
    dropdownWorkers = dictSh.ListObjects("Сотрудники").ListColumns("Сотрудники").DataBodyRange
    dropdownWorkers = twoDimArrayToOneDim(dropdownWorkers)
    With sh
        If .Cells(2, 3) = "" Or .Cells(2, 4) = "" Then
            MsgBox "Не заполнен предыдущий ввод от " & .Cells(2, 10), vbCritical, "Ошибка"
            Exit Sub
        End If
        
        .ListObjects("УчетРемонта").ListRows.Add (1)
        .Rows(2).Interior.ColorIndex = xlNone
        .Rows(2).Font.Bold = False
        .Cells(2, 1) = Date
        .Cells(2, 8) = "В работе"
        .Cells(2, 10) = Now
        ' With .Cells(2, 3).Validation
        '     .Delete
        '     .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(dropdownCars, ",")
        ' End With
        ' With .Cells(2, 5).Validation
        '     .Delete
        '     .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Да,Нет"
        ' End With
        ' With .Cells(2, 8).Validation
        '     .Delete
        '     .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="В работе,В ремонте"
        ' End With
        ' With .Cells(2, 9).Validation
        '     .Delete
        '     .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(dropdownWorkers, ",")
        ' End With
    End With
End Sub