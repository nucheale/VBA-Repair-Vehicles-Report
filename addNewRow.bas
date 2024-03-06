Function twoDimArrayToOneDim(oldArr) 'функция для преобразования двумерного массива в одномерный
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For i = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(i) = oldArr(i, 1)
    Next i
    twoDimArrayToOneDim = newArr
End Function

Private Sub addNewRow()
    On Error GoTo errorExit 'уход в конец на ошибке чтобы у клиента не было никаких всплывающих ошибок
    Set sh = ThisWorkbook.Sheets("Учет")
    Set dictSh = ThisWorkbook.Sheets("Справочник")
    dropdownCars = dictSh.ListObjects("Авто").ListColumns("Именование").DataBodyRange 'массив со списком ТС
    dropdownCars = twoDimArrayToOneDim(dropdownCars)
    dropdownWorkers = dictSh.ListObjects("Сотрудники").ListColumns("Сотрудники").DataBodyRange  'массив со списком сотрудников
    dropdownWorkers = twoDimArrayToOneDim(dropdownWorkers)
    With sh
        If .Cells(2, 3) = "" Or .Cells(2, 4) = "" Then 'проверка заполнения строки
            MsgBox "Не заполнен предыдущий ввод от " & .Cells(2, 10), vbCritical, "Ошибка"
            Exit Sub
        End If
        
        .ListObjects("УчетРемонта").ListRows.Add (1) 'добавление новой строки, заполнение стоковыми данными и сброс форматирования от верхней строки
        .Rows(2).Interior.ColorIndex = xlNone
        .Rows(2).Font.Bold = False
        .Cells(2, 1) = Date
        .Cells(2, 8) = "В работе"
        .Cells(2, 10) = Now
        
    .Cells.FormatConditions.Delete 'обновление правил условного форматирования
    Set fc1 = Columns("A:J").FormatConditions.Add(Type:=xlExpression, Formula1:="=$H1=""В ремонте""")
    fc1.SetFirstPriority
    With fc1.Interior
        .PatternColorIndex = 0
        .Color = RGB(255, 91, 91)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    fc1.StopIfTrue = False

    Set fc2 = Columns("B:B").FormatConditions.Add(Type:=xlExpression, Formula1:="=И(ДЕНЬНЕД($B1;2)>5;$B1<>"""")")
    fc2.SetFirstPriority
    With fc2.Font
        .Color = RGB(192, 0, 0)
    End With
    fc2.StopIfTrue = False

    Set fc3 = Columns("A:A").FormatConditions.Add(Type:=xlExpression, Formula1:="=И(ДЕНЬНЕД($A1;2)>5;$A1<>"""")")
    fc3.SetFirstPriority
    With fc3.Font
        .Color = RGB(192, 0, 0)
    End With
    fc3.StopIfTrue = False

    Set fc4 = Columns("F:F").FormatConditions.Add(Type:=xlExpression, Formula1:="=$E1=""Да""")
    fc4.SetFirstPriority
    With fc4.Interior
        .PatternColorIndex = 0
        .Color = 8577748
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    fc4.StopIfTrue = False

    End With
errorExit:
    On Error GoTo 0
End Sub


