Function removeDublicatesFromOneDimArr(arr) 'удаление дубликатов в одномерном массиве
    Dim coll As New Collection
    For Each e In arr
        On Error Resume Next
        coll.Add e, e
        On Error GoTo 0
    Next e
    Dim uniqueArr As Variant
    ReDim uniqueArr(1 To coll.Count)
    For i = 1 To coll.Count
        uniqueArr(i) = coll(i)
    Next i
    removeDublicatesFromOneDimArr = uniqueArr
End Function


Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'ручное изменение даты отчета
    On Error GoTo errorExit 'уход в конец на ошибке чтобы у клиента не было никаких всплывающих ошибок
    If Selection.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Cells(1, 2)) Is Nothing Then
        lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
        lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
        lastRow = WorksheetFunction.Max(lastRow1, lastRow2) 'находим макс последнюю строку из 1 и 2 столбца, т.к. в них может быть разное количество заполенных строк
        If Cells(lastRow, 1) = "В работе" Then lastRow = lastRow + 1 'если нет заполненных ТС
        Range(Cells(4, 1), Cells(lastRow, 2)).ClearContents
        Range(Cells(4, 1), Cells(lastRow, 2)).ClearFormats
        reportDate = Cells(1, 2) 'дата отчета
        reportTable = Sheets("Учет").ListObjects("УчетРемонта").DataBodyRange
        Dim datesBegin, datesEnd, cars, statuses As Variant 'массивы со столбцами дата начала ремонта, дата окончания ремонта, ТС, статус из листа Учет
        ReDim datesBegin(1 To UBound(reportTable, 1))
        ReDim datesEnd(1 To UBound(reportTable, 1))
        ReDim cars(1 To UBound(reportTable, 1))
        ReDim statuses(1 To UBound(reportTable, 1))
        For i = LBound(reportTable, 1) To UBound(reportTable, 1) 'заполнение массивов
            datesBegin(i) = reportTable(i, 1)
            datesEnd(i) = reportTable(i, 2)
            cars(i) = reportTable(i, 3)
            statuses(i) = reportTable(i, 8)
        Next i
        
        For i = LBound(datesEnd) To UBound(datesEnd) 'если дата окончания не заполнена, считаем ее как сегодняшнюю
            If datesEnd(i) = Empty Then datesEnd(i) = Date + 1
        Next i
        
        counterOk = 1
        counterBroken = 1
        Dim okCars, brokenCars As Variant
        ReDim okCars(1 To UBound(cars)) 'ТС в работе
        ReDim brokenCars(1 To UBound(cars)) 'ТС в ремонте
        For i = LBound(datesBegin) To UBound(datesBegin)
            If reportDate >= CDate(datesBegin(i)) And reportDate <= CDate(datesEnd(i)) Then 'если дата отчета между датами ремонта ТС считаем эту ТС как В ремонте
                brokenCars(counterBroken) = cars(i)
                counterBroken = counterBroken + 1
            Else
                okCars(counterOk) = cars(i)  'иначе - В работе
                counterOk = counterOk + 1
            End If
        Next i
        
        If Not brokenCars(1) = Empty Then brokenCars = removeDublicatesFromOneDimArr(brokenCars) 'удаление дубликатов
        If Not okCars(1) = Empty Then okCars = removeDublicatesFromOneDimArr(okCars) 'удаление дубликатов
        
        Dim okCarsResult As Variant
        ReDim okCarsResult(1 To UBound(okCars)) 'убираем сломанные ТС из ТС в работе
        counterOk = 1
        For i = LBound(okCars) To UBound(okCars)
            isBroken = False
            For Each car In brokenCars
                If okCars(i) = car Then
                    isBroken = True
                    Exit For
                End If
            Next car
            If Not isBroken Then
                okCarsResult(counterOk) = okCars(i)
                counterOk = counterOk + 1
            End If
        Next i

        With Sheets("Статистика") 'заполнение листа с отчетом полученными массивами и форматирование
            .Cells(4, 1).Resize(UBound(okCarsResult), 1).Value = Application.Transpose(okCarsResult)
            .Cells(4, 2).Resize(UBound(brokenCars), 1).Value = Application.Transpose(brokenCars)
            lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
            lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
            lastRow = WorksheetFunction.Max(lastRow1, lastRow2)
            Range(Cells(4, 1), Cells(lastRow, 2)).Borders.LineStyle = xlContinuous
        End With
        
    End If
errorExit:
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate() 'отчет на сегодняшний день
    On Error GoTo errorExit 'уход в конец на ошибке чтобы у клиента не было никаких всплывающих ошибок
    lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
    lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
    lastRow = WorksheetFunction.Max(lastRow1, lastRow2)
    If Cells(lastRow, 1) = "В работе" Then lastRow = lastRow + 1
    Range(Cells(4, 1), Cells(lastRow, 2)).ClearContents
    Range(Cells(4, 1), Cells(lastRow, 2)).ClearFormats
    reportDate = Date
    Cells(1, 2) = reportDate
    reportTable = Sheets("Учет").ListObjects("УчетРемонта").DataBodyRange
    Dim datesBegin, datesEnd, cars, statuses As Variant
    ReDim datesBegin(1 To UBound(reportTable, 1))
    ReDim datesEnd(1 To UBound(reportTable, 1))
    ReDim cars(1 To UBound(reportTable, 1))
    ReDim statuses(1 To UBound(reportTable, 1))
    For i = LBound(reportTable, 1) To UBound(reportTable, 1)
        datesBegin(i) = reportTable(i, 1)
        datesEnd(i) = reportTable(i, 2)
        cars(i) = reportTable(i, 3)
        statuses(i) = reportTable(i, 8)
    Next i
    
    For i = LBound(datesEnd) To UBound(datesEnd)
        If datesEnd(i) = Empty Then datesEnd(i) = Date
    Next i
    
    counterOk = 1
    counterBroken = 1
    Dim okCars, brokenCars As Variant
    ReDim okCars(1 To UBound(cars))
    ReDim brokenCars(1 To UBound(cars))
    For i = LBound(datesBegin) To UBound(datesBegin)
        If reportDate >= CDate(datesBegin(i)) And reportDate <= CDate(datesEnd(i)) Then
            brokenCars(counterBroken) = cars(i)
            counterBroken = counterBroken + 1
        Else
            okCars(counterOk) = cars(i)
            counterOk = counterOk + 1
        End If
    Next i
    
    If Not brokenCars(1) = Empty Then brokenCars = removeDublicatesFromOneDimArr(brokenCars)
    If Not okCars(1) = Empty Then okCars = removeDublicatesFromOneDimArr(okCars)
    
    Dim okCarsResult As Variant
    ReDim okCarsResult(1 To UBound(okCars)) 'убираем сломанные ТС из ТС в работе
    counterOk = 1
    For i = LBound(okCars) To UBound(okCars)
        isBroken = False
        For Each car In brokenCars
            If okCars(i) = car Then
                isBroken = True
                Exit For
            End If
        Next car
        If Not isBroken Then
            okCarsResult(counterOk) = okCars(i)
            counterOk = counterOk + 1
        End If
    Next i
        
    With Sheets("Статистика") 'заполнение листа с отчетом полученными массивами и форматирование
        .Cells(4, 1).Resize(UBound(okCarsResult), 1).Value = Application.Transpose(okCarsResult)
        .Cells(4, 2).Resize(UBound(brokenCars), 1).Value = Application.Transpose(brokenCars)
        lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
        lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
        lastRow = WorksheetFunction.Max(lastRow1, lastRow2)
        Range(Cells(4, 1), Cells(lastRow, 2)).Borders.LineStyle = xlContinuous
    End With
errorExit:
    On Error GoTo 0
End Sub

