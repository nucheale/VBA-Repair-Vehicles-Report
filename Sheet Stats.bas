Function removeDublicatesFromOneDimArr(arr)
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


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Cells(1, 2)) Is Nothing Then
        Str = Target.Row
        Stlb = Target.Column
        Calendar.Show vbModeless
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Selection.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Cells(1, 2)) Is Nothing Then
        lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
        lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
        lastRow = WorksheetFunction.Max(lastRow1, lastRow2)
        If Cells(lastRow, 1) = "В работе" Then lastRow = lastRow + 1
        Range(Cells(4, 1), Cells(lastRow, 2)).ClearContents
        Range(Cells(4, 1), Cells(lastRow, 2)).ClearFormats
        reportDate = CDate(Cells(1, 2))
        reportTable = Sheets("Учет").ListObjects("УчетРемонта").DataBodyRange
        Dim datesBegin, datesEnd, cars, statuses As Variant
        ReDim datesBegin(1 To UBound(reportTable, 1))
        ReDim datesEnd(1 To UBound(reportTable, 1))
        ReDim cars(1 To UBound(reportTable, 1))
        ReDim statuses(1 To UBound(reportTable, 1))
        'MsgBox UBound(reportTable, 1)
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
            ' If reportDate >= CDate(datesBegin(i)) And reportDate <= CDate(datesEnd(i)) Then
            '     If statuses(i) = "В работе" Then
            '         okCars(counterOk) = cars(i)
            '         counterOk = counterOk + 1
            '     Else
            '         brokenCars(counterBroken) = cars(i)
            '         counterBroken = counterBroken + 1
            '     End If
            ' End If
            If reportDate >= CDate(datesBegin(i)) And reportDate <= CDate(datesEnd(i)) Then
                brokenCars(counterBroken) = cars(i)
                counterBroken = counterBroken + 1
            Else
                okCars(counterOk) = cars(i)
                counterOk = counterOk + 1
            End If
        Next i
        brokenCars = removeDublicatesFromOneDimArr(brokenCars)
        okCars = removeDublicatesFromOneDimArr(okCars)

        With Sheets("Статистика")
            .Cells(4, 1).Resize(UBound(okCars), 1).Value = Application.Transpose(okCars)
            .Cells(4, 2).Resize(UBound(brokenCars), 1).Value = Application.Transpose(brokenCars)
            lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
            lastRow2 = Cells(Rows.Count, 2).End(xlUp).Row
            lastRow = WorksheetFunction.Max(lastRow1, lastRow2)
            Range(Cells(4, 1), Cells(lastRow, 2)).Borders.LineStyle = xlContinuous
        End With
        
    End If
    
End Sub
