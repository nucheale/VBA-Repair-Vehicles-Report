Private Sub Worksheet_Change(ByVal Target As Range)
    If Selection.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Cells(1, 2)) Is Nothing Then
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(lastRow, 1) = "Дата" Then lastRow = lastRow + 1
        Range(Cells(5, 1), Cells(lastRow, 6)).ClearContents
        Range(Cells(5, 1), Cells(lastRow, 6)).ClearFormats
        reportCar = CStr(Cells(1, 2))
        reportTable = Sheets("Учет").ListObjects("УчетРемонта").DataBodyRange
        For i = LBound(reportTable, 1) To UBound(reportTable, 1)
            If reportTable(i, 2) = Empty Then reportTable(i, 2) = "настоящее время"
        Next i
        Dim reportArr As Variant
        ReDim reportArr(1 To UBound(reportTable, 1), 1 To UBound(reportTable, 2))
        counter = 1
        For i = LBound(reportTable, 1) To UBound(reportTable, 1)
            If reportTable(i, 3) = reportCar Then
                reportArr(counter, 1) = reportTable(i, 1) & " – " & reportTable(i, 2)
                reportArr(counter, 2) = reportTable(i, 4)
                reportArr(counter, 3) = reportTable(i, 9)
                reportArr(counter, 4) = reportTable(i, 5)
                reportArr(counter, 5) = reportTable(i, 6)
                reportArr(counter, 6) = reportTable(i, 7)
                counter = counter + 1
            End If
        Next i
        
        Cells(5, 1).Resize(UBound(reportArr), UBound(reportArr, 2)).Value = reportArr
    
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        Range(Cells(5, 1), Cells(lastRow, 6)).Borders.LineStyle = xlContinuous
        Range(Cells(5, 5), Cells(lastRow, 6)).NumberFormat = "#,##0"
    End If
    
End Sub