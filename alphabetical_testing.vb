Sub Stock_test()

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim YearChange As Double
Dim PerChange As Double
Dim StockVol As Long
Dim printRow As Integer
Dim j As Integer
printRow = 2


    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        YearChange = Cells(i, 6).Value - Cells(i - j, 3)
        PerChange = (Cells(i, 6).Value - Cells(i - j, 3)) / Cells(i - j, 3)
        
        Cells(printRow, 9).Value = Cells(i, 1)
        Cells(printRow, 10).Value = YearChange
        Cells(printRow, 11).Value = FormatPercent(PerChange)
        Cells(printRow, 12).Value = Cells(printRow, 12).Value + Cells(i, 7)
        
        printRow = printRow + 1
        j = 0
        StockVal = 0
    
        Else
        Cells(printRow, 12).Value = Cells(printRow, 12).Value + Cells(i, 7)
        j = j + 1
        End If

    Next i

End Sub
