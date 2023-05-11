Sub Stock_test()

lastrow = Cells(Rows.Count, 1).End(xlUp).Row 'useful for longer spreadsheets (expecially those with over 750k entries)
Dim YearChange As Double
Dim PerChange As Double
Dim StockVol As Long
Dim printRow As Integer
Dim j As Integer
dim h as integer 'used because last row goes way further than is necessary for the summary columns
printRow = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 16).Value = Cells(1, 9).Value
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
'Loop and decision statements for the summary columns
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
        h = h + 1

        Else
        Cells(printRow, 12).Value = Cells(printRow, 12).Value + Cells(i, 7)
        j = j + 1
        End If

    Next i
    'Conditional formatting for the second column of summaries
    for i = 2 to h

        if cells(i, 10) < 0 Then

        Cells(i, 10).interior.colorindex = 3

        Else

        Cells(i, 10).interior.colorindex = 4

        end if
    
    next i
'Setting up the summary statistics columns
Cells(2, 17).value = FormatPercent(worksheetfunction.max(Range("K2:K" & h)))
Cells(3, 17).value = FormatPercent(worksheetfunction.min(Range("K2:K" & h)))
Cells(4, 17).value = worksheetfunction.max(Range("L2:L" & h))
    
    For i = 2 to h

        if cells(i, 11).value = cells(2, 17).value Then

        Cells(2, 16).value = cells(i, 9)

        elseif cells(i, 11).value = cells(3, 17).value Then

        Cells(3, 16).value = cells(i, 9)
        
        elseif cells(i, 12).value = cells(4, 17).value Then

        Cells(4, 16).value = cells(i, 9)

        end if
    
    next i
    
End Sub
