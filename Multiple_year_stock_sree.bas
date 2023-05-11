Attribute VB_Name = "Module1"
Sub ticker()
Dim s, worksheet_count As Integer
worksheet_count = ActiveWorkbook.Worksheets.Count
'MsgBox ("worksheet count " + Str(worksheet_count))

For s = 1 To worksheet_count

'MsgBox ("Executing it for " + Str(s) + " time")

ThisWorkbook.Worksheets(s).Activate

'MsgBox (ThisWorkbook.Worksheets(s).Name)

    'Adding new columns to the sheet
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percentage Change"
    Range("L1") = "Total Stock Volume"
    Range("O2") = "Greatest % increase"
    Range("O3") = "Greatest % decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"

        Dim ticker As String
        Dim openingPrice, endingPrice, yearlyChange, percentageChange, totalStockVolume, initialRecord As Double
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox ("lastRow " + Str(lastRow))
        newRow = 2
        
        For i = 2 To lastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'Assign Ticker value
                Cells(newRow, 9).Value = Cells(i, 1).Value
                'Assign Total Stock volume
                Cells(newRow, 12).Value = totalStockVolume + Cells(i, 7).Value
                'Read the closing price before breaking the loop
                closingPrice = Cells(i, 6).Value
                'Calculate the yearly change before breaking the loop
                yearlyChange = Round((closingPrice - openingPrice), 2)
                'Assign the yearly change to cell in excel sheet
                Cells(newRow, 10).Value = Format(yearlyChange, "0.00")
                If (yearlyChange > 0) Then
                    Cells(newRow, 10).Interior.ColorIndex = 4
                Else:
                    Cells(newRow, 10).Interior.ColorIndex = 3
                End If
                'calculate the percentage and assign it to the new column
                percentageChange = yearlyChange / openingPrice
                'Formatting the percentage and writing into the columns
                Cells(newRow, 11) = Format(percentageChange, "0.00%")
                'Increment the row to display the ticker in a new row
                newRow = newRow + 1
                'Set Initial record value to zero to assign the new opening price
                initialRecord = 0
                totalStockVolume = 0
            Else:
                totalStockVolume = totalStockVolume + Cells(i, 7).Value
                ' if this is the initial record, assign the opening price to a variable and increment the initial record count to 1
                If (initialRecord = 0) Then
                    initialRecord = 1
                    openingPrice = Cells(i, 3)
                End If
            End If
        Next i
        
        'Find the greatest % of increase, decrease and total stock volume
        
        'lastRow in the column Yearly Change
        lastRowYearlyChange = Cells(Rows.Count, 10).End(xlUp).Row
        Dim greatestVolume, greatestPercentageInc, greatestPercentageDec As Double
        Dim tickerInc, tickerDec, tickerVolume As String
        
        
        For j = 2 To lastRowYearlyChange
            If (Cells(j, 11).Value > greatestPercentageInc) Then
                greatestPercentageInc = Cells(j, 11).Value
                tickerInc = Cells(j, 9).Value
            End If
            
            If (Cells(j, 11).Value < greatestPercentageDec) Then
                greatestPercentageDec = Cells(j, 11).Value
                 tickerDec = Cells(j, 9).Value
            End If
            
            If (Cells(j, 12).Value > greatestVolume) Then
                greatestVolume = Cells(j, 12).Value
                 tickerVolume = Cells(j, 9).Value
            End If
            
        Next j
        
        Range("P2") = tickerInc
        Range("P3") = tickerDec
        Range("P4") = tickerVolume
        Range("Q2") = Format(greatestPercentageInc, "0.00%")
        Range("Q3") = Format(greatestPercentageDec, "0.00%")
        Range("Q4") = Format(greatestVolume, "##0.0E+0")

Next s
End Sub
