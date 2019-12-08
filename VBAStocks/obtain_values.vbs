Sub obtain_values()
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    currentFirstRow = 2
    newRow = 2
    totalStockVolume = 0
    yearlyChange = 0
    percentageChange = 0

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Volume"
    
    For i = 2 To lastRow
        totalStockVolume = totalStockVolume + Range("G" & i).Value
        If Range("A" & i).Value <> Range("A" & i + 1).Value Then
            Range("I" & newRow).Value = Range("A" & i).Value
            If Range("C" & currentFirstRow).Value = 0 Then
                yearlyChange = 0
            Else
                yearlyChange = Range("F" & i).Value - Range("C" & currentFirstRow).Value
            End If
            percentageChange = Round(yearlyChange / Range("C" & currentFirstRow).Value, 4)
            Range("J" & newRow).Value = yearlyChange
            If yearlyChange < 0 Then
                Range("J" & newRow).Interior.ColorIndex = 3
            Else
                Range("J" & newRow).Interior.ColorIndex = 4
            End If
            Range("K" & newRow).Value = percentageChange
            Range("L" & newRow).Value = totalStockVolume
            newRow = newRow + 1
            currentFirstRow = i + 1
            totalStockVolume = 0
            yearlyChange = 0
            percentageChange = 0
        End If
    Next i
End Sub