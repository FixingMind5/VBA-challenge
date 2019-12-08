Sub comparing_stock_volume()
    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    greaterStockVolume = 0
    greatestIncrease = 0
    greatestDecrease = 0

    greaterStockVolumeIndex = 0
    greatestIncreaseIndex = 0
    greatestDecreaseIndex = 0


    For i = 2 To lastRow
        If Range("L" & i).Value > greaterStockVolume Then
            greaterStockVolume = Range("L" & i).Value
            greaterStockVolumeIndex = i
        End If

        If Range("K" & i).Value > greatestIncrease Then
            greatestIncrease = Range("K" & i).Value
            greatestIncreaseIndex = i
        End If

        If Range("K" & i).Value < greatestDecrease Then
            greatestDecrease = Range("K" & i).Value
            greatestDecreaseIndex = i
        End If

    Next i

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = Range("I" & greatestIncreaseIndex).Value
    Range("Q2").Value = greatestIncrease

    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = Range("I" & greatestDecreaseIndex).Value
    Range("Q3").Value = greatestDecrease

    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = Range("I" & greaterStockVolumeIndex).Value
    Range("Q4").Value = greaterStockVolume
End Sub