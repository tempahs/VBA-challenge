Sub getTickerAndVolume()
    Call performanceOff
   Call allWorksheets
End Sub

Function allWorksheets()
    Dim f_activeSheet As Worksheet
    For Each f_activeSheet In Worksheets
        Application.StatusBar = "Working on " + ActiveSheet.Name + " sheet..."
        f_activeSheet.Select
        Call TickerVolume
    Next
    MsgBox ("Script ended")
End Function

Function TickerVolume()
    Call performanceOn
    Dim currentTicker1, currentTicker2 As String
    Dim currentVolume, totalVolume As Double
    Dim currentRow As Integer
    Dim loopCounter As Integer
    Dim lastRow As Long
    totalVolume = 0
    currentRow = 1
    loopCounter = -1
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        currentTicker1 = Cells(i, 1).Value2
        currentTicker2 = Cells(i + 1, 1).Value2
        currentVolume = Cells(i, 7).Value2
        totalVolume = totalVolume + currentVolume
        loopCounter = loopCounter + 1
        
        If (currentTicker2 <> currentTicker1) Then
            currentRow = currentRow + 1
            Cells(currentRow, 9).Value = currentTicker1
            Cells(currentRow, 12).Value = totalVolume
            Call TickerStockVolume(currentRow, i, loopCounter)
            totalVolume = 0
            loopCounter = -1
        End If
    Next i
    Application.StatusBar = "Done"
    getGreatest (currentRow)
    Call performanceOff
End Function

Function TickerStockVolume(f_currentRow, f_i, f_loopCounter)
    Dim openPrice, closePrice As Double
    openPrice = Cells(f_i - f_loopCounter, 3).Value2
    closePrice = Cells(f_i, 6).Value2
    Cells(f_currentRow, 10).Value = closePrice - openPrice
    Cells(f_currentRow, 10).NumberFormat = "0.000000000"
    If (Cells(f_currentRow, 10).Value) >= 0 Then
        Cells(f_currentRow, 10).Interior.ColorIndex = 4
    Else
        Cells(f_currentRow, 10).Interior.ColorIndex = 3
    End If
    If (openPrice = 0 Or closePrice = 0) Then
        Cells(f_currentRow, 11).Value = 0
    Else
        Cells(f_currentRow, 11).Value = closePrice / openPrice - 1
    End If
    Cells(f_currentRow, 11).NumberFormat = "0.00%"
End Function

Function getGreatest(countTotalRows)
    Dim tickerRows As Integer
    Dim currentPercent, maxPercent, currentStockVolume, maxStockVolume As Double
    tickerRows = countTotalRows
    maxPercent = 0
    minPercent = 0
    maxStockVolume = 0
    For i = 2 To tickerRows
        currentPercent = Cells(i, 11).Value2
        If currentPercent > maxPercent Then
            maxPercent = currentPercent
            Range("P2").Value = Cells(i, 9).Value2
            Range("Q2").Value = maxPercent
            Range("Q2").NumberFormat = "0.00%"
        End If
        If currentPercent < minPercent Then
            minPercent = currentPercent
            Range("P3").Value = Cells(i, 9).Value2
            Range("Q3").Value = minPercent
            Range("Q3").NumberFormat = "0.00%"
        End If
        currentStockVolume = Cells(i, 12).Value2
        If currentStockVolume > maxStockVolume Then
            maxStockVolume = currentStockVolume
            Range("P4").Value = Cells(i, 9).Value2
            Range("Q4").Value = maxStockVolume
        End If

    Next i
    Call ColumnLabels
End Function

Function ColumnLabels()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greates % Increase"
    Cells(3, 15).Value = "Greates % Decrease"
    Cells(4, 15).Value = "Greates Total Volume"
End Function

Function performanceOn()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Function

Function performanceOff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Function
