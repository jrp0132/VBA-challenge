Attribute VB_Name = "Module1"
Sub CalculateQuarterlyChangePerSheet()
    Dim wsQ1 As Worksheet
    Dim wsQ2 As Worksheet
    Dim wsQ3 As Worksheet
    Dim wsQ4 As Worksheet
    Dim tickerDataQ1 As Object
    Dim tickerDataQ2 As Object
    Dim tickerDataQ3 As Object
    Dim tickerDataQ4 As Object
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    
    Set tickerDataQ1 = CreateObject("Scripting.Dictionary")
    Set tickerDataQ2 = CreateObject("Scripting.Dictionary")
    Set tickerDataQ3 = CreateObject("Scripting.Dictionary")
    Set tickerDataQ4 = CreateObject("Scripting.Dictionary")
    
    Set wsQ1 = ThisWorkbook.Sheets(1)
    Set wsQ2 = ThisWorkbook.Sheets(2)
    Set wsQ3 = ThisWorkbook.Sheets(3)
    Set wsQ4 = ThisWorkbook.Sheets(4)
    
    ' Initialize summary values
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    ' Process data for each quarter
    ProcessQuarterData wsQ1, tickerDataQ1, maxIncrease, maxIncreaseTicker, maxDecrease, maxDecreaseTicker, maxVolume, maxVolumeTicker
    ProcessQuarterData wsQ2, tickerDataQ2, maxIncrease, maxIncreaseTicker, maxDecrease, maxDecreaseTicker, maxVolume, maxVolumeTicker
    ProcessQuarterData wsQ3, tickerDataQ3, maxIncrease, maxIncreaseTicker, maxDecrease, maxDecreaseTicker, maxVolume, maxVolumeTicker
    ProcessQuarterData wsQ4, tickerDataQ4, maxIncrease, maxIncreaseTicker, maxDecrease, maxDecreaseTicker, maxVolume, maxVolumeTicker
    
    ' Output the summarized data for each ticker
    OutputSummaryData wsQ1, tickerDataQ1, maxIncreaseTicker, maxDecreaseTicker, maxVolumeTicker
    OutputSummaryData wsQ2, tickerDataQ2, maxIncreaseTicker, maxDecreaseTicker, maxVolumeTicker
    OutputSummaryData wsQ3, tickerDataQ3, maxIncreaseTicker, maxDecreaseTicker, maxVolumeTicker
    OutputSummaryData wsQ4, tickerDataQ4, maxIncreaseTicker, maxDecreaseTicker, maxVolumeTicker
End Sub

Sub ProcessQuarterData(ws As Worksheet, ByRef tickerData As Object, ByRef maxIncrease As Double, ByRef maxIncreaseTicker As String, ByRef maxDecrease As Double, ByRef maxDecreaseTicker As String, ByRef maxVolume As Double, ByRef maxVolumeTicker As String)
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim firstRow As Long
    Dim lastClosingPrice As Double
    Dim totalVolume As Double
    Dim firstPriceSet As Boolean
    firstPriceSet = False
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    
    ' Loop through each row in the current sheet
    For i = 2 To lastRow
        ticker = ws.Cells(i, 1).Value
        If Not firstPriceSet Then
            openingPrice = ws.Cells(i, 3).Value ' First opening price
            firstPriceSet = True
        End If
        lastClosingPrice = ws.Cells(i, 6).Value ' Last closing price
        totalVolume = totalVolume + ws.Cells(i, 7).Value ' Sum volume
        
        If i = lastRow Or ws.Cells(i + 1, 1).Value <> ticker Then
            ' Calculate quarterly change
            Dim quarterlyChange As Double
            quarterlyChange = lastClosingPrice - openingPrice
            
            ' Calculate percentage change
            Dim percentChange As Double
            If openingPrice <> 0 Then
                percentChange = ((lastClosingPrice - openingPrice) / openingPrice) * 100 ' Convert to percentage
            Else
                percentChange = 0
            End If
            
            ' Save data for the ticker
            tickerData(ticker) = Array(quarterlyChange, percentChange, totalVolume)
            
            ' Summary values
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncreaseTicker = ticker
            End If
            
            If percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecreaseTicker = ticker
            End If
            
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                maxVolumeTicker = ticker
            End If
            
            firstPriceSet = False
            totalVolume = 0
        End If
    Next i
End Sub

Sub OutputSummaryData(ws As Worksheet, ByRef tickerData As Object, maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String)
    Dim outputRow As Long
    outputRow = 2

    ' Add headers to the output columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

    ' Output the data for each ticker
    For Each Key In tickerData.keys
        Dim dataArray As Variant
        dataArray = tickerData(Key)
        ws.Cells(outputRow, 9).Value = Key ' Ticker symbol
        ws.Cells(outputRow, 10).Value = dataArray(0) ' Total quarterly change
        ws.Cells(outputRow, 11).Value = dataArray(1) ' Percentage change
        ws.Cells(outputRow, 12).Value = dataArray(2) ' Total volume

        ' Conditional formatting for Quarterly Change column
        If dataArray(0) > 0 Then
            ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive quarterly change
        ElseIf dataArray(0) < 0 Then
            ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative quarterly change
        End If

        outputRow = outputRow + 1
    Next Key
    
    ' Output the summary of the greatest values in specified cells
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = maxIncreaseTicker
    ws.Cells(2, 17).Value = Format(tickerData(maxIncreaseTicker)(1), "0.00") & "%"
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = maxDecreaseTicker
    ws.Cells(3, 17).Value = Format(tickerData(maxDecreaseTicker)(1), "0.00") & "%"
    
    ws.Cells(5, 15).Value = "Greatest Total Volume"
    ws.Cells(5, 16).Value = maxVolumeTicker
    ws.Cells(5, 17).Value = tickerData(maxVolumeTicker)(2)
End Sub

