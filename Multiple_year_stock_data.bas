Attribute VB_Name = "Module1"
Sub AnalyzeStocks()

' Declare variables
Dim ws As Worksheet
Dim lastRow As Long
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As Double
Dim summaryRow As Long
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double
Dim maxIncreaseTicker As String
Dim maxDecreaseTicker As String
Dim maxVolumeTicker As String

' Loop through each worksheet in the workbook
For Each ws In Worksheets

' Set the summary row to start at row 2
summaryRow = 2

' Initialize maximum increase, decrease, and volume variables
maxIncrease = 0
maxDecrease = 0
maxVolume = 0
maxIncreaseTicker = ""
maxDecreaseTicker = ""
maxVolumeTicker = ""

' Set headers for the output table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

' Find the last row with data in column A
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Loop through each row in the worksheet
For i = 2 To lastRow

' If the ticker symbol in the current row is different from the previous row
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set the ticker symbol
ticker = ws.Cells(i, 1).Value

' Set the opening price
openPrice = ws.Cells(i, 3).Value

' Reset the total volume
totalVolume = 0

End If

' Add the current row's volume to the total volume
totalVolume = totalVolume + ws.Cells(i, 7).Value

' If the ticker symbol in the current row is different from the next row
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set the closing price
closePrice = ws.Cells(i, 6).Value

' Calculate the yearly change in price
yearlyChange = closePrice - openPrice

' Calculate the percentage change in price
If openPrice <> 0 Then
percentChange = yearlyChange / openPrice
Else
percentChange = 0
End If

' Output the ticker symbol and changes
ws.Range("I" & summaryRow).Value = ticker
ws.Range("J" & summaryRow).Value = yearlyChange
ws.Range("K" & summaryRow).Value = percentChange
ws.Range("L" & summaryRow).Value = totalVolume

' Update the maximum increase, decrease, and volume values
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

' Increment the summary row
summaryRow = summaryRow + 1

End If

Next i

' Output the maximum increase, decrease, and volume values
ws.Range("P2").Value = maxIncreaseTicker
ws.Range("P3").Value = maxDecreaseTicker
ws.Range("P4").Value = maxVolumeTicker
ws.Range("Q2").Value = maxIncrease
ws.Range("Q3").Value = maxDecrease
ws.Range("Q4").Value = maxVolume

' Format themaximum increase, decrease, and volume values
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0"

Next ws

End Sub
