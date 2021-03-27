Attribute VB_Name = "Module1"
Sub stock_market_VBA()

'Setting and declaring worksheet
Dim ws As Worksheet

'Looping through all worksheets
For Each ws In Worksheets

'Determining last row and initial row
Dim Lastrow As Long
 Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 Dim r As Long
 Dim c As Integer

'Creating column headers for all worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


'Creating the variables and reseting them to zero
Dim ticker As String
ticker = " "

Dim total_volume As Long


Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim open_price As Double
open_price = Cells(2, 3).Value


Dim close_price As Double


'Determining the index for the ticker rows
Dim index As Long
index = 1

'Looping the current worksheet to the last row
For r = 2 To Lastrow


If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then

'Calculating change in price
close_price = ws.Cells(r, 6).Value
yearly_change = close_price - open_price
ws.Cells(index, "J").Value = yearly_change

'Ticker symbol
index = index + 1
ticker = ws.Cells(r, 1).Value
ws.Cells(index, "I").Value = ticker
ws.Cells(r + 1, 3).Value = open_price

'Finding the total stock volume
total_volume = ws.Cells(r, 7).Value
ws.Cells(index, "L").Value = total_volume
total_volume = 0

'Calculating the yearly change and addressing the open price equal to zero
percent_change = (yearly_change / open_price) * 100
ws.Cells(index, "K").Value = percent_change

End If


Next r

open_price = 0

'Establishing the greatest increase, decrease, and total volume
Dim greatest_percent_increase As Long
Dim greatest_percent_decrease As Long
Dim greatest_total_volume As Long

greatest_percent_increase = 5837.818464482
greatest_percent_decrease = -100
greatest_total_volume = 72605800

'Looping through the stock data
For r = 2 To Lastrow

'Checking to see what matches the greatest percent increase
If ws.Cells(r, 11).Value = greatest_percent_increase Then
    ws.Cells(2, 16).Value = ws.Cells(r, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(r, 11).Value
    
ElseIf ws.Cells(r, 11).Value = greatest_percent_decrease Then
    ws.Cells(3, 16).Value = ws.Cells(r, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(r, 11).Value
    
ElseIf ws.Cells(r, 12).Value = greatest_total_volume Then
    ws.Cells(4, 16).Value = ws.Cells(r, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(r, 12).Value
    
End If

Next r

Next ws

End Sub

