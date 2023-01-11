Attribute VB_Name = "Module1"
Sub Data_Test():

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
'Dim stock_volume As Double
'stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet (ws)
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0
Dim price_change_percent As Double
price_change_percent = 0
Dim TickerRow As Long: TickerRow = 1

open_price = ws.Cells(2, 3).Value

'Do loop of current worksheet to Lastrow

For i = 2 To Lastrow

'Ticker symbol output
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    TickerRow = TickerRow + 1
    Ticker = ws.Cells(i, 1).Value
    ws.Cells(TickerRow, "I").Value = Ticker
    
    'Calculate change in Price
    close_price = ws.Cells(i, 6).Value
    
    ws.Cells(TickerRow, "J").Value = close_price - open_price
    
    
        If open_price = 0 Then
        price_change_percent = 0
        
        ElseIf open_price <> 0 Then
        
        price_change_percent = ((close_price - open_price) / open_price)
        
        End If
        
        ws.Cells(TickerRow, "K").Value = price_change_percent
        
        Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
        ws.Cells(TickerRow, "L").Value = Ticker_volume
        Ticker_volume = 0
        open_price = ws.Cells(i + 1, 3).Value
        
    Else: Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
    
    End If

Next i

ws.Range("K2:K" & TickerRow).NumberFormat = "0.00%"
For i = 2 To TickerRow

    If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Cells(i, 11).Value
        ws.Range("P2").Value = ws.Cells(i, 9).Value
               
    End If
        
    If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Cells(i, 11).Value
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    
               
        End If
        
    If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
    ws.Range("Q4").Value = ws.Cells(i, 12).Value
    ws.Range("P4").Value = ws.Cells(i, 9).Value
        
               
        End If
        
    If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
        Else
    
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    
        End If
       

Next i


Next ws


End Sub



