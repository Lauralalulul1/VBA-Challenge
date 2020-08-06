Attribute VB_Name = "Module1"
Sub Stock_Market()
'Attribute to Module 1

'Create a script that will loop through all the stocks for one year and output the following information.
'Define Variables
Dim ticker_name As String
ticker_name = Cells(ticker_total + 1, 9).Value

Dim ticker_total As Double

Dim LastRow As Long
Dim ticker_column As Integer
Dim summary_row As Integer

Dim ws As Worksheet

Dim total_stock_volume As Long

Dim year_change As Long
Dim open_price As Long
Dim close_price As Long
Dim start_year As Long
Dim end_year As Long

Dim percent_change As Double

 
    'Must first Activate worksheet or it won't run
    For Each ws In Worksheets
    ws.Activate
    
'Loop through the tickers -source from class activities in slack (Credit Card activity)
    ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'We skip the header row, therefore i = 2
For i = 2 To LastRow
 'if the next ticker name is different
 If Cells(i + 1, 1).Value <> ticker_name Then
 'store the total, one last time since the next row is a new ticker
     ticker_total = ticker_total + 1
  'because we are seeing the next ticker name is changed, reset the total
     ticker_total = 0
             Else
        'we are in the same ticker name, so add the value
        ticker_total = ticker_total + Cells(i, 7).Value
            End If
            
 'Column Creation made to compile
ws.Range("I1").Value = "ticker_name"
ws.Range("J1").Value = "year change"
ws.Range("k1").Value = "percent change"
ws.Range("L1").Value = "total stock volume"

'Close price for ticker name
close_price = Cells(i, 6)

'The total stock volume of the stock.Cells(ticker_total + 1, 12).Value = total_stock_volume
total_stock_volume = total_stock_volume + Cells(i, 7).Value
total_stock_volume = 0

 
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
year_change = 0
Cells(ticker_total + 1, 10).Value = year_change
open_price = 0
year_change = close_price - open_price


'You should also have conditional formatting that will highlight positive change in green and negative change in red.
    'Format year change to show drop or rise in stock
            If year_change > 0 Then
                Cells(ticker_total + 1, 10).Interior.ColorIndex = 4
                '(green - rise in stock)
                      ElseIf year_change < 0 Then
                Cells(ticker_total + 1, 10).Interior.ColorIndex = 3
                '(red - fall in stock)
                     Else: Cells(ticker_total + 1, 10).Interior.ColorIndex = 6
                '(yellow - no change)
                    End If
                    
            
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
percent_change = 0
If open_price = 0 Then
    percent_change = 0
    Else: percent_change = (year_change / open_price)
        
'Convert to percentage format
Cells("K" & summary_row).NumberFormat = "0.00%"

'Open Price to 0 when different ticker name
open_price = 0



'Total Stock Volume to 0 when diferent ticker name
tock_stock_volume = 0
 End If
 Next i
  'Variable Values for percent decreases and increases
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' skipping the header row, loop through the list of tickers.
    For i = 2 To lastRowState
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker name with the greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Ticker name with greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet.
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub
