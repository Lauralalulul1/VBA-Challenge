Attribute VB_Name = "Module11"
Sub Stock_Market()
'Attribute to Module 11

'Create a script that will loop through all the stocks for one year and output the following information.
'Declare first set of Variables
Dim ticker_name As String
Dim ticker_total As Double
Dim LastRow As Long

Dim ws As Worksheet

Dim total_stock_volume As Currency
Dim year_change As Long
Dim close_price As Currency
Dim start_year As Long
Dim end_year As Long
Dim percent_change As Integer
Dim open_price As Currency

'Declare second set of Variables
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Integer
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Currency
Dim greatest_percent_increase As Integer
Dim greatest_stock_volume_ticker As String

'beginning values for first set of Variables
ticker_name = ""
ticker_total = 0
year_change = 0
open_price = 0
percent_change = 0
total_stock_volume = 0

'Loop
    For Each ws In Worksheets
    
'Activate Worksheet
ws.Activate

'Retrieval of ticker_name
    'Source from class activities in slack (Credit Card activity)
    ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    'Skip the header row and display variable
    For i = 2 To LastRow
    ticker_name = Cells(i, 1).Value

'if the next ticker name is different
    If Cells(i + 1, 1).Value <> ticker_name Then
 'store the total, one last time since the next row is a new ticker
     ticker_total = ticker_total + 1
'store the name
     ticker_name = Cells(ticker_total + 1, 9)
     
 'Retrieval of open price and display variable
    If open_price = 0 Then
    open_price = Cells(i, 3).Value
End If

'Retrieval of close price and display variable
  close_price = Cells(i, 6)
  
  
'Retrieval of total_stock_volume and display variable
      total_stock_volume = total_stock_volume + Cells(i, 7).Value
    

'Column Creation made to compile
ws.Range("I1").Value = "ticker name"
ws.Range("J1").Value = "year change"
ws.Range("k1").Value = "percent change"
ws.Range("L1").Value = "total stock volume"


'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
year_change = close_price - open_price

Cells(ticker_total + 1, 10).Value = year_change

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
If open_price = 0 Then
    percent_change = 0
    Else:
        percent_change = (year_change / open_price)
        End If
        
'Convert to percentage format
    Cells(ticker_total + 1, 11).NumberFormat = "0.00%"
               
'Set back to 0 for next different ticker
   open_price = 0
   
'put in worksheet & adding stock volume
    total_stock_volume = Cells(ticker_total + 1, 12).Value
        
'Set back to 0 for next different ticker symbol
total_stock_volume = 0

    End If
    
    
Next i

'Column Creation to displlay second set of variables
    Range("O2").Value = "Greatest Percent Increase"
    Range("O3").Value = "Greatest Percent Decrease"
    Range("O4").Value = "Greatest Total Volume"
    

'Find Last row
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
'Values for second set of Variables
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    'Skip first row and loop
    For i = 2 To LastRow
    
        
'ticker greatest percent increase
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
'ticker greatest percent decrease
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
'Ticker greatest stock volume
        If Cells(i, 12).Value > greatest_stock_volume Then
             greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
  Next i
    
'Change to Percent Format
    Range("P2").NumberFormat = "0.00%"
    Range("Q2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("P4").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0.00%"
    
 Next ws


End Sub
