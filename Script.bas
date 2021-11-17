Attribute VB_Name = "Module1"
 Sub stock_analysis():
'create a loop through all the stocks for 1 year to output:
'1)ticker symbol
'2)yearly change from opening price @ the beginning of a given year and the closing price @ the end of the year
'3)% change from opening price at the beginning of a given year to the closing price at the end of that year
'4)total stock volume

'Activate sheet A
Sheets("A").Activate

'Set headers into designated cells
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total stock volume"

'set a variable for change from beginning to the end of year
Dim yearly_change As Double

'set a variable for % change from beginning to the end of year
Dim Percentage_change As Double

'set a variable for total stock volume
Dim Total_Stock_volume As Long

'set a variable for last row
Dim Last_Row As Long

'set a variable for the closing price
Dim finish As Double

'set a variable for the ticker to count
Dim open_ticker As Long
open_ticker = 2

'set a variable for the yearly open price
Dim yearly_open As Double

'set variable for ticker row number
Dim ticker_row_number As Integer
'set ticker row number to row 2
ticker_row_number = 2

'set variable for percentage change row to 2
Dim percentage_change_row As Integer
percentage_change_row = 2

'loop through rows in the ticker column to ID first pay and last pay
For i = 2 To 70926
   
    'find where the last A in ticker is
    If cells(i + 1, 1).Value <> cells(i, 1).Value Then
       'set yearly open value as the open ticker row and column 3
    yearly_open = cells(open_ticker, 3).Value
        'assign the value of finish
        finish = cells(i, 6).Value
        
        'calculate the yearly change by the difference of finish and yearly_open price
       yearly_change = finish - yearly_open
      
       'print the value of yearly change
       cells(ticker_row_number, 10).Value = yearly_change
       'print the value of ticker
       cells(ticker_row_number, 9).Value = cells(i, 1).Value
       
       'calculate percentage change
       Percentage_change = (finish - yearly_open) / yearly_open
       
       'print percentage change
       cells(percentage_change_row, 11).Value = Percentage_change
       
             If Percentage_change > 0 Then
       
             cells(percentage_change_row, 11).Interior.ColorIndex = 4
       
             Else
       
             cells(percentage_change_row, 11).Interior.ColorIndex = 3
       
             End If
       
        
              
      open_ticker = i + 1
             percentage_change_row = percentage_change_row + 1
       
         total_stock_volume_row = total_stock_volume_row + 1
   ticker_row_number = ticker_row_number + 1
    
    End If
    
    Total_Stock_volume = cells(i, 7).Value
   
     cells(ticker_row_number, 12) = Total_Stock_volume

Next i

End Sub
 

