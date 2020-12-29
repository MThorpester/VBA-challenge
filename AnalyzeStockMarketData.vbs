Attribute VB_Name = "AnalyzeStockMarketData"
Sub StockData()

' Declare variables
Dim current_ticker As String
Dim current_volume As Double
Dim current_open As Currency
Dim current_close As Currency
Dim next_open As Currency
Dim yearly_change_dollars As Currency
Dim yearly_change_pct As Double
Dim last_row_sheet As Long
Dim summary_table_row As Integer
Dim ws As Worksheet

' Loop through all workesheets in the workbook

For Each ws In Worksheets
    
    ' Determine the last row on the sheet
    last_row_sheet = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Insert new column headers
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "Total Stock Volume"
    ws.Range("K1").Value = "Yearly Change(dollars)"
    ws.Range("L1").Value = "Yearly Change(percent)"
    ws.Range("L2:L" & last_row_sheet).NumberFormat = "0.00%"
    
    ws.Range("O1").Value = "Yearly Summary"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
       
    ' Loop through each row on the sheet
    For i = 2 To last_row_sheet
    
        ' If this is the first row on the sheet: initialize opening price, current volume, ticker symbol & summary table row count
        If i = 2 Then
            next_open = ws.Cells(i, 3).Value
            summary_table_row = 2
            current_volume = 0
            current_ticker = " "
        End If
        
        ' Check if this row and the next of data are associated with the same ticker symbol,
        ' If it is not, then we have reached the end of a tcker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Ticker symbol, add to the total volume, save the yearly opening & closing price for the current ticker
          current_ticker = ws.Cells(i, 1).Value
          current_volume = current_volume + ws.Cells(i, 7).Value
          current_close = ws.Cells(i, 6).Value
          current_open = next_open
                
          ' Print the Ticker Symbol, yearly stock volume, yearly opening & closing price for the current ticker symbol
          ' to the summary table and add to the summary table row count
          ws.Range("I" & summary_table_row).Value = current_ticker
          ws.Range("J" & summary_table_row).Value = current_volume
          ws.Range("K" & summary_table_row).Value = current_close - current_open
          
          ' compute the percent change, guarding against a zero divide error
          If (current_open = 0 And current_close = 0) Then
            yearly_change_pct = 0
          Else
            If current_close = 0 Then
                yearly_change_pct = -100
            Else
                If current_open = 0 Then
                    yearly_change_pct = 100
                Else
                    yearly_change_pct = ((current_close - current_open) / current_open)
                End If
            End If
           End If
                  
          ws.Range("L" & summary_table_row).Value = yearly_change_pct
          
          'Format the background colors for price changes and percent changes: green for positive, red for negative
          'Set the background color to Red
           If (current_close - current_open) < 0 Then
              ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
              ws.Range("L" & summary_table_row).Interior.ColorIndex = 3
           Else
              If (current_close - current_open) > 0 Then
                ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
                ws.Range("L" & summary_table_row).Interior.ColorIndex = 4
              End If
           End If
                    
          'Prepare for the next ticker symbol: Increment the summary table row count, reset the total stock volume to zero and
          'Save the yearly opening price for the new ticker sumbol
          summary_table_row = summary_table_row + 1
          current_volume = 0
          next_open = ws.Cells(i + 1, 3).Value
    
        ' If the cell immediately following a row is the same ticker symbol...
        ' Add to the stock volume running totals for the current ticker symbol
        Else
          current_volume = current_volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    'When there are no more rows to process on the worksheet, determine and write out the significant stock indicators
    Call ExtremeStocks(summary_table_row, ws)
    
    'Autofit columns & output a message indicating that all rows on the sheet have been processed
    ws.Columns("I:Q").AutoFit
    'MsgBox ("You have processed all " & (last_row_sheet - 1) & "rows on this sheet")

Next ws


End Sub

Sub ExtremeStocks(summary_table_row, ws)

'Initialize local variables
Dim max_increase As Single
Dim max_decrease As Single
Dim max_volume As Double
Dim max_increase_ticker As String
Dim max_decrease_ticker As String
Dim max_volume_ticker As String

' Initialize variables
max_increase = 0
max_decrease = 0
max_volume = 0
summary_table_row = summary_table_row - 1

' Output a message with the total number of rows in the summary row table
'MsgBox ("The number of Summary table rows is:" & summary_table_row)

'Find the stocks with the maximum increase, maximum decrease and maximum volume for the year
For i = 2 To (summary_table_row - 1)

     If ws.Cells(i, 10) > max_volume Then
       max_volume = ws.Cells(i, 10)
       max_volume_ticker = ws.Cells(i, 9)
     End If
         
    If ws.Cells(i, 12) > max_increase Then
       max_increase = ws.Cells(i, 12)
       max_increase_ticker = ws.Cells(i, 9)
    Else
       If ws.Cells(i, 12) < max_decrease Then
       max_decrease = ws.Cells(i, 12)
       max_decrease_ticker = ws.Cells(i, 9)
       End If
    End If
Next i

'Write these to a separate section of the worksheet

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("P2").Value = max_increase_ticker
ws.Range("Q2").Value = max_increase

ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("P3").Value = max_decrease_ticker
ws.Range("Q3").Value = max_decrease

ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P4").Value = max_volume_ticker
ws.Range("Q4").Value = max_volume
      
ws.Range("Q2:Q3").NumberFormat = "0.00%"

End Sub


