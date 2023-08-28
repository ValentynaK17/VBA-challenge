Attribute VB_Name = "Module1"
Sub StockYearStatistics_SimplifiedVersion():

'this script is created with the following assumptions:
' - ticker labels are sorted alphabetically (primary sorting)
' - data are sorted by date as a secondary sorting
' - all year dates are listed per each ticker and the earliest date for the year for which we have data is always 01 January and the latest is 31 December
' - for each tab we have same rows listed in the same order

    Dim active_worksheet As Worksheet
    'define variables for knowing how many rows and columns are populated. There could be a lot, so Long type
    Dim populated_rows_count As Long
    Dim populated_columns_count As Long
     'define data to keep separate list of unique tickers and placement of newly created table data
    Dim Unique_Tickers_count As Long
    'define temporary data for keeping price, volume, % in-between calculations
    Dim temp_open_price As Currency
    Dim temp_close_price As Currency
    Dim volume_sum As Double
    Dim max_increase_row As Integer
    Dim max_decrease_row As Integer
    Dim max_volume_row As Integer
    
    'let us start our switch between sheets one by one
    For Each active_worksheet In Worksheets
    
    'delete content populated with this procedure last time
    active_worksheet.Columns("I:R").Delete Shift:=xlToLeft
    
     'let us set all the data placements' related variables
    ' let us find number of populated rows for the current sheet; source:https://www.thespreadsheetguru.com/last-row-column-vba/
    populated_rows_count = active_worksheet.Cells(active_worksheet.Rows.Count, 1).End(xlUp).Row
    ' let us find number of populated columns for the current sheet; source:https://www.thespreadsheetguru.com/last-row-column-vba/
    populated_columns_count = active_worksheet.Cells(1, active_worksheet.Columns.Count).End(xlToLeft).Column


    'populate headers for new table within the worksheet
        active_worksheet.Cells(1, 9).Value = "Ticker"
        active_worksheet.Cells(1, 10).Value = "Yearly Change"
        active_worksheet.Cells(1, 11).Value = "Percentage Change"
        active_worksheet.Cells(1, 12).Value = "Total Stock Volume"
        active_worksheet.Cells(2, 15).Value = "Greatest % increase"
        active_worksheet.Cells(3, 15).Value = "Greatest % decrease"
        active_worksheet.Cells(4, 15).Value = "Greatest total volume"
        active_worksheet.Cells(1, 16).Value = "Ticker"
        active_worksheet.Cells(1, 17).Value = "Value"
        

    'get list of tickers in separate column & list of ticker starting points assuming that all the tickers are sorted by <ticker field>
    'set defaults
    Unique_Tickers_count = 0
    volume_sum = 0
    temp_open_price = active_worksheet.Cells(2, 3).Value

    For i = 2 To populated_rows_count
     volume_sum = volume_sum + active_worksheet.Cells(i, 7).Value  'calculate the total stock volume of the stock
             If active_worksheet.Cells(i, 1).Value <> active_worksheet.Cells(i + 1, 1).Value Then
                Unique_Tickers_count = Unique_Tickers_count + 1
                active_worksheet.Cells(Unique_Tickers_count + 1, 9).Value = active_worksheet.Cells(i, 1).Value
                temp_close_price = active_worksheet.Cells(i, 6).Value
                 'calculate the Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year. And print it
                 active_worksheet.Cells(Unique_Tickers_count + 1, 10).Value = temp_close_price - temp_open_price
                'calculate the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year. And print it
                 active_worksheet.Cells(Unique_Tickers_count + 1, 11).Value = active_worksheet.Cells(Unique_Tickers_count + 1, 10).Value / temp_open_price
                'print the total stock volume of the stock
                 active_worksheet.Cells(Unique_Tickers_count + 1, 12).Value = volume_sum
                volume_sum = 0
                temp_open_price = active_worksheet.Cells(i + 1, 3).Value
              End If
    Next i
    
     'change the format of data for Yearly_Change
        active_worksheet.Columns("J:J").EntireColumn.NumberFormat = "0.00"
      'change the format of data for Percentage_Change
        active_worksheet.Columns("K:K").EntireColumn.NumberFormat = "0.00%"
    
    'color the cell depending on the value
    For i = 2 To (Unique_Tickers_count + 1)
        For j = 10 To 11
        If active_worksheet.Cells(i, j).Value > 0 Then
        active_worksheet.Cells(i, j).Interior.ColorIndex = 4
        ElseIf active_worksheet.Cells(i, j).Value < 0 Then
        active_worksheet.Cells(i, j).Interior.ColorIndex = 3
        ElseIf active_worksheet.Cells(i, j).Value = 0 Then
        active_worksheet.Cells(i, j).Interior.ColorIndex = 6
        End If
        Next j
    Next i

 
 'find the Greatest % increase, Greatest % decrease and Greatest total volume
    'set default values
    max_increase_row = 2
    max_decrease_row = 2
    max_volume_row = 2
     
    For i = 3 To Unique_Tickers_count + 1
    'look for the row with greatest volume
         If active_worksheet.Cells(i, 12).Value > active_worksheet.Cells(max_volume_row, 12).Value Then
          max_volume_row = i
         End If
    
    'look for the row with Greatest % increase and Greatest % decrease
    If active_worksheet.Cells(i, 11).Value >= 0 Then
             If active_worksheet.Cells(i, 11).Value > active_worksheet.Cells(max_increase_row, 11).Value Then
             max_increase_row = i
            End If
    ElseIf active_worksheet.Cells(i, 11).Value < 0 Then
        If active_worksheet.Cells(i, 11).Value < active_worksheet.Cells(max_decrease_row, 11).Value Then
        max_decrease_row = i
        End If
    End If
    
    Next i
    
    'Print the Greatest % increase, Greatest % decrease and Greatest total volume
    active_worksheet.Cells(2, 16) = active_worksheet.Cells(max_increase_row, 9)  'ticker with max price increase
    active_worksheet.Cells(2, 17) = active_worksheet.Cells(max_increase_row, 11)  'max price increase value
    active_worksheet.Cells(3, 16) = active_worksheet.Cells(max_decrease_row, 9) 'ticker with max price decrease
    active_worksheet.Cells(3, 17) = active_worksheet.Cells(max_decrease_row, 11) 'max decrease price value
    active_worksheet.Cells(4, 16) = active_worksheet.Cells(max_volume_row, 9) 'ticker with max volume
    active_worksheet.Cells(4, 17) = active_worksheet.Cells(max_volume_row, 12) 'max volume value
    
     'change the format of data for "Value" column
     active_worksheet.Range("Q2:Q3").NumberFormat = "0.00%"
     active_worksheet.Range("Q4").NumberFormat = "0.00E+00"
    
    'autofit cells
    active_worksheet.Columns("A:R").AutoFit

    
    Next active_worksheet


End Sub

