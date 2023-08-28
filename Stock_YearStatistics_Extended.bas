Attribute VB_Name = "Module1"
Sub StockYearStatisticsExpanded():
'expanded option with less assumptions that data are sorted, etc. Plus bonus part with merged sheet

    Dim active_worksheet As Worksheet
    Dim merged_sheet As Worksheet
    'define variables for knowing how many rows and columns are populated. There could be a lot, so Long type
    Dim populated_rows_count As Long
    Dim populated_columns_count As Long
    'define variales for the placement of source data columns. There few of them, so integer type
    Dim ticker_column_number As Integer
    Dim date_column_number As Integer
    Dim open_column_number As Integer
    Dim high_column_number As Integer
    Dim low_column_number As Integer
    Dim close_column_number As Integer
    Dim vol_column_number As Integer
    'define variable for earliest and latest date within year
    Dim earliest_date_as_number As Long
    Dim latest_date_as_number As Long
    'define data to keep separate list of unique tickers and placement of newly created table data
    Dim Unique_Tickers_count As Long
    Dim ticker_unique_column_number As Integer
    Dim Yearly_Change_column_number As Integer
    Dim Percentage_Change_column_number As Integer
    Dim Total_Stock_Volume_column_number As Integer
    Dim Column_with_headers_column_number As Integer
    Dim Ticker_max_column_number As Integer
    Dim Value_max_column_number As Integer
    'define temporary data for keeping price, volume, % in-between calculations
    Dim temp_open_price As Currency
    Dim temp_close_price As Currency
    Dim volume_sum As Double
    Dim temp_start As Long
    Dim max_increase_row As Integer
    Dim max_decrease_row As Integer
    Dim max_volume_row As Integer
    ' Create a Variable to Hold  Last Row, and Year
        Dim last_Row As Long
        Dim last_Row_Year As Long
        Dim year_Name As String
        Dim tickers_Array() As String
        Dim years_Array() As Integer
        Dim volume_Array() As Double
        Dim tickers_count As Long
        Dim years_count As Integer
        Dim volume As Long
        Dim is_unique As Boolean
    'define array to keep starting point of each ticker
    Dim Tickers_end_positions_Array() As Long

    'delete one "Merged" worksheet that was created before
    For Each active_worksheet In Worksheets
        If active_worksheet.Name = "Merged" Then
        active_worksheet.Delete
        End If
    Next active_worksheet
        
    For Each active_worksheet In Worksheets
    
    'delete content populated with this procedure last time
    active_worksheet.Columns("I:R").Delete Shift:=xlToLeft
    
    'let us set all the data placements' related variables
    'let us find number of populated rows for the current sheet; source:https://www.thespreadsheetguru.com/last-row-column-vba/
    populated_rows_count = active_worksheet.Cells(active_worksheet.Rows.Count, 1).End(xlUp).Row
    'let us find number of populated columns for the current sheet; source:https://www.thespreadsheetguru.com/last-row-column-vba/
    populated_columns_count = active_worksheet.Cells(1, active_worksheet.Columns.Count).End(xlToLeft).Column
    
    ' find the placement of source data columns
    For i = 1 To populated_columns_count
      If active_worksheet.Cells(1, i).Value = "<ticker>" Then
         ticker_column_number = i
       ElseIf active_worksheet.Cells(1, i).Value = "<date>" Then
        date_column_number = i
       ElseIf active_worksheet.Cells(1, i).Value = "<open>" Then
         open_column_number = i
       ElseIf active_worksheet.Cells(1, i).Value = "<low>" Then
           low_column_number = i
       ElseIf active_worksheet.Cells(1, i).Value = "<close>" Then
          close_column_number = i
       ElseIf active_worksheet.Cells(1, i).Value = "<vol>" Then
          vol_column_number = i
      End If
    Next i
    
'populate headers for new table within the worksheet
        ticker_unique_column_number = populated_columns_count + 2
        active_worksheet.Cells(1, ticker_unique_column_number).Value = "Ticker"
        Yearly_Change_column_number = populated_columns_count + 3
        active_worksheet.Cells(1, Yearly_Change_column_number).Value = "Yearly Change"
        Percentage_Change_column_number = populated_columns_count + 4
        active_worksheet.Cells(1, Percentage_Change_column_number).Value = "Percentage Change"
        Total_Stock_Volume_column_number = populated_columns_count + 5
        active_worksheet.Cells(1, Total_Stock_Volume_column_number).Value = "Total Stock Volume"
        Column_with_headers_column_number = populated_columns_count + 8
        active_worksheet.Cells(2, Column_with_headers_column_number).Value = "Greatest % increase"
        active_worksheet.Cells(3, Column_with_headers_column_number).Value = "Greatest % decrease"
        active_worksheet.Cells(4, Column_with_headers_column_number).Value = "Greatest total volume"
        Ticker_max_column_number = populated_columns_count + 9
        active_worksheet.Cells(1, Ticker_max_column_number).Value = "Ticker"
        Value_max_column_number = populated_columns_count + 10
        active_worksheet.Cells(1, Value_max_column_number).Value = "Value"

 'get list of tickers in separate column & list of ticker starting points assuming that all the tickers are sorted by <ticker field>
 Unique_Tickers_count = 1
 active_worksheet.Cells(Unique_Tickers_count + 1, ticker_unique_column_number).Value = active_worksheet.Cells(2, ticker_column_number).Value
 
    For i = 3 To populated_rows_count
             If active_worksheet.Cells(i, ticker_column_number).Value <> active_worksheet.Cells(Unique_Tickers_count + 1, ticker_unique_column_number).Value Then
                Unique_Tickers_count = Unique_Tickers_count + 1
                active_worksheet.Cells(Unique_Tickers_count + 1, ticker_unique_column_number).Value = active_worksheet.Cells(i, ticker_column_number).Value
                ReDim Preserve Tickers_end_positions_Array(Unique_Tickers_count - 1)
                Tickers_end_positions_Array(Unique_Tickers_count - 1) = i - 1
              End If
    Next i
    'assign position for a last ticker
                ReDim Preserve Tickers_end_positions_Array(Unique_Tickers_count)
                Tickers_end_positions_Array(Unique_Tickers_count) = populated_rows_count

'find the row with earliest and latest date of the year per ticker, their price and the difference
    'set default values, for very first ticker
    earliest_date_as_number = active_worksheet.Cells(2, date_column_number).Value
    latest_date_as_number = active_worksheet.Cells(2, date_column_number).Value
    temp_open_price = active_worksheet.Cells(2, open_column_number).Value
    temp_close_price = active_worksheet.Cells(2, close_column_number).Value
    volume_sum = active_worksheet.Cells(2, vol_column_number).Value
    temp_start = 2
    
    'per each ticker
    For j = 1 To Unique_Tickers_count
    'look for latest and earliest date of the year within this ticker range and setting respective prices
          If temp_start + 1 <= Tickers_end_positions_Array(j) Then  'just in case there is only one record fora ticker we skip comparison
            For i = temp_start + 1 To (Tickers_end_positions_Array(j))
              volume_sum = volume_sum + active_worksheet.Cells(i, vol_column_number).Value  'calculate the total stock volume of the stock.
                If (active_worksheet.Cells(i, date_column_number).Value > latest_date_as_number) Then
                latest_date_as_number = active_worksheet.Cells(i, date_column_number).Value
                temp_close_price = active_worksheet.Cells(i, close_column_number).Value
                End If
                 If (active_worksheet.Cells(i, date_column_number).Value < earliest_date_as_number) Then
                 earliest_date_as_number = active_worksheet.Cells(i, date_column_number).Value
                 temp_open_price = active_worksheet.Cells(i, open_column_number).Value
                 End If
            Next i
          End If
         'calculate the Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year. And print it
         active_worksheet.Cells(j + 1, Yearly_Change_column_number).Value = temp_close_price - temp_open_price
         'calculate the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year. And print it
         active_worksheet.Cells(j + 1, Percentage_Change_column_number).Value = active_worksheet.Cells(j + 1, Yearly_Change_column_number).Value / temp_open_price    '(temp_close_price - temp_open_price) / temp_open_price
         'print the total stock volume of the stock.
         active_worksheet.Cells(j + 1, Total_Stock_Volume_column_number).Value = volume_sum
         
         'prep default for the next ticker
          If j <> Unique_Tickers_count Then
             temp_start = Tickers_end_positions_Array(j) + 1
             earliest_date_as_number = active_worksheet.Cells(temp_start, date_column_number).Value
             latest_date_as_number = active_worksheet.Cells(temp_start, date_column_number).Value
             temp_open_price = active_worksheet.Cells(temp_start, open_column_number).Value
             temp_close_price = active_worksheet.Cells(temp_start, close_column_number).Value
             volume_sum = active_worksheet.Cells(temp_start, vol_column_number).Value
          End If
    Next j
    
      'change the format of data for Yearly_Change
        active_worksheet.Columns("J:J").EntireColumn.NumberFormat = "0.00"
      'change the format of data for Percentage_Change
        active_worksheet.Columns("K:K").EntireColumn.NumberFormat = "0.00%"
    
    'color the cell depending on the value
    For i = 2 To (Unique_Tickers_count + 1)
        For j = Yearly_Change_column_number To Percentage_Change_column_number
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
         If active_worksheet.Cells(i, Total_Stock_Volume_column_number).Value > active_worksheet.Cells(max_volume_row, Total_Stock_Volume_column_number).Value Then
          max_volume_row = i
         End If
    
    'look for the row with Greatest % increase and Greatest % decrease
    If active_worksheet.Cells(i, Percentage_Change_column_number).Value >= 0 Then
             If active_worksheet.Cells(i, Percentage_Change_column_number).Value > active_worksheet.Cells(max_increase_row, Percentage_Change_column_number).Value Then
             max_increase_row = i
            End If
    ElseIf active_worksheet.Cells(i, Percentage_Change_column_number).Value < 0 Then
        If active_worksheet.Cells(i, Percentage_Change_column_number).Value < active_worksheet.Cells(max_decrease_row, Percentage_Change_column_number).Value Then
        max_decrease_row = i
        End If
    End If
    
    Next i
    
    'Print the Greatest % increase, Greatest % decrease and Greatest total volume
    active_worksheet.Cells(2, Ticker_max_column_number) = active_worksheet.Cells(max_increase_row, ticker_unique_column_number) 'ticker with max increase
    active_worksheet.Cells(2, Value_max_column_number) = active_worksheet.Cells(max_increase_row, Percentage_Change_column_number) 'max increase value
    active_worksheet.Cells(3, Ticker_max_column_number) = active_worksheet.Cells(max_decrease_row, ticker_unique_column_number) 'ticker with max decrease
    active_worksheet.Cells(3, Value_max_column_number) = active_worksheet.Cells(max_decrease_row, Percentage_Change_column_number) 'max decrease value
    active_worksheet.Cells(4, Ticker_max_column_number) = active_worksheet.Cells(max_volume_row, ticker_unique_column_number) 'ticker with max volume
    active_worksheet.Cells(4, Value_max_column_number) = active_worksheet.Cells(max_volume_row, Total_Stock_Volume_column_number) 'max volume value
    
     'change the format of data for "Value" column
     active_worksheet.Range("Q2:Q3").NumberFormat = "0.00%"
     active_worksheet.Range("Q4").NumberFormat = "0.00E+00"
    
    'autofit cells
    active_worksheet.Columns("A:R").AutoFit
    
    Next active_worksheet
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 'Extra experiment to create merged tab and grouped table
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      
        
        'Add a sheet named "Merged"
        Sheets.Add.Name = "Merged"
        'move created sheet to the very end
        Sheets("Merged").Move After:=Sheets(Sheets.Count)
        'Specify the location of the combined sheet
        Set merged_sheet = Worksheets("Merged")
        
        'Populate headers in a new worksheet
        merged_sheet.Range("A1").Value = "Year"
        'Copy the headers from sheet 1
        merged_sheet.Range("B1:E1").Value = Sheets(1).Range("I1:L1").Value
    
        'Loop through all sheets
        For Each active_worksheet In Worksheets
         If active_worksheet.Name <> "Merged" Then
          'Find the last row of the combined sheet after each paste
          'Add 1 to get first empty row
         lastRow = merged_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
         'Find the last row of each worksheet
         'Subtract one to return the number of rows without header
         lastRowYear = active_worksheet.Cells(Rows.Count, "L").End(xlUp).Row - 1
        
         'populate first column with Year value
         merged_sheet.Range("A" & lastRow & ":A" & ((lastRowYear - 1) + lastRow)).Value = active_worksheet.Name
        
         'Copy the contents of each year sheet into the combined sheet
         merged_sheet.Range("B" & lastRow & ":E" & ((lastRowYear - 1) + lastRow)).Value = active_worksheet.Range("I2:L" & (lastRowYear + 1)).Value
         End If
        Next active_worksheet

    '''''''''''''''''''''''''''' ''''''''''''''''''''''''''''  '''''''''''''''''''''''''''' ''''''''''''''''''''''''''' ''''''''''''''''''''''''''''  '''''''''''''''''''''''''''' ''''''''''''''''''''''''''' ''''''''''''''''''''''''''''  ''''''''''''''''''''''''''''
    'Creating grouped table\/
    '''''''''''''''''''''''''''' ''''''''''''''''''''''''''''  '''''''''''''''''''''''''''' ''''''''''''''''''''''''''' ''''''''''''''''''''''''''''  '''''''''''''''''''''''''''' ''''''''''''''''''''''''''' ''''''''''''''''''''''''''''  ''''''''''''''''''''''''''''
    
   last_Row_Year = merged_sheet.Cells(Rows.Count, "A").End(xlUp).Row
        
    'find all unique tickers and years
    'set defaults
    tickers_count = 1
    years_count = 1
    ReDim Preserve tickers_Array(tickers_count)
    tickers_Array(tickers_count) = merged_sheet.Cells(2, 2).Value
    ReDim Preserve years_Array(years_count)
    years_Array(years_count) = merged_sheet.Cells(2, 1).Value
    'find the rest: tickers
    For i = 3 To last_Row_Year
        is_unique = True
        For j = 1 To tickers_count
            If merged_sheet.Cells(i, 2).Value = tickers_Array(j) Then
            is_unique = False
            End If
        Next j
        If is_unique Then
        tickers_count = tickers_count + 1
        ReDim Preserve tickers_Array(tickers_count)
        tickers_Array(tickers_count) = merged_sheet.Cells(i, 2).Value
        End If
    Next i
    'find the rest: years
    For i = 3 To last_Row_Year
        is_unique = True
        For j = 1 To years_count
            If merged_sheet.Cells(i, 1).Value = years_Array(j) Then
            is_unique = False
            End If
        Next j
        If is_unique Then
                    years_count = years_count + 1
                     ReDim Preserve years_Array(years_count)
                     years_Array(years_count) = merged_sheet.Cells(i, 1).Value
        End If
    Next i
    
     'find a volume
         For i = 1 To tickers_count
            ReDim Preserve volume_Array(i)
            volume_Array(i) = 0
             For j = 2 To last_Row_Year
                If merged_sheet.Cells(j, 2).Value = tickers_Array(i) Then
                volume_Array(i) = volume_Array(i) + merged_sheet.Cells(j, 5).Value
                End If
             Next j
          Next i

     'Populate headers and tickers for grouped table in a new worksheet, assuming that years are sorted in original table
        merged_sheet.Range("H1").Value = "Ticker"
        merged_sheet.Range("I1").Value = "Volume"
        For i = 2 To (tickers_count + 1)
        merged_sheet.Cells(i, 8).Value = tickers_Array(i - 1)
        merged_sheet.Cells(i, 9).Value = volume_Array(i - 1)
        Next i
        
        For i = 1 To years_count
        merged_sheet.Cells(1, 9 + i).Value = years_Array(i)
        Next i
        
           'populate the rest
        For i = 2 To last_Row_Year
            For j = 1 To tickers_count
                For k = 1 To years_count
                    If merged_sheet.Cells(i, 2).Value = tickers_Array(j) And merged_sheet.Cells(i, 1).Value = years_Array(k) Then
                    merged_sheet.Cells(j + 1, 9 + k).Value = merged_sheet.Cells(i, 4).Value
                    End If
                Next k
            Next j
        Next i
        
    'color the cell depending on the value
    For i = 2 To tickers_count + 1
      For j = 10 To 9 + years_count
        If merged_sheet.Cells(i, j).Value > 0 Then
        merged_sheet.Cells(i, j).Interior.ColorIndex = 4
        ElseIf merged_sheet.Cells(i, j).Value < 0 Then
        merged_sheet.Cells(i, j).Interior.ColorIndex = 3
        ElseIf merged_sheet.Cells(i, j).Value = 0 Then
        merged_sheet.Cells(i, j).Interior.ColorIndex = 6
        End If
        Next j
    Next i
    
    'sort  by volume
   Range("H1:L3001").Sort Key1:=Range("I1"), _
                     Order1:=xlDescending, _
                     Header:=xlYes
    'Autofit to display data
    merged_sheet.Columns("A:D").AutoFit
        
End Sub






