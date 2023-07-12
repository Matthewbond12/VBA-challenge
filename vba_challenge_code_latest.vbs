Attribute VB_Name = "Module1"
Option Explicit
Sub runallsheets()
    
    Dim wks As Worksheet
    Application.ScreenUpdating = False
    For Each wks In Worksheets
        wks.Select
        Call stock_data
    Next
    Application.ScreenUpdating = True
    
End Sub
Sub stock_data()
   
    'Declare constants
    Const FIRST_DATA_ROW As Integer = 2
    Const IN_TICKER_COL As Integer = 1
    Const OUT_TICKER_COL As Integer = 10
    
    'Declare variables
    Dim ticker_symbol As String
    Dim next_ticker As String
    Dim input_row As Long
    Dim lastrow As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim change_ratio As Double
    Dim output_row As Long
    Dim volume As LongLong
   
    'Assign variables
    lastrow = Cells(Rows.Count, IN_TICKER_COL).End(xlUp).Row
    output_row = FIRST_DATA_ROW
    open_price = Cells(FIRST_DATA_ROW, IN_TICKER_COL + 2).Value
    volume = 0
        
    For input_row = FIRST_DATA_ROW To lastrow
        ticker_symbol = Cells(input_row, IN_TICKER_COL).Value
        next_ticker = Cells(input_row + 1, IN_TICKER_COL).Value
        ' Last row of current stock
        If ticker_symbol <> next_ticker Then
            'Input
            close_price = Cells(input_row, IN_TICKER_COL + 5).Value
            'Calculations
            yearly_change = close_price - open_price
            change_ratio = yearly_change / open_price
            volume = volume + Cells(input_row, IN_TICKER_COL + 6).Value
            'Output
             Cells(output_row, OUT_TICKER_COL).Value = ticker_symbol
             Cells(output_row, OUT_TICKER_COL + 1).Value = yearly_change
             Cells(output_row, OUT_TICKER_COL + 2).Value = FormatPercent(change_ratio)
             Cells(output_row, OUT_TICKER_COL + 3).Value = volume
             'Prepare for next stock
             output_row = output_row + 1
             open_price = Cells(input_row + 1, 3).Value
             'Reset volume count to zero when ticker symbol changes
             volume = 0
        Else
            volume = volume + Cells(input_row, IN_TICKER_COL + 6).Value
            Cells(output_row, OUT_TICKER_COL + 3).Value = volume
        End If
    
    Next input_row
    'Declare variables for conditional formatting
   
    Dim last_output_row As Long
    Dim conditional_formatting_column As Long
    'Assign and identify last row in summary output table to end conditional format if statement
    last_output_row = Cells(Rows.Count, 12).End(xlUp).Row
    ' Loop through first row of summary table through last row
    For output_row = FIRST_DATA_ROW To last_output_row
        'Nested loop through yearly change and percent change columns
        For conditional_formatting_column = OUT_TICKER_COL To OUT_TICKER_COL + 1
        'If row in yearly_change, percent_change are greater than 0 return green, less than red
        If Cells(output_row, conditional_formatting_column + 1).Value > 0 Then
            Cells(output_row, conditional_formatting_column + 1).Interior.ColorIndex = 4
        Else
            Cells(output_row, conditional_formatting_column + 1).Interior.ColorIndex = 3
        End If
    
        Next conditional_formatting_column
        
    Next output_row
    
    ' Declare variables for Max percent increase
    Dim new_row_max As Double
    Dim new_ticker_max As String
    'Assign variables for max percent increase
    new_ticker_max = Cells(input_row, OUT_TICKER_COL).Value
    new_row_max = Cells(input_row, OUT_TICKER_COL + 1).Value
    
    For input_row = FIRST_DATA_ROW To lastrow
                 
        ' Last row of current stock
        If Cells(input_row + 1, OUT_TICKER_COL + 2).Value > new_row_max Then
        new_row_max = Cells(input_row + 1, OUT_TICKER_COL + 2)
        new_ticker_max = Cells(input_row + 1, OUT_TICKER_COL)
        Range("P2").Value = new_ticker_max
        Range("Q2").Value = FormatPercent(new_row_max)
        End If
    
    Next input_row
    
    ' Declare variables for Max percent decrease
    Dim new_row_min As Double
    Dim new_ticker_min As String
    
    'Assign max percent variables
    new_row_min = Cells(input_row, OUT_TICKER_COL)
    new_ticker_min = Cells(input_row, OUT_TICKER_COL)
    
    For input_row = FIRST_DATA_ROW To lastrow
        ' If next row is less than prior row take that value as new min
        If Cells(input_row + 1, OUT_TICKER_COL + 2).Value < new_row_min Then
        new_row_min = Cells(input_row + 1, OUT_TICKER_COL + 2).Value
        new_ticker_min = Cells(input_row + 1, OUT_TICKER_COL).Value
        'Print values to summary table
        Range("P3").Value = new_ticker_min
        Range("Q3").Value = FormatPercent(new_row_min)
        End If
    
    Next input_row
    
    'Declare variables to retreive greatest total volume
    Dim new_ticker_volume As String
    Dim new_row_volume As LongLong
    
    'Assign variables
    new_ticker_volume = Cells(input_row, OUT_TICKER_COL).Value
    new_row_volume = Cells(input_row, OUT_TICKER_COL + 3).Value
    
    For input_row = FIRST_DATA_ROW To lastrow
        'If next row is greater than previous max, make that row new max
        If Cells(input_row + 1, OUT_TICKER_COL + 3).Value > new_row_volume Then
        new_row_volume = Cells(input_row + 1, OUT_TICKER_COL + 3).Value
        new_ticker_volume = Cells(input_row + 1, OUT_TICKER_COL).Value
        'Print new ticker and max volume to summary table
        Range("P4").Value = new_ticker_volume
        Range("Q4").Value = new_row_volume
        End If
    
    Next input_row
    
     ' Print column headers
    Range("J1, P1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
        
End Sub

