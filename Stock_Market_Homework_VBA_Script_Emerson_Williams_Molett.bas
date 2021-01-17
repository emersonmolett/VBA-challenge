Attribute VB_Name = "Module1"
Sub ticker_name():

For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

    ' ticker symbol to appear in Col I
    
    Dim ticker_name As String
    
    ' location for each ticker in summary table
    
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    ' loop through all tickers
    
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastrow
    
    ' check if we are still within ticker name, if no
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    
    daily_change = Cells(i, 3).Value - Cells(i, 6).Value
     
    ticker_name = Cells(i, 1).Value
    
    ' add ticker_name_total
    
    ticker_name_total = ticker_name_total + Cells(i, 7).Value
    
    ' ticker symbol
    Range("I" & summary_table_row).Value = ticker_name
    
    summary_table_row = summary_table_row + 1
    
    ' reset ticker_name_total
    ticker_name_total = 0
    
    ' if the cell immediately following a row is the same ticker_name
    
    Else
    
    ' add to the ticker_name_total
    
        ticker_name_total = ticker_name_total + Cells(i, 7).Value
        
        
        Dim yearly_change As Double
        yearly_change = 0
    
        Dim open_price As Double
        ' open_price = Cells(2, 3).Value

        yearly_change = yearly_change + daily_change
    
        ' yearly change as a number
    
        Range("J" & summary_table_row).Value = yearly_change
        If yearly_change > 0 Then

            Range("J" & summary_table_row).Interior.ColorIndex = 4
       
        ElseIf yearly_change < 0 Then
            Range("J" & summary_table_row).Interior.ColorIndex = 3
        
            yearly_change = 0
        Else
            Range("J" & summary_table_row).Interior.ColorIndex = 0
        End If

    
        Dim percent_change As Double
        If Cells(i, 3).Value <> 0 Then
    
    
    
            percent_change = daily_change / Cells(i, 3).Value * 100
    
            Range("K" & summary_table_row).Value = percent_change
        End If
    
        ' stock vol
    
        Range("L" & summary_table_row).Value = ticker_name_total
    
 
    End If
    
    Next i
    
    
    

Next ws


End Sub



