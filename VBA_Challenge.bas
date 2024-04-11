Attribute VB_Name = "Module1"
Sub Stockdata()
'Initial variable
Dim ws As Worksheet
Dim Ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim summary_table_Row As Double

'Run through the workbook

For Each ws In ThisWorkbook.Worksheets

'Set the ranges
ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

'Set the loop

summary_table_Row = 2

'Loop through all tickers
For i = 2 To ws.UsedRange.Rows.Count
   If i = 2 Then
    year_open = ws.Cells(i, 3).Value
    End If
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'Find the value
    Ticker = ws.Cells(i, 1).Value
    year_close = ws.Cells(i, 6).Value
    vol = vol + ws.Cells(i, 7).Value
    yearly_change = year_close - year_open
    percent_change = (year_close - year_open) / year_open
    'Input value
    ws.Cells(summary_table_Row, 10).Value = Ticker
    ws.Cells(summary_table_Row, 11).Value = yearly_change
    ws.Cells(summary_table_Row, 12).Value = percent_change
    ws.Cells(summary_table_Row, 13).Value = vol
    summary_table_Row = summary_table_Row + 1
    vol = 0
    year_close = 0
    year_open = ws.Cells(i + 1, 3).Value
    Else
    vol = vol + ws.Cells(i, 7).Value
    End If
    
'finnish the Loop
    Next i
    
    ws.Columns("L").NumberFormat = "0.00%"

'Format the color
    For j = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
    If ws.Cells(j, 11).Value < 0 Then
    ws.Cells(j, 11).Interior.ColorIndex = 3
    Else
    ws.Cells(j, 11).Interior.ColorIndex = 4
    End If
        Next j
    
'Find the greatest data
    
    Dim largest_stock, Smallest_stock, large_volume_Stock As String
    Dim Great_in, great_dc, grt_total As Double
    
'set Value
    Great_in = 0
    greatt_dc_ = 0
    grt_total = 0


'Find the loop

    For k = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
    If ws.Cells(k, 12).Value > Grt_in Then
    Grt_in = ws.Cells(k, 12).Value
        largest_stock = ws.Cells(k, 10).Value
        End If
    If ws.Cells(k, 12).Value < grt_dc Then
        grt_dc = ws.Cells(k, 12).Value
        Smallest_stock = ws.Cells(k, 10).Value
        End If
    If ws.Cells(k, 13).Value > grt_total Then
        grt_total = ws.Cells(k, 13).Value
        large_volume_Stock = ws.Cells(k, 10).Value
        End If
'Destination
    ws.Cells(2, 16) = largest_stock
    ws.Cells(3, 16) = Smallest_stock
    ws.Cells(4, 16) = large_volume_Stock
    
'Value
    ws.Cells(2, 17) = Grt_in
     ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17) = grt_dc
     ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17) = grt_total
    
Next k

    
    Next ws

End Sub


