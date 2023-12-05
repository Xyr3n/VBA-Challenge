Attribute VB_Name = "Module1"
Sub stock_data()
  'define variables
    Dim ws As Worksheet
    
    Dim total_vol As LongLong
    Dim yearly_change As Double
    Dim open_price As Double
    Dim percent_change As Double
    
    Dim row_count As Integer
    Dim ticker_col As Integer
    Dim yc_col As Integer
    Dim pc_col As Integer
    Dim vol_col As Integer
    
    Dim max_inc As Variant
    Dim max_dec As Variant
    Dim max_vol As LongLong
    
    Dim rg As Range
    Dim condpositive As FormatCondition
    Dim condnegative As FormatCondition
    
    'loop through sheets
     For Each ws In ThisWorkbook.Sheets
    
        'set up
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease "
        ws.Range("O4") = "Greatest Total Volume"
        
        'Set column widths
        ws.Columns("I").ColumnWidth = 10
        ws.Columns("J:K").ColumnWidth = 15
        ws.Columns("L").ColumnWidth = 20
        ws.Columns("O").ColumnWidth = 20
        ws.Columns("P:Q").ColumnWidth = 10
        
        'initialize
        row_count = 2
        total_vol = ws.Cells(2, 7).Value
        open_price = ws.Cells(2, 3).Value
        ticker_col = 9
        yc_col = 10
        pc_col = 11
        vol_col = 12
        last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
      
        'loop through data
        For i = 3 To last_row
           'conditional
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
               'calculate
                yearly_change = ws.Cells(i - 1, 6).Value - open_price
                percent_change = yearly_change / open_price
                
                'print values
                ws.Cells(row_count, ticker_col).Value = ws.Cells(i - 1, 1).Value
                ws.Cells(row_count, yc_col).Value = Format(yearly_change, "0.00")
                ws.Cells(row_count, pc_col).Value = FormatPercent(percent_change)
                ws.Cells(row_count, vol_col).Value = total_vol
                
                'update
                row_count = row_count + 1
                total_vol = ws.Cells(i, 7).Value
                open_price = ws.Cells(i, 3).Value
                
            Else
                total_vol = total_vol + ws.Cells(i, 7).Value
            
            End If
        Next i
        
        
        'initialize
        last_row_new = ws.Cells(Rows.count, 9).End(xlUp).Row
        max_inc = ws.Cells(2, 11).Value
        max_dec = ws.Cells(2, 11).Value
        max_vol = ws.Cells(2, 12).Value
        
       'loop through percent change to find the max by comparing each number
        For i = 2 To last_row_new
            If ws.Cells(i + 1, 11).Value > max_inc Then
            max_inc = ws.Cells(i + 1, 11).Value
            ticker = ws.Cells(i + 1, 9).Value
            
            Else
            ws.Range("P2") = ticker
            ws.Range("Q2") = FormatPercent(max_inc)
            
            End If
        Next i
        
         'loop through percent change to find the min
        For i = 2 To last_row_new
            If ws.Cells(i + 1, 11).Value < max_dec Then
            max_dec = ws.Cells(i + 1, 11).Value
            ticker = ws.Cells(i + 1, 9).Value
            
            Else
            ws.Range("P3") = ticker
            ws.Range("Q3") = FormatPercent(max_dec)
            
            End If
        Next i
        
        'loop through total stock volume to find the max
        For i = 2 To last_row_new
            If ws.Cells(i + 1, 12).Value > max_vol Then
            max_vol = ws.Cells(i + 1, 12).Value
            ticker = ws.Cells(i + 1, 9).Value
            
            Else
            ws.Range("P4") = ticker
            ws.Range("Q4") = max_vol
            
            End If
        Next i
    
    
    'conditional formatting
        Set rg = ws.Range(ws.Cells(2, 10), ws.Cells(last_row_new, 11))
        rg.FormatConditions.Delete
        Set condpositive = rg.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set condnegative = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")
   
        With condpositive
        .Interior.Color = vbGreen
        .Font.Color = vbBlack
        End With
        
        With condnegative
        .Interior.Color = vbRed
        .Font.Color = vbBlack
        End With

    Next ws
End Sub

