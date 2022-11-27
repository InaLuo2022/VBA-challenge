' The module is to analyze:
    '  1. the yearly change at the beginning of a given year to the closing price at the end of that year.
    '  2. the total stock volume of the stock.
    '  3. Highlight positive change by using conditional formatting

Sub Stock_Performance()

'Analyze stock performance each year
For Each ws In Worksheets
 
  'Set columns for Ticker, Yearly Change, Percent Change and Total Stock Volume
  ws.Range("j1").ColumnWidth = 15
  ws.Range("k1").ColumnWidth = 15
  ws.Range("l1").ColumnWidth = 20
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
  ' Set Ticker Name as a variable
    Dim Ticker_Name As Integer
    Ticker_Name = 2
    
    'Find last row each worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set open price & close price each year
    Dim Open_Price As Long
    Dim Close_Price As Long
    Dim RowCount As Integer
    
    Open_Price = 2
    Close_Price = 2
    RowCount = 0
    
    ws.Cells(Ticker_Name, 12).Value = 0
    
    For Close_Price = 2 To lastrow
        
        RowCount = Close_Price - Open_Price
        
        If ws.Cells(Open_Price, 1).Value = ws.Cells(Close_Price, 1).Value Then
            
            ws.Cells(Ticker_Name, 9).Value = ws.Cells(Open_Price, 1).Value
            ws.Cells(Ticker_Name, 10).Value = ws.Cells(Open_Price + RowCount, 6).Value - ws.Cells(Open_Price, 3).Value
            ws.Cells(Ticker_Name, 11).Value = FormatPercent(ws.Cells(Ticker_Name, 10).Value / ws.Cells(Open_Price, 3).Value)
            ws.Cells(Ticker_Name, 12).Value = ws.Cells(Open_Price + RowCount, 7).Value + ws.Cells(Ticker_Name, 12).Value
            
        Else
            Ticker_Name = Ticker_Name + 1
            Open_Price = Open_Price + RowCount
            ws.Cells(Ticker_Name, 12).Value = ws.Cells(Close_Price, 7).Value + ws.Cells(Ticker_Name, 12).Value
            
        End If
    Next Close_Price
Next


'Conditional Formatting to highlight Yearly change
For Each ws In Worksheets
lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

Dim Ticker_Categorized As Integer
    
    For Ticker_Categorized = 2 To lastrow
    
        If ws.Cells(Ticker_Categorized, 10).Value < 0 Or ws.Cells(Ticker_Categorized, 10).Value = 0 Then
            
            ws.Cells(Ticker_Categorized, 10).Interior.ColorIndex = 3
                
        Else
                
            ws.Cells(Ticker_Categorized, 10).Interior.ColorIndex = 4
                
        End If
    
    Next
    
Next

End Sub




