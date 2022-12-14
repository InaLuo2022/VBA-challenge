'The VBA Script shows the Max and Min value of the difference each year in which ticker symbols.
'Also demostrate the ticker symbol award the maximum total stock volume each year.

Sub Stock_Greatest()

'Run the script in each worksheet
For Each ws In Worksheets
    
    'Find the end row of the worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    With ws
        'set columns of Ticker and Value
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Range("Q1").ColumnWidth = 20
        .Range("O1").ColumnWidth = 20
        
        'set rows for Max and Min % increase and Max total stock volume
        .Cells(2, 15).Value = "Greatest % Increase"
        .Cells(3, 15).Value = "Greatest % Decrease"
        .Cells(4, 15).Value = "Greatest Total Volume"
        
        'find Max and Min % increase and Max total stock volume
        .Cells(2, 17).Value = FormatPercent(Application.WorksheetFunction.Max(.Range(.Cells(2, 11), .Cells(lastrow, 11))))
        .Cells(3, 17).Value = FormatPercent(Application.WorksheetFunction.Min(.Range(.Cells(2, 11), .Cells(lastrow, 11))))
        .Cells(4, 17).Value = Application.WorksheetFunction.Max(.Range(.Cells(2, 12), .Cells(lastrow, 12)))
    
        'Loop to find the ticker symbol for Max and Min % increase and Max total stock volume
        Dim Ticker_Name As Long
    
        For Ticker_Name = 2 To lastrow
    
            If .Cells(Ticker_Name, 11).Value = .Cells(2, 17).Value Then
    
                .Cells(2, 16).Value = .Cells(Ticker_Name, 9).Value
        
            ElseIf .Cells(Ticker_Name, 11).Value = .Cells(3, 17).Value Then
    
                .Cells(3, 16).Value = .Cells(Ticker_Name, 9).Value
        
            ElseIf .Cells(Ticker_Name, 12).Value = .Cells(4, 17).Value Then
    
                .Cells(4, 16).Value = .Cells(Ticker_Name, 9).Value
        
            End If
    
        Next

    End With
    
Next

End Sub

