Sub StockData():
    
    'Loop through all worksheets
    For Each ws In Worksheets
        'Delcaring variables
        Dim WorksheetName As String
        Dim i As Double
        Dim j As Double
        Dim TickerCount As Double
        Dim LastRow As Double
        Dim FirstRow As Double
        Dim PerChange As Double
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatVol As Double
        
        'Creating WS output column titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Calling worksheet
        WorksheetName = ws.Name
          
        'Set to start at the first column, second row
        TickerCount = 2
        j = 2
        
        'Locate the last row of column A
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For i = 2 To LastRow
            
                'Check to see if the <ticker> values have changed then write to column I and calculate Yearly Change to column J
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'If Yearly Change value is less than 0, set to color red, if greater than 0, set to color green
                    If ws.Cells(TickerCount, 10).Value < 0 Then
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                    Else
                    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change to column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    'Formatting Percentage Change column K to a percentage
                    ws.Cells(TickerCount, 11).Value = Format(PerChange, "Percent")
                    
                    End If
                    
                'Calculate and write Total Volume to column L
                ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickerCount by 1 and start next row
                TickerCount = TickerCount + 1
                j = i + 1
                
                End If
            
            Next i
        
        'Define summary table values
        GreatVol = ws.Cells(2, 12).Value
        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value
        
        'Locate the last row of column I
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
            For i = 2 To FirstRow
            
                'Locate the Greatest Increase value and write to column Q, row 2
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                End If
                
                'Locate the Greatest Decrease value and write to colum Q, row 3
                If ws.Cells(i, 11).Value > GreatInc Then
                GreatInc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                'Formatting Great Increase to a percentage
                ws.Cells(2, 17).Value = Format(GreatInc, "Percent")
                
                End If
                
                'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatDec Then
                GreatDec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                'Formatting Great Decrease to a percentage
                ws.Cells(3, 17).Value = Format(GreatDec, "Percent")
                
                End If
            
            Next i
            
        'Autofit all columns in the worksheet
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub
