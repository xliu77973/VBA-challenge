


Sub tickersymbol()
 
'create 2nd table in each spreadsheet(Q1, Q2, Q3, Q4)-------------------------------
 
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    ws.Activate
        
        'get the last row in the table
        last_row = Cells(Rows.Count, 1).End(xlUp).row
        
        'add headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percentage Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim quarter_change As Double
        Dim percentage_change As Double
        
        Dim volume As Double
        Dim row As Double
        Dim column As Integer
        
        volume = 0
        row = 2
        column = 1
        
        
        
        'set the initial price
        open_price = Cells(2, column + 2).Value
        
        'loop through all ticker to check for mismatch
        For i = 2 To last_row
        
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
         
                'set ticker name
                ticker = Cells(i, column).Value
                'insert ticker name
                Cells(row, column + 8).Value = ticker
            
                'set close price
                close_price = Cells(i, column + 5).Value
            
                'calculate quarterly change
                quarterly_change = close_price - open_price
                Cells(row, column + 9).Value = quarterly_change
            
                'calculate percentage change
                percentage_change = quarterly_change / open_price
                Cells(row, column + 10).Value = percentage_change
                Cells(row, column + 10).NumberFormat = "0.00%"
            
                'calculate total volume each quarter
                volume = volume + Cells(i, column + 6).Value
                Cells(row, column + 11).Value = volume
            
                'iterate to next row
                row = row + 1
            
                'reset open price to next ticker
                open_price = Cells(i + 1, column + 2)
            
                'reset volume for next ticker
                volume = 0
            
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i

'coloring cells----------------------------------------

        'find last row of ticker column
        quarterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        'coloring cells
        For j = 2 To quarterly_change_last_row
            If (Cells(j, 10).Value > 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf (Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 2
            ElseIf (Cells(j, 10) < 0) Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
'create 3rd table----------------------------------------

        'insert Greatest% increase/decrease, Greatest Total Volume, Ticker, and Value
        Cells(1, 15).Value = "Tiker"
        Cells(1, 16).Value = "Value"
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        
        'find greatest % increase/decrease and total volume
        For k = 2 To quarterly_change_last_row
            If Cells(k, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row)) Then
            Cells(2, 15).Value = Cells(k, 9).Value
            Cells(2, 16).Value = Cells(k, 11).Value
            Cells(2, 16).NumberFormat = "0.00%"
        ElseIf Cells(k, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row)) Then
            Cells(3, 15).Value = Cells(k, 9).Value
            Cells(3, 16).Value = Cells(k, 11).Value
            Cells(3, 16).NumberFormat = "0.00%"
        ElseIf Cells(k, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row)) Then
            Cells(4, 15).Value = Cells(k, 9).Value
            Cells(4, 16).Value = Cells(k, 12).Value
        End If
    Next k
    Range("A1").Select
    
    Next ws

End Sub
