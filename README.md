# VBA-challenge
Week 2 Challenge
For this assignmnet the code source I used was Challenge 2 Speed Run Video. 

 
'multiple year stock data

Sub stock_analysis()
        'set dimensions
    Dim total As Double
        'double - computers read numbers with decimal points as double in programing
    Dim rowindex As Long
        'for index start or end point - column number
        'index1 will represent the column
    Dim change As Double
            'data types that allows for decimals
    Dim columnindex As Integer
        'index2 will represent the row.
    Dim start As Long
            ''longer values or numbers that the integer data type cannot hold
    Dim rowCount As Long
            'longer values or numbers that the integer data type cannot hold
    Dim percentChange As Double
            'data types that allows for decimals
    Dim days As Integer
    Dim dailyChange As Single
            'decimal values that do not exceed two-digit decimals
    Dim averageChange As Double
            ''data types that allows for decimals
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        columnindex = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
            'set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock volume"
        ws.Range("J1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greaest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
            'get the row number of the last row with data
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        
            'index1 will represent the column
            'index2 will represent the row
        For rowindex = 2 To rowCount
        
            'if ticker changes then pritn results
        If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
    
            'stores results in variable
            total = total + ws.Cells(rowindex, 7).Value
        
        If total = 0 Then
            'print the results
            
        ws.Range("i" & 2 + columnindex).Value = Cells(rowindex, 1).Value
        ws.Range("j" & 2 + columnindex).Value = 0
        ws.Range("k" & 2 + columnindex).Value = "%" & 0
        ws.Range("L" & 2 + columnindex).Value = 0
        
        Else
            If ws.Cells(start, 3) = 0 Then
             For find_value = start To rowindex
              If ws.Cells(find_value, 3).Value <> 0 Then
               start = find_value
               Exit For
            End If
            
        Next find_value
    End If
    
    change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
    percentChange = change / ws.Cells(start, 3)
    
    
    'index1 will represent the column
    'index2 will represent the row
    start = rowindex + 1
    ws.Range("i" & 2 + columnindex) = ws.Cells(rowindex, 1).Value
    ws.Range("J" & 2 + columnindex) = change
    ws.Range("J" & 2 + columnindex).NumberFormat = "0.00"
    ws.Range("K" & 2 + columnindex).Value = percentChange
    ws.Range("K" & 2 + columnindex).NumberFormat = "0.00%"
    ws.Range("L" & 2 + columnindex).Value = total
    
    Select Case change
        Case Is > 0
        ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 4
        Case Is < 0
        ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 3
        Case Else
        ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 0
        End Select
    
    
    End If
    
    total = 0
    change = 0
    columnindex = columnindex + 1
    days = 0
    dailyChange = 0
    
    
    
    Else
        'if ticker is still the same add results
        total = total + ws.Cells(rowindex, 7).Value
        
        End If
        
    Next rowindex
    
        'take the max and min and place thme in a seperate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:k" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:k" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:k" & rowCount)), ws.Range("k2:k" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:k" & rowCount)), ws.Range("k2:k" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("l2:l" & rowCount)), ws.Range("l2:l" & rowCount), 0)
        
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
    
    
    
    
    
        
    Next ws
    
    
    
    
    
    
    
    
End Sub
