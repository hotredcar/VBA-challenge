Attribute VB_Name = "Module3"
Sub Main()

Dim ticker As String
Dim total As Double

For Each ws In Worksheets

    j = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'add column names
    ws.Range("J1") = "Ticker"
    ws.Range("J1").EntireColumn.AutoFit
    
    ws.Range("K1") = "Yearly Change"
    ws.Range("K1").EntireColumn.AutoFit
    
    ws.Range("L1") = "Percent Change"
    ws.Range("L1").EntireColumn.AutoFit
    
    ws.Range("M1") = "Total Stock Volume"
    ws.Range("M1").EntireColumn.AutoFit

     For i = 2 To lastrow
     
        'if there is a change from a row to another
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
            'add current total and display ticker and total
            ticker = ws.Cells(i, 1).Value
            total = total + ws.Cells(i, 7).Value
            
            ws.Cells(j, 10).Value = ticker
            ws.Cells(j, 13).Value = total
            
            
            'calculate/diplay yearly change
            ws.Cells(j, 11).Value = ws.Cells(i, 6).Value - ws.Cells(i - n, 3).Value 'last "closing price" of the year subtract first "opening price" of the year
            
                If ws.Cells(j, 11).Value > 0 Then
                    ws.Cells(j, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(j, 11).Interior.ColorIndex = 3
                End If
                
            'calculate/display percentage
            ws.Cells(j, 12).Value = ws.Cells(j, 11).Value / ws.Cells(i - n, 3).Value
            ws.Cells(j, 12).NumberFormat = "0.00%"
            
            j = j + 1   'go to next row of summary table
            
            total = 0   'reset total
            
            n = 0       'reset number of rows
           
        'if the ticker remains the same
        Else
            total = total + ws.Cells(i, 7).Value    'countinue to sum total
            
            n = n + 1                               'continue to count rows
            
        End If
        
    Next i
    
Next ws

End Sub
