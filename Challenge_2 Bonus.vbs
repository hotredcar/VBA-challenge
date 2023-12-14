Attribute VB_Name = "Module7"
Sub Bonus():

Dim tick1 As String
Dim tick2 As String
Dim tick3 As String

Dim tic_pi As Double
Dim tic_pd As Double
Dim tic_total As Double

    
For Each ws In Worksheets

    'diplay labels
    ws.Range("Q1") = "Ticker"
    ws.Range("Q1").EntireColumn.AutoFit
    
    ws.Range("R1") = "Value"
    ws.Range("R1").EntireColumn.AutoFit
    
    
    ws.Range("P2") = "Greates % Increase"
    ws.Range("P3") = "Greates % Decrease"
    ws.Range("P4") = "Greatest Total Volume"
    ws.Range("P1").EntireColumn.AutoFit
    
    lastrow = Cells(Rows.Count, 10).End(xlUp).Row
    
    tic_pi = ws.Cells(2, 12).Value
    tic_pd = ws.Cells(2, 12).Value
    tic_total = ws.Cells(2, 13).Value
        
    'calculate/display greatest percentage increase
    For i = 2 To lastrow
        
        If ws.Cells(i, 12).Value > tic_pi Then
            tick1 = ws.Cells(i, 10).Value
            tic_pi = ws.Cells(i, 12).Value
        End If
        
        ws.Cells(2, 17).Value = tick1
        ws.Cells(2, 18).Value = tic_pi
        ws.Cells(2, 18).NumberFormat = "0.00%"
         
    Next i
    
    'calculate/display greatest percentage decrease
    For i = 2 To lastrow
        
        If ws.Cells(i, 12).Value < tic_pd Then
            tick2 = ws.Cells(i, 10).Value
            tic_pd = ws.Cells(i, 12).Value
        End If
        
        ws.Cells(3, 17).Value = tick2
        ws.Cells(3, 18).Value = tic_pd
        ws.Cells(3, 18).NumberFormat = "0.00%"
             
    Next i
   
    'calculate/display greatest total volume
    For i = 2 To lastrow
        
        If ws.Cells(i, 13).Value > tic_total Then
            tick3 = ws.Cells(i, 10).Value
            tic_total = ws.Cells(i, 13).Value
        End If
         
        ws.Cells(4, 17).Value = tick3
        ws.Cells(4, 18).Value = tic_total
        
    Next i
       
Next ws

End Sub

