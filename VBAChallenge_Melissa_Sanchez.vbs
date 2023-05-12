Attribute VB_Name = "Module14"
Sub ticker()
    
    'Counts number of worksheets
    Dim WorksheetName As String
    Dim WS_Count As Integer
    Dim row_counter As Double
    row_counter = 2
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Variables
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim greatest_name As String
    Dim decrease_Name As String
    Dim Greatest_total_volume As Double
    Dim Greatest_total_volume_name As String
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        
        Dim n As Integer
        n = 2
        
        Dim Stock_details(2) As Double
        Dim Stock_Name As String
        
        
        'We  are considering Stock_details(0)as opening price, Stock_details(1)as closing price and Stock_details(2)as the volumen of stock
        
        ' Column Names

        ws.Cells(1, 10).Value = "Stock Name"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop to move around the sheet
        For i = 2 To LastRow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'Determinate yearly change of the stock
                Stock_details(1) = ws.Cells(i, 6).Value
                Stock_details(2) = Stock_details(2) + ws.Cells(i, 7).Value
                
                Stock_Name = ws.Cells(i, 1).Value
                
                ws.Cells(n, 10).Value = Stock_Name
                ws.Cells(n, 13).Value = Stock_details(2)
                ws.Cells(n, 11).Value = Stock_details(1) - Stock_details(0)
                ws.Cells(n, 11).NumberFormat = "$0.00"
                ws.Cells(n, 12).Value = (((Stock_details(1) - Stock_details(0)) / Stock_details(0)))
                ws.Cells(n, 12).NumberFormat = "0.00%"
                
                'Reset variables for next stock
                Stock_details(0) = ws.Cells(i + 1, 3).Value
                Stock_details(2) = 0
                
                'Increment row counter
                n = n + 1
                
            Else
                'Add to the volume of the stock
                Stock_details(2) = Stock_details(2) + ws.Cells(i, 7).Value
                
                'Determinate opening price for every stock
                If Stock_details(0) = 0 Then
                    Stock_details(0) = ws.Cells(i, 3).Value
                End If
                
                'Color codes yearly change
                If ws.Cells(n, 11).Value >= 0 Then
                ws.Cells(n, 11).Interior.ColorIndex = 4
                 End If
                 
                
                If ws.Cells(n, 11).Value < 0 Then
                ws.Cells(n, 11).Interior.ColorIndex = 3
        End If
        
            End If
            
        Next i
        
        'Find stock with greatest percent increase, greatest percent decrease, and greatest total volume
        Greatest_Increase = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        Greatest_Decrease = WorksheetFunction.Min(ws.Range("L2:L" & LastRow))
        Greatest_total_volume = WorksheetFunction.Max(ws.Range("M2:M" & LastRow))
        
        greatest_name = ws.Cells(WorksheetFunction.Match(Greatest_Increase, ws.Range("L2:L" & LastRow), 0) + 1, 10).Value
        decrease_Name = ws.Cells(WorksheetFunction.Match(Greatest_Decrease, ws.Range("L2:L" & LastRow), 0) + 1, 10).Value
        Greatest_total_volume_name = ws.Cells(WorksheetFunction.Match(Greatest_total_volume, ws.Range("M2:M" & LastRow), 0) + 1, 10).Value
        
       'Print results
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatest_name
        ws.Cells(2, 17).Value = Greatest_Increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = decrease_Name
        ws.Cells(3, 17).Value = Greatest_Decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = Greatest_total_volume_name
        ws.Cells(4, 17).Value = Greatest_total_volume
        ws.Cells(4, 17).NumberFormat = "0.00"
        
   Range("J:Q").EntireColumn.AutoFit
   
    Next ws
    
End Sub
