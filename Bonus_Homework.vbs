Attribute VB_Name = "Module1"

Sub Analyze_All_Stocks()

    'Declare variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Double
    
    Dim currentWS As Worksheet
    Dim lastRow As Long
    Dim row_result_table As Integer
    Dim row_counter As Integer
    Dim max_decr As Double
    Dim max_incr As Double
    Dim max_total As Double
    
    'Loop through all the worksheets
    For Each currentWS In Worksheets
        
        'Initialize variables
        lastRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row
        row_result_table = 2
    
        'Create headers for the final table
        currentWS.Range("J1").Value = "Ticker"
        currentWS.Range("K1").Value = "Yearly Change"
        currentWS.Range("L1").Value = "Percent Change"
        currentWS.Range("M1").Value = "Total Stock Volume"
        
        'Loop through all the stocks and get needed data
        For i = 2 To lastRow
            
            If currentWS.Cells(i, 1).Value <> currentWS.Cells(i + 1, 1).Value Then
            
                'Get the ticker value
                ticker = currentWS.Cells(i, 1).Value
                currentWS.Range("J" & row_result_table).Value = ticker
                
                'Get the Yearly Change and format cells accordingly
                yearly_change = currentWS.Cells(i, 6) - currentWS.Cells(i - row_counter, 3)
                currentWS.Range("K" & row_result_table).Value = yearly_change
                If yearly_change > 0 Then
                    currentWS.Range("K" & row_result_table).Interior.ColorIndex = 4
                    Else: currentWS.Range("K" & row_result_table).Interior.ColorIndex = 3
                End If
                
                'Calculate the Percent Change, make sure th    0
                If currentWS.Cells(i - row_counter, 3) <> 0 Then
                    percent_change = WorksheetFunction.Round(yearly_change / currentWS.Cells(i - row_counter, 3), 4)
                    currentWS.Range("L" & row_result_table).Value = percent_change
                    Else: currentWS.Range("L" & row_result_table).Value = "N/A"
                End If
                
                'Calculate the Total Stock Value
                total_stock = total_stock + CDbl(currentWS.Cells(i, 7))
                currentWS.Range("M" & row_result_table).Value = total_stock
                
                'Set row_counter to 0
                row_counter = 0
                
                'Set total_stock to 0
                total_stock = 0
                
                'Increase row number in result table
                row_result_table = row_result_table + 1
                
                Else
                
                'Calculate number of rows for the current ticker
                row_counter = row_counter + 1
                
                'Calculate total stock
                total_stock = total_stock + CDbl(currentWS.Cells(i, 7))
                
            End If
            
        Next i
        
        '------PART 2 Get Greatest Increase and Decrease, and Total Values
        
        'Declare vriables
        Dim increase As Double
        Dim decrease As Double
        Dim great_total As Double
        
        Dim max_decr_row As Integer
        Dim max_incr_row As Integer
        Dim max_total_row As Integer
    
        'Set initial max values
        max_decr = CDbl(currentWS.Range("L2").Value)
        max_incr = CDbl(currentWS.Range("L2").Value)
        max_total = CDbl(currentWS.Range("M2").Value)
        
        'Create headers for the final table
        currentWS.Range("O2").Value = "Greatest % Increase"
        currentWS.Range("O3").Value = "Greatest % Decrease"
        currentWS.Range("O4").Value = "Greatest Total Volume"
        currentWS.Range("P1").Value = "Ticker"
        currentWS.Range("Q1").Value = "Value"
        currentWS.Range("M1").Value = "Total Stock Volume"
              
        'Loop through result table
        For i = 2 To row_result_table
      
            'Find Greatest % Increase
            If currentWS.Cells(i, 12).Value <> "N/A" And currentWS.Cells(i, 12).Value > max_incr Then
                max_incr = currentWS.Cells(i, 12).Value
                max_incr_row = i
            End If
            
            'Find Greatest % Decrease
            If currentWS.Cells(i, 12).Value <> "N/A" And currentWS.Cells(i, 12).Value < max_decr Then
                max_decr = currentWS.Cells(i, 12).Value
                max_decr_row = i
            End If
            
            'Find Greatest Total Volume
            If currentWS.Cells(i, 13).Value > max_total Then
                max_total = currentWS.Cells(i, 13).Value
                max_total_row = i
            End If
            
        Next i
        
        'Fill the table for greatest values
        currentWS.Range("P2").Value = currentWS.Range("J" & max_incr_row).Value
        currentWS.Range("P3").Value = currentWS.Range("J" & max_decr_row).Value
        currentWS.Range("P4").Value = currentWS.Range("J" & max_total_row).Value
        
        currentWS.Range("Q2").Value = max_incr
        currentWS.Range("Q3").Value = max_decr
        currentWS.Range("Q4").Value = max_total
                
        'Format columns
        currentWS.Range("L2:L" & row_result_table).NumberFormat = "0.00%"
        currentWS.Range("Q2:Q3").NumberFormat = "0.00%"
        currentWS.Columns("J:M").AutoFit
        currentWS.Columns("O:Q").AutoFit
              
   Next currentWS
   
End Sub
