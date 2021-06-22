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
    
    'Initialize variables
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    row_result_table = 2
    
    'Create headers for the final table
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    
    'Loop through all the stocks and get needed data
    'For Each currentWS In Worksheets
        
        For i = 2 To lastRow
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                'Get the ticker value
                ticker = Cells(i, 1).Value
                Range("J" & row_result_table).Value = ticker
                
                'Get the Yearly Change and format cells accordingly
                yearly_change = Cells(i, 6) - Cells(i - row_counter, 3)
                Range("K" & row_result_table).Value = yearly_change
                If yearly_change > 0 Then
                    Range("K" & row_result_table).Interior.ColorIndex = 4
                    Else: Range("K" & row_result_table).Interior.ColorIndex = 3
                End If
                
                'Calculate the Percent Change, make sure th    0
                If Cells(i - row_counter, 3) <> 0 Then
                    percent_change = WorksheetFunction.Round(yearly_change / Cells(i - row_counter, 3), 4)
                    Range("L" & row_result_table).Value = percent_change
                    Else: Range("L" & row_result_table).Value = "N/A"
                End If
                
                'Calculate the Total Stock Value
                total_stock = total_stock + CDbl(Cells(i, 7))
                Range("M" & row_result_table).Value = total_stock
                
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
                total_stock = total_stock + CDbl(Cells(i, 7))
                
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
        max_decr = CDbl(Range("L2").Value)
        max_incr = CDbl(Range("L2").Value)
        max_total = CDbl(Range("M2").Value)
        
        'Create headers for the final table
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("M1").Value = "Total Stock Volume"
              
        'Loop through result table
        For i = 2 To row_result_table
      
            'Find Greatest % Increase
            If Cells(i, 12).Value <> "N/A" And Cells(i, 12).Value > max_incr Then
                max_incr = Cells(i, 12).Value
                max_incr_row = i
            End If
            
            'Find Greatest % Decrease
            If Cells(i, 12).Value <> "N/A" And Cells(i, 12).Value < max_decr Then
                max_decr = Cells(i, 12).Value
                max_decr_row = i
            End If
            
            'Find Greatest Total Volume
            If Cells(i, 13).Value > max_total Then
                max_total = Cells(i, 13).Value
                max_total_row = i
            End If
            
        Next i
        
        Range("P2").Value = Range("J" & max_incr_row).Value
        Range("P3").Value = Range("J" & max_decr_row).Value
        Range("P4").Value = Range("J" & max_total_row).Value
        
        Range("Q2").Value = max_incr
        Range("Q3").Value = max_decr
        Range("Q4").Value = max_total
                
        'Format columns
        Range("L2:L" & row_result_table).NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        Columns("J:M").AutoFit
        Columns("O:Q").AutoFit
        
        
   ' Next current
End Sub
