Sub Button1_Click()

sheetValue = InputBox("Which worksheet would you like to sum up?")

Worksheets(sheetValue).Activate

    'Define variables I think I need
    Dim ticker As String
    ticker = " "
    'Dim ticker_Vol As Long or is it double
    'ticker_Vol = 0
    Dim year_open As Double
    year_open = 0
    Dim year_close As Double
    year_close = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim stock_volume As Double
    stock_volume = 0
    Dim max_percent As Double
    max_percent = 0
    Dim min_percent As Double
    min_percent = 0
    Dim max_volume As Double
    max_volume = 0
    
    
    'variable to find last row in each worksheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'set table as long for final data
    Dim Sum_Table_Row As Long
    Sum_Table_Row = 2
    
    'set headers for table
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Define first ticker's year open       
    year_open = Cells(2, 3).Value
    
        For I = 2 To lastrow
        
        
            'If a cell in the ith row is not the same as the cell in the next ith row then...
           If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

                'pull ticker value which is located in (i, 1)
                ticker = Cells(I, 1).Value
                
                'print ticker value in summary table
                Range("J" & Sum_Table_Row).Value = ticker
                                              
                'pulls the close value in the ith row
                year_close = Cells(I, 6).Value

                'finds yearly change using formual
                Yearly_Change = year_close - year_open
                
                'Prints yearly change to column k
                Range("K" & Sum_Table_Row).Value = Yearly_Change
                
                'eliminates the possibility of error when running percent change formula
                If (year_open = 0) Then

                    Percent_Change = 0

                Else
                    
                    'if percent change does not equal 0 then complete formula as normal
                    Percent_Change = Yearly_Change / year_open
                
                End If

                'print percent change to column L and complete number format to percent             
                Range("L" & Sum_Table_Row).Value = Percent_Change
                Range("L" & Sum_Table_Row).NumberFormat = "0.00%"

                'not actually sure what this does, but I know I needed it.              
                stock_volume = stock_volume + Cells(I, 7).Value
                Range("M" & Sum_Table_Row).Value = stock_volume
                                                                                                                                                                   
                'add another row to the table in order to add the next ticker
                Sum_Table_Row = Sum_Table_Row + 1

                'calculates year open using cells indicated below for tickers other than the first ticker                               
                year_open = Cells(I + 1, 3).Value
                
                'reset stock volume count
                stock_volume = 0
                       
            Else

                'total the stock volume for each ticker
                stock_volume = stock_volume + Cells(I, 7).Value
                
            End If
                                                           
    Next I
            
    'set headers for new table
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'set lastrow for new table
    lastrow_Sum_Table_Row = Cells(Rows.Count, 12).End(xlUp).Row

        For I = 2 To lastrow_Sum_Table_Row

            'calculate max for the L range
            If Cells(I, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_Sum_Table_Row)) Then

                'prints information from first table to second table for the max
                Cells(2, 16).Value = Cells(I, 10).Value
                Cells(2, 17).Value = Cells(I, 12).Value
                Cells(2, 17).NumberFormat = "0.00%"
            
            'calculate min for the L range
            ElseIf Cells(I, 12).Value = Application.WorksheetFunction.Min(Range("L2:L" & lastrow_Sum_Table_Row)) Then
                
                'prints information from first table to second table for the min
                Cells(3, 16).Value = Cells(I, 10).Value
                Cells(3, 17).Value = Cells(I, 12).Value
                Cells(3, 17).NumberFormat = "0.00%"

            'calculate max for M column    
            ElseIf Cells(I, 13).Value = Application.WorksheetFunction.Max(Range("M2:M" & lastrow_Sum_Table_Row)) Then
                
                'prints information from first table to second table
                Cells(4, 16).Value = Cells(I, 10).Value
                Cells(4, 17).Value = Cells(I, 13).Value
                       
         
            End If

        Next I
        
        'Creates color index for positive and negative numbers
        For I = 2 To lastrow_Sum_Table_Row
            If Cells(I, 11).Value > 0 Then
                Cells(I, 11).Interior.ColorIndex = 10
            Else
                Cells(I, 11).Interior.ColorIndex = 3
            End If
            
        Next I

End Sub
