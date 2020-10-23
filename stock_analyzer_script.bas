Attribute VB_Name = "Module1"
Sub AnalyzeStocks()

    For Each ws In Worksheets
    
        'Add Ticker, Yearly Change, Percent Change, and Total Stock Volume header
         ws.Cells(1, 9).Value = "Ticker"
         ws.Cells(1, 10).Value = "Yearly Change"
         ws.Cells(1, 11).Value = "Percent Change"
         ws.Cells(1, 12).Value = "Total Stock Volume"
         ws.Cells(1, 15).Value = "Ticker"
         ws.Cells(1, 16).Value = "Value"
         ws.Cells(2, 14).Value = "Greatest % Increase"
         ws.Cells(3, 14).Value = "Greatest % Decrease"
         ws.Cells(4, 14).Value = "Greatest Total Volume"
             
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Declare variables
        Dim Stock_Name As String
        Dim Opening_Price As Double
        Dim Percent_Change As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Total_Stock_Volume As Double
        Dim Summary_Table_Row As Double
        
        'Set initial value for Opening Price
        Opening_Price = ws.Cells(2, 3).Value
        
        'This number will grow over iterations, set 0 as initial value
        Total_Stock_Volume = 0

        'Keep track of the location for each stock in summary table
        Summary_Table_Row = 2
        
        'Create a script that will loop through all the stocks for one year
        For i = 2 To LastRow
            
            'Check if cells are the same ticker symbol. If they are not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the ticker symbol
                Stock_Name = ws.Cells(i, 1).Value
                
                'Yearly change from opening price at the
                'beginning of a given year to the closing price at the end of that year
                Closing_Price = ws.Cells(i, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                
                'The percent change from opening price at the
                'beginning of a given year to the closing price at the end of that year.
                If Opening_Price > 0 Then
                    Percent_Change = (Yearly_Change / Opening_Price)
                    
                Else: Percent_Change = 0
                
                End If
                
                'Add to Total Stock Volume then print to Summary Table
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                'Print this ticker symbol in Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Stock_Name
                
                'Print Yearly Change in Summary Table and apply conditional formatting
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
        
                'Print Percentage Change in Summary Table and apply conditional formatting
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                
                'Print Total Stock Volume in Summary Table
                 ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'Add one to the Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1
                    
                
                'Reset counters (Opening, Closing, Volume)
                Total_Stock_Volume = 0
                Opening_Price = ws.Cells(i + 1, 3).Value
                Closing_Price = 0
                Yearly_Change = 0
            
            'If the cell immediately following a row is the same stock
            Else

                'Add to Total Stock Volume of same ticker
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

            End If
        
        Next i
        
        'Find minimum/maximum values for greatest % increase, greatest % decrease, and greatest total volume and format values
        ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
           
        'Loop through Summary Table to find minimum/maximum values
        For i = 2 To LastRow
            
                'Greatest % increase and ticker
                If ws.Cells(i, 11).Value = ws.Cells(2, 16).Value Then
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                    
                
                'Greatest % decrease and ticker
                ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 16).Value Then
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                    
                'Greatest total volume and ticker
                ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 16).Value Then
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                
                End If
            
        Next i
        
        'Automatically resize column widths
        ws.Range("N2:N4").Columns.AutoFit
        ws.Range("P2:P4").Columns.AutoFit
    
Next ws
    
MsgBox "Done"


End Sub





