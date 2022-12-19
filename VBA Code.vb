Sub Stock_Price_Calculator()
    
    'Create variable for number of sheets
    Dim WS_Count As Integer
    
    'Find number of sheets
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    'Step through each sheet
    For i = 1 To WS_Count
    
        'Create Ticker variable
        Dim Ticker As String
        
        'Keep track of each ticker in summary table
        Summary_Table_Row = 2

        'Find last row
        lastrow = ActiveWorkbook.Worksheets(i).Range("A1").End(xlDown).Row

        'Set up Beginning_Price and Ending_Price
        Beginning_Price = 0
        Ending_Price = 0
        
        'Add Column Headers
        ActiveWorkbook.Worksheets(i).Cells(1, 9).Value = "Ticker"
        ActiveWorkbook.Worksheets(i).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(i).Cells(1, 12).Value = "Total Stock Volume"
                     
        'Step through each row to calculate required fields
        For j = 2 To lastrow
    
            'Ignore the zero rows
            On Error Resume Next
            
            'Save the beginning price
            If ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j - 1, 1).Value Then
            Beginning_Price = ActiveWorkbook.Worksheets(i).Cells(j, 3).Value
                         
            'Check if we are still in the same ticker symbol. For the last ticker of each group:
            ElseIf ActiveWorkbook.Worksheets(i).Cells(j, 1).Value <> ActiveWorkbook.Worksheets(i).Cells(j + 1, 1).Value Then
        
            'Save each ticker symbol
            Ticker = ActiveWorkbook.Worksheets(i).Cells(j, 1).Value
                     
            'Save the Ending Price
            Ending_Price = ActiveWorkbook.Worksheets(i).Cells(j, 6).Value
        
            'Output each Ticker in the summary table
            ActiveWorkbook.Worksheets(i).Range("I" & Summary_Table_Row).Value = Ticker
            
            'Output the yearly price change and apply conditional formatting
            yearly_price_change = Ending_Price - Beginning_Price
            ActiveWorkbook.Worksheets(i).Range("J" & Summary_Table_Row).Value = yearly_price_change
                If yearly_price_change >= 0 Then
                    ActiveWorkbook.Worksheets(i).Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                Else
                    ActiveWorkbook.Worksheets(i).Range("J" & Summary_Table_Row).Interior.Color = vbRed
                End If
                
            'Output the percentage price change and apply conditional formatting
            Percent_Price_Change = (Ending_Price / Beginning_Price - 1)
            ActiveWorkbook.Worksheets(i).Range("K" & Summary_Table_Row).Value = Percent_Price_Change
            ActiveWorkbook.Worksheets(i).Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        
            'Output Volume in the summary table
            Dim tickers As Range
            Set tickers = ActiveWorkbook.Worksheets(i).Range("A2:A" & lastrow)

            Dim volume As Range
            Set volume = ActiveWorkbook.Worksheets(i).Range("G2:G" & lastrow)
            
            ActiveWorkbook.Worksheets(i).Range("L" & Summary_Table_Row).Value = Application.WorksheetFunction.SumIf(tickers, Ticker, volume)
    
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
                           
            End If
        Next j
     
    'Greatest Change Bonus
    'Input titles
    ActiveWorkbook.Worksheets(i).Cells(2, 14).Value = "Greatest % Increase"
    ActiveWorkbook.Worksheets(i).Cells(3, 14).Value = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(i).Cells(4, 14).Value = "Greatest Total Volume"
    ActiveWorkbook.Worksheets(i).Cells(1, 15).Value = "Ticker"
    ActiveWorkbook.Worksheets(i).Cells(1, 16).Value = "Value"
        
    'Set variables
    Dim change As Range
    Set change = ActiveWorkbook.Worksheets(i).Range("K2:K" & lastrow)
     
    Dim vol As Range
    Set vol = ActiveWorkbook.Worksheets(i).Range("L2:L" & lastrow)
     
    Dim MatchPerfRange As Range
    Set MatchPerfRange = ActiveWorkbook.Worksheets(i).Range("K2:K" & lastrow)
       
    Dim MatchVolRange As Range
    Set MatchVolRange = ActiveWorkbook.Worksheets(i).Range("L2:L" & lastrow)
       
    Dim IndexRange As Range
    Set IndexRange = ActiveWorkbook.Worksheets(i).Range("I2:I" & lastrow)
       
    'Calculate Greatest Increase Output
    GreatestInc = Application.WorksheetFunction.Max(change)
    GreatestIncTick = Application.WorksheetFunction.Index(IndexRange, Application.WorksheetFunction.Match(GreatestInc, MatchPerfRange, 0))
    ActiveWorkbook.Worksheets(i).Cells(2, 16).Value = GreatestInc
    ActiveWorkbook.Worksheets(i).Range("P2").NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(i).Cells(2, 15).Value = GreatestIncTick
                    
    'Calculate Greatest Decrease Ouput
    GreatestDec = Application.WorksheetFunction.Min(change)
    GreatestDecTick = Application.WorksheetFunction.Index(IndexRange, Application.WorksheetFunction.Match(GreatestDec, MatchPerfRange, 0))
    ActiveWorkbook.Worksheets(i).Cells(3, 16).Value = GreatestDec
    ActiveWorkbook.Worksheets(i).Range("P3").NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(i).Cells(3, 15).Value = GreatestDecTick
     
    'Calculate Greatest Volume Output
    GreatestVol = Application.WorksheetFunction.Max(vol)
    GreatestVolTick = Application.WorksheetFunction.Index(IndexRange, Application.WorksheetFunction.Match(GreatestVol, MatchVolRange, 0))
    ActiveWorkbook.Worksheets(i).Cells(4, 16).Value = GreatestVol
    ActiveWorkbook.Worksheets(i).Cells(4, 15).Value = GreatestVolTick
         
    Next i
End Sub