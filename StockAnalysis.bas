Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

   'Loop through all worksheets
    For Each ws In Worksheets

       'Declaring and Assigning Variables
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As Double
        TotalStockVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim LastRow As Double
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
    
        'Find Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
            
            'Column Headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
            'Total Stock Volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            'Checks if same ticker name... if not
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
               'Ticker name
                Ticker = ws.Cells(i, 1).Value
                'Printing name in the table
                ws.Range("I" & SummaryTableRow).Value = Ticker
                'Printing total stock volume in the table
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                'Reset ticker
                TotalStockVolume = 0
            
                'Set name for Yearly Open, Yearly Close, and Yearly Change
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                'Percent Change
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                
                'Formatting to include % and two decimal places
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                
                'Conditional formatting positive is green / negative is red
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
                
                'Add one to table
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
          
          Next i
          
          'BONUS
          
          'Greatest % increase and decrease and Greatest total volume
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
          'Loop for final results
          For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If
                
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If
                
                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If
                
         Next i
         
         'Formatting the Greatest values
         ws.Range("Q3").NumberFormat = "0.00%"
         ws.Range("Q4").NumberFormat = "0.00E+00"
         
         'Formatting to include % and two decimal places
         ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
         ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
         
  'Formatting to autofit the columns
  ws.Columns("I:Q").AutoFit
    
    Next ws

End Sub

