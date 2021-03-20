Attribute VB_Name = "VBAChallenge"
Sub stonk_compiler():

'Function across multiple worksheets

    For Each ws In Worksheets
    
    'Define the variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Single
    'Had to definer Total Stock Volume as LongLong since it was overflowing using just long
    Dim TotalVolume As LongLong
    Dim SummaryTableRow As Integer
    Dim LastRow As Long
    Dim Volume As Long
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    
    'Setting the Start Row for compiled data
    SummaryTableRow = 2
    
    'Setting inital Value for stock volume
    Volume = 0
    TotalVolume = 0
    
    'Setting up initial value for yearly open/close and the changes
    YearlyOpen = ws.Cells(2, 3).Value
    YearlyClose = 0
    YearlyChange = 0
    
    'Setting variable for last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("I1").Font.Bold = True
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("J1").Font.Bold = True
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K1").Font.Bold = True
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("L1").Font.Bold = True
        
        For i = 2 To LastRow
        'Need to lists all distinct ticker in a column
        'Find and list all the tickers in each Sheet
        'Do the same for Total Stock Volume, Yearly and Percent Change
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Retrieve Stock Ticker Name
                Ticker = ws.Cells(i, 1).Value
                
                'Retrive Stock Volume
                Volume = ws.Cells(i, 7).Value
                
                'Add Volume to Total Stock Volume to get Final Total
                TotalVolume = TotalVolume + Volume
                
                ws.Range("I" & SummaryTableRow).Value = Ticker
                
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
                
                'Retrieve yearly closing price and calculate change
                YearlyClose = ws.Cells(i, 6).Value
                
                YearlyChange = YearlyClose - YearlyOpen
                
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                'Calculate Percent Change
                'To prevent divided by 0 errors
                If YearlyOpen <> 0 Then
                    PercentChange = (YearlyChange / YearlyOpen)
                    ws.Range("K" & SummaryTableRow).Value = PercentChange
                Else
                    ws.Range("K" & SummaryTableRow).Value = 0
                End If
                
                'Conditional Formatting of YearlyChange Cell
                    If YearlyChange > 0 Then
                        'Make it green if positive change
                        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    Else
                        'Make it red if positive change
                        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    End If
                
                'Update YearlyOpen for next ticker
                YearlyOpen = ws.Cells(i + 1, 3).Value
                
                'Move to next row for compiled data table
                SummaryTableRow = SummaryTableRow + 1
                
                'Reset Volume
                TotalVolume = 0
            
            'If it is the same ticker
            Else
                
                Volume = ws.Cells(i, 7).Value
    
                TotalVolume = TotalVolume + Volume
                
            End If
        
        Next i
'-------------------------------------------------------------------------------------------
'Bonus
'Find Ticker and Value for;
    'Greatest % Increase
    'Greatest % Decrease
    'Greatest Total Volume
    
    'Define Variables
    Dim BonusTicker As String
    Dim BonusVolume As LongLong
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim BonusLastRow As Integer
    
    'Set Initial Values
    BonusVolume = 0
    GreatestIncrease = 0
    GreatestDecrease = 0
    
    'Create Headers and Bold the text
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N2").Font.Bold = True
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N3").Font.Bold = True
    ws.Range("N4").Value = "Greatest Volume"
    ws.Range("N4").Font.Bold = True
    ws.Range("O1").Value = "Ticker"
    ws.Range("O1").Font.Bold = True
    ws.Range("P1").Value = "Value"
    ws.Range("P1").Font.Bold = True
    
    'Find Last Row of Summary Table
        '--------------------------Experimental Code--------------------------------
        '----------BonusLastRow = ws.Range("I" & Rows.Count).End(xlUp).Row----------
        '---------------------------------------------------------------------------
    BonusLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Iterate through every row of the compiled data table to find Ticker with Greatest: % Increase, % Decrease, and Stock Volume
    'Use if statements to compare cell data of each stock ticker to find the greatest of each category
    For i = 2 To BonusLastRow
    
        If ws.Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = ws.Cells(i, 11).Value
            
            'Cell P2 will keep updating until the last iteration with the current max value.
            'At last iteration, P2 will have the data of the max of the all the tickers
            ws.Range("P2").Value = GreatestIncrease
            
            'Pulling Ticker name for the max value stock ticker and writing it into Cell O2
            ws.Range("O2").Value = ws.Cells(i, 9).Value
            
        'Using same conditional to look for min while looking for max
        ElseIf ws.Cells(i, 11).Value < GreatestDecrease Then
            
            GreatestDecrease = ws.Cells(i, 11).Value
            
            ws.Range("P3").Value = GreatestDecrease
            
            ws.Range("O3").Value = ws.Cells(i, 9).Value
        
        End If
        
        'Using separate loop to find greatest stock volume
        If ws.Cells(i, 12).Value > BonusVolume Then
        
            BonusVolume = ws.Cells(i, 12).Value
            
            ws.Range("P4").Value = BonusVolume
            
            ws.Range("O4").Value = ws.Cells(i, 9).Value
        End If
    
    Next i
    
    'Format O2:P4 into the correct format
    ws.Range("P2:P3").NumberFormat = "0.00%"
    

        '-----------------------Experimental Code--------------------------------------------------
        'Cells(Count, 4)=Application.WorksheetFunction.Max(Range(Cells(m, 1),Cells(n, 1)))
        'Application.WorksheetFunction.Max(range("Data!A1:A7"))
        'ws.Range("P2").Value = Application.WorksheetFunction.Max(Range("Data!K2:K & BonusLastRow"))
        'ws.Range("P3").Value = Application.WorksheetFunction.Min(Range("Data!K2:K" & BonusLastRow))
        'ws.Range("P4").Value = Application.WorksheetFunction.Max(Range("Data!L2:L" & BonusLastRow))
        '-------------------------------------------------------------------------------------------
    Next ws
    
End Sub

