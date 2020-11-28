Attribute VB_Name = "Module1"
Sub Stock_Ticker_Loop():

'Make sure the code will run through all of the worksheets
'_________________________________________________________

For Each ws In Worksheets

    'Define Variables
    '______________________________
    
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim HighPrice As Double
    Dim LowPrice As Double
    Dim ClosePrice As Double
    Dim Volume As Double
    Dim SummaryRow As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim Lastrow As Long
    Dim i As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    'Set the initial values for the variables
    '_________________________________________
    
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    SummaryRow = 2
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    OpenPrice = ws.Cells(2, 3).Value
    
    TotalVolume = 0
    
    'Set the total row count for the worksheet
    '_________________________________________
    
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Setting Up the Loop
    '___________________
    
    For i = 2 To Lastrow
    
        'Add the volume of each iteration to the total volume value
        '___________________________________________________________
    
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
            'Use an IF statment to find when there is a new ticker symbol
            '____________________________________________________________
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
        
                'Overwrite the values for the variables
                '______________________________________
        
                YearlyChange = ClosePrice - OpenPrice
            
                If OpenPrice <> 0 Then
                    PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                 
                    Else
                    MsgBox ("For " & Ticker & " Open Price Was 0")
                 
                End If
            
                'Place the new values in summary columns
                '_______________________________________
    
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
            
                    'Use conditional formatting to change the colors of the yearly change cells
                    '__________________________________________________________________________
                
                    If YearlyChange > 0 Then
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
                        ElseIf YearlyChange < 0 Then
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                
                        ElseIf YearlyChange = 0 Then
                        ws.Cells(SummaryRow, 10).Interior.ColorIndex = 0
                                
                    End If
            
                'Finish placeing the new values in the summary columns
                '_____________________________________________________
            
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
        
                'Reset the values of the variables for the next iteration except for the values we still need
                '____________________________________________________________________________________________
        
                YearlyChange = 0
                ClosePrice = 0
        
                'Overwrite the open price for the next ticker symbol
                '___________________________________________________
        
                OpenPrice = ws.Cells(i + 1, 3).Value
        
                'Drop down to the next row in the summary chart
                '______________________________________________
        
                SummaryRow = SummaryRow + 1
            
                'Find the greatest increase and decrease stocks
                '______________________________________________
            
                If PercentChange > GreatestIncrease Then
                    GreatestIncrease = PercentChange
                    GreatestIncreaseTicker = Ticker
        
                    ElseIf PercentChange < GreatestDecrease Then
                    GreatestDecrease = PercentChange
                    GreatestDecreaseTicker = Ticker
        
                End If
            
                'Find the greatest volume stock
                '______________________________
            
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = Ticker
            
                End If
            
                'Place the new values in the "greatest" summary chart
                '____________________________________________________
            
                ws.Cells(2, 15).Value = GreatestIncreaseTicker
                ws.Cells(2, 16).Value = GreatestIncrease
                ws.Cells(3, 15).Value = GreatestDecreaseTicker
                ws.Cells(3, 16).Value = GreatestDecrease
                ws.Cells(4, 15).Value = GreatestVolumeTicker
                ws.Cells(4, 16).Value = GreatestVolume
            
                'Reset the values of the variables for the next iteration
                '________________________________________________________
            
                PercentChange = 0
                TotalVolume = 0
            
            End If
            
  Next i
  
  'Set the column and row headers for the results and the summary table on each worksheet
  '______________________________________________________________________________________
  
  ws.Cells(1, 9).Value = "TICKER"
  ws.Cells(1, 10).Value = "YEARLY CHANGE"
  ws.Cells(1, 11).Value = "PERCENT CHANGE"
  ws.Cells(1, 12).Value = "TOTAL STOCK VOLUME"
  ws.Cells(2, 14).Value = "Greatest % Increase"
  ws.Cells(3, 14).Value = "Greatest % Decrease"
  ws.Cells(4, 14).Value = "Greatest Total Volume"
  ws.Cells(1, 15).Value = "Ticker"
  ws.Cells(1, 16).Value = "Value"
  
  'Make it so the columns on each worksheet are the right size to show the data
  '____________________________________________________________________________
  
  ws.Columns("a:p").AutoFit
    
Next ws

'Indicate that the code is done running
'______________________________________

MsgBox ("Finished!")
    
End Sub
