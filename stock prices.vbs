Sub StockPrices()
    
    MsgBox ("Analyzing data...")
    
    'Time Variables'
    Dim startTimeSec As Single
    Dim finishTimeSec As Single
    
    
    startTimeSec = Timer()


    'Ticker'
    Dim ticker As String

    'Price Variables'
    Dim yearBeginPrice As Double
    Dim yearEndPrice As Double
        
    'Volume Variable'
    Dim totalVolume As Variant

    'Ticker Index'
    Dim tickerIndex As Integer

    'Record Count'
    Dim recordCount As Long
    
    'Greatest Percent Increase'
    Dim greatestPercentIncrease As Double
    Dim greatestPercentIncreaseTicker As String
    
    'Greatest Percent Decrease'
    Dim greatestPercentDecrease As Double
    Dim greatestPercentDecreaseTicker As String
    
    'Greatest Total Volume'
    Dim greatestTotalVolume As Variant
    Dim greatestTotalVolumeTicker As String
    
    'TO MAKE FUNCTION RUN ACROSS ALL SHEETS'
    For j = 2 To Sheets.Count
    
    'Let's Maintain Out Variables!'
    ticker = Sheets(j).Range("A2").Value
    yearBeginPrice = Sheets(j).Range("C2").Value
    yearEndPrice = 0
    totalVolume = Sheets(j).Range("G2").Value
    tickerIndex = 2
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    recordCount = ActiveSheet.UsedRange.Rows.Count
    
        'Output'
        Sheets(j).Range("I1").Value = "Ticker"
        Sheets(j).Range("J1").Value = "Yearly Change"
        Sheets(j).Range("K1").Value = "Percent Change"
        Sheets(j).Range("L1").Value = "Total Stock Volume"
    
        Sheets(j).Range("O2").Value = "Greatest % increase"
        Sheets(j).Range("O3").Value = "Greatest % Decrease"
        Sheets(j).Range("O4").Value = "Greatest Total Volume"
        Sheets(j).Range("P1").Value = "Ticker"
        Sheets(j).Range("Q1").Value = "Value"
    
        
    
        'Where Analysis Begins'
        For i = 3 To recordCount
    
            
            If (ticker <> Sheets(j).Cells(i, 1).Value) Or (i = recordCount) Then
    
                
                    'Ticker'
                    Sheets(j).Cells(tickerIndex, 9).Value = ticker
                    
                    'Yearly Change'
                    Sheets(j).Cells(tickerIndex, 10).Value = yearEndPrice - yearBeginPrice
                    
                    '2 decimal places'
                    Sheets(j).Cells(tickerIndex, 10).NumberFormat = "0.00"
                    
                    'Negative Change (Red)'
                    If Sheets(j).Cells(tickerIndex, 10).Value < 0 Then
                        Sheets(j).Cells(tickerIndex, 10).Interior.ColorIndex = 3
                        
                    'Positive Change (Green)'
                    ElseIf Sheets(j).Cells(tickerIndex, 10).Value >= 0 Then
                        Sheets(j).Cells(tickerIndex, 10).Interior.ColorIndex = 10
                    End If
                    
                    'Percent Change'
                    If yearBeginPrice <> 0 Then
                        Sheets(j).Cells(tickerIndex, 11).Value = (yearEndPrice - yearBeginPrice) / yearBeginPrice
                    Else
                        Sheets(j).Cells(tickerIndex, 11).Value = 0
                    End If
                    
                    '2 decimal places'
                    Sheets(j).Cells(tickerIndex, 11).NumberFormat = "0.00%"
                    
                    'Set Greatest Percent Increase'
                    If Sheets(j).Cells(tickerIndex, 11).Value > greatestPercentIncrease Then
                        greatestPercentIncrease = Sheets(j).Cells(tickerIndex, 11).Value
                        greatestPercentIncreaseTicker = ticker
                    End If
                    
                    'Set Greatest Percent Decrease'
                    If Sheets(j).Cells(tickerIndex, 11).Value < greatestPercentDecrease Then
                        greatestPercentDecrease = Sheets(j).Cells(tickerIndex, 11).Value
                        greatestPercentDecreaseTicker = ticker
                    End If
                    
                    'Total Volume'
                    Sheets(j).Cells(tickerIndex, 12).Value = totalVolume
                
                    'Set Greatest Total Volume'
                    If Sheets(j).Cells(tickerIndex, 12).Value > greatestTotalVolume Then
                        greatestTotalVolume = Sheets(j).Cells(tickerIndex, 12).Value
                        greatestTotalVolumeTicker = ticker
                    End If
                
                
                    'Ticker Index (Next)'
                    tickerIndex = tickerIndex + 1
                    
                    'Place Next Ticker'
                    ticker = Sheets(j).Cells(i, 1).Value
                    
                    'Beginning Stock Price'
                    yearBeginPrice = Sheets(j).Cells(i, 3).Value
                    
                    'Year End Price'
                    yearEndPrice = Sheets(j).Cells(i, 6).Value
                    
                    'Total Volume'
                    totalVolume = Sheets(j).Cells(i, 7).Value
    
            
            Else
                
                yearEndPrice = Sheets(j).Cells(i, 6).Value
                
                totalVolume = totalVolume + Sheets(j).Cells(i, 7).Value
    
            End If
    
        Next i
        
        
        Sheets(j).Range("P2").Value = greatestPercentIncreaseTicker
        Sheets(j).Range("Q2").Value = greatestPercentIncrease
        Sheets(j).Range("Q2").NumberFormat = "0.00%"
        
        Sheets(j).Range("P3").Value = greatestPercentDecreaseTicker
        Sheets(j).Range("Q3").Value = greatestPercentDecrease
        Sheets(j).Range("Q3").NumberFormat = "0.00%"
    
        Sheets(j).Range("P4").Value = greatestTotalVolumeTicker
        Sheets(j).Range("Q4").Value = greatestTotalVolume
    
    'Next Sheet'
    Next j
    
    finishTimeSec = Timer()
    
    
    MsgBox ("Analysis Complete!")

    
    MsgBox ("Runtime:  " & finishTimeSec - startTimeSec & " seconds")
        
End Sub
