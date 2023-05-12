Attribute VB_Name = "Module1"
Sub Stock()
    Dim ws As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim ticker, MaxIncrTicker, MaxDecrTicker, MaxVolTicker  As String
    Dim openingPrice As Double, closingPrice As Double
    Dim totalVolume As Double
    Dim yearlyChange As Double, percentChange As Double
    Dim MaxIncrease, MaxDecrease, MaxVolume As Double
    
    For Each ws In Worksheets
        outputRow = 2
        openingPrice = 0
        totalVolume = 0
                
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If
            
            ws.Range("I" & outputRow).Value = ticker
            ws.Range("J" & outputRow).Value = yearlyChange
            ws.Range("K" & outputRow).NumberFormat = "0.00%"
            ws.Range("K" & outputRow).Value = percentChange
            ws.Range("L" & outputRow).Value = totalVolume
            
            
            If yearlyChange >= 0 Then
                ws.Range("J" & outputRow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & outputRow).Interior.ColorIndex = 3
            End If
                If percentChange > MaxIncrease Then
                    MaxIncrease = percentChange
                    MaxIncrTicker = ticker
                End If
                    
                If percentChange < MaxDecrease Then
                    MaxDecrease = percentChange
                    MaxDecrTicker = ticker
                End If
                
                If totalVolume > MaxVolume Then
                    MaxVolume = totalVolume
                    MaxVolTicker = ticker
                End If

             outputRow = outputRow + 1
        Next i
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("L2:L" & lastRow).NumberFormat = "0,00"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "GreatesT % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("O2").Value = MaxIncrTicker
        ws.Range("O3").Value = MaxDecrTicker
        ws.Range("O4").Value = MaxVolTicker
        ws.Range("P1").Value = "Value"
        ws.Range("P2").Value = FormatPercent(MaxIncrease, 2)
        ws.Range("P3").Value = FormatPercent(MaxDecrease, 2)
        ws.Range("P4").Value = FormatNumber(MaxVolume)
     
    Next ws
    
End Sub

