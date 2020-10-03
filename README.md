# VBAHomework

Sub StockMarket():

For Each ws In Worksheets

    Dim TickerSymbol As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim TicketList() As Long
    Dim LastRow As Long

TickerRow = 2
TotalStockVolume = 0

ws.Range("I" & 1).Value = "Ticker"
ws.Range("J" & 1).Value = "Yearly_Change"
ws.Range("K" & 1).Value = "Percent_Change"
ws.Range("L" & 1).Value = "Total_Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            Ticker = ws.Cells(i, 1).Value
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            YearlyChange = ws.Cells(i, 6).Value - StockOpen
                
                If StockOpen = 0 Then
                PercentChange = 0
                    
                 Else
                 PercentChange = YearlyChange / StockOpen

            End If
                         
ws.Range("I" & TickerRow).Value = Ticker
ws.Range("L" & TickerRow).Value = TotalStockVolume

TotalStockVolume = 0

ws.Range("J" & TickerRow).Value = YearlyChange
ws.Range("K" & TickerRow).Value = PercentChange
ws.Range("K" & TickerRow).NumberFormat = "0.00%"

TickerRow = TickerRow + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
StockOpen = ws.Cells(i, 3).Value
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

    End If

Next i

    For i = 2 To LastRow
    
        If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4
        
        ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
            
    End If
    
Next i

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

ws.Range("O" & 2).Value = "Greatest % Increase"
ws.Range("O" & 3).Value = "Greatest % Decrease"
ws.Range("O" & 4).Value = "Greatest Total Volume"

For a = 2 To LastRow

    If ws.Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    
    End If
    
Next a

For b = 2 To LastRow

    If ws.Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    
    End If
    
Next b

For c = 2 To LastRow

    If ws.Cells(c, 12).Value > TotalStockVolume Then
        TotalStockVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = TotalStockVolume
        ws.Range("P4").Value = ws.Cells(c, 9)
        
    End If
    
Next c

Next ws
         
End Sub


