Sub StockMarket()
'Following code was a collaboration between Matthew Sanders, James Brannan, and Allison Chase
    'Define variables
    Dim Ticker As String
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Summary_Table_Row As Integer
    Dim CurrentMin As Double
    Dim CurrentMax As Double
    Dim CurrentGreatestVolume As Double
    maxVal = -999999999
    Dim openPrice As Double
    Dim closePrice As Double
    'Loop code through all worksheets in document
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        'Initialize variables
        CurrentMax = 0
        CurrentMin = 0
        CurrentGreatestVolume = 0
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        TotalStockVolume = 0
        Summary_Table_Row = 2
        'Hard code for column lables and addtional information table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        'Loop through all rows in column A
        For i = 2 To LastRow
            'Determine if the ticker value in Column A changes
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                closePrice = Cells(i, 6).Value
                QuarterlyChange = closePrice - openPrice
                PercentChange = (QuarterlyChange / openPrice)
                'Determine what the greatest increase and decrease in %
                If PercentChange > CurrentMax Then
                    CurrentMax = PercentChange
                    Cells(2, 16) = Ticker
                    Cells(2, 17) = CurrentMax
                    Cells(2, 17).NumberFormat = "0.00%"
                ElseIf PercentChange < CurrentMin Then
                    CurrentMin = PercentChange
                    Cells(3, 16) = Ticker
                    Cells(3, 17) = CurrentMin
                    Cells(3, 17).NumberFormat = "0.00%"
                End If
                'Place results in the correct cells of the information table
                Range("I" & Summary_Table_Row).Value = Ticker
                Range("J" & Summary_Table_Row).Value = QuarterlyChange
                Range("K" & Summary_Table_Row).Value = PercentChange
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Range("L" & Summary_Table_Row).Value = TotalStockVolume
                Summary_Table_Row = Summary_Table_Row + 1
                TotalStockVolume = 0
            Else
                'set the opening price for current ticker
                If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                    openPrice = Cells(i, 3).Value
                End If
                    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            End If
        Next i
        ' Loop through rows to find the greatest total volume
    Next ws
    'Format cells to turn green or red
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4 ' Green for positive 
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3 ' Red for negative 
            ElseIf IsEmpty(ws.Cells(i, 10)) Then
                ws.Cells(i, 10).Interior.ColorIndex = xlNone ' No color
            End If
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4 ' Green for positive 
            ElseIf ws.Cells(i, 11).Value < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 3 ' Red for negative 
            End If
            If IsEmpty(ws.Cells(i, 11)) Then
                ws.Cells(i, 11).Interior.ColorIndex = xlNone ' No color
            End If
        Next i
    Next ws
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        maxVal = 0 ' Initialize maxVal before looping through the rows
        For i = 2 To LastRow
            CurrentGreatestVolume = ws.Cells(i, 12).Value ' Use ws to reference the current worksheet
            If CurrentGreatestVolume > maxVal Then
                maxVal = CurrentGreatestVolume
                Ticker = ws.Cells(i, 9).Value
                ws.Cells(4, 16) = Ticker
                ws.Cells(4, 17) = maxVal
            End If
    Next i
Next ws
End Sub
