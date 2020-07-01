Sub StockReport()

For Each ws In Worksheets

' Column Titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"


' My Variables
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim Summary_Row As Integer
    Dim openIndex As Long
    Dim volume As Double
    Dim change As Double


' Variable iniation and formats
    Summary_Row = 1
    openIndex = 2
    volume = 0
    ws.Range("K:K").NumberFormat = "0.00%"


' Report with Ticker, Yearly Change, Percentage Change and Stock Volume
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                Summary_Row = Summary_Row + 1
                ws.Range("I" & Summary_Row).Value = ticker
                closePrice = ws.Cells(i, 6).Value
                openPrice = ws.Cells(openIndex, 3).Value
                openIndex = i + 1
                change = closePrice - openPrice
                ws.Range("J" & Summary_Row).Value = change
                    If openPrice = 0 Then
                        ws.Range("K" & Summary_Row).Value = 0
                    Else
                        ws.Range("K" & Summary_Row).Value = (closePrice - openPrice) / openPrice
                    End If
                volume = volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Row).Value = volume
                volume = 0
            Else
                volume = volume + ws.Cells(i, 7).Value

            End If

        Next i


' Format Yearly Change
    sRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For j = 2 To sRow
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If

    Next j

' Greatest % Increase, Greatest % Decrease and Greatest volume
' Table Titles and format
    ws.Cells(1, 14).Value = "Summary"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Range("N:N").Columns.AutoFit

' My Variables
    Dim max As Double
    Dim min As Double
    Dim volmax As Double

' Max and Min Results and Format
    max = WorksheetFunction.max(ws.Range("K:K"))
    ws.Cells(2, 15).Value = max
    ws.Cells(2, 15).NumberFormat = "0.00%"
    min = WorksheetFunction.min(ws.Range("K:K"))
    ws.Cells(3, 15).Value = min
    ws.Cells(3, 15).NumberFormat = "0.00%"
    volmax = WorksheetFunction.max(ws.Range("L:L"))
    ws.Cells(4, 15).Value = volmax
    ws.Range("O:O").Columns.AutoFit

Next ws

End Sub
