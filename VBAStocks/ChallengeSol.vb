Sub ChallengeSol()

'Define variables
Dim TickerName As String
Dim opening As Double
Dim closing As Double
Dim YearChange As Double
Dim PerChange As Double
Dim vol As Double
Dim tabrow As Integer
Dim PerInc As Integer
Dim PerDec As Integer
Dim TotalVol As Double

'--------------------------------------------
'LOOP THROUGH ALL SHEETS
'--------------------------------------------

For Each ws In Worksheets

'Create Table Headers
ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume of Stock"

ws.Cells(1, 14).Value = "Greatest % Increase"
ws.Cells(1, 15).Value = "Greatest % Decrease"
ws.Cells(1, 16).Value = "Greatest Total Volume"

'Define dynamic variables and formulas
vol = 0
tabrow = 1
LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'Loop through to extract variables

    'Loop through rows
    For r = 2 To LastRow

        'Holds value of opening cost
        If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value Then
            opening = ws.Cells(r, 3).Value
        End If

        'Searches for when the ticker symbol changes
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            closing = ws.Cells(r, 6).Value

        'Calculates formulas/retrieves values
            TickerName = ws.Cells(r, 1).Value
            YearChange = closing - opening
                If opening = 0 Then
                    ws.Cells(tabrow + 1, 11) = "N/A"
                Else
                    PerChange = YearChange / opening
                    ws.Cells(tabrow + 1, 11) = Format(PerChange, "Percent")
                End If
            vol = vol + ws.Cells(r, 4).Value

        'Print values into table
        ws.Cells(tabrow + 1, 9) = TickerName
        ws.Cells(tabrow + 1, 10) = YearChange
        ws.Cells(tabrow + 1, 12) = vol

            'Add color formatting
            If (opening = 0) Then
                ws.Cells(tabrow + 1, 11).Interior.ColorIndex = 5
            
            ElseIf (PerChange >= 0) Then
                ws.Cells(tabrow + 1, 11).Interior.ColorIndex = 4

            Else
                ws.Cells(tabrow + 1, 11).Interior.ColorIndex = 3

            End If

        'Add 1 to the tab row
        tabrow = tabrow + 1

        'Reset the vol total
        vol = 0
        
        'If Ticker name is same
        Else
            'Add to volume total
            vol = vol + ws.Cells(r, 4).Value
        End If

    Next r

'Print values into second table from first table
        PerInc = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 14) = PerInc

        PerDec = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(2, 15) = PerDec
        
        TotalVol = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(2, 16) = TotalVol

Next ws

End Sub