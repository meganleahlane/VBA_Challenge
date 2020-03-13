Sub TickerTable()

'Create Table Headers
Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume of Stock"


'Define variables and formulas
Dim TickerName As String
Dim opening As Double
Dim closing As Double
Dim YearChange As Double
Dim PerChange As Double

Dim vol As Double
vol = 0

'Set summary table boundaries
Dim tabrow As Integer
'Initially set the table row number to be 1
tabrow = 1


'Loop through to extract variables

    'Define last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through rows
    For r = 2 To LastRow

        'Holds value of opening cost
        If Cells(r - 1, 1).Value <> Cells(r, 1).Value Then
            opening = Cells(r, 3).Value
        End If

        'Searches for when the ticker symbol changes
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            closing = Cells(r, 6).Value

        'Calculates formulas/retrieves values
            TickerName = Cells(r, 1).Value
            
            YearChange = closing - opening
            
                If opening = 0 Then 
                    Cells(tabrow + 1, 11) = "N/A"
                Else
                    PerChange = YearChange / opening
                    Cells(tabrow + 1, 11) = Format(PerChange, "Percent")
                End if
            
            vol = vol + Cells(r, 4).Value

        'Print values into table
        Cells(tabrow + 1, 9) = TickerName
        Cells(tabrow + 1, 10) = YearChange
        Cells(tabrow + 1, 12) = vol

            'Add color formatting
            
            If (Opening = 0) Then
                Cells(tabrow + 1, 11).Interior.ColorIndex = 5
            
            ElseIf (PerChange >= 0) Then
                Cells(tabrow + 1, 11).Interior.ColorIndex = 4

            Else
                Cells(tabrow + 1, 11).Interior.ColorIndex = 3

            End If

        'Add 1 to the tab row
        tabrow = tabrow + 1

        'Reset the vol total
        vol = 0
        
        'If Ticker name is same
        Else

        'Add to volume total
        vol = vol + Cells(r, 4).Value

        End If

    Next r

End Sub



