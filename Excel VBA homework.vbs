Private Sub easy()

    ' Declaring variables
    Dim ticker As String
    Dim totalvol As Double
    Dim counter as Double
    Dim Stock_Data_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim lastRow as Double

    ' Initializing variables
    Total_Vol = 0
    Stock_Data_Row = 2

    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Yearly Percentage"
    Cells(1, 13).Value = "Total Stock Volume"
    Range("J" & Stock_Data_Row).Value = ticker
    Range("K" & Stock_Data_Row).Value = yearly_change
    Range("L" & Stock_Data_Row).Value = yearly_percentage
    Range("M" & Stock_Data_Row).Value = total_stock_volume

    ' Find last row in sheet
    lastRow = Cells(row.count, 1).End(xlUp).row

    For i = 2 To lastRow


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            year_open = Cells(i - counter, 3).Value
            year_close = Cells(i, 6).Value
            yearly_change = year_close - year_open
            ticker = Cells(i, 1).Value
         Else
            Stock_Data_Row = Stock_Data_Row + 1
            counter = counter + 1
        End If
    Next i
End Sub