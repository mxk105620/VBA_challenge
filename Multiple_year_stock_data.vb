Sub multi_year_stock()
For Each ws In Worksheets

Dim Ticker As String

Dim Total_stock_volume As Double
Total_stock_volume = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Yearlychange As Double
Dim closeprice As Double
Dim openprice As Double
Dim percentchange As Double
Dim lastrow As Double
Dim lastrow1 As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim a As Range
Dim greatest_volume As Double

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "yearly change"
ws.Cells(1, 11).Value = "Percent"
ws.Cells(1, 12).Value = "Total volume"

openprice = ws.Cells(2, 3).Value

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row - 1
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value
Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value


closeprice = ws.Cells(i, 6).Value


Yearlychange = (closeprice - openprice)
ws.Range("J" & Summary_Table_Row).Value = Yearlychange
If openprice = 0 Then
    percentchange = 0
Else
    percentchange = Yearlychange / openprice
End If
ws.Range("K" & Summary_Table_Row).Value = percentchange
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"


If Yearlychange > 0 Then
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
Else
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
End If
ws.Range("I" & Summary_Table_Row).Value = Ticker

ws.Range("L" & Summary_Table_Row).Value = Total_stock_volume


    Summary_Table_Row = Summary_Table_Row + 1
    Total_stock_volume = 0
    openprice = ws.Cells(i + 1, 3).Value
    
Else
Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value

End If
Next i


lastrow1 = ws.Cells(Rows.Count, 11).End(xlUp).Row - 1

greatest_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow1))
greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow1))
greatest_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow1))

ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(2, 17).Value = greatest_increase
ws.Cells(2, 17).NumberFormat = "0.00%"


ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(3, 17).Value = greatest_decrease
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Cells(4, 15).Value = "Greatest Total volume"
ws.Cells(4, 17).Value = greatest_volume

For j = 2 To lastrow1
If ws.Cells(j, 11).Value = greatest_increase Then ws.Cells(2, 16) = ws.Cells(j, 9).Value
If ws.Cells(j, 11).Value = greatest_decrease Then ws.Cells(3, 16) = ws.Cells(j, 9).Value
If ws.Cells(j, 12).Value = greatest_volume Then ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
Next j

Next ws

End Sub