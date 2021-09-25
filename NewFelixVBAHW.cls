Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
'dim what our variables will be
Dim Start_Price As Double

Dim End_Price As Double

Dim Price_Change As Double

Dim Ticker As String

Dim Percent_Change As Double

Dim Stock_Volume As Double

Stock_Volume = 0
  'dim the area we will populate
  Dim Row As Double

 Row = 2
  Dim Column As Integer
  Column = 1
   Dim i As Long
      'label the columns

    Cells(1, 9).Value = "Ticker"

    Cells(1, 10).Value = "Yearly Price Change"

    Cells(1, 11).Value = "Percent Change"

    Cells(1, 12).Value = "Total Stock Volume"

    'find out what the last populated cell is

     LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

      'choose starting price

    Start_Price = Cells(2, Column + 2).Value
        
    'for loop

    For i = 2 To LastRow

    'see if ticker matches

     If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then

    'choose name and  price

    Ticker_Name = Cells(i, Column).Value
Cells(Row, Column + 8).Value = Ticker_Name
End_Price = Cells(i, Column + 5).Value
'formula to see what the price change is
Price_Change = End_Price - Start_Price
Cells(Row, Column + 9).Value = Price_Change
 'find percents for step three and do the else part
 If (Start_Price = 0 And End_Price = 0) Then
Percent_Change = 0
ElseIf (Start_Price = 0 And End_Price <> 0) Then
 Percent_Change = 1
 Else
 Percent_Change = Price_Change / Start_Price
 Cells(Row, Column + 10).Value = Percent_Change
  'make it a percentage(google how to do this)
 Cells(Row, Column + 10).NumberFormat = "0.00%"
 End If
 'find volume for step 4

 Stock_Volume = Stock_Volume + Cells(i, Column + 6).Value

   Cells(Row, Column + 11).Value = Stock_Volume
                
  'do the reset thing for the rows

Row = Row + 1
 Start_Price = Cells(i + 1, Column + 2)
Stock_Volume = 0
 'if statment
   Else
  Stock_Volume = Stock_Volume + Cells(i, Column + 6).Value
 End If
 Next i
   'set last row of price change
  LastRowPrice = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Set the Cell Colors(10 green 3 red)
For j = 2 To LastRowPrice
   If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
    Cells(j, Column + 9).Interior.ColorIndex = 10

   ElseIf Cells(j, Column + 9).Value < 0 Then
  Cells(j, Column + 9).Interior.ColorIndex = 3
  End If
  Next j
'make the new data fit on the old work sheet(google this)
ActiveSheet.UsedRange.EntireColumn.AutoFit
 Next WS
End Sub
