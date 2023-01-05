Attribute VB_Name = "Module1"
Sub ticker()

Dim ws As Worksheet

'looping through each work sheet
For Each ws In ThisWorkbook.Worksheets
 ws.Activate
 
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
 
'declaring variables
Dim ticker As String
Dim Row As Double
Row = 2
Dim volume As Double
Dim totalvolume As Double
Dim openingprice As Double
Dim closingprice As Double


'number of rows for column A
RowRange = Range("A1").End(xlDown).Row

Dim WS_Count As Integer
Dim h As Integer


'looping through all tickers symbols
For I = 2 To 1000
    'variables used to compare previous and next variables
    ticker = Cells(I, 1).Value
    volume = Cells(I, 7).Value
    
'if ticker is not equal to the previous ticker
    If ticker <> Cells(I - 1, 1).Value Then
        'opening price of ticker
        openingprice = Cells(I, 3).Value
        'print out ticker
        Cells(Row, 9).Value = ticker
        ticker = Cells(I, 1).Value
        'variable to keep track of printed rows
        Row = Row + 1
        totalvolume = 0
    End If
    
    'if statement to keep track of the year end closing price of the ticker
     If ticker <> Cells(I + 1, 1).Value Then
        closingprice = Cells(I, 6).Value
    End If
      
    'print out total volume, percentchange and yearly change (formatted)
    totalvolume = totalvolume + volume
    Cells(Row - 1, 12).Value = totalvolume
    Cells(Row - 1, 10).Value = -(openingprice - closingprice)
    percentchange = Cells(Row - 1, 10).Value / openingprice
    Cells(Row - 1, 11).Value = Format(percentchange, "0.00%")
    
    'formatting for colored  cells
     If Cells(Row - 1, 10).Value > 0 Then
        Cells(Row - 1, 10).Interior.ColorIndex = 4
     Else
        Cells(Row - 1, 10).Interior.ColorIndex = 3
     End If
     
    Next I
    
RowRange2 = Range("I1").End(xlDown).Row

Dim greatestincrease As Double
greatestincrease = Cells(2, 10).Value
Dim greatestincreaseticker As String

'loop for finding greatest % increase
For gi = 2 To RowRange2
 If Cells(gi, 11).Value >= greatestincrease Then
    greatestincrease = Cells(gi, 11).Value
    greatestincreaseticker = Cells(gi, 9).Value
 End If
 Cells(2, 16).Value = greatestincreaseticker
 Cells(2, 17).Value = Format(greatestincrease, "0.00%")
Next gi

'loop for finding greatest % decrease
Dim greatestdecrease As Double
greatestdecrease = Cells(2, 11).Value
Dim greatestdecreaseticker As String

For gd = 2 To RowRange2
 If Cells(gd, 11).Value < greatestdecrease Then
    greatestdecrease = Cells(gd, 11).Value
    greatestdecreaseticker = Cells(gd, 9).Value
 End If
 Cells(3, 16).Value = greatestdecreaseticker
 Cells(3, 17).Value = Format(greatestdecrease, "0.00%")
Next gd

'loop for greatest volume
Dim greastestvolume As Double
greatestvolume = Cells(2, 12).Value
Dim greatestvolumeticker As String

For gv = 2 To RowRange2
 If Cells(gv, 12).Value >= greastestvolume Then
    greastestvolume = Cells(gv, 12).Value
    greatestvolumeticker = Cells(gv, 9).Value
 End If
 Cells(4, 16).Value = greatestvolumeticker
 Cells(4, 17).Value = Format(greastestvolume, "##0.0E+0")
Next gv
    
    
    Next ws
 
End Sub


