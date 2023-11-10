Sub stock()

Dim ws As Worksheet


For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percenta Change"
ws.Range("L1").Value = "Total Stock Value"

ws.Range("p1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"




Dim TickerName As String
Dim TickerVolume As Double
Dim Lastrow As Double


TickerName = " "
TickerVolume = 0

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row



Dim openprice As Double
Dim closeprice As Double
Dim pricechange As Double
Dim percentchange As Double
Dim Tickerrow As Double

Tickerrow = 2

openprice = 0
closeprice = 0
pricechange = 0
percentchange = 0

Dim i As Double
Dim j As Double
Dim k As Double



For i = 2 To Lastrow

If i = 2 Then

openprice = ws.Cells(2, 3).Value

End If


If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then


    TickerName = Cells(i, 1).Value
    
    TickerVolume = TickerVolume + Cells(i, 7).Value
    
    
    ws.Cells(Tickerrow, "I").Value = TickerName
    ws.Cells(Tickerrow, "L").Value = TickerVolume
    
    closeprice = Cells(i, 6).Value
    pricechange = (closeprice - openprice)
        
    ws.Cells(Tickerrow, "J").Value = pricechange
        
    
If (openprice = 0) Then
percentchange = 0
        
Else
    
percentchange = pricechange / openprice
        
End If
        
'percentchange = pricechange / openprice
        
ws.Cells(Tickerrow, "K").Value = percentchange
ws.Cells(Tickerrow, "K").NumberFormat = "0.00%"

Tickerrow = Tickerrow + 1

TickerVolume = 0

openprice = Cells(i + 1, 3)

Else

TickerVolume = TickerVolume + Cells(i, 7).Value


End If

Next i



Dim Lastrowinfo As Double
Lastrowinfo = Cells(Rows.Count, 9).End(xlUp).Row


For i = 2 To Lastrowinfo

If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
Next i


For i = 2 To Lastrowinfo

If Cells(i, 11).Value >= Application.WorksheetFunction.Max(Range("K2:K" & Lastrowinfo)) Then
ws.Cells(2, 16).Value = Cells(i, 9).Value
ws.Cells(2, 17).Value = Cells(i, 11).Value
ws.Cells(2, 17).NumberFormat = "0.00%"

ElseIf Cells(i, 11).Value <= Application.WorksheetFunction.Min(Range("K2:K" & Lastrowinfo)) Then
ws.Cells(3, 16).Value = Cells(i, 9).Value
ws.Cells(3, 17).Value = Cells(i, 11).Value
ws.Cells(3, 17).NumberFormat = "0.00%"

ElseIf Cells(i, 12).Value >= Application.WorksheetFunction.Max(Range("L2:L" & Lastrowinfo)) Then
ws.Cells(4, 16).Value = Cells(i, 9).Value
ws.Cells(4, 17).Value = Cells(i, 12).Value


End If


Next i



Next ws

End Sub
