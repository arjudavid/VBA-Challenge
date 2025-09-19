Sub Alphabettesting()

'Declaring the Variables
Dim i As Long
Dim ticker as String
Dim lastrow As Long
Dim rownumber as Double
Dim openprice as Double
Dim closeprice as Double
Dim totalvolume as Double
Dim quarterlychange as Double
Dim percentagechange as Double
Dim greatestincrease as Double
Dim lowestincrease as Double
Dim totalvolumeincrease as Double
Dim ws as Worksheet
Dim sheetnames as Variant
Dim sheetindex as Integer

'Assigning a last row and a resultrow 
resultrow = 2
sheetnames = Array("Q1", "Q2", "Q3", "Q4")


lastrow = Cells(Rows.Count, "A").End(xlUp).Row


'Looping through data on each worksheet to find each ticker quarterlychange, percentchange, and totalvolume and outputting it on a new table of information
For sheetindex = LBound(sheetnames) To UBound(sheetnames)
    Set ws = ThisWorkbook.Sheets(sheetnames(sheetindex))
    
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastrow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            openprice = ws.Cells(i, 3).Value
            totalvolume = 0
        End If

        totalvolume = totalvolume + ws.Cells(i, 7).Value

        If ws.Cells(i + 1, 1).Value <> ticker Or i = lastrow Then
            closeprice = ws.Cells(i, 6).Value
            quarterlychange = closeprice - openprice
            If openprice <> 0 Then
                percentchange = (quarterlychange / openprice) * 100
            Else
                percentchange = 0
            End If

            ws.Cells(resultrow, 9).Value = ticker
            ws.Cells(resultrow, 10).Value = quarterlychange
            ws.Cells(resultrow, 11).Value = Format(percentchange, "0.00") & "%"
            ws.Cells(resultrow, 12).Value = totalvolume
            resultrow = resultrow + 1
        End If
    Next i

    ' Getting the maximum, lowest increase from the data that's been looped, and the total volume as well
    greatestincrease = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
    rownumber = WorksheetFunction.Match(greatestincrease, ws.Range("K2:K" & lastrow), 0) + 1
    tickerSymbol1 = ws.Cells(rownumber, "I").Value

    lowestincrease = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
    rownumber = WorksheetFunction.Match(lowestincrease, ws.Range("K2:K" & lastrow), 0) + 1
    tickerSymbol2 = ws.Cells(rownumber, "I").Value

    totalvolumeincrease = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    rownumber = WorksheetFunction.Match(totalvolumeincrease, ws.Range("L2:L" & lastrow), 0) + 1
    tickerSymbol3 = ws.Cells(rownumber, "I").Value

    ' Outputting the results
    ws.Cells(2, "N").Value = "Greatest % Increase"
    ws.Cells(2, "O").Value = tickerSymbol1
    ws.Cells(2, "P").Value = greatestincrease

    ws.Cells(3, "N").Value = "Lowest % Increase"
    ws.Cells(3, "O").Value = tickerSymbol2
    ws.Cells(3, "P").Value = lowestincrease

    ws.Cells(4, "N").Value = "Total Volume Increase"
    ws.Cells(4, "O").Value = tickerSymbol3
    ws.Cells(4, "P").Value = totalvolumeincrease

    resultrow = 2 ' Reset resultrow for the next sheet
Next sheetindex

End Sub
