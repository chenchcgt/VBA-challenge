Sub stock()

For Each Sheet In Worksheets

Dim symbol As String
Dim yearchange As Double
Dim volume As Double
Dim counter As Double
Dim percentchange As Double

'declare values

lastrow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
lastcol = Sheet.Cells(1, Columns.Count).End(xlToLeft).Column
Sheet.Cells(1, 10).Value = "Ticker"
Sheet.Cells(1, 11).Value = "Yearly Change"
Sheet.Cells(1, 12).Value = "Percent Change"
Sheet.Cells(1, 13).Value = "Total Volume"
Sheet.Cells(1, 16).Value = "ticker"
Sheet.Cells(1, 17).Value = "value"
Sheet.Cells(2, 15).Value = "Greatest % increase"
Sheet.Cells(3, 15).Value = "Greatest % decrease"
Sheet.Cells(4, 15).Value = "Greatest total volume"

Row = 2
counter = 0
volume = 0

For i = 2 To lastrow

'compare and caluclate change
If Sheet.Cells(i, 1).Value <> Sheet.Cells(i + 1, 1) Then

symbol = Sheet.Cells(i, 1).Value
Sheet.Cells(Row, 10).Value = symbol


firstvalue = Sheet.Cells(i - (counter), 3).Value
lastvalue = Sheet.Cells(i, 6).Value
yearchange = lastvalue - firstvalue
Sheet.Cells(Row, 11).Value = yearchange

If yearchange < 0.01 Then
Sheet.Cells(Row, 11).Interior.ColorIndex = 3
Else
Sheet.Cells(Row, 11).Interior.ColorIndex = 4
End If

'only calculate when denominator is not zero
If firstvalue <> 0 Then
percentchange = ((lastvalue - firstvalue) / firstvalue) * 100
Sheet.Cells(Row, 12).Value = percentchange
End If

'add volume
volume = volume + Sheet.Cells(i, 7)
Sheet.Cells(Row, 13) = volume

Row = Row + 1
counter = 0
i = i + 1
volume = 0
End If


'Debug.Print (firstvalue)

counter = counter + 1
volume = volume + Sheet.Cells(i, 7)

Next i

'initialization for max decrease,max increase,max volume
i = 2
maxincrease = Sheet.Cells(i, 12).Value
Name = Sheet.Cells(i, 10).Value
maxvolume = Sheet.Cells(i, 13).Value
Namevolume = Sheet.Cells(i, 10).Value
maxdecrease = Sheet.Cells(i, 12).Value
Namedecease = Sheet.Cells(i, 10).Value

For i = 2 To lastrow
'compare increase

If Sheet.Cells(i + 1, 12).Value > maxincrease Then
maxincrease = Sheet.Cells(i + 1, 12).Value
Name = Sheet.Cells(i + 1, 10).Value

End If

'compare decrease
If Sheet.Cells(i + 1, 12).Value < maxdecrease Then
maxdecrease = Sheet.Cells(i + 1, 12).Value
Namedecrease = Sheet.Cells(i + 1, 10).Value

End If

'compare volume
If Sheet.Cells(i + 1, 13).Value > maxvolume Then
maxvolume = Sheet.Cells(i + 1, 13).Value
Namevolume = Sheet.Cells(i + 1, 10).Value

End If

'compute results
Sheet.Cells(2, 17).Value = maxincrease
Sheet.Cells(2, 16).Value = Name

Sheet.Cells(3, 17).Value = maxdecrease
Sheet.Cells(3, 16).Value = Namedecrease

Sheet.Cells(4, 17).Value = maxvolume
Sheet.Cells(4, 16).Value = Namevolume


Next i
Next Sheet
End Sub




