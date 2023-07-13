Attribute VB_Name = "Module1"
Sub Stock()

Dim numrows As Double
numrows = Range("A1").End(xlDown).Row

Dim ticker As String
ticker = Cells(2, 1).Value

Dim ychange As Double
Dim pchange As Double
Dim stockvol As Double
Dim numcom As Double
numcom = 1
stockvol = 0
Dim openp As Double
openp = Cells(2, 3).Value

Cells(2, 9).Value = ticker

For i = 2 To numrows
    If Cells(i, 1).Value = ticker Then
        stockvol = stockvol + Cells(i, 7).Value
        ychange = Cells(i, 6).Value - openp
        pchange = ychange / openp
        Cells(numcom + 1, 10).Value = ychange
        Cells(numcom + 1, 11).Value = pchange
        Cells(numcom + 1, 12).Value = stockvol
    Else
        numcom = numcom + 1
        ticker = Cells(i, 1).Value
        openp = Cells(i, 3).Value
        stockvol = 0
        stockvol = stockvol + Cells(i, 7).Value
        ychange = Cells(i, 6).Value - openp
        pchange = ychange / openp
        Cells(numcom + 1, 9).Value = ticker
        Cells(numcom + 1, 10).Value = ychange
        Cells(numcom + 1, 11).Value = pchange
        Cells(numcom + 1, 12).Value = stockvol
    End If
Next i
Dim maxchange As Double
Dim minchange As Double
Dim maxstock As Double
maxchange = Cells(2, 11).Value
minchange = Cells(2, 11).Value
maxstock = Cells(2, 12).Value
Dim maxcticket As String
Dim mincticker As String
Dim maxsticket As String


For i = 2 To numcom
    If Cells(i, 11).Value > maxchange Then
        maxchange = Cells(i, 11).Value
        maxcticket = Cells(i, 9).Value
    End If
    If Cells(i, 11).Value < minchange Then
        minchange = Cells(i, 11).Value
        mincticket = Cells(i, 9).Value
    End If
    If Cells(i, 12) > maxstock Then
        maxstock = Cells(i, 12)
        maxsticket = Cells(i, 9).Value
    End If
Next i

Cells(2, 16).Value = maxcticket
Cells(3, 16).Value = mincticket
Cells(4, 16).Value = maxsticket
Cells(2, 17).Value = maxchange
Cells(3, 17).Value = minchange
Cells(4, 17).Value = maxstock

End Sub
