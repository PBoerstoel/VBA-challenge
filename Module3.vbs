Attribute VB_Name = "Module3"
Sub Stock()
For Each ws In Worksheets

    Dim numrows As Double
    numrows = ws.Range("A1").End(xlDown).Row
    
    Dim ticker As String
    ticker = ws.Cells(2, 1).Value
    
    Dim ychange As Double
    Dim pchange As Double
    Dim stockvol As Double
    Dim numcom As Double
    numcom = 1
    stockvol = 0
    Dim openp As Double
    openp = ws.Cells(2, 3).Value
    
    ws.Cells(2, 9).Value = ticker
    
    For i = 2 To numrows
        If ws.Cells(i, 1).Value = ticker Then
            stockvol = stockvol + ws.Cells(i, 7).Value
            ychange = ws.Cells(i, 6).Value - openp
            pchange = ychange / openp
            ws.Cells(numcom + 1, 10).Value = ychange
            If ychange < 0 Then
                ws.Cells(numcom + 1, 10).Interior.ColorIndex = 3
            ElseIf ychange > 0 Then
                ws.Cells(numcom + 1, 10).Interior.ColorIndex = 4
            End If
            ws.Cells(numcom + 1, 11).Value = pchange
            If pchange < 0 Then
                ws.Cells(numcom + 1, 11).Interior.ColorIndex = 3
            ElseIf pchange > 0 Then
                ws.Cells(numcom + 1, 11).Interior.ColorIndex = 4
            End If
            ws.Cells(numcom + 1, 12).Value = stockvol
        Else
            numcom = numcom + 1
            ticker = ws.Cells(i, 1).Value
            openp = ws.Cells(i, 3).Value
            stockvol = 0
            stockvol = stockvol + ws.Cells(i, 7).Value
            ychange = ws.Cells(i, 6).Value - openp
            pchange = ychange / openp
            ws.Cells(numcom + 1, 9).Value = ticker
            ws.Cells(numcom + 1, 10).Value = ychange
            If ychange < 0 Then
                ws.Cells(numcom + 1, 10).Interior.ColorIndex = 3
            ElseIf ychange > 0 Then
                ws.Cells(numcom + 1, 10).Interior.ColorIndex = 4
            End If
            ws.Cells(numcom + 1, 11).Value = pchange
            If pchange < 0 Then
                ws.Cells(numcom + 1, 11).Interior.ColorIndex = 3
            ElseIf pchange > 0 Then
                ws.Cells(numcom + 1, 11).Interior.ColorIndex = 4
            End If
            ws.Cells(numcom + 1, 12).Value = stockvol
        End If
    Next i
    Dim maxchange As Double
    Dim minchange As Double
    Dim maxstock As Double
    maxchange = ws.Cells(2, 11).Value
    minchange = ws.Cells(2, 11).Value
    maxstock = ws.Cells(2, 12).Value
    Dim maxcticket As String
    Dim mincticker As String
    Dim maxsticket As String
    
    
    For i = 2 To numcom
        If Cells(i, 11).Value > maxchange Then
            maxchange = ws.Cells(i, 11).Value
            maxcticket = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < minchange Then
            minchange = ws.Cells(i, 11).Value
            mincticket = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12) > maxstock Then
            maxstock = ws.Cells(i, 12)
            maxsticket = ws.Cells(i, 9).Value
        End If
    Next i
    
    ws.Cells(2, 16).Value = maxcticket
    ws.Cells(3, 16).Value = mincticket
    ws.Cells(4, 16).Value = maxsticket
    ws.Cells(2, 17).Value = maxchange
    ws.Cells(3, 17).Value = minchange
    ws.Cells(4, 17).Value = maxstock
Next ws
End Sub
