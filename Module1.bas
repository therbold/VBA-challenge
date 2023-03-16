Attribute VB_Name = "Module1"
Sub stockcalc()

Dim WS As Worksheet
    For Each WS In Worksheets
    
    Dim name As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim Volume As Double
    Volume = 0
    Dim sumrow As Integer
    sumrow = 2
    
    Cells(1, "i").Value = "Name"
    Cells(1, "j").Value = "Yeary Change"
    Cells(1, "k").Value = "Percent Change"
    Cells(1, "l").Value = "Volume"
    Cells(2, "n").Value = "Greatest % Increase"
    Cells(3, "n").Value = "Greatest % Decrease"
    Cells(4, "n").Value = "Greatest Volume"
    Cells(1, "o").Value = "Name"
    Cells(1, "p").Value = "Value"
            
    Dim lastrow As Long
    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
                  
    Dim i As Long
    
    openprice = Cells(2, 3).Value
       
    For i = 2 To lastrow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        name = Cells(i, 1).Value
        closeprice = Cells(i, 6).Value
        yearchange = closeprice - openprice
        percentchange = (closeprice - openprice) / openprice
        Volume = Volume + Cells(i, 7).Value
        
        Cells(sumrow, 9).Value = name
        Cells(sumrow, 10).Value = yearchange
        Cells(sumrow, 11).Value = percentchange
        Cells(sumrow, 11).NumberFormat = "0.00%"
        Cells(sumrow, 12).Value = Volume
        sumrow = sumrow + 1
        
        End If
        
        Next i
        
        'If Cells(i, 10) < 0 Then Cells(i, 10).Interior.ColorIndex = 3
        'ElseIf Cells(i, 10) > 0 Then Cells(i, 10).Interior.ColorIndex = 4
        'End If

     If Cells(i, 11).Value = WorksheetFunction.Max(WS.Range("k2:k" & lastrow)) Then
        Cells(2, 15).Value = Cells(i, 9).Value
        Cells(2, 16).Value = Cells(i, 11).Value
    'ElseIf Cells(i, 11).Value = Worksheetfuntion.Min(WS.Range("k2:k" & lastrow)) Then
        Cells(3, 15).Value = Cells(i, 9).Value
        Cells(3, 16).Value = Cells(i, 11).Value
    If Cells(i, 11).Value = Worksheetfuntion.Max(WS.Range("l2:l" & lastrow)) Then
        Cells(3, 15).Value = Cells(i, 9).Value
        Cells(3, 16).Value = Cells(i, 11).Value
    End If
    
        Cells(2, 16).NumberFormat = "o.oo%"
        Cells(3, 16).NumberFormat = "o.oo%"
     
Next WS

     
End Sub
