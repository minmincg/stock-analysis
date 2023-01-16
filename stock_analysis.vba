Sub stockchange()

For Each ws In Worksheets

    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim ticker As String
    Dim currentRow As Integer
    Dim stockvol As Double
    Dim openval As Double
    Dim closeval As Double
    Dim changeval As Double

    
    currentRow = 2
    openval = Cells(2, "C").Value
    maxval = 0
  
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "$ change"
  ws.Range("K1").Value = "percentage change"
  ws.Range("L1").Value = "total stock vol"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  ws.Range("O2").Value = "greatest % "
  ws.Range("O3").Value = "lowest %"
  ws.Range("O4").Value = "max stock volume"
  
    For I = 2 To lastRow
        If ws.Cells(I, "A").Value <> ws.Cells(I + 1, "A").Value Then 'cambiamos de ticker
        'stop comparing and change ticker
            
            closeval = ws.Cells(I, 6).Value
            
            ticker = ws.Cells(I, "A").Value
            changeval = closeval - openval
            percentclose = closeval / openval
            percentopen = 1
            percentchange = percentclose - percentopen
            stockvol = stockvol + ws.Cells(I, "G").Value
            
            ws.Cells(currentRow, "I").Value = ticker
            ws.Cells(currentRow, "J").Value = changeval
            ws.Cells(currentRow, "K").Value = FormatPercent(percentchange)
            ws.Cells(currentRow, "L").Value = stockvol
            
        If percentchange < 0 Then
             ws.Cells(currentRow, "J").Interior.ColorIndex = 3
        Else
            ws.Cells(currentRow, "J").Interior.ColorIndex = 4
        End If
            
            currentRow = currentRow + 1
            tickerchange = 0
            stockvol = 0
            openval = ws.Cells(I + 1, "C").Value
            
        
        Else 'siguen siendo iguales
        'add stock volume
        
        stockvol = stockvol + ws.Cells(I, "G").Value
        
        
        End If
    Next I
    currentRow = 2
    
    For I = 2 To lastRow
       
       If maxval < ws.Cells(I, "K").Value Then
        
                maxval = ws.Cells(I, "K").Value
                maxticker = ws.Cells(I, "I").Value
        
      ElseIf minval > ws.Cells(I, "K").Value Then
      
                    minval = ws.Cells(I, "K").Value
                    minticker = ws.Cells(I, "I").Value
       
       End If
       
                    ws.Cells(currentRow, "P").Value = maxticker
                    ws.Cells(currentRow, "Q").Value = FormatPercent(maxval)
                    
                    ws.Cells(3, "P").Value = minticker
                    ws.Cells(3, "Q").Value = FormatPercent(minval)
    
    If maxstock < ws.Cells(I, "L").Value Then
        
                    maxstock = ws.Cells(I, "L").Value
                    maxstocktick = ws.Cells(I, "I").Value
        
    End If
                    ws.Cells(4, "P").Value = maxstocktick
                    ws.Cells(4, "Q").Value = maxstock
        
        
       
    Next I
    
Next ws
End Sub

