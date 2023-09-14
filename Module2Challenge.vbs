Sub Module2Challenge()
   
    'This makes all following actions take place on
    'every worksheet in the woorkbook
        For Each ws In Worksheets
        ws.Activate
    
    'Column names
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
                    
       
    'This is where I want to put an aditional functionality
    'that summarizes my data from above
    
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

    'Tells script where to stop counting
    
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        
    'More Variables
     
        totalvolume = 0
        openprice = Cells(2, "C").Value
        beginning = 2
        
    'This is my for loop
    
        For i = 2 To lastrow
       
    'These are my if/else statments

        If Cells(i, "A").Value = Cells(i + 1, "A").Value Then
               totalvolume = totalvolume + Cells(i, "G").Value

        Else
               totalvolume = totalvolume + Cells(i, "G").Value
               closeprice = Cells(i, "F").Value
               yearlychange = closeprice - openprice
               
        If openprice <> 0 Then
               
               percentchange = yearlychange / openprice * 100
               openprice = Cells(i + 1, "C").Value
               
        End If
                
               Cells(beginning, "I").Value = Cells(i, "A").Value
               Cells(beginning, "J").Value = yearlychange
               Cells(beginning, "K").Value = "%" & percentchange
               Cells(beginning, "L").Value = totalvolume
               
    'These are my fomratting scripts
      
        If yearlychange > 0 Then
                Range("J" & beginning).Interior.Color = vbGreen
        ElseIf yearlychange < 0 Then
                Range("J" & beginning).Interior.Color = vbRed
        Else
                Range("J" & beginning).Interior.Color = vbWhite
                   
    'Ends it all if "these" conditions are met
     
        End If
               totalvolume = 0
               openprice = Cells(i + 1, "C").Value
               beginning = beginning + 1

        End If
           
    ' Check for greatest % increase
        If percentchange > greatestIncrease Then
                greatestIncrease = percentchange
                tickerGreatestIncrease = Cells(i, "A").Value
        End If
            
    ' Check for greatest % decrease
        If percentchange < greatestDecrease Then
                greatestDecrease = percentchange
                tickerGreatestDecrease = Cells(i, "A").Value
        End If
            
    ' Check for greatest total volume
        If totalvolume > greatestVolume Then
                greatestVolume = totalvolume
                tickerGreatestVolume = Cells(i, "A").Value
        End If
           
        Next i
       
    ' Outputs the greatest values to the summary section
                ws.Range("P2").Value = tickerGreatestIncrease
                ws.Range("Q2").Value = greatestIncrease / 100
                ws.Range("P3").Value = tickerGreatestDecrease
                ws.Range("Q3").Value = greatestDecrease / 100
                ws.Range("P4").Value = tickerGreatestVolume
                ws.Range("Q4").Value = greatestVolume
       
       
        Next ws
        MsgBox ("*imagine R2D2 noise*")

End Sub


