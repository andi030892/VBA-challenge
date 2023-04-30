Sub Challenge2()

    For Each ws In Worksheets
    
'create variables

    Dim tickersymbol As String
    Dim yearlychange As Double
    Dim totalstockvolume As Double
    Dim percentchange As Double
    Dim newticker As Integer
    Dim worksheet_count As Integer
    Dim count As Double
    Dim endingprice As Double
    Dim startingprice As Double
    Dim lastrow As Long
    Dim lastvalue As Long
    
    
    
'set values
    
    
    percentchange = 0
    totalstockvolume = 0
    newticker = 2
    count = 0

  
   
'find last row

    lastrow = ws.Cells(Rows.count, "A").End(xlUp).Row
    
    
    lastvalue = ws.Cells(Rows.count, "K").End(xlUp).Row
    

    
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
    
    
    startingprice = ws.Cells(2, 3).Value

  
'row loop
        
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                tickersymbol = ws.Cells(i, 1).Value
                totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
                endingprice = ws.Cells(i, 6).Value
                

'calculate yearly change and percent change
                
                
                yearlychange = endingprice - startingprice
            If startingprice <> 0 Then
                    percentchange = yearlychange / startingprice
                Else
                    percentchange = 0
            End If
                
'update starting price for next ticker symbol
                
                
                startingprice = ws.Cells(i + 1, 3).Value
    
    If yearlychange < 0 Then
          ws.Range("J" & newticker).Interior.ColorIndex = 3
        Else
          ws.Range("J" & newticker).Interior.ColorIndex = 4
    End If
    
    
        ws.Range("I" & newticker).Value = tickersymbol
        ws.Range("J" & newticker).Value = yearlychange
        ws.Range("K" & newticker).Value = percentchange
        ws.Range("K" & newticker).NumberFormat = "0.00%"
        ws.Range("L" & newticker).Value = totalstockvolume
    
    
    
    newticker = newticker + 1
    totalstockvolume = 0
    
    
    Else
    
    totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
    count = count + 1





End If

Next i


' Find the greatest percent increase, percent decrease, and total volume

Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double
Dim max_increase_ticker As String
Dim max_decrease_ticker As String
Dim max_volume_ticker As String

max_increase = ws.Cells(2, 11).Value
max_decrease = ws.Cells(2, 11).Value
max_volume = ws.Cells(2, 12).Value

For i = 2 To lastvalue
    ' Find the greatest percent increase
    If ws.Cells(i, 11).Value > max_increase Then
        max_increase = ws.Cells(i, 11).Value
        max_increase_ticker = ws.Cells(i, 9).Value
    End If
    
    ' Find the greatest percent decrease
    If ws.Cells(i, 11).Value < max_decrease Then
        max_decrease = ws.Cells(i, 11).Value
        max_decrease_ticker = ws.Cells(i, 9).Value
    End If
    
    ' Find the greatest total volume
    If ws.Cells(i, 12).Value > max_volume Then
        max_volume = ws.Cells(i, 12).Value
        max_volume_ticker = ws.Cells(i, 9).Value
    End If
Next i

' Output the results to the sheet

ws.Range("P2").Value = max_increase_ticker
ws.Range("Q2").Value = max_increase
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("P3").Value = max_decrease_ticker
ws.Range("Q3").Value = max_decrease
ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("P4").Value = max_volume_ticker
ws.Range("Q4").Value = max_volume



Next ws



End Sub


