Attribute VB_Name = "TickerSumAlfa"
Sub Ticker():

    Dim Ticker As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Total_St_Vol As Double
    Dim R As Long
    Dim SumR As Long
    
    Range("j1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    
    
    R = 2
    SumR = 2
    
    Ticker = Cells(2, 1).Value
    
    Opening_price = CDbl(Cells(2, 3).Value)
       
   'Ctrl + Shift + Down (Range should be first cell in data set)
     Do Until IsEmpty(Cells(R, 1).Value)
        If Cells(R, 1).Value = Ticker Then
            Total_St_Vol = Total_St_Vol + Cells(R, 7).Value
        Else
        ' Closing ticker group
            Closing_price = Cells(R - 1, 6).Value
            Cells(SumR, 10).Value = Ticker
            Cells(SumR, 11).Value = Closing_price - Opening_price
            If Cells(SumR, 11).Value = 0 Then
                Cells(SumR, 12).Value = "0.00%"
            Else
                Cells(SumR, 12).Value = Round((Opening_price / Cells(SumR, 11).Value), 2) & "%"
            End If
            Cells(SumR, 13).Value = Total_St_Vol
            SumR = SumR + 1
            
        'Opening new ticker group
            Ticker = Cells(R, 1).Value
            Opening_price = Cells(R, 3).Value
            Total_St_Vol = Cells(R, 7).Value
         End If
         
         R = R + 1
         
    Loop
    Closing_price = Cells(R - 1, 6).Value
    Cells(SumR, 10).Value = Ticker
    Cells(SumR, 11).Value = Closing_price - Opening_price
    Cells(SumR, 12).Value = Round((Opening_price / Cells(SumR, 11).Value), 2) & "%"
    Cells(SumR, 13).Value = Total_St_Vol
    
    MsgBox "All done! Tickers Processed " & SumR
        
End Sub
