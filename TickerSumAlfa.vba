Sub Ticker():

    Dim Ticker As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Total_St_Vol As Double
    Dim R As Long
    Dim SumR As Long
    Dim Pct As Double
    Dim Delta_Price As Double
    
    For Each ws In Worksheets
    
        With ws
            .Range("J1").Value = "Ticker"
            .Range("K1").Value = "Yearly Change"
            .Range("L1").Value = "Percent Change"
            .Range("M1").Value = "Total Stock Volume"
            
            SumR = 1
            
            Opening_price = .Cells(2, 3).Value
                    
            LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
                
            For R = 2 To LastRow
                             
                Total_St_Vol = Total_St_Vol + .Cells(R, 7).Value
                 
                If .Cells(R, 1).Value <> .Cells(R + 1, 1).Value Then
                    SumR = SumR + 1
                    Ticker = .Cells(R, 1).Value
            
                    'Print rows in the summary table
                    .Range("J" & SumR).Value = Ticker
                    .Range("M" & SumR).Value = Total_St_Vol
                    
                    Closing_price = .Cells(R, 6).Value
                    Delta_Price = Closing_price - Opening_price
                   
                    .Range("K" & SumR).Value = Delta_Price
              
                    'avoid division by zero
                    
                    .Range("L" & SumR).NumberFormat = "0.00%"
                    If Opening_price = 0 Then
                        .Range("L" & SumR).Value = 0
                        .Range("L" & SumR).Font.Color = vbBlue
                    Else
                        Pct = Delta_Price / Opening_price
                        .Range("L" & SumR).Value = Pct
                        If Pct > 0 Then
                            .Range("L" & SumR).Interior.Color = vbGreen
                        Else
                            .Range("L" & SumR).Interior.Color = vbRed
                            
                        End If
                                                
                    End If
                                          
                    ' Reset Total_St_Vol
                    
                    Total_St_Vol = 0
                    Opening_price = .Cells(R + 1, 3).Value
                End If
            
            Next R
            
            .Columns("J:M").AutoFit
        End With
    Next ws
    
End Sub
      
   
    
    


