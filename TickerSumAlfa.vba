Sub Ticker():

    Dim Ticker As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Total_St_Vol As Double
    Dim R As Long
    Dim SumR As Long
    Dim Pct As Double
    Dim Delta_Price As Double
    Dim Min_Pct As Double
    Dim Max_Pct As Double
    Dim Max_Total As Double
    Dim Max_Ticker, Min_Ticker, Max_Tot_Ticker As String
    
    
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
                        Pct = 0
                    Else
                        Pct = Delta_Price / Opening_price
                        .Range("L" & SumR).Value = Pct
                        If Pct > 0 Then
                            .Range("L" & SumR).Interior.Color = vbGreen
                        Else
                            .Range("L" & SumR).Interior.Color = vbRed
                            
                        End If
                                                
                    End If
                    If Pct > Max_Pct Then
                        Max_Pct = Pct
                        Max_Ticker = Ticker
                    End If
                    If Pct < Min_Pct Then
                        Min_Pct = Pct
                        Min_Ticker = Ticker
                    End If
                    If Total_St_Vol > Max_Total Then
                        Max_Total = Total_St_Vol
                        Max_Tot_Ticker = Ticker
                    End If
                                                              
                    ' Reset Total_St_Vol
                    
                    Total_St_Vol = 0
                    Opening_price = .Cells(R + 1, 3).Value
                End If
            
            Next R
            .Range("P2").Value = "Greatest % Increase"
            .Range("P3").Value = "Greatest % Decrease"
            .Range("P4").Value = "Greatest Total Volume"
            .Range("Q1").Value = "Ticker"
            .Range("R1").Value = "Value"
            
            .Range("Q2").Value = Max_Ticker
            .Range("R2").Value = Max_Pct
            .Range("R2").NumberFormat = "0.00%"
            .Range("Q3").Value = Min_Ticker
            .Range("R3").Value = Min_Pct
            .Range("R3").NumberFormat = "0.00%"
            .Range("Q4").Value = Max_Tot_Ticker
            .Range("R4").Value = Max_Total
            
             
            .Columns("J:R").AutoFit
        End With
    Next ws
    
End Sub
      



