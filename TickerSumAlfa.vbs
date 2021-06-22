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
    
		ws.Range("J1").Value = "Ticker"
		ws.Range("K1").Value = "Yearly Change"
		ws.Range("L1").Value = "Percent Change"
		ws.Range("M1").Value = "Total Stock Volume"
		
		SumR = 1
                Max_Pct = 0
                Min_Pct = 0
                Max_Total = 0		
		
		Opening_price = ws.Cells(2, 3).Value
				
		LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
			
		For R = 2 To LastRow
						 
			Total_St_Vol = Total_St_Vol + ws.Cells(R, 7).Value
			 
			If ws.Cells(R, 1).Value <> ws.Cells(R + 1, 1).Value Then
				SumR = SumR + 1
				Ticker = ws.Cells(R, 1).Value
		
				'Print rows in the summary table
				ws.Range("J" & SumR).Value = Ticker
				ws.Range("M" & SumR).Value = Total_St_Vol
				
				Closing_price = ws.Cells(R, 6).Value
				Delta_Price = Closing_price - Opening_price
			   
				ws.Range("K" & SumR).Value = Delta_Price
		  
				'avoid division by zero
				
				ws.Range("L" & SumR).NumberFormat = "0.00%"
				If Opening_price = 0 Then
					ws.Range("L" & SumR).Value = 0
					ws.Range("L" & SumR).Font.Color = vbBlue
					Pct = 0
				Else
					Pct = Delta_Price / Opening_price
					ws.Range("L" & SumR).Value = Pct
					If Pct > 0 Then
						ws.Range("L" & SumR).Interior.Color = vbGreen
					Else
						ws.Range("L" & SumR).Interior.Color = vbRed
						
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
				Opening_price = ws.Cells(R + 1, 3).Value
			End If
		
		Next R
		ws.Range("P2").Value = "Greatest % Increase"
		ws.Range("P3").Value = "Greatest % Decrease"
		ws.Range("P4").Value = "Greatest Total Volume"
		ws.Range("Q1").Value = "Ticker"
		ws.Range("R1").Value = "Value"
		
		ws.Range("Q2").Value = Max_Ticker
		ws.Range("R2").Value = Max_Pct
		ws.Range("R2").NumberFormat = "0.00%"
		ws.Range("Q3").Value = Min_Ticker
		ws.Range("R3").Value = Min_Pct
		ws.Range("R3").NumberFormat = "0.00%"
		ws.Range("Q4").Value = Max_Tot_Ticker
		ws.Range("R4").Value = Max_Total
		
		 
		ws.Columns("J:R").AutoFit
	Next ws
    
End Sub
