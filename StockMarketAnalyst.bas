Attribute VB_Name = "Module1"
Sub StockMarketAnalyst()

    Dim Sum As Double
    Dim SumHolder As Integer
    
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    
    Dim ChangePrice As Double
    Dim PercentChange As Double
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVol As Double
    
    
    SumHolder = 2
    Sum = 0
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    'MsgBox (lastRow)
    
    ' Placeholder for first price in list. If OpenPrice = -100000 then it's the first price
    OpenPrice = -100000
    
    
    ' Loop through rows in the column
    For i = 2 To lastRow

        Sum = Sum + Cells(i, 7)
        
        ' If the the price is the first one, then copy what is in the cell as the opening price
        If OpenPrice = -100000 Then
            OpenPrice = Cells(i, 3)
            'MsgBox ("OpenPrice " + Str(OpenPrice))
        End If
        
      ' Searches for when the value of the next cell is different than that of the current cell
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            Cells(SumHolder, 9) = Cells(i, 1)
            Cells(SumHolder, 12) = Sum
            Sum = 0

            
            ClosingPrice = Cells(i, 6)
            
            ChangePrice = OpenPrice - ClosingPrice
            
            Cells(SumHolder, 10) = ChangePrice
            
            If OpenPrice > ClosingPrice Then
                Cells(SumHolder, 10).Interior.ColorIndex = 3
            ElseIf ClosingPrice > OpenPrice Then
                Cells(SumHolder, 10).Interior.ColorIndex = 4
            End If
            
        
            If ChangePrice = 0 Then
                PercentChange = 0
            ' If OpenPrice is zero, then set percent change as zero, otherwise we get a divide by zero error.
            ElseIf OpenPrice = 0 Then
                PercentChange = 0
            Else
            
                PercentChange = ChangePrice / OpenPrice
            End If
            
            Cells(SumHolder, 11) = PercentChange
            Cells(SumHolder, 11).NumberFormat = "0.00%"
            
            OpenPrice = -100000
            
            SumHolder = SumHolder + 1
        
        
        End If

    Next i


    lastRowSummary = Cells(Rows.Count, "I").End(xlUp).Row + 1

    GreatestIncrease = Range("K2")
    GreatestDecrease = Range("K2")
    GreatestVol = Range("L2")
    
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    
    Range("O2") = Range("I1")
    Range("O3") = Range("I1")
    Range("O4") = Range("I1")
    Range("P2") = Range("K1")
    Range("P3") = Range("K1")
    Range("P4") = Range("L1")
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
    ' Loop through the summary table section to find the minimum/maximum/greatest volume
    For j = 3 To lastRowSummary

        ' Check if there is a new GreatestIncrease
        If Cells(j, 11) > GreatestIncrease Then
            GreatestIncrease = Cells(j, 11)
            Cells(2, 16) = Cells(j, 9)
            Cells(2, 17) = Cells(j, 11)
        End If
        
        ' Check for new GreatestDecrease
        If Cells(j, 11) < GreatestDecrease Then
            GreatestDecrease = Cells(j, 11)
            Cells(3, 16) = Cells(j, 9)
            Cells(3, 17) = Cells(j, 11)
        End If

        ' Check for new GreatestVolume
        If Cells(j, 12) > GreatestVol Then
            GreatestVol = Cells(j, 12)
            Cells(4, 16) = Cells(j, 9)
            Cells(4, 17) = Cells(j, 12)
        End If


    Next j


End Sub

