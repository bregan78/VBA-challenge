Attribute VB_Name = "Module2"
Sub newhw():
    Dim LastRow As Long
    Dim CounttestStart As Long
    Dim CounttestEnd As Long
    Dim ChangeInValue As Long
    Dim CountTotalVolume As Long
    Dim TotalVolume As Double
    Dim WhileCounter As Long
    Dim countrows As Long
    Dim countrowsPercent As Long
    Dim maxPercent As Long
    Dim maxDecrease As Long
    Dim maxTotalVolume As Double
    Dim countRowsVolume As Long
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
    
    CounttestStart = 1
    CounttestEnd = 1
    CountTotalVolume = 2
    WhileCounter = 2
    WhileTotalVolumnRowCount = 2
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Begining Price"
    Cells(1, 12).Value = "Ending Price"
    Cells(1, 13).Value = "Yearly Change"
    Cells(1, 14).Value = "Percent Change"
    Cells(1, 15).Value = "Total Stock Volume"
    Cells(2, 18).Value = "Greatest % Increase"
    Cells(3, 18).Value = "Greatest % Decrease"
    Cells(4, 18).Value = "Greatest Total Volume"
    Cells(1, 19).Value = "Ticket"
    Cells(1, 20).Value = "Value"
    
    For i = 1 To LastRow
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then 'finds first row of new ticker
        
            CounttestStart = 1 + CounttestStart 'keeps count of new rows for ticker
            Cells(Str(CounttestStart), 10) = Cells(i + 1, 1).Value 'gets ticker value
            Cells(Str(CounttestStart), 11) = Cells(i + 1, 3).Value 'gets begining price
            
            
        ElseIf Cells(i + 1, 1).Value <> Cells(i + 2, 1).Value Then 'finds last row of ticker
            CounttestEnd = 1 + CounttestEnd 'keeps track of ticker row
            
            Cells(Str(CounttestEnd), 12) = Cells(i + 1, 6).Value 'gets close value and populates 12th column
            Cells(Str(CounttestStart), 13).Value = Cells(Str(CounttestEnd), 12) - Cells(Str(CounttestStart), 11) 'calculates yearly change
                If Cells(CounttestEnd, 13).Value >= 0 Then
                    Cells(CounttestEnd, 13).Interior.ColorIndex = 4
                Else
                    Cells(CounttestEnd, 13).Interior.ColorIndex = 3
                End If
                
                    
            Cells(Str(CounttestStart), 14).Value = (1 - (Cells(Str(CounttestStart), 11).Value / Cells(Str(CounttestStart), 12).Value)) * 100 ' calculates percent change
        
        End If
        
                        

        
    Next i
    
    
    countrows = Cells(Rows.Count, 10).End(xlUp).Row 'counts all ticker rows
    
    For m = 2 To countrows 'starts at row 2 of ticker column in J
   
        Do While Not IsNull(Cells(m, 10).Value)
            TotalVolume = 0
            Do While Cells(Str(WhileCounter), 1).Value = Cells(m, 10).Value
                WhileCounter = 1 + WhileCounter
                TotalVolume = TotalVolume + Cells(Str(WhileCounter), 7).Value
                Cells(m, 15).Value = TotalVolume
            Loop
            If Cells(Str(WhileCounter), 1).Value <> Cells(m, 10).Value Then
                Exit Do
            End If
            
        
        Loop
    
    Next m
    
    countrowsPercent = Cells(Rows.Count, 14).End(xlUp).Row
    
    maxPercent = 0
    maxDecrease = 0
   
    For i = 2 To countrowsPercent
        If Cells(i, 14) > maxPercent Then
            maxPercent = Cells(i, 14)
            ticket = Cells(i, 10)
            
            Cells(2, 19).Value = ticket
            Cells(2, 20).Value = maxPercent
        ElseIf Cells(i, 14) < maxDecrease Then
            maxDecrease = Cells(i, 14)
            ticket = Cells(i, 10)
            Cells(3, 19).Value = ticket
            Cells(3, 20).Value = maxDecrease
        End If
    Next i
    countRowsVolume = Cells(Rows.Count, 15).End(xlUp).Row
    maxTotalVolume = 0
    
    For v = 2 To countRowsVolume
        If Cells(v, 15) > maxTotalVolume Then
            maxTotalVolume = Cells(v, 15)
            ticketvolume = Cells(v, 10)
            Cells(4, 19).Value = ticketvolume
            Cells(4, 20).Value = maxTotalVolume
        End If
    Next v
        'For i = 1 To LastRow
            'Do While Cells(i, 1).Value = Cells(m, 10).Value
                'WhileCounter = 1 + WhileCounter
                'TotalVolumn = 0
                'TotalVolumn = TotalVolumn + Cells(Str(WhileCounter), 7).Value
                'Cells(Str(WhileTotalVolumnRowCount), 15).Value = TotalVolumn
                'If IsNull(Cells(Str(WhileCounter), 1).Value) Then
                    'Exit Do
                'End If
                
            'Loop
            
          'Next i
     'Next m
        
        'TotalVolumn = 0
End Sub
