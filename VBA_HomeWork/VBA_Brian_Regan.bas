Attribute VB_Name = "VBA_Brian_Regan"
Sub worksheetloopnewhw():
    Dim LastRow As Long
    Dim CounttestStart As Long
    Dim CounttestEnd As Long
    Dim ChangeInValue As Long
    Dim CountTotalVolume As Long
    Dim TotalVolume As Double
    Dim WhileCounter As Long
    Dim countrows As Long
    Dim countrowsPercent As Long
    Dim maxpercent As Double
    Dim maxDecrease As Double
    Dim maxTotalVolume As Double
    Dim countRowsVolume As Long
    Dim current As Worksheet
    
 For Each current In Worksheets
   
    
    LastRow = current.UsedRange.Rows.Count
    
    
    CounttestStart = 1
    CounttestEnd = 1
    CountTotalVolume = 2
    WhileCounter = 2
    WhileTotalVolumnRowCount = 2
    
    current.Cells(1, 10).Value = "Ticker"
    current.Cells(1, 11).Value = "Begining Price"
    current.Cells(1, 12).Value = "Ending Price"
    current.Cells(1, 13).Value = "Yearly Change"
    current.Cells(1, 14).Value = "Percent Change"
    current.Cells(1, 15).Value = "Total Stock Volume"
    current.Cells(2, 18).Value = "Greatest % Increase"
    current.Cells(3, 18).Value = "Greatest % Decrease"
    current.Cells(4, 18).Value = "Greatest Total Volume"
    current.Cells(1, 19).Value = "Ticker"
    current.Cells(1, 20).Value = "Value"
    
    For i = 1 To LastRow
        maxpercent = 0
        maxDecrease = 0
        If current.Cells(i, 1).Value <> current.Cells(i + 1, 1).Value Then 'finds first row of new ticker
        
            CounttestStart = 1 + CounttestStart 'keeps count of new rows for ticker
          
            
            current.Cells(Str(CounttestStart), 10) = current.Cells(i + 1, 1).Value 'gets ticker value
            current.Cells(Str(CounttestStart), 11) = current.Cells(i + 1, 3).Value 'gets begining price
            
            
        ElseIf current.Cells(i + 1, 1).Value <> current.Cells(i + 2, 1).Value Then 'finds last row of ticker
            CounttestEnd = 1 + CounttestEnd 'keeps track of ticker row
            
            current.Cells(Str(CounttestEnd), 12) = current.Cells(i + 1, 6).Value 'gets close value and populates 12th column
            current.Cells(Str(CounttestStart), 13).Value = current.Cells(Str(CounttestEnd), 12) - current.Cells(Str(CounttestStart), 11) 'calculates yearly change
                If current.Cells(CounttestEnd, 13).Value >= 0 Then
                    current.Cells(CounttestEnd, 13).Interior.ColorIndex = 4 'updates fill of cell to green
                Else
                    current.Cells(CounttestEnd, 13).Interior.ColorIndex = 3 'updates fill of cell to red
                End If
                
            If current.Cells(Str(CounttestStart), 11).Value > 0 And current.Cells(Str(CounttestStart), 12).Value > 0 Then
                current.Cells(Str(CounttestStart), 14).Value = (current.Cells(Str(CounttestStart), 13).Value / current.Cells(Str(CounttestStart), 11).Value) ' calculates percent change
                
                   
                   
                         
            Else
                current.Cells(Str(CounttestStart), 14).Value = 0
                
            End If
            current.Cells(Str(CounttestStart), 14).NumberFormat = "0.00%" 'updates cells to percent
        End If
        
                        

        
    Next i
    
    
    countrows = current.Cells(Rows.Count, 10).End(xlUp).Row 'counts all ticker rows
    
    For m = 2 To countrows 'starts at row 2 of ticker column in J
   
        Do While Not IsNull(current.Cells(m, 10).Value) 'loops while there is a value in ticker column
            TotalVolume = 0
            Do While current.Cells(Str(WhileCounter), 1).Value = current.Cells(m, 10).Value 'loops through each ticker while = to the new ticker column value
                WhileCounter = 1 + WhileCounter
                TotalVolume = TotalVolume + current.Cells(Str(WhileCounter), 7).Value 'adds up the volume
                current.Cells(m, 15).Value = TotalVolume 'updates the new volume column for that ticker
            Loop
            If current.Cells(Str(WhileCounter), 1).Value <> current.Cells(m, 10).Value Then 'breaks out of counting volumne for current ticker
                Exit Do
            End If
            
        
        Loop
    
    Next m
    
    countrowsPercent = current.Cells(Rows.Count, 14).End(xlUp).Row 'gets a count of all rows for percent change
    
    maxpercent = 0
    maxDecrease = 0
   
    For i = 2 To countrowsPercent 'iterates through the count
        If current.Cells(i, 14) > maxpercent Then 'determines if the new row is greater than maxpercent
            maxpercent = current.Cells(i, 14).Value 'finds greatest % increase
            ticket = current.Cells(i, 10)
            
            current.Cells(2, 19).Value = ticket
            current.Cells(2, 20).Value = maxpercent
            current.Cells(2, 20).NumberFormat = "0.00%"
            
        ElseIf current.Cells(i, 14) < maxDecrease Then 'finds the largest decrease %
            maxDecrease = current.Cells(i, 14)
            ticket = current.Cells(i, 10)
            current.Cells(3, 19).Value = ticket
            current.Cells(3, 20).Value = maxDecrease
            current.Cells(3, 20).NumberFormat = "0.00%"
        End If
    Next i
    countRowsVolume = current.Cells(Rows.Count, 15).End(xlUp).Row
    maxTotalVolume = 0
    
    For v = 2 To countRowsVolume 'finds the ticker with the highest total volume
        If current.Cells(v, 15) > maxTotalVolume Then
            maxTotalVolume = current.Cells(v, 15)
            ticketvolume = current.Cells(v, 10)
            current.Cells(4, 19).Value = ticketvolume
            current.Cells(4, 20).Value = maxTotalVolume
        End If
    Next v
       
  current.Cells.EntireColumn.AutoFit ' updates all columns to fit all cells characters
  Next
End Sub
