Sub challenge2():
    
    For Each ws In Worksheets
    
        'declare array to store stock info 0-ticker, 1-start date, 2-open, 3-end date, 4-close, 5-volume
        Dim stock(5) As Variant
        
        'declare arrays to store greatests
        Dim increase(1) As Variant
        Dim decrease(1) As Variant
        Dim volume(1) As Variant
        
        'declare row counters and max count
        Dim i, imax, prntrw As Integer
        
        'assign value to counters and max counts
        imax = ws.Cells(Rows.Count, 1).End(xlUp).Row
        prntrw = 2
        
        'set up headers and labels for output
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'initialize array with first row of data
        stock(0) = ws.Cells(2, 1).Value
        stock(1) = ws.Cells(2, 2).Value
        stock(2) = ws.Cells(2, 3).Value
        stock(3) = ws.Cells(2, 2).Value
        stock(4) = ws.Cells(2, 6).Value
        stock(5) = ws.Cells(2, 7).Value
    
        'at each row, compare ticker
        For i = 2 To imax
        
            'if same, compare dates and possibly override dates and open/close, add to volume
            If stock(0) = ws.Cells(i + 1, 1).Value Then
            
                If stock(1) > ws.Cells(i + 1, 2).Value Then
                    stock(1) = ws.Cells(i + 1, 2).Value
                    stock(2) = ws.Cells(i + 1, 3).Value
                End If
                
                If stock(3) < ws.Cells(i + 1, 2).Value Then
                    stock(3) = ws.Cells(i + 1, 2).Value
                    stock(4) = ws.Cells(i + 1, 6).Value
                End If
                
                stock(5) = stock(5) + ws.Cells(i + 1, 7).Value
                
            'if different, print info to row, check if a greatest, increase row count, reset array
            Else
                ws.Range("I" & prntrw).Value = stock(0)
                ws.Range("J" & prntrw).Value = stock(4) - stock(2)
                If ws.Range("J" & prntrw).Value > 0 Then
                    ws.Range("J" & prntrw).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & prntrw).Interior.ColorIndex = 3
                End If
                ws.Range("K" & prntrw).Value = (stock(4) - stock(2)) / stock(2)
                ws.Range("K" & prntrw).NumberFormat = "#.##%"
                ws.Range("L" & prntrw).Value = stock(5)
                
                If (stock(4) - stock(2)) / stock(2) > increase(1) Then
                    increase(0) = stock(0)
                    increase(1) = (stock(4) - stock(2)) / stock(2)
                End If
                
                If (stock(4) - stock(2)) / stock(2) < decrease(1) Then
                    decrease(0) = stock(0)
                    decrease(1) = (stock(4) - stock(2)) / stock(2)
                End If
                
                If stock(5) > volume(1) Then
                    volume(0) = stock(0)
                    volume(1) = stock(5)
                End If
                
                prntrw = prntrw + 1
                
                stock(0) = ws.Cells(i + 1, 1).Value
                stock(1) = ws.Cells(i + 1, 2).Value
                stock(2) = ws.Cells(i + 1, 3).Value
                stock(3) = ws.Cells(i + 1, 2).Value
                stock(4) = ws.Cells(i + 1, 6).Value
                stock(5) = ws.Cells(i + 1, 7).Value
                
            End If
        
        Next i
        
        ws.Range("P2").Value = increase(0)
        ws.Range("Q2").Value = increase(1)
        ws.Range("P3").Value = decrease(0)
        ws.Range("Q3").Value = decrease(1)
        ws.Range("P4").Value = volume(0)
        ws.Range("Q4").Value = volume(1)
        ws.Range("Q2:Q3").NumberFormat = "#.##%"
        ws.Range("Q4").NumberFormat = "#.##E+##"
        
    Next ws

End Sub