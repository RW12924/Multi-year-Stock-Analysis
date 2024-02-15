Attribute VB_Name = "Module1"
Sub Analysis():
    Dim ticker As String
    Dim OpenValue, CloseValue, change, percent, volume, tickrow As Integer
    Dim GVolume, PercentInc, PercentDec As Double
    
    For Each ws In Worksheets
    
    OpenValue = ws.Cells(2, 3).Value
    tickrow = 2
    volume = 0
    PercentInc = -100
    PercentDec = 100
    GVolume = 0
    
        For Row = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        
            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
            
                ticker = ws.Cells(Row, 1).Value
                ws.Cells(tickrow, 9).Value = ticker
                
                CloseValue = ws.Cells(Row, 6).Value
                change = CloseValue - OpenValue
                ws.Cells(tickrow, 10).Value = change
                
                percent = (change / OpenValue)
                ws.Cells(tickrow, 11).Value = percent
                
                volume = volume + ws.Cells(Row, 7).Value
                ws.Cells(tickrow, 12).Value = volume
                volume = 0
                
                OpenValue = ws.Cells(Row + 1, 3).Value
                tickrow = tickrow + 1
                
            Else
            
                volume = volume + ws.Cells(Row, 7).Value
                
            End If
            
        Next Row
        
        For Row = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            
            If ws.Cells(Row, 10).Value > 0 Then
            
                ws.Cells(Row, 10).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(Row, 10).Value < 0 Then
            
                ws.Cells(Row, 10).Interior.ColorIndex = 3
                
            End If
            
            If ws.Cells(Row, 11).Value > 0 Then
            
                ws.Cells(Row, 11).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(Row, 11).Value < 0 Then
            
                ws.Cells(Row, 11).Interior.ColorIndex = 3
                
            End If
        
            If PercentInc < ws.Cells(Row, 11).Value Then
            
                PercentInc = ws.Cells(Row, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(Row, 1).Value
                ws.Cells(2, 17).Value = PercentInc
                
            End If
            
            If PercentDec > ws.Cells(Row, 11).Value Then
            
                PercentDec = ws.Cells(Row, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(Row, 1).Value
                ws.Cells(3, 17).Value = PercentDec
                
            End If
    
            If GVolume < ws.Cells(Row, 12).Value Then
            
                GVolume = ws.Cells(Row, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(Row, 1).Value
                ws.Cells(4, 17).Value = GVolume
                
            End If
            
        Next Row
    
    Next ws
    
End Sub

