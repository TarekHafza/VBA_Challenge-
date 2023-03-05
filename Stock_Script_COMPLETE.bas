Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    For Each WS In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowH As Long
        Dim LastRowI As Long
      
        Dim PercentChange As Double
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        Dim GreatVolume As Double
        
        Dim TickerName As String
        
        Dim Current As Worksheet
        
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 11).Value = "Total Stock Volume"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 12).Value = "Percent Change"
        WS.Cells(2, 14).Value = "Greatest Increase"
        WS.Cells(3, 14).Value = "Greatest Decrease"
        WS.Cells(4, 14).Value = "Greatest Volume"
        WS.Cells(1, 15).Value = "Ticker"
        WS.Cells(1, 16).Value = "Value"
        
        WS.Range("I:O").EntireColumn.AutoFit
      
        TickCount = 2
        j = 2
        LastRowA = WS.Cells(Rows.Count, 1).End(xlUp).Row
        TickerName = WS.Cells(2, 1).Value
                
        For i = 2 To LastRowA
            If WS.Cells(i, 1).Value <> TickerName Then
                WS.Cells(TickCount, 9).Value = TickerName
                TickerName = WS.Cells(i, 1).Value
                WS.Cells(TickCount, 10).Value = WS.Cells(i - 1, 6).Value - WS.Cells(j, 3).Value
                
                If WS.Cells(TickCount, 10).Value < 0 Then
                    WS.Cells(TickCount, 10).Interior.ColorIndex = 3
                Else
                    WS.Cells(TickCount, 10).Interior.ColorIndex = 4
                End If
            
                If WS.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((WS.Cells(i - 1, 6).Value - WS.Cells(j, 3).Value) / WS.Cells(j, 3).Value)
                    WS.Cells(TickCount, 12).Value = Format(PercentChange, "Percent")
                Else

                WS.Cells(TickCount, 12).Value = Format(0, "Percent")
            End If

                WS.Cells(TickCount, 11).Value = WorksheetFunction.Sum(Range(WS.Cells(j, 7), WS.Cells(i, 7)))
        
                j = i + 1
                TickCount = TickCount + 1
            End If
            
            
        Next i
    
        LastRowI = WS.Cells(Rows.Count, 10).End(xlUp).Row
        
        GreatVolume = WS.Cells(Rows.Count, 11)
        GreatIncrease = WS.Cells(Rows.Count, 12)
        GreatDecrease = WS.Cells(Rows.Count, 12)
    
        For i = 2 To LastRowI
            If WS.Cells(i, 11).Value > GreatVolume Then
                GreatVolume = WS.Cells(i, 11).Value
                WS.Cells(4, 15).Value = WS.Cells(i, 9).Value 
            Else     
                GreatVolume = GreatVolume
            End If
        
            If WS.Cells(i, 12).Value > GreatIncrease Then
                GreatIncrease = WS.Cells(i, 12).Value
                WS.Cells(2, 15).Value = WS.Cells(i, 9).Value
            Else
                GreatIncrease = GreatIncrease
            End If

            If WS.Cells(i, 12).Value < GreatDecrease Then
                GreatDecrease = WS.Cells(i, 12).Value
                WS.Cells(3, 15).Value = WS.Cells(i, 9).Value
            Else
                GreatDecrease = GreatDecrease
            End If
            
            
            WS.Cells(4, 16).Value = Format(GreatVolume, "Scientific")
            WS.Cells(2, 16).Value = Format(GreatIncrease, "Percent")
            WS.Cells(3, 16).Value = Format(GreatDecrease, "Percent")              
           
        Next i
       
    Next WS
        
End Sub
 

