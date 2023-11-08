Attribute VB_Name = "Module1"
Sub Stocks()

For Each ws In Worksheets

Dim YearlyChange As Double
Dim results As Integer
Dim Ticker As String

Dim PercentChange As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim vol As Double
Dim GreatestIncrease, GreatestDecrease, TotalVolume As Double


results = 2
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

OpenPrice = ws.Cells(2, 3).Value
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'heading
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "PercentChnage"
ws.Cells(1, 12).Value = "TotalVolume"


    For i = 2 To LastRow
    
       
       
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ClosePrice = ws.Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                    ws.Cells(results, 11).Value = PercentChange
                Else
                    PercentChange = 0
                End If
            OpenPrice = ws.Cells(i + 1, 3).Value
            ws.Cells(results, 9).Value = Ticker
            ws.Cells(results, 10).Value = YearlyChange
            ws.Cells(results, 11).Value = PercentChange
                If YearlyChange < 0 Then
                    ws.Cells(results, 10).Interior.ColorIndex = 3
                    ws.Cells(results, 11).Interior.ColorIndex = 3
        
                Else
                    ws.Cells(results, 10).Interior.ColorIndex = 4
                    ws.Cells(results, 11).Interior.ColorIndex = 4
                End If
            
         End If
    
        
            If ws.Cells(results, 11).Value > GreatestIncrease Then
                GreatestIncrease = PercentChange
                ws.Cells(2, 14).Value = ws.Cells(i, 1).Value
                ws.Cells(2, 15).Value = GreatestIncrease
                'ws.Cells(2, 15).Value = FormatPercent(GreatestIncrease)
            ElseIf PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                ws.Cells(3, 14).Value = ws.Cells(i, 1).Value
                ws.Cells(3, 15).Value = GreatestDecrease
                'ws.Cells(3, 15).Value = Formatnum(GreatestDecrease)
                
            End If
            
            If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                ws.Cells(4, 14).Value = ws.Cells(i, 1).Value
                ws.Cells(4, 15).Value = GreatestVolume
            
             
            End If
        
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Cells(results, 12).Value = TotalVolume
            TotalVolume = 0
            
               
          results = results + 1
        Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
           

    Next i


Next ws


End Sub


