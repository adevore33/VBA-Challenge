Attribute VB_Name = "Module1"
Sub tickers()

    Dim ws As Worksheet
    
    For Each ws In Worksheets

    ws.Range("J1").Value = "Ticker"
    
    ws.Range("K1").Value = "Yearly Change"
    
    ws.Range("L1").Value = "Percent Change"
    
    ws.Range("M1").Value = "Total Stock Volume"
    
    ws.Range("Q1").Value = "Ticker"
    
    ws.Range("R1").Value = "Value"
    
    ws.Range("P2").Value = "Greatest % Increase"
    
    ws.Range("P3").Value = "Greatest % Decrease"
    
    ws.Range("P4").Value = "Greatest Total Volume"

    Dim Ticker_Name As String
    
    Dim Volume_Total As Double
    Volume_Total = 0
    
    Dim Summary_Table As Integer
    Summary_Table = 2
    
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    
    Dim Closed_Price As Double
    Closed_Price = Cells(Rows.Count, "F").End(xlDown)
    
    For i = 2 To 759001
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker_Name = ws.Cells(i, 1).Value
            
            Closed_Price = ws.Cells(i, 6).Value
            
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
            ws.Range("J" & Summary_Table).Value = Ticker_Name
            
            ws.Range("M" & Summary_Table).Value = Volume_Total
            
            ws.Range("K" & Summary_Table).Value = Closed_Price - Open_Price
            
            ws.Range("L" & Summary_Table).Value = ((Closed_Price - Open_Price) / Open_Price)
                    
            Summary_Table = Summary_Table + 1
            
            Volume_Total = 0
            
            Open_Price = ws.Cells(i + 1, 3).Value
            
           
        Else
        
            Volume_Total = Volume_Total + Cells(i, 7)
            
        End If
        
    Next i
    
    For j = 2 To 3001
    
        If ws.Cells(j, 11).Value >= 0 Then
        
        ws.Cells(j, 11).Interior.ColorIndex = 4
        
        Else
        
        ws.Cells(j, 11).Interior.ColorIndex = 3

        
        End If

    Next j
    
    ws.Range("L2:L3001").NumberFormat = "0.00%"
    
    Dim Max_Percent_Change As Double
    Max_Percent_Change = WorksheetFunction.Max(ws.Range("L2:L3001"))
    
    Dim Min_Percent_Change As Double
    Min_Percent_Change = WorksheetFunction.Min(ws.Range("L2:L3001"))
    
    Dim Max_Volume_Change As Double
    Max_Volume_Change = WorksheetFunction.Max(ws.Range("M2:M3001"))
    
    For k = 2 To 3001
    
        If ws.Cells(k, 12).Value = Max_Percent_Change Then
        Ticker_Name = ws.Cells(k, 10).Value
        ws.Range("Q2").Value = Ticker_Name
        ws.Range("R2").Value = WorksheetFunction.Max(ws.Range("L2:L3001"))
        
        
        End If
    Next k
    
    For l = 2 To 3001
    
        If ws.Cells(l, 12).Value = Min_Percent_Change Then
        Ticker_Name = ws.Cells(l, 10).Value
        ws.Range("Q3").Value = Ticker_Name
        ws.Range("R3").Value = WorksheetFunction.Min(ws.Range("L2:L3001"))
        
        End If
        
    Next l
        
    For m = 2 To 3001
    
        If ws.Cells(m, 13).Value = Max_Volume_Change Then
        Ticker_Name = ws.Cells(m, 10).Value
        ws.Range("Q4").Value = Ticker_Name
        ws.Range("R4").Value = WorksheetFunction.Max(ws.Range("M2:M3001"))
        
        End If
        
    Next m
    
    
    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Columns("A:R").AutoFit
    
    Next
    
End Sub
