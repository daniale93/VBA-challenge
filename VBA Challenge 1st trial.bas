Attribute VB_Name = "Module1"
Sub StockMarket()

For Each ws In Worksheets
    Dim WorksheetName As String
    Dim ticker As String
    Dim YrChange As Double
   Dim PcChange As Double
    Dim Volume As Double
    
    Dim Summary_Table_Row As Long
    Dim FirstOpenPrice As Double
   Dim LastClosePrice As Double
    Dim lastrow As Long
    
    Volume = 0
    Summary_Table_Row = 2
    
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
     ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
       ticker = ws.Cells(i, 1).Value
        
          Volume = Volume + ws.Cells(i, 7).Value
        LastClosePrice = ws.Cells(i, 6).Value
        YrChange = LastClosePrice - FirstOpenPrice
        
        If FirstOpenPrice = 0 Then
        PcChange = 0
        Else
                
        PcChange = (LastClosePrice / FirstOpenPrice) - 1
        End If
         ws.Range("I" & Summary_Table_Row).Value = ticker
       ws.Range("L" & Summary_Table_Row).Value = Volume
        ws.Range("J" & Summary_Table_Row).Value = YrChange
        ws.Range("K" & Summary_Table_Row).Value = PcChange
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        If PcChange >= 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        
        If PcChange > ws.Cells(2, 17).Value Then
        
        ws.Cells(2, 16).Value = ticker
        ws.Cells(2, 17).Value = PcChange
        Else
        End If
          
        If PcChange < ws.Cells(3, 17).Value Then
        
        ws.Cells(3, 16).Value = ticker
        ws.Cells(3, 17).Value = PcChange
        Else
        End If
        
        If Volume > ws.Cells(4, 17).Value Then
        
        ws.Cells(4, 16).Value = ticker
        ws.Cells(4, 17).Value = Volume
        Else
        End If
        
           Summary_Table_Row = Summary_Table_Row + 1
        
        Volume = 0
        
        
     ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
     
     FirstOpenPrice = ws.Cells(i, 3).Value
     
        
        
        Volume = ws.Cells(i, 7).Value
    Else
    
    Volume = Volume + ws.Cells(i, 7).Value
    
           
        End If
    
 Next i
    
        Next ws
        
        
        
        
    
End Sub
