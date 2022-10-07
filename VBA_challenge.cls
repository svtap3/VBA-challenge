VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim Ticker As Long
        Dim PercentageChng As Double
        Dim LastA As Long
        Dim LastI As Long
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
                
       WorksheetName = ws.Name
                     
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
       
       ws.Cells(1, 16).Value = "Ticker"
       ws.Cells(1, 17).Value = "Value"
       ws.Cells(2, 15).Value = "Greatest % Increase"
       ws.Cells(3, 15).Value = "Greatest % Decrease"
       ws.Cells(4, 15).Value = "Greatest Total Voume"
         
       
        Ticker = 2
        
        j = 2
        
        LastA = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
            For i = 2 To LastA
            
            'Check for change in ticker
            If ws.Cells(i + 1, "A").Value <> ws.Cells(i, 1).Value Then
            
            ws.Cells(Ticker, "I").Value = ws.Cells(i, 1).Value
            
            ws.Cells(Ticker, "J").Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
            If ws.Cells(Ticker, "J").Value < 0 Then
                            
                'Red
                    ws.Cells(Ticker, "J").Interior.ColorIndex = 3
                
            Else
                'Green
                    ws.Cells(Ticker, "J").Interior.ColorIndex = 4
                    
            End If
            
            If ws.Cells(j, 3).Value <> 0 Then
                                
                PercentageChng = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                   
                ws.Cells(Ticker, "K").Value = Format(PercentageChng, "Percent")
                    
                Else
                    
                ws.Cells(Ticker, "K").Value = Format(0, "Percent")
                    
                End If
                                            
            ws.Cells(Ticker, "L").Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            Ticker = Ticker + 1
            
            j = i + 1
            
            End If
            
        Next i
        
    LastI = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    GreatestVolume = ws.Cells(2, 12).Value
    GreatestIncrease = ws.Cells(2, 11).Value
    GreatestDecrease = ws.Cells(2, 11).Value
   
    
        For i = 2 To LastI
    
                        
                If ws.Cells(i, 12).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestVolume = GreatestVolume
                
                End If
                
                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestIncrease = GreatestIncrease
                
                End If
                
                If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestDecrease = GreatestDecrease
                
                End If
    
        ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
        ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
        ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
        
        Next i
        
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
                    
    Next ws
               
        
      
        
End Sub



