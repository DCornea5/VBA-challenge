
Sub Worksheet1()

For Each ws In Worksheets
Worksheets(ws.Name).Activate
Call ticker
Next ws

End Sub



'bring unique ticker on same sheet
Sub ticker()
        
    ' Set an initial variable for holding the ticker symbol and values
    Dim Ticker_Symbol As String
    Dim Op As Double
    Dim Cl As Double
    Dim PrChange As Long
    Dim YrChange As Long
    Dim PercentChange As Integer
    Dim Closing As Long
    Dim Opening As Long
    Dim GreatestValue As Double
    Dim TickerV As String
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
      
    
    
    
    
    ' Set an initial variable for holding Total Stock Volume
    Dim TotalStockVolume As Double
    TotalStockVolume = 0
               
    ' Keep track of the location for each ticker symbol in the summary table
    Dim Stock_Summary_Row As Integer
    Stock_Summary_Row = 2
    
    'Last row
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    ' Loop through all ticker symbols
    TickerSymbol = Cells(2, 1)
    
        For i = 1 To RowCount
        
            ' Check if we are still within the same ticker symbol, if it is not...
            If i = 1 Then
            
                Op = Cells(2, 3).Value
            
                        
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' Set the Ticker symbol
                Ticker_Symbol = Cells(i, 1).Value
                
                ' Print the Ticker Symbol in the summary table
                Range("I" & Stock_Summary_Row).Value = Ticker_Symbol
                
                'Add to the Total Stock Volume
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                                                
                'Print the TotalStockVolume to the summary table
                Range("L" & Stock_Summary_Row).Value = TotalStockVolume
                
                'Cl, Op, Yearly Change balance
                Cl = Cells(i, 6).Value
                Op = Cells(i + 1, 3).Value
                
                Range("J" & Stock_Summary_Row - 1).Value = Range("N" & Stock_Summary_Row - 1).Value - Range("M" & Stock_Summary_Row - 1).Value
                


                
                              
                'Print cl, op  balance
                Range("N" & Stock_Summary_Row).Value = Cl
                Range("M" & Stock_Summary_Row + 1).Value = Op
                Cells(2, 13).Value = Cells(2, 3).Value
               
                    
                ' Add one to the summary table row
                Stock_Summary_Row = Stock_Summary_Row + 1
                
                'Reset TotalStockVolume
                TotalStockVolume = 0
                
                                             
                'If next row has different ticker
                Else
                
                'Add to total stock
                TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                Cl = Cells(i + 1, 6).Value
                                
                YrChange = Cl - Op
                                
                                                     
            End If
            
        Next i
        Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        Range("O2:O2").Value = Array("Greatest % Increase")
        Range("O3:O3").Value = Array("Greatest % Decrease")
        Range("O4:O4").Value = Array("Greatest Total Value")
        Range("Q1:P1").Value = Array("Ticker", "Volume")
        Range("M1:N1").Value = Array("Opening", "Closing")
        Range("I1:Q1").Font.Bold = True
        Range("I:P").Columns.AutoFit
        Range("J" & Stock_Summary_Row).Value = YrChange
                
        
             'get percentage change
             
             lastrow = Cells(Rows.Count, "I").End(xlUp).Row
             
              For k = 2 To lastrow
              
              If Cells(k, 13).Value <> 0 Then
                                               
                 Cells(k, 11).Value = (Cells(k, 14).Value - Cells(k, 13)) / Cells(k, 13).Value
                 
                Else: Cells(k, 13).Value = 0
                End If
            
                Next k
                
                Range("K:K").NumberFormat = ("0.00%")
                
                
                 
            'get greatest value of ticker volume
             
             GreatesVaule = Cells(2, 12).Value
             TickerV = Cells(2, 9).Value
             For g = 2 To lastrow
             
             If Cells(g, 12).Value > GreatestValue Then
             GreatestValue = Cells(g, 12).Value
             TickerV = Cells(g, 9).Value
             
             End If
             Next g
             Cells(4, 17) = GreatestValue
             Cells(4, 16) = TickerV
             
             'get greatest increase of percent change
             
             
             GreatesIncrease = Cells(2, 11).Value
             TickerV = Cells(2, 9).Value
             For h = 2 To lastrow
             
             If Cells(h, 11).Value > GreatestIncrease Then
             GreatestIncrease = Cells(h, 11).Value
             TickerV = Cells(h, 9).Value
             
             End If
             Next h
             Cells(2, 17) = GreatestIncrease
             Cells(2, 16) = TickerV
             Range("Q2").NumberFormat = "0.00%"
             
             
             'get greatest decrease of percent change
             
             
             GreatesDecrease = Cells(2, 11).Value
             TickerV = Cells(2, 9).Value
             For d = 2 To lastrow
             
             If Cells(d, 11).Value < GreatestDecrease Then
             GreatestDecrease = Cells(d, 11).Value
             TickerV = Cells(d, 9).Value
             
             End If
             Next d
             Cells(3, 17) = GreatestDecrease
             Cells(3, 16) = TickerV
             Range("Q3").NumberFormat = "0.00%"
             
             
            'color formatting
            Dim c As Integer
            For c = 2 To lastrow
            If Cells(c, 10).Value > 0 Then
            Cells(c, 10).Interior.Color = vbGreen
            ElseIf Cells(c, 10).Value < 0 Then
                       
            Cells(c, 10).Interior.Color = vbRed
            Else
            
            Cells(c, 10).Interior.Color = vbnone
            
            End If
            Next c
            
          
              
End Sub



     
            
                
    
                





