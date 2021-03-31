Attribute VB_Name = "Stocks_Ticker"
Sub Stocks_Ticker()

' Variables
Dim ticker_Symbol
Dim yearlyChange
Dim percentChange
Dim totalStockVolume As Double
Dim openPrice
Dim closePrice
Dim summary_Table_Row
Dim yearStart
Dim wsCount As Integer
Dim Total_Ticker_Volume As Double
Total_Ticker_Volume = 0

' hard challange variables
Dim greatestIncreaseTicker
Dim greatestDecreaseTicker
Dim greatestVolumeTicker
Dim greatestIncreaseValue
Dim greatestDecreaseValue
Dim greatestVolumeValue

greatestIncreaseTicker = "temp"
greatestDecreaseTicker = "temp"
greatestVolumeTicker = "temp"
greatestIncreaseValue = 0
greatestDecreaseValue = 0
greatestVolumeValue = 0

'worksheet iterate
wsCount = ActiveWorkbook.Worksheets.Count

'Lastrow = Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
summary_Table_Row = 2

For ws = 1 To wsCount
    ' format worksheet by adding colomns
    Worksheets(ws).Range("I1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly Change"
    Worksheets(ws).Range("K1") = "Percent Change"
    Worksheets(ws).Range("L1") = "Total Stock Volume"
    
    'Hard Solution formatting
     Worksheets(ws).Range("P1") = "Ticker"
     Worksheets(ws).Range("Q1") = "Value"
     Worksheets(ws).Range("O2") = "Greatest % Increase"
     Worksheets(ws).Range("O3") = "Greatest % Decrease"
     Worksheets(ws).Range("O4") = "Greatest Total Volume"
     Worksheets(ws).Range("Q2:Q3").NumberFormat = "0.00%"
     
     Lastrow = Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
    summary_Table_Row = 2
    
    
    For i = 2 To Lastrow
    
        ' set ticker_Symbol and begin stock volume increment
        ticker_Symbol = Worksheets(ws).Cells(i, 1)
        
    
        
        'set opening price
        If openPrice = "" Then
            openPrice = Worksheets(ws).Cells(i, 3)
        End If
      
        
        'iterate
        If ticker_Symbol <> Worksheets(ws).Cells((i + 1), 1) Then
        
        'set close price
            closePrice = Worksheets(ws).Cells(i, 6)
        'calc yearly change
            yearlyChange = closePrice - openPrice
            
    
         ' Add to the Ticker name total volume
         Total_Ticker_Volume = Total_Ticker_Volume + Worksheets(ws).Cells(i, 7).Value
        
        ' ticker_Symbol output to worksheet
            Worksheets(ws).Range("I" & summary_Table_Row).Value = ticker_Symbol
        
        ' yearly change output to worksheet with formatting
            Worksheets(ws).Range("J" & summary_Table_Row).Value = yearlyChange
       If yearlyChange > 0 Then
             Worksheets(ws).Range("J" & summary_Table_Row).Interior.ColorIndex = 4 'Green
       Else
            Worksheets(ws).Range("J" & summary_Table_Row).Interior.ColorIndex = 3 'Red
       End If
        
        ' percent changed output to worksheet with formatting
        If openPrice <> 0 Then
        
            percentChange = (yearlyChange / openPrice)
            
        Else
            percentChange = 0
       End If
        Worksheets(ws).Range("K" & summary_Table_Row).Value = percentChange
        Worksheets(ws).Range("K" & summary_Table_Row).NumberFormat = "0.00%"
        
        ' stock volume output to worksheet
    
        Worksheets(ws).Range("L" & summary_Table_Row).Value = Total_Ticker_Volume
        
        
        
             'hard check greatest volume
        If Total_Ticker_Volume > greatestVolumeValue Then
        
            
            greatestVolumeValue = Total_Ticker_Volume
            greatestVolumeTicker = ticker_Symbol
        End If
        
            'hard greatest increase/decrease %
        If percentChange > greatestIncreaseValue Then
            greatestIncreaseValue = percentChange
            greatestIncreaseTicker = ticker_Symbol
        ElseIf percentChange < greatestDecreaseValue Then
            greatestDecreaseValue = percentChange
            greatestDecreaseTicker = ticker_Symbol
        End If
    
        
      
      summary_Table_Row = summary_Table_Row + 1
      closePrice = 0
      openPrice = Worksheets(ws).Cells(i + 1, 3).Value
         
   
      Total_Ticker_Volume = 0
      
      Else
        ' Increase the Total Ticker Volume
        Total_Ticker_Volume = Total_Ticker_Volume + Worksheets(ws).Cells(i, 7).Value
     
      End If
   
    
    Next i
  
    
    
      'output hard challange
    Worksheets(ws).Range("P2") = greatestIncreaseTicker
    Worksheets(ws).Range("Q2") = greatestIncreaseValue
    Worksheets(ws).Range("P3") = greatestDecreaseTicker
    Worksheets(ws).Range("Q3") = greatestDecreaseValue
    Worksheets(ws).Range("P4") = greatestVolumeTicker
    Worksheets(ws).Range("Q4") = greatestVolumeValue

    'reset hard variables
    greatestIncreaseTicker = "temp"
    greatestDecreaseTicker = "temp"
    greatestVolumeTicker = "temp"
    greatestIncreaseValue = 0
    greatestDecreaseValue = 0
    greatestVolumeValue = 0
     
    Worksheets(ws).Cells.EntireColumn.AutoFit
    Worksheets(ws).Cells.EntireRow.AutoFit

Next ws



End Sub




