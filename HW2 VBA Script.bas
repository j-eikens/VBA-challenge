Attribute VB_Name = "Module1"
Sub Homework2():
                        
    For Each ws In Worksheets
               
    Dim ticker As String
    Dim total_stock_volume As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_open As Double
    Dim stock_close As Double
    Dim opening_date As Long
    Dim closing_date As Long
    Dim red As Integer
    Dim green As Integer
    Dim white As Integer
      
    
    total_stock_volume = 0
    yearly_change = 0
    percent_change = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    red = 3
    green = 4
    white = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Create column labels
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Cell formating
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2", "Q3").NumberFormat = "0.00%"
       
    
    For i = 2 To LastRow
    
        'Check if cell is equal to proceeding cell
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set ticker name
            ticker = ws.Cells(i, 1).Value
        
            'Add total_stock_volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                       
            'Printing unique tickers
            ws.Range("I" & Summary_Table_Row).Value = ticker
        
            'Print total_stock_volume
            ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
            
        
            'Reset total_stock_volume
            total_stock_volume = 0
                        
            'Adds row to summary table
            Summary_Table_Row = Summary_Table_Row + 1
        
            'MsgBox (ticker)
            
            stock_close = ws.Cells(i, 6).Value
         
            yearly_change = stock_close - stock_open
            
            percent_change = (stock_close - stock_open) / stock_open
            
            ws.Range("J" & Summary_Table_Row - 1).Value = yearly_change
            
            ws.Range("K" & Summary_Table_Row - 1).Value = percent_change

                          
        Else
            
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

                      
        End If
        
        'find stock_open
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            stock_open = ws.Cells(i, 3).Value
            
       
        End If
             
    Next i
    
    For i = 2 To LastRow
  
        If ws.Cells(i, 10) > 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = green
        
        ElseIf ws.Cells(i, 10) < 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = red
            
        Else
            
            ws.Cells(i, 10).Interior.ColorIndex = white
        
        
        End If
    
    Next i
    
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestTot As Double
        
    greatestInc = 0
    greatestDec = 0
    greatestTot = 0
    
    For i = 2 To LastRow
        
        If ws.Cells(i, 11).Value > greatestInc Then
            
            greatestInc = ws.Cells(i, 11).Value
            
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
            
            ws.Range("P2").Value = ws.Cells(i, 9).Value

            
        End If
        
        If ws.Cells(i, 11).Value < greatestDec Then
        
            greatestDec = ws.Cells(i, 11).Value
            
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
            
            ws.Range("P3").Value = ws.Cells(i, 9).Value

                        
        End If
        
        If ws.Cells(i, 12).Value > greatestTot Then
        
            greatestTot = ws.Cells(i, 12).Value
            
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
            
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            
        End If
                
    Next i
   
   'MsgBox ActiveWorkbook.Worksheets(ws).Name
   
   Next ws
   
End Sub




