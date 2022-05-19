Attribute VB_Name = "Module1"
Sub stock_analysis()

    Dim ticker As String
        
    Dim total_stock_volume As Double
    total_stock_volume = 0
    
    Dim yearly_change As Double
    
    yearly_change = 0
    
    Dim percent_change As Double
    percent_change = 0

  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To lastrow

   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      

    total_stock_volume = total_stock_volume + Cells(i, 7).Value
    
    ticker = Cells(i, 1).Value
    
    
    yearly_change = Cells(i, 5).Value - Cells(i + 1, 3).Value
    
    
    percent_change = yearly_change / Cells(i + 1, 3).Value
      
    Range("I" & Summary_Table_Row).Value = ticker
   
    Range("L" & Summary_Table_Row).Value = total_stock_volume
    
    Range("J" & Summary_Table_Row).Value = yearly_change
    
    Range("K" & Summary_Table_Row).Value = "%" & percent_change
    
   
    Summary_Table_Row = Summary_Table_Row + 1
        
    total_stock_volume = 0

      
    total_stock_volume = total_stock_volume + Cells(i, 7).Value
      
    yearly_change = Cells(i, 5).Value - Cells(i + 1, 3).Value
    
    percent_change = yearly_change / Cells(i + 1, 3).Value
    
    End If
    
    
    
    
    If Cells(i, 10) > 0 Then
    
    Cells(i, 10).Interior.ColorIndex = 4
    
    Else
    
    Cells(i, 10).Interior.ColorIndex = 3

    End If
    


  Next i

End Sub

