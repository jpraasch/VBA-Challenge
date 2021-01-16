Attribute VB_Name = "Module2"
Sub Stockchecker2()
     
     'Each Sheet setup
    For Each WS In Worksheets
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = WS.Cells(1, Columns.Count).End(xlToLeft).Column
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
  
    'Screen Stocks section
     
    'Setting where things go and are
    Dim Stock_Name As String
    Dim Stock_Vol As LongPtr
        Stock_Vol = 0
    Dim SummaryTbl As Integer
        SummaryTbl = 2
    Dim openStock As Double
    Dim closeStock As Double
    Dim yearchange As Double
    Dim percentchange As Double
   
  openStock = Cells(2, 3).Value
   
For i = 2 To LastRow
             
   'Screen the stocks and input into new column
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        
        'Set the Stock name,add Stock Volume
        Stock_Name = WS.Cells(i, 1).Value
        Stock_Vol = Stock_Vol + WS.Cells(i, 7).Value
        
        'finding closing stock
        closeStock = WS.Cells(i, 6).Value
                
        'finding yearly change
        yearchange = closeStock - openStock
        ' Enter Year Change
        WS.Range("J" & SummaryTbl).Value = yearchange
        
        'Input Stock Volume
        WS.Range("L" & SummaryTbl).Value = Stock_Vol
        WS.Range("I" & SummaryTbl).Value = Stock_Name
        
        'finding percent change
        If openStock <> 0 Then
            percentchange = yearchange / openStock
            Else
            percentchange = 0
             WS.Range("K" & SummaryTbl).Value = percentchange
             WS.Range("K" & SummaryTbl).Style = "Percent"
             End If
            
        'Adding into new column
        SummaryTbl = SummaryTbl + 1
        Stock_Vol = 0
        
        'reset openStock
        openStock = WS.Cells(i + 1, 3)
        
        Else
        Stock_Vol = Stock_Vol + WS.Cells(i, 7).Value
                      
        End If
            
Next i
          
     'Conditional Formatting
     For i = 2 To LastRow
        If WS.Cells(i, 10).Value > 0 Then
        WS.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf WS.Cells(i, 10).Value < 0 Then
        WS.Cells(i, 10).Interior.ColorIndex = 3
        
      
      'autofit column width for stock totals *in every worksheet*
        Columns("L").ColumnWidth = 16
        Columns("K").ColumnWidth = 13
        Columns("J").ColumnWidth = 13
        End If
        Next i
  
  Next WS
    
    
End Sub
