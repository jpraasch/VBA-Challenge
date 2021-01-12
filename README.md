Sub Stockchecker()
     
     'Each Sheet setup
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Yearly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        
    'turn off screen updates (Saw online to speed up performance)
    Application.ScreenUpdating = False
        
    'Format Changes
   ws.Range("J2:J" & LastRow).Style = "Percent"
   
           
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
   
   For i = 2 To LastRow
             
   'Screen the stocks and input into new column
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set the Stock name,add Stock Volume
        Stock_Name = ws.Cells(i, 1).Value
        Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
        
        'Input Stock Volume
        ws.Range("K" & SummaryTbl).Value = Stock_Vol
        ws.Range("H" & SummaryTbl).Value = Stock_Name
        
        'Adding into new column
        SummaryTbl = SummaryTbl + 1
        Stock_Vol = 0
        
        Else
        Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
        
        
        End If
        
        'find the Open Price
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        openStock = ws.Cells(i, 3).Value
        End If
        
        'finding the Closing Price
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        closeStock = ws.Cells(i - 1, 6).Value
        End If
                                
         'finding yearly change
        yearchange = closeStock - openStock
             
        'finding percent change
        If openStock <> 0 Then
            percentchange = yearchange / openStock
            Else
            percentchange = 0
            End If
                       
       ' Enter Year Change
        ws.Range("I" & SummaryTbl).Value = yearchange
           
        'Enter Percent change
        ws.Range("J" & SummaryTbl).Value = percentchange
       
        
        
        
    Next i
     
     'Conditional Formatting
     For i = 2 To LastRow
        If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        
      
      'autofit column width for stock totals *in every worksheet*
        Columns("J").NumberFormat = "0.00"
        Columns("K").AutoFit
        Columns("L").AutoFit
        End If
        Next i
  
  Next ws
    
    'Turning the Screen Update back on (Recommened from same online forums)
    Application.ScreenUpdating = True
    
End Sub

    
