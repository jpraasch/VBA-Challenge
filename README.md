Sub Stockchecker()
     
     'Each Sheet setup
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("H1").EntireColumn.Insert
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Yearly Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Format Changes
    For i = 2 To LastRow
        For j = 10 To LastColumn
              ws.Cells(i, j).Style = "Percent"
        Next j
    Next i
           
    'Screen Stocks
    
    'Setting where things go and are
    Dim Stock_Name As String
    Dim Stock_Vol As LongLong
        Stock_Vol = 0
    Dim SummaryName As Integer
        SummaryName = 8
    Dim SummaryVol As Integer
        SummaryVol = 11
        
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Stock_Name = ws.Cells(i, 1).Value
        Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
        Range("H" & SummaryName).Value = Stock_Name
        Range("K" & SummaryVol).Value = Stock_Vol
        
        SummaryName = SummaryName + 1
        SummaryVol = SummaryVol + 1
        
        Stock_Vol = 0
        
        Else
        Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
        
        End If
    Next i
    
