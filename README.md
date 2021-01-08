# VBA-Challenge
Sub Stockchecker()


    For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Range("H1").EntireColumn.Insert
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Yearly Change"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"
    Next ws
    
End Sub
