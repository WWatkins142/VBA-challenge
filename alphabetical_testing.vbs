Attribute VB_Name = "Module1"
Sub StockTicker():
    ' Variable to hold the ticker name
    Dim tickerName As String
    
    ' Variable to hold Total Volume
    Dim totalVol As LongLong
    totalVol = 0
    
    ' Variable to hold Yearly Change
    Dim yearlyChange As Double
    
    ' Variable to hold Percent Change
    Dim percentChange As Double
    
    ' summary table row
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    '   loop through all of the worksheets in Excel Workbook
    For Each ws In Worksheets
    
        '   to get the last row
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '   Add the ticker to column I (maybe this can be removed)
       ws.Range("I1").EntireColumn.Insert
        
        '   Add Ticker header to cell "I1"
        ws.Range("I1").Value = "Ticker"
        
        '   Add Yearly Change header to cell "J1"
        ws.Range("J1").Value = "Yearly Change"
        
        '   Add Percent Change header to cell "K1"
        ws.Range("K1").Value = "Percent Change"
        
        '   Add Total Stock Volume to cell "L1"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' loop through all of the ticker name rows.
        For Row = 2 To lastRow
            
            ' check to see if we are in the same Ticker Name
            If (ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value) Then
                
                ' Set/Reset the Ticker name
                tickerName = ws.Range("A" & Row).Value
                
                ' Add to the Total Stock Volume before changing the Ticker Name
                totalVol = totalVol + ws.Range("G" & Row).Value
                
                ' add the values to the summary table
                ' add the Ticker name to Column I on the current summary table row
                ws.Range("I" & summaryTableRow).Value = tickerName
                ' add Yearly Change to Column J on the current summary table row
                
                ' add Percent Change to Column K on the current summary table row
                
                ' add the Total Stock Volume to column L on the current summary table row.
                ws.Range("L" & summaryTableRow).Value = totalVol
                
                ' once the summary table is populated, then add one to the summary row count
                 summaryTableRow = summaryTableRow + 1
                ' then reset the brand total to 0
                brandTotal = 0
                
            Else
                'if we are in the same Ticker Name, add on to the running total
                totalVol = totalVol + ws.Range("G" & Row).Value
                
            End If
                
        '   Add the Ticker name to all of the rows
        ' ws.Range("I2:I" & lastRow).Value = tickerName
         
         
        Exit For
    
    Next Row
    
  Next ws
    
    
End Sub
