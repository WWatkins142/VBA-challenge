Attribute VB_Name = "Module1"
Sub StockTicker()
   ' define worksheet variable
    Dim ws As Worksheet
    ' loop through all of the worksheets in Excel Workbook
    For Each ws In Worksheets
    ws.Activate
    
        ' Variable to hold the ticker name
        Dim tickerName As String
        
        ' Variable to hold Total Volume
        Dim totalVol As LongLong
        totalVol = 0
        
        ' Variable to hold Yearly Change
        Dim yearChange As Double
        
        ' Variable to hold Percent Change
        Dim perChange As Double
        
        ' variable to hold open and closing values
        Dim stockClose As Double
        
        Dim stockOpen As Double
        
        ' Set opening price
        stockOpen = 0
        
        ' summary table row
        Dim summaryTableRow As Integer
        summaryTableRow = 2
    
        '   to get the last row
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
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
                ' start of the year stock price
                If stockOpen = 0 Then
                    stockOpen = ws.Cells(Row, 3).Value
                End If
                
                ' Set closing price
                stockClose = ws.Cells(Row, 6).Value
                
                ' calculate yearly change
                yearChange = stockClose - stockOpen
                
                ' add to summary table
                ws.Range("J" & summaryTableRow).Value = yearChange
                
                'calculate percent change
                If (stockOpen = 0 And stockClose = 0) Then
                    perChange = 0
                ElseIf (stockOpen = 0 And stockClose <> 0) Then
                    perChange = 1
                Else
                    perChange = yearChange / stockOpen
                    ' add to summary table
                    ws.Range("K" & summaryTableRow).Value = perChange
                    ' format percent change to percentage
                    ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                End If
                
                ' format to highlight positive change in green and negative change in red
                If perChange > 0 Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                 Else
                     ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                End If
                ' add the Total Stock Volume to column L on the current summary table row
                ws.Range("L" & summaryTableRow).Value = totalVol
                
                ' once the summary table is populated, then add one to the summary row count
                summaryTableRow = summaryTableRow + 1
                 
                ' then reset the Total Stock Volume to 0
                totalVol = 0
               
               ' reset opening price
               stockOpen = 0
                
            Else
                'if we are in the same Ticker Name, add on to the running total
                totalVol = totalVol + ws.Range("G" & Row).Value
                
            End If
                    
        Next Row
        
        
    Next ws
    
End Sub
