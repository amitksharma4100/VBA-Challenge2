Attribute VB_Name = "Module1"
Sub TickerData()

 For Each ws In Worksheets

    ' Define variables
    Dim lastRow As Long
    Dim Ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableIndex As Integer
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    Dim rng As Range
    


    
    
    ' Define summary table headers
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    ' Find the last row of data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Initialize summary table index
    summaryTableIndex = 2
    
    ' Loop through each row of data
    For i = 2 To lastRow
        
        ' Check if ticker symbol has changed
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
            ' Set ticker symbol
            
            Ticker = ws.Cells(i, 1).Value
            
            ' Set opening price
        
            
            openingPrice = ws.Cells(i, 3).Value
            
            ' Reset total volume
            totalVolume = 0
            
        End If
        
        ' Add to total volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        ' Check if we have reached the end of a ticker's data
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            ' Set closing price
            closingPrice = ws.Cells(i, 6).Value
            
            ' Calculate yearly change
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate percent change
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If
            
            ' Clear summary table
            
            ws.Cells(summaryTableIndex, 10).Clear
            ws.Cells(summaryTableIndex, 11).Clear
            ws.Cells(summaryTableIndex, 12).Clear
            ws.Cells(summaryTableIndex, 13).Clear
            
            ' Output results to summary table
            
            ws.Cells(summaryTableIndex, 10).Value = Ticker
            ws.Cells(summaryTableIndex, 11).Value = yearlyChange
            ws.Cells(summaryTableIndex, 12).Value = percentChange
            ws.Cells(summaryTableIndex, 13).Value = totalVolume
            
            ' Format percent change as percentage
            
            ws.Cells(summaryTableIndex, 12).NumberFormat = "0.00%"
            
            ' Highlight positive yearly change in green and negative yearly change in red
            If yearlyChange > 0 Then
                ws.Cells(summaryTableIndex, 11).Interior.ColorIndex = 4 ' Green
            ElseIf yearlyChange < 0 Then
                ws.Cells(summaryTableIndex, 11).Interior.ColorIndex = 3 ' Red
            Else
                ws.Cells(summaryTableIndex, 11).Interior.ColorIndex = 0 ' No color
            End If
            
            ' Increment summary table index
            summaryTableIndex = summaryTableIndex + 1
            
        End If
        
    Next i
    
    ' Define second summary Table call outs
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Clear cells containing maximum value
    ws.Cells(2, 17).Clear
    ws.Cells(2, 16).Clear
    ws.Cells(3, 17).Clear
    ws.Cells(3, 16).Clear
    ws.Cells(4, 17).Clear
    ws.Cells(4, 16).Clear
    
   ' Find greatest percent Increase, decrease and greatest total volume
    
    Set rng = ws.Range("L2:L3001")
    
    maxPercentDecrease = Application.WorksheetFunction.Min(rng)
    ws.Range("Q3").Value = maxPercentDecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    
    
    maxPercentIncrease = Application.WorksheetFunction.Max(rng)
    ws.Range("Q2").Value = maxPercentIncrease
    ws.Range("Q2").NumberFormat = "0.00%"
    
    
    maxTotalVolume = Application.WorksheetFunction.Max(ws.Range("M2:M3001"))
    ws.Range("Q4").Value = maxTotalVolume
    ws.Range("Q2:Q4").Columns.AutoFit
    
    'Get Ticker values
    ' Loop through each row of summary table"
    Dim J As Long
    Dim PercentChangeRng As Range
    Dim TickerRng As Range
    Dim VolumeRng As Range
    For Each W In ActiveWorkbook.Worksheets
    W.Activate
    
    Set PercentChangeRng = ws.Range("L2:L3001")
    Set TickerRng = ws.Range("J2:J3001")
    Set VolumeRng = ws.Range("M2:M3001")
    
    For J = 2 To 3
    
    ws.Range("P" & J).Value = Application.WorksheetFunction.Index(TickerRng, Application.WorksheetFunction.Match(ws.Range("Q" & J).Value, PercentChangeRng, 0))
    Next J
    
   ws.Cells(4, 16).Value = Application.WorksheetFunction.Index(TickerRng, Application.WorksheetFunction.Match(ws.Cells(4, 17).Value, VolumeRng, 0))
      
Next W
Next ws
 
End Sub


