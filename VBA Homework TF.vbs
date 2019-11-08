Dim i As Long
Dim Ticker As String
Dim r As Integer
Dim Volume As Variant
Dim LastRow As Long
Dim LastColumn As Integer
Dim WorksheetName As String
Dim SummaryRow As Integer
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim GreatestChangeUp As Double
Dim GreatestChangeDown As Double
Dim GreatestTotalVolume As Variant





Sub Stocks()

wscount = ActiveWorkbook.Worksheets.Count


For Each ws In Worksheets
    i = 0

    SummaryRow = 2

    'Determine the last row for each ws
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Grab the ws name
    WorksheetName = ws.Name

    'Determine last column number
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    'Add New Headers
    ws.Cells(1, LastColumn + 2).Value = "Ticker"
    ws.Cells(1, LastColumn + 3).Value = "Yearly Change"
    ws.Cells(1, LastColumn + 4).Value = "Percent Change"
    ws.Cells(1, LastColumn + 8).Value = "Ticker"
    ws.Cells(1, LastColumn + 9).Value = "Value"
    ws.Cells(2, LastColumn + 7).Value = "Greatest % Increase"
    ws.Cells(3, LastColumn + 7).Value = "Greatest % Decrease"
    ws.Cells(4, LastColumn + 7).Value = "Greatest Total Volume"
    
    
    OpeningPrice = ws.Cells(2, 3).Value

    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Get the ticket Name
            Ticker = ws.Cells(i, 1).Value
            
            'Take the final volume number
            Volume = Volume + ws.Cells(i, 7).Value
            ClosingPrice = ws.Cells(i, 6).Value
            

            'Add summary values
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = ClosingPrice - OpeningPrice
            
            'If statement for the color formatting of price change
            If ws.Cells(SummaryRow, 10).Value > 0 Then
                ws.Cells(SummaryRow, 10).Interior.Color = rgbGreen
            Else
                ws.Cells(SummaryRow, 10).Interior.Color = rgbRed
            End If
            
            
            ws.Cells(SummaryRow, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
            ws.Cells(SummaryRow, 12).Value = Volume

            'Add to the summary for the next ticker
            SummaryRow = SummaryRow + 1

            'Reset the volume
            Volume = 0

            'Get Opening Price for next run
            OpeningPrice = ws.Cells(i + 1, 3).Value
        
        'If the cell prior is the same ticker
        Else

            Volume = Volume + ws.Cells(i, 7).Value
            

        End If
    
    Next i
    
    'Get the min/max/total for each ws
    GreatestChangeUp = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    GreatestChangeDown = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    GreatestTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    
    ws.Range("P2").Value = GreatestChangeUp
    ws.Range("P3").Value = GreatestChangeDown
    ws.Range("P2:P3").NumberFormat = "0.00%"
    ws.Range("P4").Value = GreatestTotalVolume
    ws.Range("P4").Style = "Currency"
    
    'Vlookup for the ticker name
    'ws.Range("O2").Value = Application.WorksheetFunction.VLookup(GreatestChangeUp, , 3, False)
    
    
    
    
Next ws

End Sub

Sub ClearAll()
    For Each ws In Worksheets
        ws.Range("I1:Z10000").ClearContents
    Next ws

End Sub

