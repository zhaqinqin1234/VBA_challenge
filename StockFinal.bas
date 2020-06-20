Attribute VB_Name = "Module1"
Sub StockChange()
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    
    Dim tickerSummaryRow As Integer
    Dim tickerStartRow As Double
    Dim lRow As Double
    Dim total As Double
    Dim openPrice As Double
    Dim closePrice As Double
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
   
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "YearlyChange"
    ws.Range("K1") = "Change%"
    ws.Range("L1") = "TotalVolume"
    tickerSummaryRow = 2
    tickerStartRow = 2
    total = 0
   
    'Find the last non-blank cell in column A(1)
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            'Retrieve ticker
            ws.Cells(tickerSummaryRow, 9).Value = ws.Cells(i, 1).Value
            'Retrive total
            total = total + ws.Cells(i, 7).Value
            'Write total to cells
            ws.Cells(tickerSummaryRow, 12).Value = total
            
            'Retrieve yearly change
            openPrice = ws.Cells(tickerStartRow, 3).Value
            closePrice = ws.Cells(i, 6).Value
            ws.Cells(tickerSummaryRow, 10).Value = closePrice - openPrice
            'Retrieve percentage change
            'ws.Cells(tickerSummaryRow, 11).Value = (ws.Cells(tickerSummaryRow, 10).Value) / (ws.Cells(tickerStartRow, 3).Value) * 100
            If openPrice = 0 Or IsEmpty(ws.Cells(tickerStartRow, 3).Value) Then
                ws.Cells(tickerSummaryRow, 11).Value = "Null"
            Else: ws.Cells(tickerSummaryRow, 11).Value = ws.Cells(tickerSummaryRow, 10).Value / openPrice
            End If
            
                       
            
            tickerSummaryRow = tickerSummaryRow + 1
            tickerStartRow = i + 1
            total = 0
         Else:
            total = total + ws.Cells(i, 7).Value
        End If
    Next i
    Next
        starting_ws.Activate
    

End Sub

Sub conditionalFormatting():

For Each ws In Worksheets

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(i, 11).NumberFormat = "0.00%"
        'Formatting increased change to green and decreased change to red
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
Next ws
End Sub


Sub Summary():
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate
    Dim rng As Range
    Dim Maximum As Double
    Dim Minimum As Double
    Dim totalMax As Double
    Dim SummaryTicker1 As String
    Dim SummaryTicker2 As String
    Dim SummaryTicker3 As String
    Set rng = ws.Range("K:K")
    Set rng1 = ws.Range("L:L")
    'Worksheet function MAX/Min returns the largest/smallest value in a  range
    Maximum = Application.WorksheetFunction.Max(rng)
    Minimum = Application.WorksheetFunction.Min(rng)
    totalMax = Application.WorksheetFunction.Max(rng1)
    
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 17).Value = Maximum
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 15).Value = "Greatest % Increase"
    
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 11).Value = Maximum Then
        SummaryTicker1 = ws.Cells(i, 9).Value
        ws.Cells(2, 16).Value = SummaryTicker1
        End If
    Next i
    
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(3, 17).Value = Minimum
ws.Cells(3, 17).NumberFormat = "0.00%"

    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 11).Value = Minimum Then
        SummaryTicker2 = ws.Cells(i, 9).Value
        ws.Cells(3, 16).Value = SummaryTicker2
        End If
    Next i
    
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(4, 17).Value = totalMax
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 12).Value = totalMax Then
        SummaryTicker3 = ws.Cells(i, 9).Value
        ws.Cells(4, 16).Value = SummaryTicker3
        End If
    Next i
    
    ws.Columns("A:Q").AutoFit
    
Next ws


End Sub
