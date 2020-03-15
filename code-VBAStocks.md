#code-VBA Challenge

Sub Stock():
Dim ws As Worksheet
Dim i As Variant
Dim LastRow As Variant
Dim p As Variant
Dim Total As Variant


'Loop through all Worksheet
For Each ws In Worksheets

'calculating last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
SumRow = 2
Total = 0
p = 2

ws.Range("J1").Value = "TotalStockVolume"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percentage Change"
ws.Range("M1").Value = "Ticker"

ws.Range("H1").Value = "Opening Value"
ws.Range("I1").Value = "Closing Value"

ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

i = 2

ws.Cells(p, 8).Value = ws.Cells(2, 3).Value 'opening Value

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Populating Ticker
        ws.Cells(SumRow, 13).Value = ws.Cells(i, 1).Value
        
        'Calculating Total Stock Volume
        Total = Total + ws.Cells(i, 7).Value
        ws.Cells(SumRow, 10).Value = Total
        SumRow = SumRow + 1
        Total = 0
        
        'Calculating change in Price
        ws.Cells(p, 9).Value = ws.Cells(i, 6).Value 'closing Value
        ws.Cells(p, 11).Value = ws.Cells(p, 9).Value - ws.Cells(p, 8).Value
                
            If ws.Cells(p, 11).Value > 0 Then
                ws.Cells(p, 11).Interior.Color = vbGreen
            ElseIf ws.Cells(p, 11).Value < 0 Then
                ws.Cells(p, 11).Interior.Color = vbRed
            End If
        
        'Calculating Percent Change
        'PercentChange = (ChangePrice / OpenPrice)
        ws.Cells(p, 12).Value = ws.Cells(p, 11).Value / ws.Cells(p, 8).Value
        ws.Cells(p, 12).Style = "Percent"
        
        p = p + 1
        
        'Resetting the Opening Price for each ticker
        ws.Cells(p, 8).Value = ws.Cells(i + 1, 3).Value
        
    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            Total = Total + ws.Cells(i, 7).Value
    End If

Next i

ws.Range("Q2").Value = ws.Application.WorksheetFunction.Max(ws.Range("L1:L3169").Value)
ws.Range("Q2").Style = "Percent"
ws.Range("P2").Value = ws.Application.WorksheetFunction.VLookup(ws.Range("Q2").Value, ws.Range("L1:M3169"), 2, False)

ws.Range("Q3").Value = ws.Application.WorksheetFunction.Min(ws.Range("L1:L3169").Value)
ws.Range("Q3").Style = "Percent"
ws.Range("P3").Value = ws.Application.WorksheetFunction.VLookup(ws.Range("Q3").Value, ws.Range("L1:M3169"), 2, False)

ws.Range("Q4").Value = ws.Application.WorksheetFunction.Max(ws.Range("J1:J3169").Value)
ws.Range("P4").Value = ws.Application.WorksheetFunction.VLookup(ws.Range("Q4").Value, ws.Range("J1:M3169"), 4, False)

Next ws

MsgBox ("Analyis done")

End Sub


