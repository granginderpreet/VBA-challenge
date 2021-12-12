<<<<<<< HEAD
Attribute VB_Name = "Module1"
Sub StockAnalysis():
Dim OpenPrice, ClosePrice, ChangeinPrice, totalVolume As Double

Dim symbolcount As Integer


'Cycle through the sheets
For Each ws In Worksheets
symbolcount = 0
worksheetName = ws.Name
'Find the last row

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Initialize the Openproice for first symbol in sheet
OpenPrice = ws.Cells(2, 3)

'Fill the top row

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "% Change"
Range("L1").Value = "Total Stock volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Cycle through the sheet from 2nd row to last

For i = 2 To lastRow
  
  'If there is a change in  ticker symbol, need to save some stats
  If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
  'Increment symbol count
  symbolcount = symbolcount + 1
  'Update the total Volume
  totalVolume = totalVolume + ws.Cells(i, 7)
  'save symbol
  ws.Cells(1 + symbolcount, 9) = ws.Cells(i, 1).Value
  'save closing price for that symbol
  ClosePrice = ws.Cells(i, 6)
  'find change in price
  ChangeinPrice = ClosePrice - OpenPrice
  If OpenPrice <> 0 Then
  PercentChangeinPrice = ChangeinPrice / OpenPrice
  Else
  PercentChangeinPrice = "NaN"
  End If
  
   ws.Cells(1 + symbolcount, 10) = ChangeinPrice
   ws.Cells(1 + symbolcount, 11) = PercentChangeinPrice
   ws.Range("K" & (1 + symbolcount)).NumberFormat = "0.00%"
   ws.Cells(1 + symbolcount, 12) = totalVolume
   If ChangeinPrice < 0 Then
   ws.Cells(1 + symbolcount, 10).Interior.Color = vbRed
   
   ElseIf (ChangeinPrice = 0) Then
    ws.Cells(1 + symbolcount, 10).Interior.Color = vbYellow
    Else
    ws.Cells(1 + symbolcount, 10).Interior.Color = vbGreen
   End If
   
   'Set the openprice for next symbol
  OpenPrice = Cells(i + 1, 3)
  'Reset volume to 0
  totalVolume = 0
  
  Else
  totalVolume = totalVolume + ws.Cells(i, 7)
  End If
  

Next i

' Use excel worksheet functions to get max and min
'Find greatest increase

ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(2, 17) = Application.WorksheetFunction.Max(Range("K:K"))
iCount = WorksheetFunction.Match(ws.Cells(2, 17), Range("K:K"), 0)
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 16) = ws.Cells(iCount, 9)

'Find greatest decrease

ws.Cells(3, 15) = "Greatest % Decrease"

ws.Cells(3, 17) = Application.WorksheetFunction.Min(ws.Range("K:K"))
iCount = WorksheetFunction.Match(ws.Cells(3, 17), ws.Range("K:K"), 0)
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16) = ws.Cells(iCount, 9)

'Find max volume
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(4, 17) = Application.WorksheetFunction.Max(ws.Range("L:L"))
iCount = WorksheetFunction.Match(ws.Cells(4, 17), ws.Range("L:L"), 0)
ws.Cells(4, 17).NumberFormat = "0"
ws.Cells(4, 16) = ws.Cells(iCount, 9)
Next ws

End Sub
Sub StockAnalysisClear()


'Cycle through the sheets
For Each ws In Worksheets
ws.Range("I:Q").Clear

Next ws


End Sub

=======
Attribute VB_Name = "Module1"
Sub StockAnalysis():
Dim OpenPrice, ClosePrice, ChangeinPrice, totalVolume As Double

Dim symbolcount As Integer


'Cycle through the sheets
For Each ws In Worksheets
symbolcount = 0
worksheetName = ws.Name
'Find the last row

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Initialize the Openproice for first symbol in sheet
OpenPrice = ws.Cells(2, 3)

'Fill the top row

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "% Change"
Range("L1").Value = "Total Stock volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Cycle through the sheet from 2nd row to last

For i = 2 To lastRow
  
  'If there is a change in  ticker symbol, need to save some stats
  If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
  'Increment symbol count
  symbolcount = symbolcount + 1
  'Update the total Volume
  totalVolume = totalVolume + ws.Cells(i, 7)
  'save symbol
  ws.Cells(1 + symbolcount, 9) = ws.Cells(i, 1).Value
  'save closing price for that symbol
  ClosePrice = ws.Cells(i, 6)
  'find change in price
  ChangeinPrice = ClosePrice - OpenPrice
  If OpenPrice <> 0 Then
  PercentChangeinPrice = ChangeinPrice / OpenPrice
  Else
  PercentChangeinPrice = "NaN"
  End If
  
   ws.Cells(1 + symbolcount, 10) = ChangeinPrice
   ws.Cells(1 + symbolcount, 11) = PercentChangeinPrice
   ws.Range("K" & (1 + symbolcount)).NumberFormat = "0.00%"
   ws.Cells(1 + symbolcount, 12) = totalVolume
   If ChangeinPrice < 0 Then
   ws.Cells(1 + symbolcount, 10).Interior.Color = vbRed
   
   ElseIf (ChangeinPrice = 0) Then
    ws.Cells(1 + symbolcount, 10).Interior.Color = vbYellow
    Else
    ws.Cells(1 + symbolcount, 10).Interior.Color = vbGreen
   End If
   
   'Set the openprice for next symbol
  OpenPrice = Cells(i + 1, 3)
  'Reset volume to 0
  totalVolume = 0
  
  Else
  totalVolume = totalVolume + ws.Cells(i, 7)
  End If
  

Next i

' Use excel worksheet functions to get max and min
'Find greatest increase

ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(2, 17) = Application.WorksheetFunction.Max(Range("K:K"))
iCount = WorksheetFunction.Match(ws.Cells(2, 17), Range("K:K"), 0)
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 16) = ws.Cells(iCount, 9)

'Find greatest decrease

ws.Cells(3, 15) = "Greatest % Decrease"

ws.Cells(3, 17) = Application.WorksheetFunction.Min(ws.Range("K:K"))
iCount = WorksheetFunction.Match(ws.Cells(3, 17), ws.Range("K:K"), 0)
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16) = ws.Cells(iCount, 9)

'Find max volume
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(4, 17) = Application.WorksheetFunction.Max(ws.Range("L:L"))
iCount = WorksheetFunction.Match(ws.Cells(4, 17), ws.Range("L:L"), 0)
ws.Cells(4, 17).NumberFormat = "0"
ws.Cells(4, 16) = ws.Cells(iCount, 9)
Next ws

End Sub
Sub StockAnalysisClear()


'Cycle through the sheets
For Each ws In Worksheets
ws.Range("I:Q").Clear

Next ws


End Sub

>>>>>>> db41a21b6330e7c7f989f88f5d0e0a7901c5b224
