Sub homework02easychallenge()

For Each ws In Worksheets

  Dim Ticker As String
  Dim Ticker_Total As Double
  Ticker_Total = 0
  
  ' For summary table
  Dim Table_Row As Integer
  Table_Row = 2
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ws.Cells(1, 10).Value = "Ticker"
  ws.Cells(1, 11).Value = "Volume"

  ' Loop
   
  For i = 2 To lastrow

    ' test for similarity with next cell, else...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' set ticker tag
      Ticker = ws.Cells(i, 1).Value

      ' add corresponding ticket total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

      ' include ticker number at corresponding place
      ws.Range("J" & Table_Row).Value = Ticker

      ' include corresponding ticker total
      ws.Range("K" & Table_Row).Value = Ticker_Total

      ' add one row for the next ticker iteration
      Table_Row = Table_Row + 1
      
      ' reset totatal counter for next iteration
      Ticker_Total = 0

    Else

      ' Add to ticker total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
  Next ws

End Sub

