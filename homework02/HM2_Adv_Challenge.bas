Sub homework02advancedchallenge()

For Each ws In Worksheets

  Dim Ticker As String
  Dim Ticker_Total, Open_Value, Delta, Final_Value, Difference As Double
  Dim Count As Integer
  Ticker_Total = 0
  Count = 0
  
  ' For summary table
  Dim Table_Row As Integer
  Table_Row = 2
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ws.Cells(1, 10).Value = "Ticker"
  ws.Cells(1, 11).Value = "Year Change"
  ws.Cells(1, 12).Value = "Percentage Change"
  ws.Cells(1, 13).Value = "Stock Volume"

  ' Loop
   
  For i = 2 To lastrow

    ' test for similarity with next cell, else...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' set ticker tag
      Ticker = ws.Cells(i, 1).Value
      
      Final_Value = ws.Cells(i, 6).Value
      Open_Value = ws.Cells(i - Count, 3).Value
      
      Difference = Final_Value - Open_Value
        If (Open_Value = 0) Then
            Delta = "Null"
            Else
            Delta = (Difference / Open_Value) * 100 & "%"
        End If
    
      ' add corresponding ticket total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

      ' include ticker number at corresponding place
      ws.Range("J" & Table_Row).Value = Ticker
      ws.Range("K" & Table_Row).Value = Difference
      
      If (Difference < 0) Then
            ws.Range("K" & Table_Row).Interior.ColorIndex = 3
            Else
            ws.Range("K" & Table_Row).Interior.ColorIndex = 4
      End If
      
      ws.Range("L" & Table_Row).Value = Delta
      ws.Range("M" & Table_Row).Value = Ticker_Total
      

      ' add one row for the next ticker iteration
      Table_Row = Table_Row + 1
      
      ' reset totatal counter for next iteration
      Ticker_Total = 0
      Count = 0

    Else

      ' Add to ticker total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
      Count = Count + 1

    End If

  Next i
  
Dim rngvolume, rngper As Range
Dim volmax, deltaMin, deltaMax  As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ws.Cells(1, 17).Value = "Ticker"
  ws.Cells(1, 18).Value = "Value"
  ws.Cells(2, 16).Value = "Greatest % Increase"
  ws.Cells(3, 16).Value = "Greatest % Decrease"
  ws.Cells(4, 16).Value = "Greatest Total Volume"

Set rngvolume = ws.Range("M2:M" & lastrow)
Set rngper = ws.Range("L2:L" & lastrow)
Set tabletotal = ws.Range("J2:M" & lastrow)

deltaMax = Application.WorksheetFunction.Max(rngper)
deltaMin = Application.WorksheetFunction.Min(rngper)
volmax = Application.WorksheetFunction.Max(rngvolume)

ws.Cells(2, 18).Value = (deltaMax * 100) & "%"
ws.Cells(3, 18).Value = (deltaMin * 100) & "%"
ws.Cells(4, 18).Value = volmax
  
deltamaxrow = Application.WorksheetFunction.Match(deltaMax, rngper, 0)
deltaminrow = Application.WorksheetFunction.Match(deltaMin, rngper, 0)
volmaxrow = Application.WorksheetFunction.Match(volmax, rngvolume, 0)
  
ws.Range("Q2").Value = ws.Cells(deltamaxrow + 1, 10)
ws.Range("Q3").Value = ws.Cells(deltaminrow + 1, 10)
ws.Range("Q4").Value = ws.Cells(volmaxrow + 1, 10)
  
Next ws
  
End Sub
