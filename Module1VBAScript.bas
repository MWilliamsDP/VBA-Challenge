Attribute VB_Name = "Module1"


    
    
Sub VBA_Homework()


For Each ws In Worksheets
   


' Add headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Difference"
    ws.Range("K1").Value = "Percent Difference"
    ws.Range("L1").Value = "Total Volume"

'Name Variables and Counters

    Dim Ticker As String
    Dim Volume As Double
    Volume = 0
    Dim TickerCounter As Integer
    TickerCounter = 2
    Dim Counter2 As Double
    Counter2 = 2
    Dim OpenValue As Double
    Dim CloseValue As Double


'loop through ticker symbols

 For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            Volume = Volume + ws.Cells(I, 7).Value
            Ticker = ws.Cells(I, 1).Value
            OpenValue = ws.Cells(Counter2, 3)
   
   
    
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    CloseValue = ws.Cells(I, 6)
    ws.Cells(TickerCounter, 9).Value = Ticker
    ws.Cells(TickerCounter, 10).Value = CloseValue - OpenValue

   'correct for 0 value
    
    If OpenValue = 0 Then
    ws.Cells(TickerCounter, 11).Value = 0

Else
    ws.Cells(TickerCounter, 11).Value = (CloseValue - OpenValue) / OpenValue
    
End If

ws.Cells(TickerCounter, 12).Value = Volume

'Color

If ws.Cells(TickerCounter, 10).Value > 0 Then
    ws.Cells(TickerCounter, 10).Interior.Color = vbGreen
    Else
        ws.Cells(TickerCounter, 10).Interior.Color = vbRed
        
        End If
        
ws.Cells(TickerCounter, 11).NumberFormat = "0.00%"


'reset
Volume = 0
TickerCounter = TickerCounter + 1
Counter2 = I + 1

End If
Next I

'Format

        ws.Columns("J").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit

Next ws
  
End Sub

