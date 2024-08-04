VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Wk()

Dim ws As Worksheet
     

For Each ws In ThisWorkbook.Worksheets
Dim Tickname As String
Dim Row As Long
Dim openV As Double
Dim closeV As Double
Dim stockV As LongLong






Row = 2
openV = ws.Cells(2, 3).Value
closeV = 0
stockV = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 


ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Quarterly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % decrease"
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(1, 16) = "TIcker"
ws.Cells(1, 17) = "Value"





  

  For i = 2 To lastrow
  
   
   
   
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
  closeV = ws.Cells(i, 6).Value
  
  
  
  
  ws.Cells(Row, 9).Value = ws.Cells(i, 1).Value
  
  ws.Cells(Row, 10).Value = closeV - openV
  
  ws.Cells(Row, 11).Value = ((closeV - openV) / openV)
  
  openV = ws.Cells(i + 1, 3).Value

  ws.Cells(Row, 11).NumberFormat = "0.00%"
  stockV = stockV + ws.Cells(i, 7).Value
  ws.Cells(Row, 12).Value = stockV
  
  
  
  
  Row = Row + 1
  stockV = 0
  


Else
stockV = stockV + ws.Cells(i, 7).Value
   
  
  End If



Next i

For i = 2 To lastrow

If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K1600")) Then

ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
ws.Cells(2, 17).NumberFormat = "0.00%"

ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K1600")) Then
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
ws.Cells(3, 17).NumberFormat = "0.00%"

ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L1600")) Then
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
ws.Cells(4, 17).Value = ws.Cells(i, 12).Value

End If
Next i

   
For i = 2 To lastrow
  If ws.Cells(i, 10).Value > 0 Then
  ws.Cells(i, 10).Interior.ColorIndex = 4
  
  ElseIf ws.Cells(i, 10).Value < 0 Then
  ws.Cells(i, 10).Interior.ColorIndex = 3
  Else
  ws.Cells(i, 10).Interior.ColorIndex = 42
  
    
  End If
  Next i
  
  
  
  
  
  
  
  
  


Next ws

End Sub



