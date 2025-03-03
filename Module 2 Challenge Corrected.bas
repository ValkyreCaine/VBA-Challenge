Attribute VB_Name = "Module1"
Sub Module_2_Challenge()
'Variable for WS
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
Dim i As Double


'Variables
Dim Ticker As String
Dim Data_Set_Table As Long
Dim opening As Double
Dim closing As Double
Dim Quarterly As Double
Dim Percent As Double
Dim Stock As Double
Dim k As Double
'Variable Initialization
opening = 0
closing = 0
Quarterly = 0
Percent = 0
Stock = 0

Data_Set_Table = 2

' Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quaterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
   
'Forloop for All Queries

LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow

'If conditionals
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    opening = Cells(i, 3).Value
    Stock = Stock + Cells(i, 7).Value

ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    closing = Cells(i, 6).Value
    Quarterly = closing - opening
    Percent = ((closing - opening) / opening)

'Percent Formatting
Range("K:K").NumberFormat = "0.00%"

'Stock Value add in
Stock = Stock + Cells(i, 7).Value

Range("i" & Data_Set_Table).Value = Ticker
Range("j" & Data_Set_Table).Value = Quarterly
Range("k" & Data_Set_Table).Value = Percent
Range("l" & Data_Set_Table).Value = Stock
Data_Set_Table = Data_Set_Table + 1

'Closing and Stock Resets for next i
closing = 0
Stock = 0
Else
Stock = Stock + Cells(i, 7).Value

End If
If ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 10
ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
End If
Next i

' set Ticker, Value, Greatest %, Increase, % Decrease, and Total volume headers
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
' Volume Last Row
volume_lastrow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
'Greatest % Increase
Greatest_inc = Application.WorksheetFunction.Max(Range("K2:K" & volume_lastrow))
Greatest_inc_row = Application.WorksheetFunction.Match(Greatest_inc, Range("K2:K" & volume_lastrow), 0) + 1 ' Find corresponding row
        
' Greatest % Decrease
Greatest_dec = Application.WorksheetFunction.Min(Range("K2:K" & volume_lastrow))
Greatest_dec_row = Application.WorksheetFunction.Match(Greatest_dec, Range("K2:K" & volume_lastrow), 0) + 1 ' Find corresponding row
        
' Greatest Total Volume
Greatest_vol = Application.WorksheetFunction.Max(Range("L2:L" & volume_lastrow))
Greatest_vol_row = Application.WorksheetFunction.Match(Greatest_vol, Range("L2:L" & volume_lastrow), 0) + 1 ' Find corresponding row
        
Range("P4").Value = Cells(Greatest_vol_row, 9).Value ' Ticker of greatest volume
Range("Q4").Value = Cells(Greatest_vol_row, 12).Value ' Greatest  Volume
Range("P2").Value = Cells(Greatest_inc_row, 9).Value ' Ticker of greatest increase
Range("Q2").Value = Cells(Greatest_inc_row, 11).Value ' Percent Change of greatest increase
Range("P3").Value = Cells(Greatest_dec_row, 9).Value ' Ticker of greatest decrease
Range("Q3").Value = Cells(Greatest_dec_row, 11).Value ' Percent Change of greatest decrease

'Format Percents
Range("Q2:Q3").NumberFormat = "0.00%"
 
   
   
   Next ws
   
End Sub
