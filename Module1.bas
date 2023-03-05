Attribute VB_Name = "Module1"
Sub Stock()

Dim ws As Worksheet
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim volume As Double
volume = 0
Dim yearly_change As Double
Dim percent As Double

For Each ws In Sheets

'Add Header to columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Value"

ws.Cells(1, 14).Value = "Ticker."
ws.Cells(1, 15).Value = "Value"
ws.Cells(2, 13).Value = "Greatest % Increase"
ws.Cells(3, 13).Value = "Greatest % Decrease"
ws.Cells(4, 13).Value = "Greatest Total Volume"

'Keep track of the location for eack ticker
Dim Summary_table_row As Integer
Summary_table_row = 2

'Determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through rows to find tickers
    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ticker = ws.Cells(i, 1).Value
        volume = volume + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_table_row).Value = ticker
        ws.Range("L" & Summary_table_row) = volume
        Summary_table_row = Summary_table_row + 1
        volume = 0
    Else

    volume = volume + Cells(i, 7).Value


        End If
     Next i
     
Summary_table_row = 2

'assign opening and closing price, and calculate yearly change
    For i = 2 To LastRow

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        close_price = ws.Cells(i, 6).Value

        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
        open_price = ws.Cells(i, 3).Value
        End If

        If open_price > 0 And close_price > 0 Then
        yearly_change = close_price - open_price
        percent = yearly_change / open_price

        ws.Range("J" & Summary_table_row).Value = yearly_change
        ws.Range("K" & Summary_table_row).Value = FormatPercent(percent)
        Summary_table_row = Summary_table_row + 1
        close_price = 0
        open_price = 0
        yearly_change = 0

        End If

    Next i


'Loop through column "Yearly change" and insert red or green interior
    For i = 2 To Summary_table_row - 1
    
        If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        Else: ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    Next i
    
'Finding and adding Greatest % Increase, Greatest % decrease, and Greatest Total Value, from the columns in to the assigned cells

   Dim greatest_increase As Double
   Dim greatest_decrease As Double
   Dim greatest_total As Double


   greatest_increase = WorksheetFunction.Max(ws.Columns("K"))
   greatest_decrease = WorksheetFunction.Min(ws.Columns("K"))
   greatest_total = WorksheetFunction.Max(ws.Columns("L"))
   
   ws.Cells(2, 15).Value = FormatPercent(greatest_increase)
   ws.Cells(3, 15).Value = FormatPercent(greatest_decrease)
   ws.Cells(4, 15).Value = greatest_total
   
    For i = 2 To LastRow
    
        If greatest_increase = ws.Cells(i, 11).Value Then
        ws.Cells(2, 14).Value = ws.Cells(i, 9).Value
    
        ElseIf greatest_decrease = ws.Cells(i, 11).Value Then
        ws.Cells(3, 14).Value = ws.Cells(i, 9).Value
        
        ElseIf greatest_total = ws.Cells(i, 12).Value Then
        ws.Cells(4, 14).Value = ws.Cells(i, 9).Value
    
        End If
    
    Next i
   
'  highlight positive change in green and negative change in red

    For i = 2 To LastRow
    
        If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    
        ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    
        End If
    
    Next i
           
Next ws

End Sub


Sub color()
'Loop through column "Yearly change" and insert red or green interior
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow
    
        Cells(i, 10).Interior.ColorIndex = 0
        
        
    Next i
End Sub


