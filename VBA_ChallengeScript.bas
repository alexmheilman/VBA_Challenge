Attribute VB_Name = "Module1"
Sub ticker():
Dim ticker_name As String
Dim summary_table_row As Integer
summary_table_row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_name = Cells(i, 1).Value
        Range("J" & summary_table_row).Value = ticker_name
        summary_table_row = summary_table_row + 1
    End If
Next i
End Sub

Sub stock_volume():
Dim stock_volume As Double
stock_volume = 0
Dim summary_table_row As Integer
summary_table_row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    stock_volume = stock_volume + Cells(i, 7).Value
    Range("M" & summary_table_row).Value = stock_volume
    summary_table_row = summary_table_row + 1
    stock_volume = 0
Else
    stock_volume = stock_volume + Cells(i, 7).Value
End If
Next i
End Sub
Sub yearly_change():
Dim summary_table_row As Integer
summary_table_row = 2
Dim change As Double
change = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
Dim stockPriceAlreadyCaptured As Boolean
For i = 2 To lastrow
    If stockPriceAlreadyCaptured = False Then
         'Set opening price
         Dim Opening_Price As Double
         Opening_Price = Cells(i, 3).Value

         'ensures no future prices captured until condition met.
         stockPriceAlreadyCaptured = True
End If
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    change = Cells(i, 6).Value - Opening_Price
    Range("K" & summary_table_row).Value = change
    Percent_Change = (change / Opening_Price) * 100
    Range("L" & summary_table_row).Value = Percent_Change
    summary_table_row = summary_table_row + 1
    change = 0
    stockPriceAlreadyCaptured = False
End If
Next i
End Sub
Sub max():
Dim max_value As Double
max_value = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 12).Value <> Cells(i, 12).Value And Cells(i + 1, 12).Value < Cells(i, 12).Value Then
    max_value = Cells(i, 12).Value
        For j = 2 To lastrow
            If Cells(j, 12).Value > max_value Then
            max_value = Cells(j, 12).Value
            Range("Q2").Value = max_value
            Range("R2").Value = Cells(j, 10).Value
        End If
        Next j
End If
Next i
End Sub
Sub min():
Dim min_value As Double
min_value = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 12).Value <> Cells(i, 12).Value And Cells(i + 1, 12).Value > Cells(i, 12).Value Then
    min_value = Cells(i, 12).Value
        For j = 2 To lastrow
            If Cells(j, 12).Value < min_value Then
            min_value = Cells(j, 12).Value
            Range("Q3").Value = min_value
            Range("R3").Value = Cells(j, 10).Value
        End If
        Next j
End If
Next i
End Sub
Sub max_stock_volume():
Dim max_value As Double
max_value = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
    If Cells(i + 1, 13).Value <> Cells(i, 13).Value And Cells(i + 1, 13).Value < Cells(i, 13).Value Then
    max_value = Cells(i, 13).Value
        For j = 2 To lastrow
            If Cells(j, 13).Value > max_value Then
            max_value = Cells(j, 13).Value
            Range("Q4").Value = max_value
            Range("R4").Value = Cells(j, 10).Value
        End If
        Next j
End If
Next i
End Sub
Sub each_year():
   Dim ws As Worksheet
   
   '** SET The Sheet Names - MUST Reflect Each Sheet Name Exactly!
   WkSheets = Array("2018", "2019", "2020")
   
   For Each ws In Sheets(Array("2018", "2019", "2020"))
      ws.Select
      Call ticker
      Call stock_volume
      Call yearly_change
      Call max
      Call min
      Call max_stock_volume
   Next ws
End Sub


