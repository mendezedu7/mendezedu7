Sub StSummary()

Dim ticker As String
Dim total_stock As Double
Dim table As Long
Dim pre_table As Long
Dim yearly_change As Double
Dim yeatly_open As Double
Dim yearly_close As Double
Dim percent As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total As Double
Dim Value As Double


Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly change"
Cells(1, 11).Value = "percent change"
Cells(1, 12).Value = "total stock volume"
Cells(1, 16).Value = "ticker"
Cells(1, 17).Value = "value"
Cells(2, 15).Value = "greatest % increase"
Cells(3, 15).Value = "geatest % decrease"
Cells(4, 15).Value = "greatest total value"

'last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

total_stock = 0
table = 2
pre_table = 2
greatest_increase = 0
greatest_decrease = 0

For i = 2 To lastrow

total_stock = total_stock + Cells(i, 7).Value

    If Cells(i + 1, 1).Value Then
    ticker = Cells(i, 1).Value
    
    Range("I" & table).Value = ticker
    Range("L" & table).Value = total_stock
    
    total_stock = 0
    yearly_open = Range("C" & pre_table)
    yearly_close = Range("F" & i)
    yearly_change = yearly_close - yearly_change
    Range("J" & table).Value = yearly_change
    
    If yearly_open = 0 Then
    percent = 0
    
    Else
    yearly_open = Range("C" & pre_table)
    percent = yearly_change / yearly_open
    
    End If
    
    
Range("K" & table).NumberFormat = "0.00%"
Range("K" & table).Value = percent

    If Range("J" & table).Value >= 0 Then
    Range("J" & table).Interior.ColorIndex = 4
    
    Else
    Range("J" & table).Interior.ColorIndex = 3
    
    End If
    
table = table + 1
pre_table = i + 1

    End If
    
    
    
Next i

lastrow_Value = Cells(Rows.Count, 11).End(xlUp).Row
    Range("Q2").NumberFormat = "0.00&"
    Range("Q3").NumberFormat = "0.00%"
    
    For j = 2 To lastrow_Value
    
    If Range("K" & j).Value > greatest_increase Then
    greatest_increase = Range("K" & j).Value
    Range("Q2").Value = greatest_increase
    Range("P2").Value = Range("I" & j).Value
    
    End If
    
    
 
    If Range("K" & j).Value < greatest_decrease Then
    greatest_deacrease = Range("K" & j).Value
    Range("Q3").Value = greatest_decrease
    Range("P3").Value = Range("I" & j).Value
    
    End If
    
    If Range("L" & j).Value > greatest_total Then
    greatest_total = Range("L" & j).Value
    Range("Q4").Value = greatest_total
    Range("P4").Value = Range("I" & j).Value
    
    End If
    
  Next j


End Sub











