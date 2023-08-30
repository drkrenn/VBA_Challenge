Attribute VB_Name = "Module1"
Sub Stocks()

Dim stock As String
Dim first As Double
Dim last As Double
Dim change As Double
Dim volume As Double
Dim perc As Double



volume = 0
first = Cells(2, 3)
last = 0
change = 0
Dim summary_table_row As Integer
summary_table_row = 2

For i = 2 To 753001

If Cells(i + 1, 1).Value <> Cells(i, 1) Then
stock = Cells(i, 1).Value
volume = volume + Cells(i, 7).Value
Range("I" & summary_table_row).Value = stock
Range("L" & summary_table_row).Value = volume
last = Cells(i, 6).Value
change = last - first
Range("J" & summary_table_row).Value = change

perc = change / first
Range("K" & summary_table_row).Value = perc

summary_table_row = summary_table_row + 1

volume = 0

first = Cells(i + 1, 3).Value

Else
volume = volume + Cells(i, 7).Value

End If

Next i

For i = 2 To 3001

If Cells(i, 10) < 0 Then
Cells(i, 10).Interior.ColorIndex = 3



Else
Cells(i, 10).Interior.ColorIndex = 4

End If

Next i


Dim inc As Double
Dim incnm As String

inc = 0

For m = 2 To 3001

If inc < Cells(m, 11) Then
inc = Cells(m, 11).Value
incnm = Cells(m, 9).Value

End If
Next m


Dim dec As Double
Dim decnm As String

dec = 0

For n = 2 To 3001

If dec > Cells(n, 11) Then
dec = Cells(n, 11).Value
decnm = Cells(n, 9).Value

End If
Next n


Dim vol As Double
Dim volnm As String

vol = 0

For p = 2 To 3001

If vol < Cells(p, 12) Then
vol = Cells(p, 12).Value
volnm = Cells(p, 9).Value

End If
Next p

Cells(2, 17) = inc
Cells(2, 16) = incnm

Cells(3, 17) = dec
Cells(3, 16) = decnm

Cells(4, 17) = vol
Cells(4, 16) = volnm

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"

End Sub

