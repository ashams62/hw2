Attribute VB_Name = "Module1"
Sub HW2():

Dim i As Long
Dim j As Integer
Dim c_1 As Long
Dim c_9 As Long
Dim m As Long
Dim k As Long
Dim l As Long
Dim d As Long
Dim t As Long
Dim y As Long

'loop multiple sheets

WS_Count = ActiveWorkbook.Worksheets.Count

For y = 1 To WS_Count

'Headers

Worksheets(y).Cells(1, 9).Value = "Ticker"
Worksheets(y).Cells(1, 10).Value = "Yearly Change"
Worksheets(y).Cells(1, 11).Value = "Percent Change"
Worksheets(y).Cells(1, 12).Value = "Total Stock Volume"
Worksheets(y).Cells(1, 16).Value = "Ticker"
Worksheets(y).Cells(1, 17).Value = "Value"

j = 2

'count Range of data

c_1 = Worksheets(y).Cells(Rows.Count, 1).End(xlUp).Row


'find change in Tickers and put data in column I

For i = 2 To c_1
    If Worksheets(y).Cells(i, 1).Value <> Worksheets(y).Cells(i + 1, 1).Value Then
        Worksheets(y).Cells(j, 9).Value = Worksheets(y).Cells(i, 1).Value
        Worksheets(y).Cells(j, 20).Value = i
        Worksheets(y).Cells(j + 1, 19).Value = i + 1
        j = j + 1
        
    End If
Next i

Worksheets(y).Cells(2, 19).Value = 2


' Calc yearly change & percentage & Sum of toral stock for each ticker

Total = 0
c_9 = Worksheets(y).Cells(Rows.Count, 9).End(xlUp).Row

For m = 2 To c_9
    k = Worksheets(y).Cells(m, 19).Value
    l = Worksheets(y).Cells(m, 20).Value
    
    For d = k To l
       Total = Total + Worksheets(y).Cells(d, 7).Value
       Worksheets(y).Cells(d, 12) = Total
    Next d

    Worksheets(y).Cells(m, 10).Value = Worksheets(y).Cells(l, 6).Value - Worksheets(y).Cells(k, 3).Value
    Worksheets(y).Cells(m, 11).Value = Worksheets(y).Cells(m, 10).Value * 100 / Worksheets(y).Cells(k, 3).Value
    Worksheets(y).Cells(m, 11).Value = Round(Worksheets(y).Cells(m, 11).Value, 2)
    Worksheets(y).Cells(m, 12) = Total
    Total = 0
    
    ' Add color
    
    If Worksheets(y).Cells(m, 10).Value >= 0 Then
       Worksheets(y).Cells(m, 10).Interior.ColorIndex = 4
       Worksheets(y).Cells(m, 11).Interior.ColorIndex = 4
    Else
       Worksheets(y).Cells(m, 10).Interior.ColorIndex = 3
       Worksheets(y).Cells(m, 11).Interior.ColorIndex = 3
    End If
    
Next m


'Find Max Min of change & Max Toal stock

Worksheets(y).Cells(2, 17).Value = 0
Worksheets(y).Cells(3, 17).Value = 0
Worksheets(y).Cells(4, 17).Value = 0

For t = 2 To c_9
    If Worksheets(y).Cells(2, 17).Value < Worksheets(y).Cells(t, 11).Value Then
        Worksheets(y).Cells(2, 17).Value = Worksheets(y).Cells(t, 11).Value
    End If
    
    If Worksheets(y).Cells(3, 17).Value > Worksheets(y).Cells(t, 11).Value Then
        Worksheets(y).Cells(3, 17).Value = Worksheets(y).Cells(t, 11).Value
    End If
    
    If Worksheets(y).Cells(4, 17).Value < Worksheets(y).Cells(t, 12).Value Then
        Worksheets(y).Cells(4, 17).Value = Worksheets(y).Cells(t, 12).Value
    End If
    
    If Worksheets(y).Cells(t, 11) = Worksheets(y).Cells(2, 17).Value Then
        Worksheets(y).Cells(2, 16).Value = Worksheets(y).Cells(t, 9).Value
    End If

    If Worksheets(y).Cells(t, 11) = Worksheets(y).Cells(3, 17).Value Then
        Worksheets(y).Cells(3, 16).Value = Worksheets(y).Cells(t, 9).Value
    End If
    
    If Worksheets(y).Cells(t, 12) = Worksheets(y).Cells(4, 17).Value Then
        Worksheets(y).Cells(4, 16).Value = Worksheets(y).Cells(t, 9).Value
    End If
Next t

' add label for max and min

Worksheets(y).Cells(2, 15).Value = "Greatest % Increase"
Worksheets(y).Cells(3, 15).Value = "Greatest % Decrease"
Worksheets(y).Cells(4, 15).Value = "Greatest Total Volume"

Worksheets(y).Columns(19).ClearContents
Worksheets(y).Columns(20).ClearContents
Next y

End Sub



