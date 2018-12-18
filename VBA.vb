<<<<<<< HEAD
Sub HWVBA()

Loops ("2016")
Loops ("2015")
Loops ("2014")
FormatSheets
MaxSheets
End Sub

Sub Loops(sheetname)

Sheets(sheetname).Select

Dim ticker As String
Dim vol As Integer
Dim nrow As Double
Dim int_summary_rows As Integer
Dim total As Double
Dim LastRow As Double

LastRow = ActiveSheet.Range("A1").End(xlDown).row
int_summary_rows = 2
For nrow = 2 To LastRow
If Cells(nrow, 1) <> Cells(nrow + 1, 1) Then
    Cells(int_summary_rows, 9) = Cells(nrow, 1)
    Cells(int_summary_rows, 13) = total + 1  '<- ORIGINAL
    
    'Cells(int_summary_rows, 12) = total +
    'VolSum = VolSum + Cells(nrow, 7)
    Cells(int_summary_rows, 12) = VolSum + Cells(nrow, 7)
    'Up to that was for total volume.
    int_summary_rows = int_summary_rows + 1
    total = 0
    last_ticker = Cells(nrow, 6).Value
    last_ticker = last_ticker
    'Cells(int_summary_rows - 1, 10) = last_ticker
    Cells(int_summary_rows - 1, 10) = last_ticker - open_ticker
    Change = Cells(int_summary_rows - 1, 10)
    If open_ticker = 0 Then Cells(int_summary_rows - 1, 11) = 0 Else Cells(int_summary_rows - 1, 11) = Cells(int_summary_rows - 1, 10) / open_ticker
    Cells(int_summary_rows - 1, 11).NumberFormat = "0%"
Else
    VolSum = VolSum + Cells(nrow, 7)
    total = total + 1
    If total = 1 Then open_ticker = Cells(nrow, 3).Value
    open_ticker = open_ticker
End If
Next
End Sub
Sub FormatSheets()
Formatting ("2016")
Formatting ("2015")
Formatting ("2014")
End Sub

Sub Formatting(sheet)
'Conditional formatting that will highlight positive change in green and negative change in red.
'Dim ncol As Integer
Sheets(sheet).Select
Dim nrow As Integer
Dim LastRow As Integer

LastRow = ActiveSheet.Range("I2").End(xlDown).row
nrow = 2
For nrow = 2 To LastRow
If Cells(nrow, 10).Value < 0 Then Cells(nrow, 10).Interior.Color = RGB(216, 18, 79) Else Cells(nrow, 10).Interior.Color = RGB(71, 185, 76)
Next nrow
End Sub
Sub MaxSheets()
FindMax ("2014")
FindMax ("2015")
FindMax ("2016")
End Sub


Sub FindMax(sheet)
Dim MaxTicker As String
Dim Max As Double
Dim Min As Double
Dim MaxTickerVol As String
Sheets(sheet).Select

firstRow = 2
ncol = 11
Max = ActiveSheet.Cells(firstRow, ncol)
ncolticker = 9
LastRow = ActiveSheet.Range("A1").End(xlDown).row

'Greatest % increase.
For i = firstRow + 1 To LastRow
    If ActiveSheet.Cells(i, ncol) > Max Then
    Max = ActiveSheet.Cells(i, ncol)
    MaxTicker = Cells(i, ncolticker)
    Else
    End If
Next

Range("P2") = MaxTicker
Range("Q2") = Max
Min = ActiveSheet.Cells(firstRow, ncol)
'Greatest % decrease.
For i = firstRow + 1 To LastRow
    If ActiveSheet.Cells(i, ncol) < Min Then
     Min = ActiveSheet.Cells(i, ncol)
    MinTicker = Cells(i, ncolticker)
    Else
    End If
Next

Range("P3") = MinTicker
Range("Q3") = Min

'Greatest volume.
ncolvol = 12
Max = ActiveSheet.Cells(firstRow, ncolvol)
LastRow = ActiveSheet.Range("I1").End(xlDown).row
For i = firstRow + 1 To LastRow
    If ActiveSheet.Cells(i, ncolvol) > Max Then
    Max = ActiveSheet.Cells(i, ncolvol)
    MaxTickerVol = ActiveSheet.Cells(i, ncolticker)
    Else
    End If
Next

Range("P4") = MaxTickerVol
Range("Q4") = Max
'ActiveSheet.Cells(1, 14) = Max
'ActiveSheet.Cells(1, 13) = MaxTickerVol

End Sub
    
Sub Original()
    Selection = Range("P2")
    Range("Q2") = WorksheetFunction.Max(Range("K2:K" & LastRow))
    Range("Q3") = WorksheetFunction.Min(Range("K2:K" & LastRow))
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))
End Sub


=======
Sub HWVBA()

Loops ("2016")
Loops ("2015")
Loops ("2014")
FormatSheets
MaxSheets
End Sub

Sub Loops(sheetname)

Sheets(sheetname).Select

Dim ticker As String
Dim vol As Integer
Dim nrow As Double
Dim int_summary_rows As Integer
Dim total As Double
Dim LastRow As Double

LastRow = ActiveSheet.Range("A1").End(xlDown).row
int_summary_rows = 2
For nrow = 2 To LastRow
If Cells(nrow, 1) <> Cells(nrow + 1, 1) Then
    Cells(int_summary_rows, 9) = Cells(nrow, 1)
    Cells(int_summary_rows, 13) = total + 1  '<- ORIGINAL
    
    'Cells(int_summary_rows, 12) = total +
    'VolSum = VolSum + Cells(nrow, 7)
    Cells(int_summary_rows, 12) = VolSum + Cells(nrow, 7)
    'Up to that was for total volume.
    int_summary_rows = int_summary_rows + 1
    total = 0
    last_ticker = Cells(nrow, 6).Value
    last_ticker = last_ticker
    'Cells(int_summary_rows - 1, 10) = last_ticker
    Cells(int_summary_rows - 1, 10) = last_ticker - open_ticker
    Change = Cells(int_summary_rows - 1, 10)
    If open_ticker = 0 Then Cells(int_summary_rows - 1, 11) = 0 Else Cells(int_summary_rows - 1, 11) = Cells(int_summary_rows - 1, 10) / open_ticker
    Cells(int_summary_rows - 1, 11).NumberFormat = "0%"
Else
    VolSum = VolSum + Cells(nrow, 7)
    total = total + 1
    If total = 1 Then open_ticker = Cells(nrow, 3).Value
    open_ticker = open_ticker
End If
Next
End Sub
Sub FormatSheets()
Formatting ("2016")
Formatting ("2015")
Formatting ("2014")
End Sub

Sub Formatting(sheet)
'Conditional formatting that will highlight positive change in green and negative change in red.
'Dim ncol As Integer
Sheets(sheet).Select
Dim nrow As Integer
Dim LastRow As Integer

LastRow = ActiveSheet.Range("I2").End(xlDown).row
nrow = 2
For nrow = 2 To LastRow
If Cells(nrow, 10).Value < 0 Then Cells(nrow, 10).Interior.Color = RGB(216, 18, 79) Else Cells(nrow, 10).Interior.Color = RGB(71, 185, 76)
Next nrow
End Sub
Sub MaxSheets()
FindMax ("2014")
FindMax ("2015")
FindMax ("2016")
End Sub


Sub FindMax(sheet)
Dim MaxTicker As String
Dim Max As Double
Dim Min As Double
Dim MaxTickerVol As String
Sheets(sheet).Select

firstRow = 2
ncol = 11
Max = ActiveSheet.Cells(firstRow, ncol)
ncolticker = 9
LastRow = ActiveSheet.Range("A1").End(xlDown).row

'Greatest % increase.
For i = firstRow + 1 To LastRow
    If ActiveSheet.Cells(i, ncol) > Max Then
    Max = ActiveSheet.Cells(i, ncol)
    MaxTicker = Cells(i, ncolticker)
    Else
    End If
Next

Range("P2") = MaxTicker
Range("Q2") = Max
Min = ActiveSheet.Cells(firstRow, ncol)
'Greatest % decrease.
For i = firstRow + 1 To LastRow
    If ActiveSheet.Cells(i, ncol) < Min Then
     Min = ActiveSheet.Cells(i, ncol)
    MinTicker = Cells(i, ncolticker)
    Else
    End If
Next

Range("P3") = MinTicker
Range("Q3") = Min

'Greatest volume.
ncolvol = 12
Max = ActiveSheet.Cells(firstRow, ncolvol)
LastRow = ActiveSheet.Range("I1").End(xlDown).row
For i = firstRow + 1 To LastRow
    If ActiveSheet.Cells(i, ncolvol) > Max Then
    Max = ActiveSheet.Cells(i, ncolvol)
    MaxTickerVol = ActiveSheet.Cells(i, ncolticker)
    Else
    End If
Next

Range("P4") = MaxTickerVol
Range("Q4") = Max
'ActiveSheet.Cells(1, 14) = Max
'ActiveSheet.Cells(1, 13) = MaxTickerVol

End Sub
    
Sub Original()
    Selection = Range("P2")
    Range("Q2") = WorksheetFunction.Max(Range("K2:K" & LastRow))
    Range("Q3") = WorksheetFunction.Min(Range("K2:K" & LastRow))
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & LastRow))
End Sub


>>>>>>> e6ca9a0cc5bd4a73f6822d8b43aea7f47ed7e310
