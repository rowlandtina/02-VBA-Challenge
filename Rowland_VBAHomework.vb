Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
    Next ws
End Sub
Sub SetTitle()
    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "TotalStock Volume"
End Sub
Sub CalculateSummary()
    ' Start writing your code here
    'Define variables
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim Volume As Double
Dim RowCount As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim SummaryRow As Integer
'store values into variables
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
SummaryRow = 2
OpenPrice = Cells(2, 3).Value
TotalStockVolume = 0
'loop through the values
For i = 2 To RowCount
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Print the ticker values
        Ticker = Cells(i, 1).Value
        Range("I" & SummaryRow).Value = Ticker
        'Calculate yearly change
        ClosePrice = Cells(i, 6).Value
        YearlyChange = (ClosePrice - OpenPrice)
        Range("J" & SummaryRow).Value = YearlyChange
        'Conditional Formatting for Yearly Change
        If YearlyChange > 0 Then
        Range("J" & SummaryRow).Interior.ColorIndex = 4
        Else
        Range("J" & SummaryRow).Interior.ColorIndex = 3
        End If
        'Calculate percent change
        If (OpenPrice = 0 And ClosePrice = 0) Then
            PercentChange = 0
        ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
            PercentChange = 1
        Else
            PercentChange = (YearlyChange / OpenPrice)
            Range("K" & SummaryRow).Value = PercentChange
            Range("K" & SummaryRow).NumberFormat = "0.00%"
        End If
        'Calculate total stock volume
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        Range("L" & SummaryRow).Value = TotalStockVolume
        SummaryRow = SummaryRow + 1
        'reset stock volume & open price
        TotalStockVolume = 0
        'This makes it take the first open price of the next ticker occurance
        OpenPrice = Cells(i + 1, 3).Value
    Else
        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    End If
Next i
    'Debug.Print ActiveSheet.Name
    Call SetTitle
End Sub

