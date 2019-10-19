Attribute VB_Name = "Module1"
Sub SummaryStock()
Attribute SummaryStock.VB_ProcData.VB_Invoke_Func = "S\n14"
'
' SummaryStock Macro
' loop through all the stocks for one year
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
    ' Find the last row in worksheet
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
' Cell Identification and Column Headings
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"

' Assign value to variables
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Ticker_Name As String
Dim Percent_Change As Double
Dim Volume As Double

' Set Volume variable as zero
Volume = 0

' Set what row to start looping
Dim Row As Double
Row = 2

' Set the column to start looping
Dim Column As Integer
Column = 1

' Long integer designation
Dim i As Long

' Set initial open price (third column)
    Open_Price = Cells(2, Column + 2).Value
' Looping through all tickers
    For i = 2 To LastRow

' Write the ticker symbol and if the cell has a different ticker
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    ' Have it choose the right ticker name
    Ticker_Name = Cells(i, Column).Value
    Cells(Row, Column + 8).Value = Ticker_Name

    ' Show the closing price in cell
        Close_Price = Cells(i, Column + 5).Value

' Write the yearly change
    Yearly_Change = Close_Price - Open_Price
    ' Show the yearly change in cell
    Cells(Row, Column + 9).Value = Yearly_Change

' Write the percent change
    If (Open_Price = 0 And Close_Price = 0) Then
        Percent_Change = 0
    ElseIf (Open_Price = 0 And Close_Price <> 0) Then
        Percent_Change = 1
        'setting initial value of percent change to 1
    Else
        Percent_Change = Yearly_Change / Open_Price
        ' show cell with percent change
        Cells(Row, Column + 10).Value = Percent_Change
        ' show cell with percentage number format
        Cells(Row, Column + 10).NumberFormat = "0.00%"
         End If

' Write the total stock volume
    ' designate place for volume in table
    Volume = Volume + Cells(i, Column + 6).Value
    ' designate value in cell showing volume
    Cells(Row, Column + 11).Value = Volume
    ' Add one to the summary table row
    Row = Row + 1
    ' Open Price reset
    Open_Price = Cells(i + 1, Column + 2)
    ' Volume Total reset
    Volume = 0

' if next cell is same ticker
    Else
        Volume = Volume + Cells(i, Column + 6).Value
    End If

Next i


' Finding last row per WS to get yearly change
YearlyChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row

' Set conditional formatting
For j = 2 To YearlyChangeLastRow
    If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
        Cells(j, Column + 9).Interior.ColorIndex = 10
    ElseIf Cells(j, Column + 9).Value < 0 Then
        Cells(j, Column + 9).Interior.ColorIndex = 3
    End If
Next j

' Challenge
' returning stock with "Greatest Percent Increase"
    Cells(2, Column + 14).Value = "Greatest Percent Increase"
' returning stock with "Greatest Percent Decrease
    Cells(3, Column + 14).Value = "Greatest Percent Decrease"
' returning stock with Greatest Total Volume
    Cells(4, Column + 14).Value = "Greatest Total Volume"
    
' getting ticker name
    Cells(1, Column + 15).Value = "Ticker"
' getting ticker value
    Cells(1, Column + 16).Value = "Value"

' finding greatest value and that stock's ticker through the rows
    For Z = 2 To YearlyChangeLastRow
    If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
        Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
        Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
        Cells(2, Column + 16).NumberFormat = "0.00%"
    ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
        Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
        Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
        Cells(3, Column + 16).NumberFormat = "0.00%"
    ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
        Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
        Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
    End If
Next Z

Next WS

End Sub
