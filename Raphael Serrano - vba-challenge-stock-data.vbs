' modified sorting code from https://trumpexcel.com/sort-data-vba/
' just in case the script was ran on unsorted data ;)
Private Sub SortMultipleColumns()

Dim maxRange As Long
Dim cellSize As String

maxRange = WorksheetFunction.CountA(Range("A:A"))
cellSize = "A1" & ":G" & maxRange

' MsgBox (cellSize)

WorksheetFunction.CountA (Range("A:A"))

With ActiveSheet.Sort
     .SortFields.Add Key:=Range("A1"), Order:=xlAscending
     .SortFields.Add Key:=Range("B1"), Order:=xlAscending
     .SetRange Range(cellSize)
     .Header = xlYes
     .Apply
End With

End Sub

' -----------------------------------------------------------
' modified conditional formatting from https://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/
Private Sub formatCells()
Dim rg As Range
Dim cond1, cond2 As FormatCondition
Set rg = Range("J2", Range("J2").End(xlDown))

' format column K and Q to %
Range("K2", Range("K2").End(xlDown)).NumberFormat = "0.00%"
Range("Q2", "Q3").NumberFormat = "0.00%"

'clear any existing conditional formatting
rg.FormatConditions.Delete

'define the rule for each conditional format
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")

'define the format applied for each conditional format
With cond1
.Interior.Color = vbGreen
.Font.Color = vbBlack
End With

With cond2
.Interior.Color = vbRed
.Font.Color = vbBlack
End With

End Sub

' -----------------------------------------------------------
' private subprocess for calculating the ticker stats
Private Sub ComputeTicker()

Dim maxRange As Long            ' variable for the maximum non-blank row
Dim currentTicker As String     ' ticker of the current record
Dim workingTicker As String     ' ticker of the previous record
Dim tickVolSum As Double          ' sum of volume for the ticker
Dim earliestOpen As Double      ' open value for the earliest date
Dim lastClose As Double         ' close value for the last date
Dim resultingDataRow As Long    ' row number for the data results

' initialise variables
workingTicker = Cells(2, 1).Value ' first working ticker is the first ticker instance
resultingDataRow = 1
earliestOpen = 0

' get maximum non-blank rows
maxRange = WorksheetFunction.CountA(Range("A:A"))

' initialise results header
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("l1") = "Total Stock Volume"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

' sort columns
Call SortMultipleColumns

' loop through each row
For Row = 2 To maxRange
    ' assign the current ticker value
    currentTicker = Cells(Row, 1).Value
    
    ' replace working ticker (previous ticker) with current ticker if the ticker changes
    If currentTicker <> workingTicker Then
        
        ' increment row results
        resultingDataRow = resultingDataRow + 1
        
        ' get the close value of the previous row
        lastClose = Cells(Row - 1, 6).Value
        
        ' write the ticker data to the spreadsheet
        Cells(resultingDataRow, 9).Value = workingTicker
        ' Yearly Change
        Cells(resultingDataRow, 10).Value = lastClose - earliestOpen
        
        ' percent change PREVENT DIVIDE BY ZERO ERROR
        If earliestOpen = 0 Or lastClose = 0 Then
            ' prevent divide by zero error
            Cells(resultingDataRow, 11).Value = 0
        Else
            Cells(resultingDataRow, 11).Value = (lastClose - earliestOpen) / earliestOpen
        End If
        
        
        ' total stock volume
        Cells(resultingDataRow, 12).Value = tickVolSum
        
        
        ' update working ticker and initialise values
        workingTicker = currentTicker
        tickVolSum = 0
        earliestOpen = 0
        lastClose = 0
        
    End If
    
    ' check data columns of the row
    If earliestOpen = 0 Then
        earliestOpen = Cells(Row, 3).Value
    End If
        
    ' sum up ticker volume
    tickVolSum = tickVolSum + CLng(Cells(Row, 7).Value)
    
Next Row

' enter last ticker details
' -----------------------------------------------------------
' increment row results
resultingDataRow = resultingDataRow + 1
        
' get the close value of the previous row
lastClose = Cells(maxRange, 6).Value

' write the ticker data to the spreadsheet
Cells(resultingDataRow, 9).Value = workingTicker
Cells(resultingDataRow, 10).Value = lastClose - earliestOpen ' Yearly Change
Cells(resultingDataRow, 11).Value = (lastClose - earliestOpen) / earliestOpen ' percent change
Cells(resultingDataRow, 12).Value = tickVolSum ' total stock volume
' -----------------------------------------------------------

End Sub

' -----------------------------------------------------------
'Private subprocess for analyzing the results of ticker processing
Private Sub findgreatest()

Dim maxIncRow, maxDecRow, maxVolRow As Long

maxIncRow = 2
maxDecRow = 2
maxVolRow = 2

' get number of result rows
maxRange = WorksheetFunction.CountA(Range("I:I"))

For Row = 3 To maxRange

    ' get the row number fo the largest increase
    If Cells(Row, 11).Value > Cells(maxIncRow, 11).Value Then
        maxIncRow = Row
    End If
    
    ' get the row number fo the largest decrease
    If Cells(Row, 11).Value < Cells(maxDecRow, 11).Value Then
        maxDecRow = Row
    End If
    
    ' get the row number for the maximum volume
    If Cells(Row, 12).Value > Cells(maxVolRow, 12).Value Then
        maxVolRow = Row
    End If

Next Row

' enter the details in the table
' 16 17
Cells(2, 16).Value = Cells(maxIncRow, 9).Value
Cells(3, 16).Value = Cells(maxDecRow, 9).Value
Cells(4, 16).Value = Cells(maxVolRow, 9).Value

Cells(2, 17).Value = Cells(maxIncRow, 11).Value
Cells(3, 17).Value = Cells(maxDecRow, 11).Value
Cells(4, 17).Value = Cells(maxVolRow, 12).Value

End Sub

' -----------------------------------------------------------
' main subprocess for analyzing stocks
Sub Main()

Dim tabCount As Integer

tabCount = ActiveWorkbook.Worksheets.Count

' loop through all worksheets
For ws = 1 To tabCount

    Worksheets(ws).Select
    ' run logic for each worksheet
    Call ComputeTicker
    Call findgreatest
    Call formatCells

Next ws

End Sub

