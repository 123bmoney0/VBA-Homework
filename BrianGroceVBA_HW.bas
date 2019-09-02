Attribute VB_Name = "Module1"
Sub VBA_Homework()
'Adding All Worksheets
For Each ws In Worksheets

'Variables
Dim Ticker As String
Dim YC As Double
Dim PC As Double
Dim TSV As Double
    TSV = 0
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
Dim OpenPrice As Double
Dim ClosePrice As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Column Headers & Cell Labels
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'Fit Columns
ws.Columns("I:P").EntireColumn.AutoFit

'Loop
For x = 2 To lastrow
    
If ws.Cells(x, 1).Value <> ws.Cells(x - 1, 1).Value Then
    OpenPrice = ws.Cells(x, 3).Value
End If
    
If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
    Ticker = ws.Cells(x, 1).Value
    TSV = TSV + ws.Cells(x, 7).Value
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    ws.Range("L" & Summary_Table_Row).Value = TSV
        ws.Range("L" & Summary_Table_Row).NumberFormat = "#,##0"
    ClosePrice = ws.Cells(x, 6).Value
    YC = ClosePrice - OpenPrice
    ws.Range("J" & Summary_Table_Row).Value = YC
        ws.Range("J" & Summary_Table_Row).NumberFormat = "$0.00"
            If YC < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf YC > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
    If OpenPrice = 0 Then
        PC = 0
    Else
        PC = YC / OpenPrice
    End If
    ws.Range("K" & Summary_Table_Row).Value = PC
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    Summary_Table_Row = Summary_Table_Row + 1
    TSV = 0
Else
    TSV = TSV + ws.Cells(x, 7).Value
End If

Next x

'Bonus Variables
columnk = ws.Range("K2:K" & lastrow)
columnL = ws.Range("L2:L" & lastrow)
Max = Application.WorksheetFunction.Max(columnk)
Min = Application.WorksheetFunction.Min(columnk)
MaxVol = Application.WorksheetFunction.Max(columnL)

ws.Range("P2") = Max
    ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3") = Min
    ws.Range("P3").NumberFormat = "0.00%"
ws.Range("P4") = MaxVol
    ws.Range("P4").NumberFormat = "#,##0"
    
For x = 2 To lastrow
    If ws.Cells(x, 11).Value = Max Then
        ws.Range("O2").Value = ws.Cells(x, 9).Value
    ElseIf ws.Cells(x, 11).Value = Min Then
        ws.Range("O3").Value = ws.Cells(x, 9).Value
    ElseIf ws.Cells(x, 12).Value = MaxVol Then
        ws.Range("O4").Value = ws.Cells(x, 9).Value
    End If
    
 Next x
Next ws


End Sub

