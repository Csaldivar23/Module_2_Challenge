# Module_2_Challenge
Module_2_Challenge

Sub quarterly_check()

Dim ws As Worksheet
Dim brand As String
Dim brand_total As Double
Dim summary_table_row As Integer
Dim openprice As Double
Dim closeprice As Double
Dim quarter As Double
Dim percent As Double
Dim lastrow As Long
Dim Row As Long

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ws.Columns("I").ColumnWidth = 20
    ws.Columns("J").ColumnWidth = 20
    ws.Columns("K").ColumnWidth = 20
    ws.Columns("L").ColumnWidth = 20
    ws.Columns("O").ColumnWidth = 20

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    openprice = ws.Cells(2, 3).Value
    summary_table_row = 2
    brand_total = 0

For Row = 2 To lastrow

    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then

        brand = ws.Cells(Row, 1).Value
        closeprice = ws.Cells(Row, 6).Value
        quarter = closeprice - openprice
        percent = (closeprice - openprice) / openprice

        brand_total = ws.Cells(Row, 7).Value + brand_total

        ws.Cells(summary_table_row, 9).Value = brand
        ws.Cells(summary_table_row, 10).Value = quarter
        ws.Cells(summary_table_row, 11).Value = percent
        ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
        ws.Cells(summary_table_row, 12).Value = brand_total

    If quarter < 0 Then
        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
    ElseIf quarter > 0 Then
        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
    End If
    
     If x Then
        maxIncrease = percent
        maxDecrease = percent
        maxIncreaseTicker = brand
        maxDecreaseTicker = brand
        x = False
    Else
        If percent > maxIncrease Then
            maxIncrease = percent
            maxIncreaseTicker = brand
        ElseIf percent < maxDecrease Then
            maxDecrease = percent
            maxDecreaseTicker = brand
        End If
    End If
    
    If brand_total > maxVolume Then
        maxVolume = brand_total
        maxVolumeTicker = brand
    End If
    
        summary_table_row = summary_table_row + 1
        brand_total = 0

    If Row + 1 <= lastrow Then
        openprice = ws.Cells(Row + 1, 3).Value
    End If
    
    Else

        brand_total = brand_total + ws.Cells(Row, 7).Value

    End If

Next Row

Next ws

    Cells(2, 16).Value = maxIncreaseTicker
    Cells(2, 17).Value = maxIncrease
    Cells(3, 16).Value = maxDecreaseTicker
    Cells(3, 17).Value = maxDecrease
    Cells(4, 16).Value = maxVolumeTicker
    Cells(4, 17).Value = maxVolume

End Sub

