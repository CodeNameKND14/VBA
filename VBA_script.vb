Sub stock_analysis():

    ' Setting the dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Single
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Single
    Dim averageChange As Single
    Dim ws As Worksheet

    For Each ws In Worksheets
    ' Setting the values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0

    ' Setting the title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

    ' Gettting the row number of the last row with data
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Storing results in variables
                total = total + ws.Cells(i, 7).Value

                ' Handling zero total volume
                If total = 0 Then
                    '  Printing the results
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0

                Else
                    ' Finding the First non zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If

                    ' Calculating Changes
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = Round((change / ws.Cells(start, 3) * 100), 2)

                    ' Starting the next stock ticker
                    start = i + 1

                    ' Printing the results to a separate worksheet
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & percentChange
                    ws.Range("L" & 2 + j).Value = total

                    ' Coloring positives green and negatives red
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select

                End If

                ' Resetting the variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0

            ' If ticker is still the same adding results
            Else
                total = total + ws.Cells(i, 7).Value

            End If

        Next i

        ' Taking the max and min and place them in a separate part in the worksheet
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

        ' Returning one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

        ' Final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)

    Next ws

End Sub
