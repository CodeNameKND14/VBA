

1. 
Sub Code_one()
Dim Ticker As String
Dim Total_Volume As Double
Dim Summary_Table As Integer
Total_Volume = 0
Summary_Table = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
    
        Range("I" & Summary_Table).Value = Ticker
        Range("J" & Summary_Table).Value = Total_Volume
    
        Summary_Table = Summary_Table + 1
        Total_Volume = 0

    Else
        Total_Volume = Total_Volume + Cells(i, 7).Value

    End If
Next i

End Sub



2. 

Sub Code_two()
Dim Ticker As String
Dim Total_Volume As Double
Dim Summary_Table As Integer
Dim Yearly_Change As Double
Dim Percent_Change As Double 
Dim openstock As Double
Dim closingstock As Double
Dim NewLine As Boolean
Total_Volume = 0
Summary_Table = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
NewLine = True

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

On Error Resume Next

For i = 2 To LastRow

    If Cells(i + 1, 1).Value = Cells(i, 1).Value And NewLine = True Then
        openstock = Cells(i, 3).Value
        NewLine = False

    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
        closingstock = Cells(i, 6).Value
        Yearly_Change = closingstock - openstock
        Percent_Change = (Yearly_Change / (openstock)) * 100    ' I think this i right

        Range("I" & Summary_Table).Value = Ticker
        Range("J" & Summary_Table).Value = Yearly_Change
        Range("K" & Summary_Table).Value = Percent_Change
        Range("L" & Summary_Table).Value = Total_Volume
        If Range("J" & Summary_Table).Value < 0 Then
            Range("J" & Summary_Table).Interior.ColorIndex = 3
        Else
           Range("J" & Summary_Table).Interior.ColorIndex = 4
        End If

        Summary_Table = Summary_Table + 1
        Total_Volume = 0
        
        NewLine = True

    Else
        Total_Volume = Total_Volume + Cells(i, 7).Value

    End If
Next i

End Sub
