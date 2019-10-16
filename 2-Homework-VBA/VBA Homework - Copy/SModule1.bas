Attribute VB_Name = "Module2"
Sub Total_Stock_Volume()

' Varibles for holding Tickers and their Total Volume around a year
Dim Tickers As String
Dim Total_Volume_Yr As Double

' Count of Total Volume

Total_Volume_Yr = 0

'Loop Through all Tickers
For i = 2 To 705714

' Keep track of summary table location
Dim Summary_Table As Integer
Summary_Table = 2

If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
Tickers = Cells(i, 1).Value

Total_Volume_Yr = Total_Volume_Yr + Cells(i, 7).Value

' Print Tickers and Total Volume in the summary table
Tickers = Range("I" & Summary_Table).Value
Total_Volume_Yr = Range("J" & Summary_Table).Value

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Total_Volume_Yr = 0

Total_Volume_Yr = Total_Volume_Yr + Cells(i, 7).Value

End If

Next i





End Sub
