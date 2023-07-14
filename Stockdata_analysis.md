Sub Stockdata_analysis()

 ' Loop through all sheets
 Dim ws As Worksheet
   For Each ws In ThisWorkbook.Worksheets
   ws.Activate
   
   'Print row names
   Range("I1") = "Ticker Name"
   Range("J1") = "Yearly change"
   Range("K1") = "Percentage change"
   Range("L1") = "Total Volume"
   Range("N2") = "Greatest % increase"
   Range("N3") = "Greatest % decrease"
   Range("N4") = " Greatest total volume"
   
  ' Set an initial variable for holding the ticker name
  Dim ticker_name As String
  Dim Total_Volume As Double
  
  Total_Volume = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock values
  Last_row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Last_row
If Cells(i - 1, 1) <> Cells(i, 1) Then
        Opening_price = Cells(i, 3)

    ' Check if we are still within the same ticker if it is not...
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      ticker_name = Cells(i, 1).Value

      ' Add to the Brand Total
      Total_Volume = Total_Volume + Cells(i, 7).Value
      Cells(i, 12).Value = Total_Volume
      Closing_price = Cells(i, 6).Value
      
      Yearly_change = Closing_price - Opening_price
      Percentage_change = Yearly_change / Opening_price

      ' Print in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker_name
      Range("L" & Summary_Table_Row).Value = Total_Volume
      Range("J" & Summary_Table_Row).Value = Yearly_change
      Range("K" & Summary_Table_Row).Value = Percentage_change
      Columns("K:K").NumberFormat = "0.00%"
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total volume for the next ticker
      Total_Volume = 0

    ' If the cell immediately following a row is the same ticker
    Else

      ' Add to the Brand Total
      Total_Volume = Total_Volume + Cells(i, 7).Value

    End If
    Next i
   Dim greatest_percentage_increase As Long
   Dim greatest_percentage_decrease As Long
   Dim greatest_total_volume As Long
   Dim lastrow As Long
    'Color index
    For j = 2 To lastrow_summary
    lastrow_summary = Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(j, 1).Value >= 0 Then
    Cells(j, 1).Interior.ColorIndex = 3
    ElseIf Cells(j, 1).Value < 0 Then
    Cells(j, 1).Interior.ColorIndex = 4
    End If
    'Greatest percentage increase
If Cells(j, 11) > greatest_percentage_increase Then
 greatest_percentage_increase = Cells(j, 11)
 Cells(2, 17) = greatest_percentage_increase
 Cells(2, 17).NumberFormat = "0.00%"
 Cells(2, 16) = Cells(j, 9)
 End If
    'Greatest percentage decrease
If Cells(j, 11) > greatest_percentage_decrease Then
 greatest_percentage_decrease = Cells(j, 11)
 Cells(3, 17) = greatest_percentage_decrease
 Cells(3, 17).NumberFormat = "0.00%"
 Cells(3, 16) = Cells(j, 9)
 End If
    'Greatest total volume
 If Cells(j, 12) > greatest_total_volume Then
 greatest_total_volume = Cells(j, 12)
 Cells(4, 17) = greatest_total_volume
 Cells(4, 16) = Cells(j, 9)
 End If
 
 Next j
 Next
 
End Sub

