Attribute VB_Name = "Module1"
Sub VBA_HW():
 ' Declare Current as a worksheet object variable.
         Dim ws As Worksheet

 ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets

' Set an initial variable for holding the brand name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per credit card brand
  Dim Total_Volume As Double
  Total_Volume = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  
   ' Keep track of the location for each credit card brand in the summary table
  Dim Open_Price_Row As Double
  Open_Price_Row = 2
  
  Open_Price = ws.Cells(2, 3).Value

  ' Loop through all credit card purchases

    ' Ticker
        ws.Range("I1").Value = "Ticker"
    ' Quarterly Change
        ws.Range("J1").Value = "Quarterly Change"
    ' Percent Change
        ws.Range("K1").Value = "Percent Change"
    ' Total Stock Volume
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ' Loop through from numbers 1 through 20
    For i = 2 To RowCount
    
     ' Check if we are still within the same credit card brand, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Brand name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Brand Total
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      Quarterly_Change = ws.Cells(i, 6).Value - Open_Price
      Percentage_Change = Quarterly_Change / Open_Price
      
      ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
      ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
      ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

      ' Print the Brand Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Volume
      
      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      
      
      
      
      
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Total_Volume = 0
      Open_Price = ws.Cells(i + 1, 3).Value
      Percentage_Change = 0
      Quarterly_Change = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    End If
    
    
    Next i
     ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
     ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & RowCount))
     ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
     
     ws.Range("Q2").NumberFormat = "0.00%"
     ws.Range("Q3").NumberFormat = "0.00%"
    
    max_index = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & RowCount), 0)
    min_index = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & RowCount), 0)
    max_volume_index = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & RowCount), 0)
    
    ws.Range("P2").Value = ws.Cells(max_index + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(min_index + 1, 9).Value
    ws.Range("P4").Value = ws.Cells(max_volume_index + 1, 9).Value
    
    
    
    Next ws
End Sub

