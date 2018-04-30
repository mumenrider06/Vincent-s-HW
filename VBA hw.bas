Attribute VB_Name = "Module1"
Sub Ticker()
  ' Set an initial variable for holding the brand name
  Dim Ticker As String
  ' Set an initial variable for holding the total per Ticker name
  Dim Ticker_Total As Double
  Ticker_Total = 0
  ' Keep track of the location for each Ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
'Establishes Last row
Lastrow = Range("A" & Rows.Count).End(xlUp).Row
'Counts Worksheets in Workbook
WS_Count = ActiveWorkbook.Worksheets.Count
For NxtWS = 1 To WS_Count
    Worksheets(NxtWS).Activate
  ' Loop through all Tickers
  For i = 2 To Lastrow
    ' Check if we are still within the same Ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the Ticker name
      Ticker = Cells(i, 1).Value
      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      ' Print the Ticker name in the Summary Table
      Range("J" & Summary_Table_Row).Value = Ticker
      ' Print the Ticker Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = Ticker_Total
      Range("K" & Summary_Table_Row).Style = "Currency"
      
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Ticker_Total = 0
    ' If the cell immediately following a row is the same Ticker name...
    Else
      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
    End If
  Next i
    Next NxtWS
End Sub
