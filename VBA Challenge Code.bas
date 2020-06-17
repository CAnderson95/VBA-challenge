Attribute VB_Name = "Module1"
Sub yearly_change()
  ' Set an initial variable for holding the Ticker Name
  Dim Ticker As String

  ' Set an initial variable for holding the Year Change per ticker
  Dim YRCHNG As Double
  YRCHNG = 0
  
    ' Set an initial variable for holding the total volume per ticker
  Dim VolTotal As Double
  VolTotal = 0
  
  ' Set an initial variable for holding the % Change per ticker
  Dim PCTCHNG As Double
  PCTCHNG = 0
  
  Dim vMax

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Loop through all Tickers
  For i = 2 To 705714

    ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker = Cells(i, 1).Value
      
       ' Add to the Volume Total
      VolTotal = VolTotal + Cells(i, 7).Value

      ' Subtract final close price from first open price
      YRCHNG = Cells(i, 3).Value - Cells(i + 261, 6).Value
      
      ' Divide final close price from first open price
      PCTCHNG = (Cells(i, 3).Value / Cells(i + 261, 6).Value)
      
      ' Print the Ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Year Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = YRCHNG
      
        ' Print the Volume Total to the Summary Table
      Range("L" & Summary_Table_Row).Value = VolTotal
      
      ' Print the Year Change to the Summary Table
      Range("K" & Summary_Table_Row).Value = PCTCHNG
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Yearly Change Total
      YRCHNG = 0
      
         ' Reset the Brand Total
      VolTotal = 0
      
       ' Reset the Yearly Change Total
      PCTCHNG = 0

    ' If the cell immediately following a row is the same Ticker...
    Else
    End If
    Next
    End Sub
