Sub VBA_Challenge()

' Create a variable to hold the counter
  Dim i As Long

   ' Set a variable for worksheet
  Dim ws As Worksheet

  ' Set an initial variable for holding the Ticker Symbol
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the First Open Price
  Dim First_Open As Double

  ' Set an initial variable for holding the Last Closing Price
  Dim Last_Close As Double

  ' Set an initial variable for holding the total volume
  Dim Total_Volume As Double

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Long

  ' Set a variable for and determine Last Row
  Dim LastRow As Long
  
  ' Declare variables for Quarterly Change and Percent Change
  Dim Quarterly_Change As Double
  Dim Percent_Change As Double

  For Each ws In ActiveWorkbook.Worksheets
  
    Summary_Table_Row = 2
    Total_Volume = 0
    First_Open = ws.Cells(2, 3).Value
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through all stock price data
    For i = 2 To LastRow
    Debug.Print i

      ' Check if we are still within the same ticker symbol, if we are not...
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Create Colums Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

         'Create Added Functionality Headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Set, Print the Ticker Symbol
        Ticker_Symbol = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

        ' Set the Last Close
        Last_Close = ws.Cells(i, 6).Value

        ' Calculate, Print Quartely Change (last closing price minus first opening price) and conditional formatting
        Quarterly_Change = Last_Close - First_Open
          If Quarterly_Change >= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
          Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
          End If
        ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change

        ' Calculate, Print Percent Change
        Percent_Change = Quarterly_Change / First_Open
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change

        ' Add to, Print the Total Volume (by Ticker Symbol)
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
      
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1

        ' Reset the Total Volume
        Total_Volume = 0
        
        ' Reset First_open
        First_Open = ws.Cells(i + 1, 3).Value

      ' If the cell immediately following a row is the same ticker...
      Else

        ' Add to the Total Volume
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value

      End If

    Next i
    
  ' Declare variables for percentages and volume
    Dim Greatest_Percent_Increase as Double
    Dim Greatest_Percent_Decrease as Double
    Dim Greatest_Total_Volume as Double

    'Declare variables for match function
    Dim Greatest_Increase_Match As Integer
    Dim Greatest_Decrease_Match As Integer
    Dim Greatest_Volume_Match As Integer

    ' Find, print greatest increase and match the corresponding ticker
    Greatest_Percent_Increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 17) = Greatest_Percent_Increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    Greatest_Increase_Match = Application.WorksheetFunction.Match(Greatest_Percent_Increase, ws.Range("K:K"), 0)
    ws.Cells(2, 16) = ws.Cells(Greatest_Increase_Match, 9)

    ' Find, print greatest decrease and match the corresponding ticker
    Greatest_Percent_Decrease = application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 17) = Greatest_Percent_Decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    Greatest_Decrease_Match = Application.WorksheetFunction.Match(Greatest_Percent_Decrease, ws.Range("K:K"), 0)
    ws.Cells(3, 16) = ws.Cells(Greatest_Decrease_Match, 9)

    ' Find, print greatest total volume and match the corresponding ticker
    Greatest_Total_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 17) = Greatest_Total_Volume
    Greatest_Volume_Match = Application.WorksheetFunction.Match(Greatest_Total_Volume, ws.Range("L:L"), 0)
    ws.Cells(4, 16) = ws.Cells(Greatest_Volume_Match, 9)

    ' Autofit cells with new information    
    ws.Columns("I:Q").AutoFit

  Next ws
  
End Sub