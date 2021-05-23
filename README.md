# VBA-Challenge
Final VBA - Challenge
Sub Solve_the_challenge()

 Dim wsCount As Integer
 wsCount = ActiveWorkbook.Worksheets.Count
 
 For ws = 1 To wsCount
 
  Print_the_headers ws
  Analyze_the_stock ws
 
 Next ws
 
End Sub

Sub Analyze_the_stock(currentWorksheet)

 Dim ws As Worksheet
 Set ws = Worksheets(currentWorksheet)
 
 Dim tickerRow As Integer: tickerRow = 2
 
 ' Ignoring the header the data starts from row 2
 Dim currentRow As Double: currentRow = 2
 ' Set the current ticker to the 1st ticker found
 Dim currentTicker As String: currentTicker = ws.Cells(currentRow, 1).Value
 ' Set the year open to the 1st open found
 Dim yearOpen As Double: yearOpen = ws.Cells(currentRow, 3).Value
                      
 Dim yearClose As Double
 
 Dim yearlyChange As Double
 Dim percentChange As Double
 Dim totalStockVolume As Double
 
 'Bonus questions
 Dim greatestPercentIncreaseTicker As String
 Dim greatestPercentIncreaseValue As Double
 Dim greatestPercentDecreaseTicker As String
 Dim greatestPercentDecreaseValue As Double
 Dim greatestTotalVolumeTicker As String
                             
 Dim greatestTotalVolumeValue As Double
 
 ' Continue reading rows until we reach a row with no ticker in the 1st column
 Do Until IsEmpty(ws.Cells(currentRow, 1).Value)
 
  ' We've read the 1st non-header row already so we can proceed from row 3
  currentRow = currentRow + 1
  ticker = ws.Cells(currentRow, 1).Value
  yearClose = ws.Cells(currentRow, 6).Value
 
  If ticker = currentTicker Then
   
   ' We're still on the same ticker so we only add to our stock volume running total
   totalStockVolume = totalStockVolume + ws.Cells(currentRow, 7).Value
 
  Else
 
   ' We found a new ticker! Calculate and write out the sums and percentages first
   yearlyChange = yearClose - yearOpen
   
   ' Print out the ticker row
   ws.Cells(tickerRow, 9).Value = currentTicker
   ws.Cells(tickerRow, 10).Value = yearlyChange
   ws.Cells(tickerRow, 12).Value = totalStockVolume
   
   If yearOpen = 0 Then
   
    ws.Cells(tickerRow, 11).Value = "NaN"
   
   Else
   
    percentChange = (yearlyChange / yearOpen) * 100
    ws.Cells(tickerRow, 11).Value = percentChange
   
   End If
   
   ' Increment ticker row
   tickerRow = tickerRow + 1
   
   ' Set the current ticker and year open to the new values found
   currentTicker = ws.Cells(currentRow, 1).Value
   yearOpen = ws.Cells(currentRow, 3).Value
   
                   
    'Print the Ticker Name in the Ticker Row, Column I
    ws.Range("I" & tickerRow).Value = Ticker_Name
    ' Print the Ticker Name in the Ticker row, Column I
    ws.Range("J" & tickerRow).Value = yearlyChange
    ' Fill "Yearly Change", i.e. yearlyChange with Green and Red colors
    If (yearlyChange > 0) Then
    'Fill column with GREEN color - good
    ws.Range("J" & tickerRow).Interior.ColorIndex = 4
    ElseIf (yearlyChange <= 0) Then
    'Fill column with RED color - bad
    ws.Range("J" & tickerRow).Interior.ColorIndex = 3
     End If
 
  End If
 
 Loop
 
End Sub

Sub Print_the_headers(currentWorksheet)
 
 Dim ws As Worksheet
 Set ws = Worksheets(currentWorksheet)
 
 ws.Range("I1") = "Ticker"
 ws.Range("J1") = "Yearly Change"
 ws.Range("K1") = "Percent Change"
 ws.Range("L1") = "Total Stock Volume"
   
                                                               
                                                               
                                                                 
 ws.Range("P1") = "Ticker"
 ws.Range("Q1") = "Value"
 ws.Range("O2") = "Greatest % Increase"
 ws.Range("O3") = "Greatest % Decrease"
 ws.Range("O4") = "Greatest Total Volume"
 ws.Range("Q2:Q3").NumberFormat = "0.00%"
