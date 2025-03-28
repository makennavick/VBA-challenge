Attribute VB_Name = "Module1"
Sub Looper()


Dim ws As Worksheet
For Each ws In Worksheets

' SET INITIAL VARIABLES
' -------------------

Dim maxVol As Variant
Dim maxVolTicker As String
Dim maxInc As Double
Dim maxIncTicker As String
Dim maxDecr As Double
Dim maxDecrTicker As String

' Set initial variable for the stock ticker
Dim currentTicker As String
currentTicker = "AAF"

' Create variable for row counter
Dim row As Long

' Set initial variable for the stock's quarterly change & percent change
Dim quarterlyChange As Double
quarterlyChange = 0
Dim percentChange As Double
percentChange = 0

' Set initial var for total vol
Dim totalVol As Variant
totalVol = 0

' start of stock
Dim firstRow As Long
firstRow = 2

' Keep track of the location for each ticker in the summary table
Dim summaryTableRow As Integer
summaryTableRow = 1

' Loop through all rows
For row = 2 To 95000

    ' if first stock...
     If row = 2 Then
         ' Add to the total volume & add to summary table
          totalVol = totalVol + CLng(ws.Cells(row, 7).Value)
          firstRow = row
    
    ' IF NEW STOCK...
    ' Check if the next row uses the same ticker. If it's a NEW ticker..
    ' ElseIf ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
    ElseIf currentTicker <> ws.Cells(row, 1).Value Then
      
      ' Add one to the summary table row
      summaryTableRow = summaryTableRow + 1
      ws.Cells(summaryTableRow, "I").Value = currentTicker
      
      ' Add to the previous stock's total volume & add to summary table
      ws.Cells(summaryTableRow, "L").Value = totalVol
      If totalVol > maxVol Then
        maxVol = totalVol
        maxVolTicker = currentTicker
      End If
     
      
      ' Quarterly change
      quarterlyChange = ws.Cells(row - 1, 6).Value - ws.Cells(firstRow, 3).Value
      ws.Cells(summaryTableRow, "J").Value = quarterlyChange
      
      ' Change color
      If quarterlyChange > 0 Then
        ws.Cells(summaryTableRow, "J").Interior.ColorIndex = 4
      ElseIf quarterlyChange < 0 Then
        ws.Cells(summaryTableRow, "J").Interior.ColorIndex = 3
      End If
      
      ' Set percent change
      percentChange = quarterlyChange / ws.Cells(firstRow, 3).Value
      ws.Cells(summaryTableRow, "K").Value = FormatPercent(percentChange)
      
      If percentChange > maxInc Then
        maxInc = percentChange
        maxIncTicker = currentTicker
      ElseIf percentChange < maxDecr Then
        maxDecr = percentChange
        maxDecrTicker = currentTicker
      End If
      
      ' PREP FOR NEW TICKER
      totalVol = ws.Cells(row, 7).Value
      firstRow = row
      currentTicker = ws.Cells(row, 1).Value

    ' If the cell immediately following a row is the same stock...
    Else
        ' add to total
        totalVol = totalVol + ws.Cells(row, 7).Value
        
    End If

  Next row
  
  
ws.Cells(2, 16).Value = maxIncTicker
ws.Cells(2, 17).Value = maxInc
ws.Cells(3, 16).Value = maxDecrTicker
ws.Cells(3, 17).Value = FormatPercent(maxDecr)
ws.Cells(4, 16).Value = maxVolTicker
ws.Cells(4, 17).Value = maxVol
  
  
Next ws
End Sub
