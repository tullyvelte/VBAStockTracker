Sub combine_tracker_three()

' make it global?
Dim ws As Worksheet
For Each ws In Worksheets

'Dim corruptRowCount As Double
 ' corruptRowCount = 0
  ' Initialize with an arbitrary length of 25 to keep track of example corrupt rows
  'Dim corruptRowExamples As String
  'For Each Row In Range("C").Rows
    ' Set a flag to check if any column value is empty. Initialize as false.
    'Dim emptyValue As Boolean
    'emptyValue = False
    ' Iterate through each column value for a given row.
    'For Each cell In Row.Cells
      ' If empty, IsEmpty is a built in function (BiF), mark as true
     ' If IsEmpty(cell.Value) Then
      '  emptyValue = True
        ' After we mark one column as empty, we no longer need to check the rest of the columns.
        ' Exit this nested for loop.
       ' Exit For
      'End If
    'Next cell
    'If emptyValue Then
      ' Color mark the range on the Excel sheet.
      ' Visually inspecting the rows may tell you what's wrong.
     ' Row.Interior.Color = RGB(255, 0, 0)
      ' Update the total corruptRowCount if the flag was set to true.
      'corruptRowCount = corruptRowCount + 1
      ' Update the array with the row number
'      corruptRowExamples(corruptRowCount) = Str(Row.Row)
    'End If
  'Next Row

    ' Make variables for ticker name, and volume, stock start and end price
    Dim ticker_name As String
    Dim stock_volume As Double
    stock_volume = 0
    Dim yearly_change As Double
    yearly_change = 0
    Dim start As Double
    start = 2
    Dim start_cost As Double
    Dim close_cost As Double
    Dim percent_change As Double
    percent_change = 0
   ' Dim dblMin As Double
   ' Dim dblMax As Double
    
    

    'Set up table location for logging ticker and volume (stealing from CC exercise)
    Dim summary_table_row As Double
    summary_table_row = 2

    ' Determine the Last Row
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker_name = Cells(i, 1).Value
            stock_volume = stock_volume + Cells(i, 7).Value
            'yearly_change = Cells(i, 6).Value - Cells(i, 3).Value
            start_cost = Cells(start, 3).Value
            close_cost = Cells(i, 6).Value
            yearly_change = close_cost - start_cost
            If start_cost = 0 Then start_cost = 0.001
            percent_change = (yearly_change / start_cost) * 100
            
        

            ' print the totals
            Range("I" & summary_table_row).Value = ticker_name
            Range("J" & summary_table_row).Value = yearly_change
            Range("K" & summary_table_row).Value = percent_change
            Range("L" & summary_table_row).Value = stock_volume
            ' Range("N" & summary_table_row).Value = dblMin
        
            ' formatting?
        
            Set ChangeRange = Range("J" & summary_table_row)
                For Each cell In ChangeRange
                 Select Case cell.Value
                    Case Is >= 0
                    cell.Interior.ColorIndex = 4
                    Case Else
                    cell.Interior.ColorIndex = 3
                    End Select
            
                Next

            summary_table_row = summary_table_row + 1
            start = i + 1

            stock_volume = 0
            yearly_change = 0
            percent_change = 0

            Else
            stock_volume = stock_volume + Cells(i, 7).Value
            yearly_change = Cells(i, 6).Value - Cells(i, 3).Value

        End If

    Next i
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Vol"
   
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest Percent Increase"
    ws.Cells(2, 17) = WorksheetFunction.Max(ws.Range("k2:k" & LastRow))
    ws.Cells(3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(3, 17) = WorksheetFunction.Min(ws.Range("k2:k" & LastRow))
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 17) = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))

    ' match ticker to values
    incrMatch = Application.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
    dcrMatch = Application.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
    volMatch = Application.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)

   'ws.Cells(2, 16) = ws.Cells(incrMatch + 1, 1)
   'ws.Cells(3, 16) = ws.Cells(dcrMatch + 1, 1)
   'ws.Cells(4, 16) = ws.Cells(volMatch + 1, 1)


    'Worksheet function MIN returns the smallest value in a range
    'ws.Cells(2, 14) = WorksheetFunction.Max(ws.Range("J2:J" & rowCount))
    'dblMin = Application.WorksheetFunction.Min(Rng)

  Next ws

End Sub
