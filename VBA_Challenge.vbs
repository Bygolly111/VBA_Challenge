Sub stock_data()

For Each ws In Worksheets

  ' Set up Headings in each Tab
  Dim Stock_Name As String
  Dim Heading1 As String
  Dim Heading2 As String
  Dim Heading3 As String
  Dim Heading4 As String
  Dim Heading5 As String
  Dim Heading6 As String
  Dim Heading7 As String
  Dim Heading8 As String
  
  
  Heading1 = "Ticker"
  Heading2 = "Yearly Change"
  Heading3 = "Percent Change"
  Heading4 = "Total Stock Volume"
  Heading5 = "Value"
  Heading6 = "Greatest % Increase"
  Heading7 = "Greatest % Decrease"
  Heading8 = "Greatest Total Volume"
  
  
  ws.Cells(1, 9) = Heading1
  ws.Cells(1, 10) = Heading2
  ws.Cells(1, 11) = Heading3
  ws.Cells(1, 12) = Heading4
  ws.Cells(1, 17) = Heading1
  ws.Cells(1, 18) = Heading5
  ws.Cells(2, 16) = Heading6
  ws.Cells(3, 16) = Heading7
  ws.Cells(4, 16) = Heading8

  ' Set an initial variable for holding details of Stock
  Dim Stock_End As Double
  Dim Stock_Total As Double
  Dim Stock_Beg As Double
  Dim Stock_YTD As Double
  Dim Count_rec As Double
  Dim Percent_change As Double
  Stock_Total = 0
  Count_rec = 0

  ' Keep track of the location for each Stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all Stock movement
  
  For i = 2 To lastrow

    ' Check if we are still within the same Stock Ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Stock name
      Stock_Name = ws.Cells(i, 1).Value

      ' Add to the Stock Total
      
      Stock_End = ws.Cells(i, 6).Value
      Stock_Beg = ws.Cells(i - Count_rec, 3).Value
      Stock_YTD = Stock_End - Stock_Beg
      Percent_change = (Stock_YTD / Stock_Beg) * 100
      Percent_change = Application.WorksheetFunction.RoundUp(Percent_change, 2)
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value

      ' Print the Stock name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Stock_Name

           
      ' Print the Stock Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Stock_YTD
      
      If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
      
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      
      Else
      
      ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      
      End If
      
       ' Print the Stock Amount to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_change & "%"
      
       ' Print the Stock Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total
     
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Total
      Stock_End = 0
      Stock_Beg = 0
      Count_rec = 0
      Stock_Total = 0

    ' If the cell immediately following a row is the same Stock...
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
      Count_rec = Count_rec + 1
      
      

    End If

  Next i
  
  lastrow_i = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
  MyMax = WorksheetFunction.Max(Range(Cells(2, 11), Cells(lastrow_i, 11)))
  MyRow = WorksheetFunction.Match(MyMax, Range(Cells(2, 11), Cells(lastrow_i, 11)), 0) + 2 - 1
  
  ws.Cells(2, 17) = ws.Cells(MyRow, 9)
  ws.Cells(2, 18) = ws.Cells(MyRow, 11) * 100 & "%"
  
    
  MyMin = WorksheetFunction.Min(Range(Cells(2, 11), Cells(lastrow_i, 11)))
  MyRow = WorksheetFunction.Match(MyMin, Range(Cells(2, 11), Cells(lastrow_i, 11)), 0) + 2 - 1
  
  ws.Cells(3, 17) = ws.Cells(MyRow, 9)
  ws.Cells(3, 18) = ws.Cells(MyRow, 11) * 100 & "%"
  
  lastrow_ii = ws.Cells(Rows.Count, 12).End(xlUp).Row
  
  MyMax_i = WorksheetFunction.Max(Range(Cells(2, 12), Cells(lastrow_ii, 12)))
  MyRow_i = WorksheetFunction.Match(MyMax_i, Range(Cells(2, 12), Cells(lastrow_ii, 12)), 0) + 2 - 1
  
  ws.Cells(4, 17) = ws.Cells(MyRow_i, 9)
  ws.Cells(4, 18) = ws.Cells(MyRow_i, 12)
  
    
  
  
     Next ws

End Sub
