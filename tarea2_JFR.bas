Attribute VB_Name = "Module1"
Sub tkr()

Dim days, lastrow, yearchange, perchange, vol, Summary_Table_Row, openp, closep, ginc, gdec, valvol As Double
 
 
Sheets("2016").Select
  
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total"
   
 
'find last row
lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
Summary_Table_Row = 2
'forloop to review ticker
 
For i = 2 To lastrow
 
' Check ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ' Set the ticker
     ticker = Cells(i, 1).Value
 
    ' find stock volume and days observed
      vol = vol + Cells(i, 7).Value
      days = days + 1
         
   'find year change and percentage change
     openp = Cells(i - days + 1, 3).Value
     closep = Cells(i, 6).Value
     yearchange = closep - openp
     If openp = 0 Then
        perchange = 0
        Else
        perchange = closep / openp - 1
      End If
                 
      ' Print Ticker name in the summary table
      Range("I" & Summary_Table_Row).Value = ticker
 
      ' Print the yearchange to the Summary Table
      Range("J" & Summary_Table_Row).Value = yearchange
      If yearchange < 0 Then
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        ElseIf yearchange > 0 Then
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
           
  
      ' Print the percentagechange to the Summary Table
      Range("K" & Summary_Table_Row).Value = perchange
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
 
      ' Print the ticker volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = vol
 
      'Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
      ' Reset the ticker Total and volume totals
      vol = 0
      days = 0
   
    ' If the cell immediately following a row is the same ticker...
    Else
 
      ' Add to the tikcer Total
      vol = vol + Cells(i, 7).Value
      days = days + 1
 
' Print Open Price in the summary table
 
 
    End If
 
  Next i
 
Range("Q2:Q3").NumberFormat = "0.00%"
 
'find min and max percentage changes and greates total
ginc = WorksheetFunction.Max(Range("K:K").Value)
gdec = WorksheetFunction.Min(Range("K:K").Value)
valvol = WorksheetFunction.Max(Range("L:L").Value)
Range("q2").Value = ginc
Range("q3").Value = gdec
Range("q4").Value = valvol
 
'loop throuigh Tkr
 
For i = 2 To lastrow
 
        If Cells(i, 11).Value = ginc Then
            Range("p2").Value = Cells(i, 9).Value
        End If
       
        If Cells(i, 11).Value = gdec Then
            Range("p3").Value = Cells(i, 9).Value
        End If
       
        If Cells(i, 12).Value = valvol Then
            Range("p4").Value = Cells(i, 9).Value
        End If
       
 
Next i
 
 
Sheets("2015").Select
 
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total"
 
lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
Summary_Table_Row = 2
For i = 2 To lastrow
   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
     ticker = Cells(i, 1).Value
     vol = vol + Cells(i, 7).Value
     days = days + 1
     openp = Cells(i - days + 1, 3).Value
     closep = Cells(i, 6).Value
     yearchange = closep - openp
     If openp = 0 Then
        perchange = 0
        Else
        perchange = closep / openp - 1
      End If
                 
      Range("I" & Summary_Table_Row).Value = ticker
      Range("J" & Summary_Table_Row).Value = yearchange
        If yearchange < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        ElseIf yearchange > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
      Range("K" & Summary_Table_Row).Value = perchange
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("L" & Summary_Table_Row).Value = vol
     
      Summary_Table_Row = Summary_Table_Row + 1
     
      vol = 0
      days = 0
   
    Else
     
      vol = vol + Cells(i, 7).Value
      days = days + 1
 
   End If
 
  Next i
 
Range("Q2:Q3").NumberFormat = "0.00%"
 
'find min and max percentage changes and greates total
ginc = WorksheetFunction.Max(Range("K:K").Value)
gdec = WorksheetFunction.Min(Range("K:K").Value)
valvol = WorksheetFunction.Max(Range("L:L").Value)
Range("q2").Value = ginc
Range("q3").Value = gdec
Range("q4").Value = valvol
 
'loop throuigh Tkr
 
For i = 2 To lastrow
 
        If Cells(i, 11).Value = ginc Then
            Range("p2").Value = Cells(i, 9).Value
        End If
       
        If Cells(i, 11).Value = gdec Then
            Range("p3").Value = Cells(i, 9).Value
        End If
       
        If Cells(i, 12).Value = valvol Then
            Range("p4").Value = Cells(i, 9).Value
        End If
      
 
Next i
 
Sheets("2014").Select
 
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percentage Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total"
 
lastrow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
Summary_Table_Row = 2
For i = 2 To lastrow
   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
     ticker = Cells(i, 1).Value
     vol = vol + Cells(i, 7).Value
     days = days + 1
     openp = Cells(i - days + 1, 3).Value
     closep = Cells(i, 6).Value
     yearchange = closep - openp
     If openp = 0 Then
        perchange = 0
        Else
        perchange = closep / openp - 1
      End If
            
      Range("I" & Summary_Table_Row).Value = ticker
      Range("J" & Summary_Table_Row).Value = yearchange
        If yearchange < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        ElseIf yearchange > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        End If
      Range("K" & Summary_Table_Row).Value = perchange
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("L" & Summary_Table_Row).Value = vol
     
      Summary_Table_Row = Summary_Table_Row + 1
     
      vol = 0
      days = 0
   
    Else
     
      vol = vol + Cells(i, 7).Value
      days = days + 1
 
    End If
 
  Next i
 
Range("Q2:Q3").NumberFormat = "0.00%"
 
'find min and max percentage changes and greates total
ginc = WorksheetFunction.Max(Range("K:K").Value)
gdec = WorksheetFunction.Min(Range("K:K").Value)
valvol = WorksheetFunction.Max(Range("L:L").Value)
Range("q2").Value = ginc
Range("q3").Value = gdec
Range("q4").Value = valvol
 
'loop throuigh Tkr
 
For i = 2 To lastrow
 
        If Cells(i, 11).Value = ginc Then
            Range("p2").Value = Cells(i, 9).Value
        End If
       
        If Cells(i, 11).Value = gdec Then
            Range("p3").Value = Cells(i, 9).Value
        End If
       
        If Cells(i, 12).Value = valvol Then
            Range("p4").Value = Cells(i, 9).Value
        End If
       
 
Next i

End Sub
