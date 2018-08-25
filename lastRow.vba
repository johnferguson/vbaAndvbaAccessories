Sub lastRow()
  Dim wb as workbook
  Dim ws as worksheet
  Dim y as long
  
  Set wb = worbookPath & "/" & workbookName
  Set ws = wb.worksheets(worksheetName)
  y = 2 'change this to the row after 
  
  Do While ws.Cells(y, 1) <> ""
    y = y + 1
  Loop
  
  debug.print y ' this will be the first empty row

End Sub


'sometimes text files have multiple pages and the header/footer takes up space
'this function can keep going after a blank
Sub lastRowWithGaps()
  Dim wb as workbook
  Dim ws as worksheet
  Dim y as long
  
  Set wb = worbookPath & "/" & workbookName
  Set ws = wb.worksheets(worksheetName)
  y = 2 'change this to the row after
  acceptableGap
  
  Do While ws.Cells(y, 1) <> ""
    y = y + 1
  Loop

End Sub
