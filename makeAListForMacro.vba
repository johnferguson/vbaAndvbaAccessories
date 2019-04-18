'this makes a list of lists for a mainframe macro
Sub list()
    'l is your target worksheet
    Set lWS = ThisWorkbook.Sheets("Sheet2")
    Set ws = ActiveSheet
    l = 0 'increments the list of lists
    y = 3 'starting y value on the data worksheet
    
    currStr = ""
    Do While ws.Cells(y, 1).Value <> "" 'loop through every row until you get to an empty row
        'if there is already anything in the string, if it is empty don't add one
        If Len(currStr) > 1 Then
            currStr = currStr + " "
        End If
        'concat everything you need with spaces inbetween
        currStr = currStr + ws.Cells(y, 4).Value & " " & ws.Cells(y, 5).Value & " " & ws.Cells(y, 11).Value
        
        'the mainframe program will only allow 3000 characters per row, start another sub array if we get past 2500 characters
        If Len(currStr) > 2500 Then
            lWS.Cells(l + 1, 1).Value = "pnList[" & l & "] = '" & currStr & "'"
            currStr = ""
            l = l + 1
        End If
        
        y = y + 1
    Loop
    
End Sub
