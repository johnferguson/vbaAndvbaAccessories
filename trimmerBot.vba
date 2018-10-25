Sub trimmerBot()
    For Each cell In Selection
        If cell.Value <> "" Then
            cell.Value = Trim(cell.Value)
        End If
    Next
End Sub

'Depending on how your report is run you may have some trailing spaces on strings
'This macro will trim every string in the selection
'The if statment speeds it up, but can still take a while if you choose too large of an area
