Sub chooseFolder()
    Dim chosenFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        If .Show = -1 Then
            chosenFolder = .SelectedItems(1)
        End If
    End With
    
    Debug.Print "Chosen Folder: " & chosenFolder
    
    If chosenFolder <> "" Then
        Dim currentFile As String

        currentFile = Dir(chosenFolder & "\*")
        Do While Len(currentFile) > 0
            Call copyPNG(chosenFolder & "\" & currentFile)
            
            currentFile = Dir
        Loop
    End If
End Sub

'Source:
'Vba Select Folder with Msofiledialogfolderpicker
'Ryan Wells - https://wellsr.com/vba/2016/excel/vba-select-folder-with-msoFileDialogFolderPicker/
