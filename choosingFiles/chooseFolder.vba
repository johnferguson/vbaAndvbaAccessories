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
            'place what you want to do here
            Debug.Print "Current File: " & chosenFolder & "\" & currentFile
            
            currentFile = Dir
        Loop
    End If
End Sub

'Choose a folder and loop through each file
'Change the *placeholder in line 15 to only look at specific file types or names

'Sources:
'Vba Select Folder with Msofiledialogfolderpicker
'Ryan Wells - https://wellsr.com/vba/2016/excel/vba-select-folder-with-msoFileDialogFolderPicker/
'Loop Through Files in a Folder Using Vba?
'https://stackoverflow.com/questions/10380312/loop-through-files-in-a-folder-using-vba
