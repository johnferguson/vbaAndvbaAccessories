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
        ' *********************
        ' put your code in here
        ' *********************
    End If
End Sub

'Source:
'Vba Select Folder with Msofiledialogfolderpicker
'Ryan Wells - https://wellsr.com/vba/2016/excel/vba-select-folder-with-msoFileDialogFolderPicker/
