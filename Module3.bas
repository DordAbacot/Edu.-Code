Attribute VB_Name = "Module3"
Sub Create_Folders()
'Create the 5 needed folders

Dim fldr As FileDialog 'Dialog to retrieve path from user
Dim strTargetPath As String 'String with Target Path where folders should be created
Dim varFolderName As Variant 'List of folders to create

varFolderName = Array("1. Facility", "2. Instructor", "3. Pre-Course", "4. Post-Course", "5. Facility Comp") 'Folders to create
    
    'Show dialog box to get path from user
'    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
'        With fldr
'            .AllowMultiSelect = False
'            .Title = "Select target folder to create subfolders"
'        End With
    
    'If a selection was made, store in strTargetPath
'    If fldr.Show = -1 Then
'        strTargetPath = fldr.SelectedItems(1)
        
    'If no selection was made, exit procedure
'    Else
'        MsgBox "No folder has been selected"
'        Exit Sub
'    End If
    
    strTargetPath = Range("C14")


         
    'Loop to create folders
    For Each Item In varFolderName
        
        'If folder does not exist, then create; otherwiese, skip and create next folder
        If Dir(strTargetPath & "\" & Item, vbDirectory) = "" _
            Then
                MkDir strTargetPath & "\" & Item
                'MsgBox "Folder " & Item & " has been created"
            'Else
                'MsgBox "Folder " & Item & " already exists"
        End If
    Next
    
End Sub
