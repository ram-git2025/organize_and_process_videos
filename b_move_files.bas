Attribute VB_Name = "b_move_files"
Dim fso As Object
Sub move_videos_to_respective_folder()
Set fso = CreateObject("Scripting.FileSystemObject")

source_folder_path = Sheet1.Cells(2, 3)
target_folder_path = Sheet1.Cells(4, 3)

Set obj_source_folder_path = fso.GetFolder(source_folder_path)

If fso.FolderExists(obj_source_folder_path) Then
 'reading files
 For Each f_name In obj_source_folder_path.Files
  recursive_files f_name, target_folder_path
 Next
Else
 Debug.Print "Source folder doesn't exist."
End If
End Sub

Sub recursive_files(ByVal f_name As String, ByVal destinationFolder As String)
 Set obj_destinationFolder = fso.GetFolder(destinationFolder)
 
If fso.FolderExists(obj_destinationFolder) Then
 For Each f_des_name In obj_destinationFolder.subfolders
  recursive_file_name_check f_name, f_des_name
 Next
 recursive_file_name_check f_name, obj_destinationFolder
Else
 Debug.Print "Destination folder doesn't exist."
End If
End Sub

Sub recursive_file_name_check(ByVal file_validate As Variant, ByVal file_path As Variant)
 For Each f_name In file_path.Files
  If LCase(fso.GetBaseName(file_validate)) = LCase(fso.GetBaseName(f_name)) Then
   SetAttr file_validate, vbNormal
   fso.MoveFile file_validate, fso.GetParentFolderName(f_name) & "\"
   fso.deletefile f_name
   Exit Sub
  End If
 Next
End Sub
