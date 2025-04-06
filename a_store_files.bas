Attribute VB_Name = "a_store_files"
Dim fso As Object
Dim dict As Object
Sub store_videos_file_in_folder()
Set fso = CreateObject("Scripting.FileSystemObject")
Set dict = CreateObject("Scripting.Dictionary")

store_folder_path = Sheet1.Cells(3, 3)
target_folder_path = Sheet1.Cells(4, 3)

Set obj_target_folder_path = fso.GetFolder(target_folder_path)

With dict
 .Add "mp4", "" 'Most widely used; compatible with almost all devices and platforms.
 .Add "mkv", "" 'Matroska (Supports multiple audio/subtitle tracks; often used for HD video.)
 .Add "avi", "" 'Audio Video Interleave (Older format by Microsoft; larger file sizes.)
 .Add "mov", "" 'Developed by Apple; great quality, large files.
 .Add "wmv", "" 'Windows Media Video
 .Add "flv", "" 'Flash Video (Used for web video streaming (e.g., old YouTube).)
 .Add "webm", "" 'Open format; great for web use and streaming.
 .Add "mpeg ", "" 'Common for DVDs and early web videos.
 .Add "mpg", "" 'Common for DVDs and early web videos.
 .Add "m4v", "" 'MPEG-4 Video (Similar to .mp4; often used for Apple devices.)
 .Add "vob", "" 'DVD Video Object
End With

If fso.FolderExists(obj_target_folder_path) Then
 'reading subfolder
 For Each f_file In obj_target_folder_path.subfolders
  recursive_folder f_file, store_folder_path
 Next
 'reading currentfolder
 recursive_folder target_folder_path, store_folder_path
Else
 Debug.Print "Source folder doesn't exist."
End If
End Sub

Sub recursive_folder(ByVal f_path As String, ByVal destinationFile As String)
 Set obj_f_path = fso.GetFolder(f_path)
 
If fso.FolderExists(obj_f_path) Then
 For Each f_name In obj_f_path.Files
  If dict.exists(fso.GetExtensionName(f_name)) Then
   moving_filename = fso.GetBaseName(f_name)
   fso.MoveFile f_name, destinationFile
   Set textFile = fso.createtextfile(obj_f_path & "\" & moving_filename & ".txt")
  End If
 Next
Else
 Debug.Print "Destination folder doesn't exist."
End If
End Sub
