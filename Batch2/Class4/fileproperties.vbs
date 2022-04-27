Dim objfso,myfile

Set objfso=CreateObject("Scripting.FileSystemObject")

Set myfile=objfso.GetFile("D:\Saritha\Saritha\SFJBS\Batch2\Class4\file.txt")

Msgbox myfile.name

Msgbox myfile.DateCreated

Msgbox myfile.DateLastAccessed
Msgbox myfile.DateLastModified
Msgbox myfile.Drive
Msgbox myfile.ParentFolder
Msgbox myfile.path
Msgbox myfile.size
Msgbox myfile.Type