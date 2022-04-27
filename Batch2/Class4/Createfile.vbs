
filepath="D:\Saritha\orange.txt"
Set objfso=CreateObject("Scripting.FileSystemObject")
Set myfile=objfso.CreateTextFile(filepath,True)
if objfso.FileExists(filepath) then
msgbox "file created"
end if
