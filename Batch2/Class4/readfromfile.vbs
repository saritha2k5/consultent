
filepath="D:\Saritha\orange.txt"
despath="D:\Saritha\Saritha"
Set objfso=CreateObject("Scripting.FileSystemObject")
Set myfile =objfso.Opentextfile("D:\Saritha\Saritha\orange.txt",1,True)

Do while myfile.AtEndofStream <> true
data=myfile.Readline()

msgbox data

Loop

