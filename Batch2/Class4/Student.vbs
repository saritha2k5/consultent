
filepath="D:\Saritha\orange.txt"
despath="D:\Saritha\Saritha"
Set objexcel=CreateObject("Excel.Application")
objexcel.visible=True
objexcel.workbooks.open "‪D:\Saritha\Saritha\Login.xls"

Set obj1=objexcel.Activeworkbook.WorkSheets(2)
obj1.Cells(1,1)="bangalore"
obj1.Cells(2,1)="mumbai"
obj1.Cells(3,1)="pune"


set obj1=Nothing
Set objexcel=nothing
