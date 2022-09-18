Set objexcel = CreateObject("excel.application")
objexcel.Application.Run "'Filepath\Filename.xlsm'!module1.test"
objexcel.DisplayAlerts = False
objexcel.Application.Quit
Set objexcel = Nothing
MsgBox ("Task Completed...!")