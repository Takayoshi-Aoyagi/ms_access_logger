Dim userName
Dim objFs
Dim objFile
Dim action
Dim ret

' Include logging.vbs
Set objFs = CreateObject("Scripting.FileSystemObject")
Set objFile = objFs.OpenTextFile("C:\Windows\logging.vbs")
ExecuteGlobal objFile.ReadAll()
objFile.Close
Set objFile = Nothing

userName = ""
action = "LOGOUT"
ret = AccessLog(userName, action)