' アクセスログ記録用関数
'
Function AccessLog(userName, action)
	Dim objFSO
	Dim currentDate
	Dim logFile

	logFile = "C:\Windows\access.log"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	currentDate = Now

	Set objFile= objFSO.OpenTextFile(logFile, 8, True)
	objFile.WriteLine currentDate & vbtab & action & vbtab & userName
	objFile.Close
End Function