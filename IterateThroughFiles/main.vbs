call main

sub main
	Set oFSO = CreateObject("Scripting.FileSystemObject")
		
	' Dim oShellApp : Set oShellApp = CreateObject("Shell.Application")
	' Dim oWScriptShell : Set oWScriptShell = CreateObject("WScript.Shell")
		
	scriptDirPath = oFSO.GetParentFolderName(WScript.ScriptFullName) 'no backslash at the end	
	'msgbox scriptdir

	inDirPath = "c:\test"

	logFilePath = scriptDirPath & "\run.log"
	'ForAppending - 8, ForWriting - 2
	Set logFile = oFSO.OpenTextFile(logFilePath,8,true)
	'logFile.WriteLine("Hello")
	
	For Each oFile In oFSO.GetFolder(inDirPath).Files				
		InFileName =  oFile.Name
		' msgbox (InFileName)		
		InFilePath =  oFSO.GetAbsolutePathName(oFile)
		If UCase(oFSO.GetExtensionName(InFileName)) = "TXT" Then
			logFile.WriteLine(InFilePath)
			'msgbox (InFileName)
		End if		
	Next 

	logFile.Close
	Set logFile = Nothing

	msgbox "Done"
end sub
