Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

vbsPath = WScript.ScriptFullName
vbsDir = objFSO.GetParentFolderName(vbsPath)

ps1Path = objFSO.BuildPath(vbsDir, "unsplash4lockscreen.ps1")
command = "powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy ByPass -File """ & ps1Path & """"

objShell.Run command, 0
