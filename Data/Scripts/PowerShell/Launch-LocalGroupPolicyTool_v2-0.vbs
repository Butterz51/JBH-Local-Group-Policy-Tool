Set shellApp = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "\\PowerShell\LocalGroupPolicyTool_v2-0.ps1")

If Not fso.FileExists(scriptPath) Then
    MsgBox "Script not found:" & vbCrLf & scriptPath, vbCritical, "JBH Services"
    WScript.Quit 1
End If

shellApp.ShellExecute "powershell.exe", _
    "-NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """", _
    "", _
    "runas", _
    1
