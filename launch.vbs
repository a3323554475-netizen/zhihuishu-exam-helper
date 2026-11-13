' Launch Start-ExamHelper.ps1 without a console window.
' Called by 启动.cmd via wscript (wscript itself is a GUI host, so no console).
' The script force-shows its own windows, so launching with style 0 is safe.
Option Explicit

Dim fso, shell, base, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

base = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & base & "\Start-ExamHelper.ps1"""

' 0 = hidden, False = do not wait for it
shell.Run cmd, 0, False
