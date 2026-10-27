' tools\Run-Hidden.vbs
' Starts the widget without a PowerShell console window.
' Double-click this file to run the widget.
' (ASCII-only comments on purpose: cscript reads .vbs as ANSI.)
Option Explicit
Dim fso, sh, base, ps1, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
ps1  = base & "\ProdCalendar.ps1"
cmd  = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """"
sh.Run cmd, 0, False
