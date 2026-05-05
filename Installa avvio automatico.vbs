Option Explicit

Dim shell, fso, projectRoot, startupFolder, shortcutPath, targetPath, shortcut

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

projectRoot = fso.GetParentFolderName(WScript.ScriptFullName)
startupFolder = shell.SpecialFolders("Startup")
shortcutPath = startupFolder & "\Oratorio Carlo Acutis.lnk"
targetPath = projectRoot & "\Apri Oratorio Carlo Acutis.vbs"

Set shortcut = shell.CreateShortcut(shortcutPath)
shortcut.TargetPath = targetPath
shortcut.WorkingDirectory = projectRoot
shortcut.Description = "Avvia automaticamente il gestionale Oratorio Carlo Acutis"
shortcut.IconLocation = "%SystemRoot%\System32\SHELL32.dll,220"
shortcut.Save

MsgBox "Avvio automatico attivato. Al prossimo accesso a Windows il gestionale si aprira da solo.", vbInformation, "Oratorio Carlo Acutis"
