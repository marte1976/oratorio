Option Explicit

Dim shell, fso, startupFolder, shortcutPath

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

startupFolder = shell.SpecialFolders("Startup")
shortcutPath = startupFolder & "\Oratorio Carlo Acutis.lnk"

If fso.FileExists(shortcutPath) Then
    fso.DeleteFile shortcutPath, True
    MsgBox "Avvio automatico disattivato.", vbInformation, "Oratorio Carlo Acutis"
Else
    MsgBox "Nessun avvio automatico attivo da rimuovere.", vbInformation, "Oratorio Carlo Acutis"
End If
