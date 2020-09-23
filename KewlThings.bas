Attribute VB_Name = "KewlThings"
Sub ShortCut()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

Dim MyShortcut, MyDesktop, DesktopPath

' Read desktop path using WshSpecialFolders object
DesktopPath = WSHShell.specialfolders("Desktop")

' Create a shortcut object on the desktop
Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\Bar Code Generator.lnk")

' Set shortcut object properties and save it
MyShortcut.TargetPath = WSHShell.ExpandEnvironmentStrings(App.Path & "\" & App.EXEName)
MyShortcut.WorkingDirectory = WSHShell.ExpandEnvironmentStrings(App.Path)
MyShortcut.WindowStyle = 4
MyShortcut.IconLocation = WSHShell.ExpandEnvironmentStrings((App.Path & "\" & App.EXEName & ".exe") & ", 0")
MyShortcut.Save
MsgBox "Finish", vbSystemModal
End Sub
