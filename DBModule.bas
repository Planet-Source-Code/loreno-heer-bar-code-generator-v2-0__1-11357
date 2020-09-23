Attribute VB_Name = "DBModule"
Public Type DBcode
   FullCode             As String * 13
   Country              As String * 15
   Manufacturer         As String * 20
   Product              As String * 20
   PicPath              As String * 512
End Type
Function GetPosition(text As String) As Long
GetPosition = CLng(Mid$(text, 2, (InStr(1, text, ">") - 2)))
End Function

