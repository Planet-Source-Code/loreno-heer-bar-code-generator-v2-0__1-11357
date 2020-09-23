VERSION 5.00
Begin VB.Form menues 
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4215
   Icon            =   "menues.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Begin VB.Menu Tile 
         Caption         =   "Tile"
      End
      Begin VB.Menu Cascade 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu C 
      Caption         =   "C"
      Begin VB.Menu AddToDatabase 
         Caption         =   "Add to DB"
      End
      Begin VB.Menu ChangeTitle 
         Caption         =   "Change Title"
      End
      Begin VB.Menu Resize 
         Caption         =   "Resize"
      End
      Begin VB.Menu plus 
         Caption         =   "+"
      End
      Begin VB.Menu minus 
         Caption         =   "- "
      End
   End
End
Attribute VB_Name = "menues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Position As Long
Public LastRecord As Long
Private Sub AddToDatabase_Click()
Dim FileNum As Integer, RecLength As Long, CodeDB As DBcode
RecLength = LenB(CodeDB)
FileNum = FreeFile
Open toolfrm.Text2.text For Random As FileNum Len = RecLength
Position = 1
Get FileNum, 1, CodeDB
If IsNumeric(CLng(Trim$(CodeDB.FullCode))) = True Then
    LastRecord = CLng(Trim$(CodeDB.FullCode))
Else
    LastRecord = 1
End If
Do
TheX = InputBox("Country:", "Enter Country")
Loop While TheX = ""
CodeDB.Country = TheX
TheX = ""
Do
TheX = InputBox("Manufacturer:", "Enter Manufacturer")
Loop While TheX = ""
CodeDB.Manufacturer = TheX
TheX = ""
Do
TheX = InputBox("Product:", "Enter Product name")
Loop While TheX = ""
CodeDB.Product = TheX
TheX = ""
CodeDB.FullCode = Mainfrm.ActiveForm.Label1.Caption & Mainfrm.ActiveForm.Label2.Caption & Mainfrm.ActiveForm.Label3.Caption
CodeDB.PicPath = ""
Put FileNum, LastRecord + 1, CodeDB
LastRecord = LastRecord + 1
CodeDB.FullCode = LastRecord
Put FileNum, 1, CodeDB
Close FileNum
End Sub

Private Sub ChangeTitle_Click()
Title = InputBox("Set a new title:", "New Title", Mainfrm.ActiveForm.Caption)
If Title <> "" Then
    Mainfrm.ActiveForm.Caption = Title
End If
End Sub

Private Sub minus_Click()
If Mainfrm.ActiveForm.isEAN = True Then
    Mainfrm.Text1.text = Subt(Mainfrm.Text1.text)
    Mainfrm.ActiveForm.Width = (options.Text1.text)
    Mainfrm.ActiveForm.Cls
    Mainfrm.ActiveForm.ScaleMode = 3
    Mainfrm.ActiveForm.Label1.Visible = True
    PaintCode Mainfrm.ActiveForm, Mid$(Mainfrm.Text1.text, 1, 1), Mid$(Mainfrm.Text1.text, 2, 6), Mid$(Mainfrm.Text1.text, 8, 6)
    Mainfrm.ActiveForm.Label1.Caption = Mid$(Mainfrm.Text1.text, 1, 1)
    Mainfrm.ActiveForm.Label2.Caption = Mid$(Mainfrm.Text1.text, 2, 6)
    Mainfrm.ActiveForm.Label3.Caption = Mid$(Mainfrm.Text1.text, 8, 6)
    Mainfrm.ActiveForm.isEAN = True
    Mainfrm.ActiveForm.Refresh
End If
End Sub

Private Sub plus_Click()
If Mainfrm.ActiveForm.isEAN = True Then
Mainfrm.Text1.text = Add(Mainfrm.Text1.text)
    Mainfrm.ActiveForm.Width = (options.Text1.text)
    Mainfrm.ActiveForm.Cls
    Mainfrm.ActiveForm.ScaleMode = 3
    Mainfrm.ActiveForm.Label1.Visible = True
    PaintCode Mainfrm.ActiveForm, Mid$(Mainfrm.Text1.text, 1, 1), Mid$(Mainfrm.Text1.text, 2, 6), Mid$(Mainfrm.Text1.text, 8, 6)
    Mainfrm.ActiveForm.Label1.Caption = Mid$(Mainfrm.Text1.text, 1, 1)
    Mainfrm.ActiveForm.Label2.Caption = Mid$(Mainfrm.Text1.text, 2, 6)
    Mainfrm.ActiveForm.Label3.Caption = Mid$(Mainfrm.Text1.text, 8, 6)
    Mainfrm.ActiveForm.isEAN = True
    Mainfrm.ActiveForm.Refresh
End If
End Sub

Private Sub Resize_Click()
FWidth = InputBox("Set a new width:", "New Width", Mainfrm.ActiveForm.Width)
If FWidth <> "" Then
    Mainfrm.ActiveForm.Width = FWidth
End If
End Sub

Private Sub Tile_Click()
Mainfrm.Arrange (1)
Unload Me
End Sub
Private Sub Cascade_Click()
Mainfrm.Arrange (0)
Unload Me
End Sub
