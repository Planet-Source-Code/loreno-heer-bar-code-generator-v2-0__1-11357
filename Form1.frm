VERSION 5.00
Begin VB.Form BarCodefrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "BCode -1-"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   90
   End
End
Attribute VB_Name = "BarCodefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isEAN As Boolean
Private Sub Form_DblClick()
Mainfrm.ActiveForm.Width = InputBox("Set a new width", "New Width", Mainfrm.ActiveForm.Width)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Mainfrm.ActiveForm.isEAN = True Then
If KeyCode = vbKeyAdd Then
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
ElseIf KeyCode = vbKeySubtract Then
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
ElseIf KeyCode = vbKeyReturn Then
    newBarCode
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
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu menues.C, , X, Y
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FState(Me.Tag).Deleted = True
End Sub
