VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Mainfrm 
   BackColor       =   &H8000000C&
   Caption         =   "Bar Code Generator"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Neu"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Drucken"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Kopieren"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zeichnen"
            Object.ToolTipText     =   "Paint"
            ImageKey        =   "Drawing"
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Tools..."
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   30
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "C.System:"
         Top             =   50
         Width           =   780
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         Height          =   210
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Code:"
         Top             =   50
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "MDIForm1.frx":0442
         Left            =   4800
         List            =   "MDIForm1.frx":044C
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         MaxLength       =   13
         TabIndex        =   2
         Text            =   "2501007661990"
         Top             =   20
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4035
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10865
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "15:40"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":045B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":056D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":067F
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0791
            Key             =   "Drawing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    Text1.MaxLength = 13
ElseIf Combo1.ListIndex = 1 Then
    Text1.MaxLength = 0
End If
End Sub

Private Sub Command1_Click()
toolfrm.Show
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu menues.Popup, , X, Y
End If
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
TheX = MsgBox("Do you realy want to exit?", vbYesNo + vbSystemModal + vbQuestion, "Exit?")
If TheX <> vbYes Then
    Cancel = True
Else
    Unload menues
    Unload options
End If
End Sub



Private Sub Text1_Change()
If Len(Text1.text) = 13 And Combo1.ListIndex = 0 Then
    If CheckCode(Text1.text) = False Then
        If repair(Text1.text) <> 0 Then
        StatusBar1.Panels(1).text = "Wrong Checknumber! Maybe you mean: " & repair(Text1.text)
        Else
        StatusBar1.Panels(1).text = "Wrong Code! Only use numbers in EAN-Codes!"
        End If
    Else
        StatusBar1.Panels(1).text = "This Code is Correct"
    End If
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Combo1.ListIndex = 0 Then
            If Len(Text1.text) = 13 Then
                Mainfrm.ActiveForm.Width = (options.Text1.text)
                Mainfrm.ActiveForm.Cls
                Mainfrm.ActiveForm.ScaleMode = 3
                Mainfrm.ActiveForm.Label1.Visible = True
                PaintCode Mainfrm.ActiveForm, Mid$(Text1.text, 1, 1), Mid$(Text1.text, 2, 6), Mid$(Text1.text, 8, 6)
                Mainfrm.ActiveForm.Label1.Caption = Mid$(Text1.text, 1, 1)
                Mainfrm.ActiveForm.Label2.Caption = Mid$(Text1.text, 2, 6)
                Mainfrm.ActiveForm.Label3.Caption = Mid$(Text1.text, 8, 6)
                Mainfrm.ActiveForm.isEAN = True
                Mainfrm.ActiveForm.Refresh
            Else
                MsgBox "Error: The lenght of the code is wrong", vbCritical, "Error"
            End If
            ElseIf Combo1.ListIndex = 1 Then
                Mainfrm.ActiveForm.Cls
                Mainfrm.ActiveForm.ScaleMode = 1
                Mainfrm.ActiveForm.Label1.Visible = False
                Code3of9 Text1.text, Mainfrm.ActiveForm, Mainfrm.ActiveForm.Label1
                Mainfrm.ActiveForm.Label2 = ""
                Mainfrm.ActiveForm.Label3 = ""
                Mainfrm.ActiveForm.isEAN = False
                Mainfrm.ActiveForm.Refresh
            End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Neu"
            newBarCode
            If Len(Text1.text) = 13 Then
                PaintCode Mainfrm.ActiveForm, Mid$(Text1.text, 1, 1), Mid$(Text1.text, 2, 6), Mid$(Text1.text, 8, 6)
                Mainfrm.ActiveForm.Label1.Caption = Mid$(Text1.text, 1, 1)
                Mainfrm.ActiveForm.Label2.Caption = Mid$(Text1.text, 2, 6)
                Mainfrm.ActiveForm.Label3.Caption = Mid$(Text1.text, 8, 6)
                Mainfrm.ActiveForm.isEAN = True
                Mainfrm.ActiveForm.Refresh
            Else
                MsgBox "Error: The lenght of the code is wrong", vbCritical, "Error"
            End If
        Case "Drucken"
            Printfrm.Show
        Case "Kopieren"
            'nichts
        Case "Zeichnen"
            If Combo1.ListIndex = 0 Then
            If Len(Text1.text) = 13 Then
                Mainfrm.ActiveForm.Width = (options.Text1.text)
                Mainfrm.ActiveForm.Cls
                Mainfrm.ActiveForm.ScaleMode = 3
                Mainfrm.ActiveForm.Label1.Visible = True
                PaintCode Mainfrm.ActiveForm, Mid$(Text1.text, 1, 1), Mid$(Text1.text, 2, 6), Mid$(Text1.text, 8, 6)
                Mainfrm.ActiveForm.Label1.Caption = Mid$(Text1.text, 1, 1)
                Mainfrm.ActiveForm.Label2.Caption = Mid$(Text1.text, 2, 6)
                Mainfrm.ActiveForm.Label3.Caption = Mid$(Text1.text, 8, 6)
                Mainfrm.ActiveForm.isEAN = True
                Mainfrm.ActiveForm.Refresh
            Else
                MsgBox "Error: The lenght of the code is wrong", vbCritical, "Error"
            End If
            ElseIf Combo1.ListIndex = 1 Then
                Mainfrm.ActiveForm.Cls
                Mainfrm.ActiveForm.ScaleMode = 1
                Mainfrm.ActiveForm.Label1.Visible = False
                Code3of9 Text1.text, Mainfrm.ActiveForm, Mainfrm.ActiveForm.Label1
                Mainfrm.ActiveForm.Label2 = ""
                Mainfrm.ActiveForm.Label3 = ""
                Mainfrm.ActiveForm.isEAN = False
                Mainfrm.ActiveForm.Refresh
            End If
    End Select
End Sub
Private Sub MDIForm_Load()
ReDim bcode(1)
ReDim FState(1)
bcode(1).Tag = 1
FState(1).Dirty = False
Combo1.ListIndex = 0
PaintCode Mainfrm.ActiveForm, Mid$(Text1.text, 1, 1), Mid$(Text1.text, 2, 6), Mid$(Text1.text, 8, 6)
Mainfrm.ActiveForm.Label1.Caption = Mid$(Text1.text, 1, 1)
Mainfrm.ActiveForm.Label2.Caption = Mid$(Text1.text, 2, 6)
Mainfrm.ActiveForm.Label3.Caption = Mid$(Text1.text, 8, 6)
Mainfrm.ActiveForm.isEAN = True
Mainfrm.ActiveForm.Refresh
TheX = GetSetting("Bar Code Generator", "Settings", "BGPicture", "")
If TheX <> "" Then
    Me.Picture = LoadPicture(TheX)
End If
Text2.BackColor = menues.BackColor
Text3.BackColor = menues.BackColor
End Sub
