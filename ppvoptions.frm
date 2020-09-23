VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ppvoptions 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Print Preview Options"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton setfg 
      Caption         =   "Set"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton setbg 
      Caption         =   "Set"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore defaults"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox showtitle 
      Caption         =   "Show title under each code"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox rowtext 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "17"
      Top             =   120
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      Height          =   255
      Left            =   120
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "FG Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      Height          =   255
      Left            =   120
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "BG Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Rows:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "ppvoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Shape1.BackColor = vbWhite
Shape2.BackColor = vbBlack
rowtext.text = 17
showtitle.Value = 0
End Sub

Private Sub Form_Load()
Shape1.BackColor = GetSetting("Bar Code Generator", "PPV Options", "BackColor", vbWhite)
Shape2.BackColor = GetSetting("Bar Code Generator", "PPV Options", "ForeColor", vbBlack)
rowtext.text = GetSetting("Bar Code Generator", "PPV Options", "Rows", 17)
showtitle.Value = GetSetting("Bar Code Generator", "PPV Options", "ShowTitle", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsNumeric(rowtext.text) = True Then
    SaveSetting "Bar Code Generator", "PPV Options", "Rows", rowtext.text
End If
SaveSetting "Bar Code Generator", "PPV Options", "BackColor", Shape1.BackColor
SaveSetting "Bar Code Generator", "PPV Options", "ForeColor", Shape2.BackColor
End Sub

Private Sub rowtext_Change()
If IsNumeric(rowtext.text) = True Then
    SaveSetting "Bar Code Generator", "PPV Options", "Rows", rowtext.text
End If
End Sub

Private Sub setbg_Click()
CD1.Color = Shape1.BackColor
CD1.ShowColor
Shape1.BackColor = CD1.Color
SaveSetting "Bar Code Generator", "PPV Options", "BackColor", Shape1.BackColor
Changed
End Sub

Private Sub setfg_Click()
CD1.Color = Shape2.BackColor
CD1.ShowColor
Shape2.BackColor = CD1.Color
SaveSetting "Bar Code Generator", "PPV Options", "ForeColor", Shape2.BackColor
Changed
End Sub
Sub Changed()
printpreview.BackColor = ppvoptions.Shape1.BackColor
printpreview.ForeColor = ppvoptions.Shape2.BackColor
End Sub

Private Sub showtitle_Click()
SaveSetting "Bar Code Generator", "PPV Options", "ShowTitle", showtitle.Value
End Sub
