VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form options 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Options"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog CD 
      Left            =   2280
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Bitmap"
      Filter          =   "All files(*.*) | *.*"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Set Background Picture"
      Height          =   735
      Left            =   2040
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create shortcut on the desktop"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove all registry entries and exit the program."
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore defaults"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "1755"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Extras:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      X1              =   0
      X2              =   3600
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zentimeter"
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   1095
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zoll"
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   735
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Twips"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   370
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard form width for EAN codes:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.text = 1755
End Sub

Private Sub Command2_Click()
TheX = MsgBox("sure?", vbYesNo, "Bar Code Generator")
If TheX = vbYes Then
    DeleteSetting "Bar Code Generator"
    End
End If
End Sub

Private Sub Command3_Click()
ShortCut
End Sub

Private Sub Command4_Click()
On Error Resume Next
CD.ShowOpen
TheX = CD.FileName
If TheX <> "" Then
    Mainfrm.Picture = LoadPicture(TheX)
    SaveSetting "Bar Code Generator", "Settings", "BGPicture", TheX
Else
    Mainfrm.Picture = Nothing
    SaveSetting "Bar Code Generator", "Settings", "BGPicture", TheX
End If
End Sub

Private Sub Form_Load()
Text1.text = GetSetting("Bar Code Generator", "Settings", "Form Width", 1755)
Text2.text = Text1.text / 1440
Text3.text = Text1.text / 567
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "Bar Code Generator", "Settings", "Form Width", Text1.text
End Sub

Private Sub Text1_Change()
Text2.text = Text1.text / 1440
Text3.text = Text1.text / 567
End Sub

Private Sub Text2_Change()
Text1.text = Text2.text * 1440
Text3.text = Text1.text / 567
End Sub

Private Sub Text3_Change()
Text1.text = Text3.text * 567
Text2.text = Text1.text / 1440
End Sub
