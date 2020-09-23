VERSION 5.00
Begin VB.Form printpreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Preview"
   ClientHeight    =   5940
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   8670
   Icon            =   "printpreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   578
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton Command3 
      Caption         =   "Options..."
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   90
   End
End
Attribute VB_Name = "printpreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = True
Me.PrintForm
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
End Sub

Private Sub Command2_Click()
Unload Me
Unload ppvoptions
End Sub

Private Sub Command3_Click()
ppvoptions.Show 1, Me
End Sub

Private Sub Form_Load()
Me.ScaleMode = 1
Command1.Top = 0
Command1.Left = printpreview.Width - Command1.Width
Command2.Top = Command1.Height
Command2.Left = Command1.Left
Command3.Top = Command1.Height * 2
Command3.Left = Command1.Left
Me.ScaleMode = 3
Me.BackColor = ppvoptions.Shape1.BackColor
Me.ForeColor = ppvoptions.Shape2.BackColor
End Sub

Private Sub Form_Resize()
Me.ScaleMode = 1
Command1.Top = 0
Command1.Left = printpreview.Width - Command1.Width
Command2.Top = Command1.Height
Command2.Left = Command1.Left
Command3.Top = Command1.Height * 2
Command3.Left = Command1.Left
Me.ScaleMode = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload ppvoptions
End Sub
