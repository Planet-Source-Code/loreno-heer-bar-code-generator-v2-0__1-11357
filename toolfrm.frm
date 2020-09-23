VERSION 5.00
Begin VB.Form toolfrm 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Tools"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command5 
      Caption         =   "Options..."
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open..."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "C:\programm files\bar code\bc.db"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "P"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Repair"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   13
      TabIndex        =   1
      Text            =   "2501007661990"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      Height          =   195
      Left            =   2212
      TabIndex        =   14
      Top             =   1455
      Width           =   390
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Left            =   1057
      TabIndex        =   13
      Top             =   1455
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      Height          =   195
      Left            =   652
      TabIndex        =   12
      Top             =   1450
      Width           =   345
   End
   Begin VB.Label Label5 
      Caption         =   "mailto:borg@bluewin.ch"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Loreno Heer"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Programed by:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   3240
      Y1              =   710
      Y2              =   710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EAN"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   0
      X2              =   3240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   3240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   0
      X2              =   3240
      Y1              =   2175
      Y2              =   2175
   End
End
Attribute VB_Name = "toolfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Len(Text1.text) >= 12 And IsNumeric(Text1.text) = True Then
Text1.text = Mid$(Text1.text, 1, 12)
Text1.text = Text1.text + 1
Text1.text = Text1.text & "0"
Do Until CheckCode(Text1.text) = True
Text1.text = Mid$(Text1.text, 1, 12) & (Mid$(Text1.text, 13, 1) + 1)
Loop
End If
End Sub

Private Sub Command2_Click()
If Len(Text1.text) >= 12 And IsNumeric(Text1.text) = True Then
Text1.text = Mid$(Text1.text, 1, 12)
Text1.text = Text1.text - 1
Text1.text = Text1.text & "0"
Do Until CheckCode(Text1.text) = True
Text1.text = Mid$(Text1.text, 1, 12) & (Mid$(Text1.text, 13, 1) + 1)
Loop
End If
End Sub

Private Sub Command3_Click()
If Len(Text1.text) >= 12 And IsNumeric(Text1.text) = True Then
Text1.text = Mid$(Text1.text, 1, 12)
Text1.text = Text1.text & "0"
Do Until CheckCode(Text1.text) = True
Text1.text = Mid$(Text1.text, 1, 12) & (Mid$(Text1.text, 13, 1) + 1)
Loop
End If
End Sub

Private Sub Command4_Click()
    If Len(Text1.text) = 13 Then
        Mainfrm.ActiveForm.Width = (options.Text1.text)
        Mainfrm.ActiveForm.Cls
        Mainfrm.ActiveForm.ScaleMode = 3
        Mainfrm.ActiveForm.Label1.Visible = True
        PaintCode Mainfrm.ActiveForm, Mid$(Text1.text, 1, 1), Mid$(Text1.text, 2, 6), Mid$(Text1.text, 8, 6)
        Mainfrm.ActiveForm.Label1.Caption = Mid$(Text1.text, 1, 1)
        Mainfrm.ActiveForm.Label2.Caption = Mid$(Text1.text, 2, 6)
        Mainfrm.ActiveForm.Label3.Caption = Mid$(Text1.text, 8, 6)
        Mainfrm.ActiveForm.Refresh
    Else
        MsgBox "Error: The lenght of the code is wrong", vbCritical, "Error"
    End If
End Sub

Private Sub Command5_Click()
options.Show
End Sub

Private Sub Command6_Click()
DBfrm.Show
End Sub

Private Sub Form_Load()
Dim FileNum As Integer, RecLength As Long, CodeDB As DBcode
RecLength = LenB(CodeDB)
FileNum = FreeFile
Text2.text = App.Path & "\BarCode.db"
Open Text2.text For Random As FileNum Len = RecLength
Label7.Caption = LOF(FileNum)
Close FileNum
End Sub

