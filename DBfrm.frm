VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DBfrm 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Database"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "DBfrm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Upd 
      Caption         =   "Update"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Product 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   11
      Text            =   "<none>"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Manufacturer 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "<none>"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Gen 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save as new"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Country 
      Height          =   285
      Left            =   120
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "<none>"
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox FCode 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4080
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Image"
      Filter          =   "Alle dateien | *.*"
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Punkt
      X1              =   216
      X2              =   120
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Punkt
      X1              =   120
      X2              =   216
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Punkt
      X1              =   216
      X2              =   216
      Y1              =   128
      Y2              =   48
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Punkt
      X1              =   120
      X2              =   120
      Y1              =   128
      Y2              =   48
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Code:"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Product:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Manufacturer:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "No Picture"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "97x81"
      Height          =   195
      Left            =   2310
      TabIndex        =   9
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label ppath 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "DBfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Position As Long
Public LastRecord As Long

Private Sub FCode_Click()
Dim FileNum As Integer, RecLength As Long, CodeDB As DBcode
RecLength = LenB(CodeDB)
FileNum = FreeFile
Open toolfrm.Text2.text For Random As FileNum Len = RecLength
Get FileNum, GetPosition(FCode.text) + 1, CodeDB
Country.text = Trim$(CodeDB.Country)
Manufacturer.text = Trim$(CodeDB.Manufacturer)
Product.text = Trim$(CodeDB.Product)
If Trim$(CodeDB.PicPath) <> "" Then
    ppath.Caption = Trim$(CodeDB.PicPath)
    Image1.Picture = LoadPicture(Trim$(CodeDB.PicPath))
    Label5.Visible = False
    Label6.Visible = False
Else
    ppath.Caption = ""
    Label5.Visible = True
    Label6.Visible = True
End If
Close FileNum
End Sub

Private Sub Form_Load()
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
If LastRecord > 1 Then
For Position = 2 To LastRecord
Get FileNum, Position, CodeDB
FCode.AddItem Trim$("<" & (Position - 1) & "> " & CodeDB.FullCode)
Next
End If
Close FileNum
End Sub

Private Sub Gen_Click()
newBarCode
Mainfrm.ActiveForm.Caption = Product.text
PaintCode Mainfrm.ActiveForm, Mid$(Right$(FCode.text, 13), 1, 1), Mid$(Right$(FCode.text, 13), 2, 6), Mid$(Right$(FCode.text, 13), 8, 6)
Mainfrm.ActiveForm.Label1.Caption = Mid$(Right$(FCode.text, 13), 1, 1)
Mainfrm.ActiveForm.Label2.Caption = Mid$(Right$(FCode.text, 13), 2, 6)
Mainfrm.ActiveForm.Label3.Caption = Mid$(Right$(FCode.text, 13), 8, 6)
Mainfrm.ActiveForm.isEAN = True
Mainfrm.ActiveForm.Refresh
End Sub

Private Sub Image1_Click()
On Error Resume Next
CD1.ShowOpen
If CD1.FileName <> "" Then
    Image1.Picture = LoadPicture(CD1.FileName)
    ppath.Caption = CD1.FileName
    End If
End Sub

Private Sub Save_Click()
Dim FileNum As Integer, RecLength As Long, CodeDB As DBcode
RecLength = LenB(CodeDB)
FileNum = FreeFile
Open toolfrm.Text2.text For Random As FileNum Len = RecLength
CodeDB.Country = Country.text
CodeDB.Manufacturer = Manufacturer.text
CodeDB.Product = Product.text
CodeDB.FullCode = Right$(FCode.text, 13)
CodeDB.PicPath = ppath.Caption
Put FileNum, LastRecord + 1, CodeDB
LastRecord = LastRecord + 1
CodeDB.FullCode = LastRecord
Put FileNum, 1, CodeDB
Close FileNum
End Sub

Private Sub Upd_Click()
Dim FileNum As Integer, RecLength As Long, CodeDB As DBcode
RecLength = LenB(CodeDB)
FileNum = FreeFile
Open toolfrm.Text2.text For Random As FileNum Len = RecLength
CodeDB.Country = Country.text
CodeDB.Manufacturer = Manufacturer.text
CodeDB.Product = Product.text
CodeDB.FullCode = Right$(FCode.text, 13)
CodeDB.PicPath = ppath.Caption
Put FileNum, (GetPosition(FCode.text) + 1), CodeDB
CodeDB.FullCode = LastRecord
Put FileNum, 1, CodeDB
Close FileNum
End Sub

