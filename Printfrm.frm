VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Printfrm 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Print"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2520
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pages"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
      Begin VB.OptionButton Oppc 
         Caption         =   "One Page per code"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Opfa 
         Caption         =   "One Page for all"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton Selected 
         Caption         =   "Selected"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton All 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Options..."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Preview..."
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Printfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub All_Click()
If Selected.Value = True Then
    Opfa.Enabled = False
    Oppc.Enabled = False
Else
    Opfa.Enabled = True
    Oppc.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Selected.Value = True Then
    Mainfrm.ActiveForm.PrintForm
ElseIf All.Value = True Then
    If Oppc.Value = True Then
        For A = 1 To UBound(bcode)
            If FState(bcode(A).Tag).Deleted <> True Then
                bcode(A).PrintForm
            End If
        Next
    ElseIf Opfa.Value = True Then
        printpreview.Show
        Dim reihe
        Dim z
        Dim B
        Dim D
        ghez = 0
        pl = 0
        For A = 1 To UBound(bcode)
StartX:
        If bcode(A).isEAN = False Then
            MsgBox "Print preview currently only supports EAN-Codes", vbCritical, "Error"
            A = A + 1
            If A > UBound(bcode) Then
                Exit For
            End If
            GoTo StartX
        ElseIf FState(bcode(A).Tag).Deleted = True Then
            A = A + 1
            If A > UBound(bcode) Then
                Exit For
            End If
            GoTo StartX
        End If
        fi = bcode(A).Label1.Caption
        se = bcode(A).Label2.Caption
        th = bcode(A).Label3.Caption
        If A = 1 Then
            printpreview.Label1(0).Caption = fi
            printpreview.Label2(0).Caption = se
            printpreview.Label3(0).Caption = th
        ElseIf A > 1 Then
            Load printpreview.Label1(A - 1)
            Load printpreview.Label2(A - 1)
            Load printpreview.Label3(A - 1)
            If ppvoptions.showtitle.Value = True Then
                Load printpreview.Label4(A - 1)
                printpreview.Label4(A - 1).Caption = bcode(A).Caption
            End If
            printpreview.Label1(A - 1).Caption = fi
            printpreview.Label2(A - 1).Caption = se
            printpreview.Label3(A - 1).Caption = th
            If (A - 1) Mod ppvoptions.rowtext.text <> 0 Then
                printpreview.Label1(A - 1).Left = printpreview.Label1(A - 2).Left
                printpreview.Label2(A - 1).Left = printpreview.Label2(A - 2).Left
                printpreview.Label3(A - 1).Left = printpreview.Label3(A - 2).Left
                printpreview.Label1(A - 1).Top = printpreview.Label1(A - 2).Top + 50
                printpreview.Label2(A - 1).Top = printpreview.Label2(A - 2).Top + 50
                printpreview.Label3(A - 1).Top = printpreview.Label3(A - 2).Top + 50
                If ppvoptions.showtitle.Value = True Then
                    printpreview.Label4(A - 1).Left = printpreview.Label4(A - 2).Left
                    printpreview.Label4(A - 1).Top = printpreview.Label4(A - 2).Top + 50
                    printpreview.Label4(A - 1).Visible = True
                End If
            Else
                printpreview.Label1(A - 1).Left = printpreview.Label1(A - 2).Left + 120
                printpreview.Label2(A - 1).Left = printpreview.Label2(A - 2).Left + 120
                printpreview.Label3(A - 1).Left = printpreview.Label3(A - 2).Left + 120
                printpreview.Label1(A - 1).Top = 20
                printpreview.Label2(A - 1).Top = 20
                printpreview.Label3(A - 1).Top = 20
                If ppvoptions.showtitle.Value = True Then
                    printpreview.Label4(A - 1).Left = printpreview.Label4(A - 2).Left + 120
                    printpreview.Label4(A - 1).Top = 32
                    printpreview.Label4(A - 1).Visible = True
                End If
            End If
            printpreview.Label1(A - 1).Visible = True
            printpreview.Label2(A - 1).Visible = True
            printpreview.Label3(A - 1).Visible = True
        End If
        reihe = 0
        z = 0
        B = 0
        D = 0
        printpreview.Line (1 + 10 + ghez, 0 + pl)-(1 + 10 + ghez, 25 + pl) 'Paint the First two lines on the begin of the Code
        printpreview.Line (3 + 10 + ghez, 0 + pl)-(3 + 10 + ghez, 25 + pl)
        reihe = code(fi)
        For z = 1 To 6 'Use A and B code to Decode the Barcode 'For each 6 numbers use 7 Lines 6 * 7 = 47 Lines
            If Mid(reihe, z, 1) = "A" Then 'Code A
                B = CodeAToByte(Mid(se, z, 1))
                For D = 1 To 7 'Paint the 7 Lines (A Code)
                    If Mid(B, D, 1) = 1 Then 'On all 7 numbers Check if it is a 1 or a 0 and Paint a Black or a White Line
                        printpreview.Line ((z - 1) * 7 + D + 3 + 10 + ghez, 0 + pl)-((z - 1) * 7 + D + 3 + 10 + ghez, 20 + pl), &H0 'Black Line
                    Else
                        printpreview.Line ((z - 1) * 7 + D + 3 + 10 + ghez, 0 + pl)-((z - 1) * 7 + D + 3 + 10 + ghez, 20 + pl), &HFFFFFF 'White Line
                    End If
                Next
            ElseIf Mid(reihe, z, 1) = "B" Then 'Code B
                B = CodeBToByte(Mid(se, z, 1))
                For D = 1 To 7 'Paint the 7 Lines (B Code)
                    If Mid(B, D, 1) = 1 Then 'On all 7 numbers Check if it is a 1 or a 0 and Paint a Black or a White Line
                        printpreview.Line ((z - 1) * 7 + D + 3 + 10 + ghez, 0 + pl)-((z - 1) * 7 + D + 3 + 10 + ghez, 20 + pl), &H0 'Black Line
                    Else
                        printpreview.Line ((z - 1) * 7 + D + 3 + 10 + ghez, 0 + pl)-((z - 1) * 7 + D + 3 + 10 + ghez, 20 + pl), &HFFFFFF 'White Line
                    End If
                Next
            End If
        Next
        printpreview.Line (6 * 7 + 5 + 10 + ghez, 0 + pl)-(6 * 7 + 5 + 10 + ghez, 25 + pl) 'Paint the middle two lines of the Code
        printpreview.Line (6 * 7 + 7 + 10 + ghez, 0 + pl)-(6 * 7 + 7 + 10 + ghez, 25 + pl)
            For z = 1 To 6 'Use C code to Decode the Barcode 'For each 6 numbers use 7 Lines 6 * 7 = 47 Lines
                B = CodeCToByte(Mid(th, z, 1)) ' Code C
                For D = 1 To 7 'Paint the 7 Lines (C Code)
                    If Mid(B, D, 1) = 1 Then 'On all 7 numbers Check if it is a 1 or a 0 and Paint a Black or a White Line
                        printpreview.Line ((z - 1) * 7 + D + 50 + 10 + ghez, 0 + pl)-((z - 1) * 7 + D + 50 + 10 + ghez, 20 + pl), &H0 'Black Line
                    Else
                        printpreview.Line ((z - 1) * 7 + D + 50 + 10 + ghez, 0 + pl)-((z - 1) * 7 + D + 50 + 10 + ghez, 20 + pl), &HFFFFFF 'White Line
                    End If
                Next
            Next
        printpreview.Line (94 + 9 + ghez, 0 + pl)-(94 + 9 + ghez, 25 + pl) 'The Last two lines
        printpreview.Line (96 + 9 + ghez, 0 + pl)-(96 + 9 + ghez, 25 + pl)
        pl = pl + 50
        If A Mod ppvoptions.rowtext.text = 0 Then
            pl = 0
            ghez = ghez + 120
        End If
        Next
    End If
End If
End Sub

Private Sub Command2_Click()
Unload ppvoptions
Unload Me
End Sub

Private Sub Command3_Click()
CD1.ShowPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload ppvoptions
End Sub

Private Sub Opfa_Click()
Command1.Caption = "Preview..."
End Sub

Private Sub Oppc_Click()
Command1.Caption = "Print"
End Sub

Private Sub Selected_Click()
If Selected.Value = True Then
    Opfa.Enabled = False
    Oppc.Enabled = False
Else
    Opfa.Enabled = True
    Oppc.Enabled = True
End If
End Sub
