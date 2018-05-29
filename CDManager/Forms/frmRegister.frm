VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Register My Personal CD Manager"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   375
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Please Enter Your Registration details"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   6015
      Begin VB.TextBox ID7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox ID6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox ID1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox ID5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox ID4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox ID3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox ID2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtCompany 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtregName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
      Begin VB.Line Line6 
         X1              =   5280
         X2              =   5400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   4680
         X2              =   4800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line4 
         X1              =   4080
         X2              =   4200
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         X1              =   3480
         X2              =   3600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   3000
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         X1              =   2280
         X2              =   2400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Registered Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label txtRegI1 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration ID   :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Label lblEvaluation 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   3480
      Width           =   6015
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080C0FF&
      Height          =   1215
      Left            =   0
      TabIndex        =   27
      Top             =   3240
      Width           =   6975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration must be done inorder to use the software."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label11 
      BackColor       =   &H000040C0&
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   0
      Width           =   6615
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   2160
      Width           =   6015
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   2640
      Width           =   6975
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   1575
      Left            =   360
      TabIndex        =   17
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H000040C0&
      Height          =   2055
      Left            =   5880
      TabIndex        =   25
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
If called_by = False Then
    If expired = False And reg_done = False Then
            Call SaveSettings
            Call LoadSettings
            Load frmSplash
            frmSplash.Show
    ElseIf reg_done = True Then
            Call SaveSettings
            Call LoadSettings
            Load frmpassword
            frmpassword.Show
    ElseIf expired = True Then
            ans = MsgBox("cannot continue with out Registering" & vbCrLf & "Do you want to register JewelBox 2004 ??", vbCritical + vbYesNo, "EVALUATION EXPIRED")
            If ans = vbYes Then
                frmRegister.Show
            Else
                End
            End If
            Call SaveSettings
            Call LoadSettings
    End If
End If
End Sub
Private Sub cmdRegister_Click()
    reg_key = ID1.Text & "-" & ID2.Text & "-" & ID3.Text & "-" & ID4.Text & "-" & ID5.Text & "-" & ID6.Text & "-" & ID7.Text
    reg_company = Trim(txtCompany.Text)
    reg_user = Trim(txtregName.Text)
    first_reg = True
    Call SaveSettings
    Call LoadSettings
    If reg_done = True Then
        MsgBox "Registration Success", vbInformation + vbOKOnly, "Registration Success"
        Unload Me
        If called_by = False Then
            Load frmSplash
            frmSplash.Show
        End If
    Else
       MsgBox "INVALID REGistration KEY", vbCritical + vbOKOnly, "INVALID CDKEY"
    End If
End Sub

Private Sub Form_Load()
cmdRegister.Enabled = True
If reg_done = True Then
    Me.Caption = "Registered To " & reg_user
    cmdRegister.Enabled = False
    txtCompany.Enabled = False
    txtregName.Enabled = False
    ID1.Enabled = False
    ID2.Enabled = False
    ID2.Enabled = False
    ID3.Enabled = False
    ID4.Enabled = False
    ID5.Enabled = False
    ID6.Enabled = False
    ID7.Enabled = False
    ID1.Text = "####"
    ID2.Text = "####"
    ID3.Text = "####"
    ID4.Text = "####"
    ID5.Text = "####"
    ID6.Text = "####"
    ID7.Text = "####"
    txtCompany.Text = reg_company
    txtregName.Text = reg_user
    lblEvaluation.Caption = "Registered Version"
ElseIf reg_done = False Then
    lblEvaluation.Caption = "You have only " & 100 - try_day & " Trys are Left,Register it immediately"
    If try_day > 100 Then
       expired = True
       Call SaveSettings
       Call LoadSettings
       lblEvaluation.Caption = "Your Evaluation usage days are over," & vbCrLf & "Register CDManager 2004 inorder to continue using it."
    End If
End If
End Sub
Private Sub txtregName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Not txtregName.Text = "" Then
      txtCompany.SetFocus
   Else
      MsgBox "Registration Name Cannot be Empty"
      txtregName.SetFocus
   End If
End If
End Sub
Private Sub txtCompany_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Not txtCompany.Text = "" Then
      ID1.SetFocus
   Else
      MsgBox "Company Name Cannot be Empty"
      txtCompany.SetFocus
   End If
End If
End Sub

Private Sub ID1_KeyPress(KeyAscii As Integer)
If Len(ID1.Text) = 3 Then
       ID2.SetFocus
End If
End Sub
Private Sub ID2_KeyPress(KeyAscii As Integer)
If Len(ID2.Text) = 3 Then
       ID3.SetFocus
End If
End Sub
Private Sub ID3_KeyPress(KeyAscii As Integer)
If Len(ID3.Text) = 3 Then
       ID4.SetFocus
End If
End Sub
Private Sub ID4_KeyPress(KeyAscii As Integer)
If Len(ID4.Text) = 3 Then
       ID5.SetFocus
End If
End Sub
Private Sub ID5_KeyPress(KeyAscii As Integer)
If Len(ID5.Text) = 3 Then
       ID6.SetFocus
End If
End Sub
Private Sub ID6_KeyPress(KeyAscii As Integer)
If Len(ID6.Text) = 3 Then
       ID7.SetFocus
End If
End Sub
Private Sub ID7_KeyPress(KeyAscii As Integer)
If Len(ID7.Text) = 3 Then
       cmdRegister.SetFocus
End If
End Sub
