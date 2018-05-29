VERSION 5.00
Begin VB.Form frmChgPass 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   ControlBox      =   0   'False
   Icon            =   "frmChgPass.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000040C0&
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000040C0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      BackColor       =   &H000040C0&
      Caption         =   "&ACCEPT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "New Password"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   4455
      Begin VB.TextBox txtNewPass2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   885
         Width           =   2265
      End
      Begin VB.TextBox txtNewPass1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   405
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   2085
         TabIndex        =   15
         Top             =   840
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   2085
         TabIndex        =   13
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PASSWORD:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   405
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "REENTER PASSWORD :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   930
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Login and Old Password"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   885
         Width           =   2265
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   10
         Top             =   405
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   2085
         TabIndex        =   11
         Top             =   840
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   2085
         TabIndex        =   3
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "OLD PASSWORD :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   930
         Width           =   1575
      End
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Only Logged-on users can change password."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Width           =   5415
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080C0FF&
      Height          =   855
      Left            =   0
      TabIndex        =   30
      Top             =   4200
      Width           =   5895
   End
   Begin VB.Label Label17 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H000040C0&
      Height          =   2775
      Left            =   5400
      TabIndex        =   26
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label14 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label11 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   -120
      TabIndex        =   21
      Top             =   3600
      Width           =   6015
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Height          =   3135
      Left            =   0
      TabIndex        =   18
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Height          =   135
      Left            =   960
      TabIndex        =   19
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Height          =   2775
      Left            =   480
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label18 
      BackColor       =   &H000080FF&
      Height          =   4815
      Left            =   5880
      TabIndex        =   29
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmChgPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccept_Click()
newpass1 = Trim(txtNewPass1.Text)
newpass2 = Trim(txtNewPass2.Text)
ss1 = "select * from login where LOGINID = '" & LCase(loginuser) & "'"
Set rs = conn.Execute(ss1)
If Not rs.EOF Then
   rs.MoveFirst
   If decrypt_pass(rs!Password) = loginpass Then
    If Trim(txtNewPass1.Text) = Trim(txtNewPass2.Text) Then
        If rsLogin.EOF Or rsLogin.BOF Then
           rsLogin.MoveFirst
        End If
        While Not rsLogin.EOF
            If rsLogin!LOGINID = LCase(Trim(txtUser.Text)) Then
                 Call updatepass
                 Exit Sub
            Else
                 rsLogin.MoveNext
            End If
        Wend
''ss5 = "update LOGIN set PASSWORD = '" & Trim(txtNewPass1.Text) & "' " & _
''             "where LOGINID = '" & Trim(LCase(txtUser.Text)) & "' ;"
''        conn.Execute (ss5)
           
    Else
        MsgBox "New Passwords & Retyped passwords doesn't match"
    End If
   Else
      MsgBox "Password incorrect"
   End If
Else
   MsgBox "User Name incorrect"
End If
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Sub updatepass()
rsLogin!LOGINID = LCase(Trim(txtUser.Text))
rsLogin!Password = encrypt_pass(Trim(newpass1))
rsLogin.Update
Call CommitDB
MsgBox "Password Changed Successfully"
pass_changed = True
loginpass = txtNewPass1.Text
End Sub

Private Sub cmdOK_Click()
If pass_changed = False Then
    Call cmdAccept_Click
Else
     pass_changed = False
End If
Unload Me
End Sub

Private Sub Form_Load()
txtUser.Text = loginuser
txtUser.Enabled = False
Me.Top = 400
Me.Left = 3000
pass_changed = False
End Sub

