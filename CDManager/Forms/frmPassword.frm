VERSION 5.00
Begin VB.Form frmpassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Login"
   ClientHeight    =   3240
   ClientLeft      =   3105
   ClientTop       =   3165
   ClientWidth     =   5655
   ForeColor       =   &H00404040&
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H000040C0&
      Caption         =   "&New User"
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
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H000040C0&
      Caption         =   "&LogIn"
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
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000040C0&
      Caption         =   "&Cancel"
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
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Login Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   885
         Width           =   2265
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   1725
         TabIndex        =   6
         Top             =   840
         Width           =   2265
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "USER ID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   0
         Left            =   1725
         TabIndex        =   2
         Top             =   315
         Width           =   2265
      End
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail : nithinmohantk@gmail.com"
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
      Left            =   2400
      TabIndex        =   21
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Nithin Mohan.T.K"
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
      Left            =   1080
      TabIndex        =   20
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By :"
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CAUTION : Enter at your own Risk"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H000040C0&
      Height          =   2295
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Height          =   1935
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   840
      TabIndex        =   10
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim Y As Integer
Dim isNew As Boolean
Private Sub cmdCancel_Click()
try_day = try_day + 1
Call SaveSettings
End
End Sub

Private Sub cmdLogin_Click()
loginuser = LCase(Trim(txtUser.Text))
loginpass = Trim(txtPass.Text)
sql = "select * from Login where LOGINID = '" & LCase(Trim(txtUser.Text)) & "'"
    If rs.State = 1 Then rs.Close
    Set rs = conn.Execute(sql)
    If rs.EOF Then
        MsgBox "Invalid UserName : " & UCase(txtUser.Text), vbCritical, "Invalid UserName"
        Y = Y + 1
    Else
        rs.MoveFirst
        If decrypt_pass(rs!Password) = Trim(txtPass.Text) Then
            MsgBox "Access granted  " + txtUser.Text, vbInformation, "Message"
            Unload frmpassword
            Load frmMain
            frmMain.Show
            If UseWizard = True Then
             Load frmWizard
               frmWizard.Show
            End If
        Else
            MsgBox "Invalid password  for " & txtUser.Text, vbCritical, "Message"
            txtPass.Text = ""
            X = X + 1
            txtPass.SetFocus
        End If
    End If
    If X = 3 Or Y = 3 Then
     MsgBox "Sorry  " & txtUser.Text & "  you have exceeded the retry level,three trials only", vbInformation, "Message"
   End
   End If
End Sub

Private Sub cmdNew_Click()
isNew = True
Load frmNewUser
frmNewUser.Show
frmpassword.Hide
End Sub

Private Sub Form_Load()
isNew = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveFormSettings Me
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        sql = "select * from Login where LOGINID = '" & LCase(Trim(txtUser.Text)) & "'"
        If rs.State = 1 Then rs.Close
        Set rs = conn.Execute(sql)
        If rs.EOF Then
            MsgBox "Invalid UserName : " & UCase(txtUser.Text), vbCritical, "Invalid UserName"
        Else
            rs.MoveFirst
            If rs!LOGINID = LCase(Trim(txtUser.Text)) Then
                txtPass.SetFocus
            End If
        End If
        Y = Y + 1
   If Y = 3 Then
     MsgBox "Sorry  " & txtUser.Text & "  you have exceeded the retry level,three trials only", vbInformation, "Message"
   End
   End If
 End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
loginuser = LCase(Trim(txtUser.Text))
loginpass = Trim(txtPass.Text)
If KeyAscii = 13 Then
  If rs.BOF Or rs.EOF Then
    rs.MoveFirst
  End If
  If LCase(txtPass.Text) = decrypt_pass(rs!Password) Then
      MsgBox "Access granted  " + txtUser.Text, vbInformation, "Message"
      Unload frmpassword
      Load frmMain
      frmMain.Show
      If UseWizard = True Then
        Load frmWizard
        frmWizard.Show
      End If
  Else
      MsgBox "Invalid password  for " & txtUser.Text, vbCritical, "Message"
      txtPass.Text = ""
  X = X + 1
  Exit Sub
  End If
If X = 3 Then
MsgBox "Sorry  " & txtUser.Text & "  three trials only", vbInformation, "Message"
End
End If

End If
End Sub

