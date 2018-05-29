VERSION 5.00
Begin VB.Form frmNewUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter New User Information"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass1 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtPass2 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "&Accept"
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
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H000040C0&
      Height          =   2535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter New User Login Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackColor       =   &H000040C0&
      Height          =   615
      Left            =   600
      TabIndex        =   14
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   1440
      Width           =   15
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Height          =   1935
      Left            =   600
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Reenter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "New User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   840
      TabIndex        =   10
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAccept_Click()
If Not txtUser.Text = "" Then
    If Not txtPass1.Text = "" And Not txtPass2.Text = "" Then
         newpass1 = Trim(txtPass1.Text)
         newpass2 = Trim(txtPass2.Text)
         Call newuser
    Else
       MsgBox "Because of Security reasons EMPTY passwords are not allowed"
       txtPass1.SetFocus
    End If
Else
    MsgBox "You Left UserName field Empty"
End If
End Sub

Private Sub cmdAccept_GotFocus()
If Not txtPass2.Text = "" Then
     cmdAccept.SetFocus
  Else
     MsgBox "Because of Security reasons EMPTY passwords are not allowed"
     txtPass2.SetFocus
  End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub txtPass1_GotFocus()
 If Not txtUser.Text = "" Then
    txtPass1.SetFocus
 Else
    MsgBox "User Name cannot be empty"
 End If
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Not txtPass1.Text = "" Then
     txtPass2.SetFocus
  Else
     MsgBox "Because of Security reasons EMPTY passwords are not allowed"
     txtPass1.SetFocus
  End If
End If
End Sub

Private Function validate() As Boolean
validate = False
If Not rsLogin.RecordCount < 1 Then
If rsLogin.EOF Or rsLogin.BOF Then
  rsLogin.MoveFirst
End If
While Not rsLogin.EOF
   If rsLogin!LOGINID = Trim(LCase(txtUser.Text)) Then
         validate = True
   End If
   rsLogin.MoveNext
Wend
End If
End Function

Private Sub newuser()
Call CommitDB
    If validate = False Then
        If newpass1 = newpass2 Then
            rsLogin.AddNew
            rsLogin!LOGINID = LCase(Trim(txtUser.Text))
            rsLogin!Password = encrypt_pass(Trim(txtPass1.Text))
            rsLogin.Update
            rsLogin.Close
            Call CommitDB
            MsgBox "New User " & Trim(UCase(txtUser.Text)) & " Successfully Added"
            frmNewUser.Hide
            frmpassword.Show
        Else
            MsgBox "Passwords not match,please reenter it"
            txtPass1.SetFocus
        End If
    ElseIf validate = True Then
        MsgBox "UserName Already Exists"
    End If
End Sub



Private Sub txtPass2_GotFocus()
If Not txtPass1.Text = "" Then
     txtPass2.SetFocus
  Else
     MsgBox "Because of Security reasons EMPTY passwords are not allowed"
     txtPass1.SetFocus
  End If
End Sub

Private Sub txtPass2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Not txtPass2.Text = "" Then
     cmdAccept.SetFocus
  Else
     MsgBox "Because of Security reasons EMPTY passwords are not allowed"
     txtPass2.SetFocus
  End If
End If
End Sub



Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Not txtUser.Text = "" Then
    txtPass1.SetFocus
 Else
    MsgBox "User Name cannot be empty"
 End If
End If
End Sub
