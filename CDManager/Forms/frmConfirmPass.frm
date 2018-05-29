VERSION 5.00
Begin VB.Form frmConfirmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm ADMININSTRATOR Password"
   ClientHeight    =   1695
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4680
   Icon            =   "frmConfirmPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1001.463
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.267
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1725
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H000040C0&
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Administrator Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   4320
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmConfirmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
If confirm_pass(Trim(txtPassword.Text)) = True Then
        Call del_all
Else
   MsgBox "Invalid Administrator Password"
End If
Me.Hide
End Sub

Private Sub Form_Load()
Me.Top = 2000
Me.Left = 2500
End Sub

