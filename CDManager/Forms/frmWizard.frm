VERSION 5.00
Begin VB.Form frmWizard 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome to MyPersonal CD Manager"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   FillColor       =   &H00800080&
   ForeColor       =   &H00000080&
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmWizard.frx":08CA
   ScaleHeight     =   6030
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   5400
      Picture         =   "frmWizard.frx":17FF6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton optCDBor 
      BackColor       =   &H00E0E0E0&
      DisabledPicture =   "frmWizard.frx":188C0
      DownPicture     =   "frmWizard.frx":1918A
      DragIcon        =   "frmWizard.frx":19A54
      Height          =   615
      Left            =   480
      Picture         =   "frmWizard.frx":1A31E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton optFind 
      BackColor       =   &H00E0E0E0&
      DisabledPicture =   "frmWizard.frx":1ABE8
      DownPicture     =   "frmWizard.frx":1B4B2
      DragIcon        =   "frmWizard.frx":1BD7C
      Height          =   615
      Left            =   480
      Picture         =   "frmWizard.frx":1C646
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton optBorrowList 
      BackColor       =   &H00E0E0E0&
      DisabledPicture =   "frmWizard.frx":1CF10
      DownPicture     =   "frmWizard.frx":1D7DA
      DragIcon        =   "frmWizard.frx":1E0A4
      Height          =   615
      Left            =   480
      Picture         =   "frmWizard.frx":1E96E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton optCDCol 
      BackColor       =   &H00E0E0E0&
      DisabledPicture =   "frmWizard.frx":1F238
      DownPicture     =   "frmWizard.frx":1FB02
      DragIcon        =   "frmWizard.frx":203CC
      Height          =   615
      Left            =   480
      Picture         =   "frmWizard.frx":20C96
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox chkWizard 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Compact Discs"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Load This on StartUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   5520
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome user,you are in the world of CDManager.You can use this software instantly using this WIZARD."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Set and watch CDRoms you own"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Add and watch Borrower  details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Search for CDRoms of your Need"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Watch and Print Reports."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Borrowers Entry"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search for CDRoms"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "View Reports"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   3960
      Width           =   3015
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkWizard_Click()
If chkWizard.Value = 1 Then
  UseWizard = True
ElseIf chkWizard.Value = 0 Then
   UseWizard = False
End If
Call SaveSettings
End Sub

Private Sub cmdCancel_Click()
Call SaveSettings
Call LoadSettings
Unload Me
End Sub

Private Sub Form_Load()
Call LoadSettings
chkWizard.Enabled = True
If UseWizard = True Then
    chkWizard.Value = 1
ElseIf UseWizard = False Then
    chkWizard.Value = 0
End If
Me.Top = 20
Me.Left = 3000
End Sub

Private Sub optBorrowList_Click()
Load frmReport
frmReport.Show
End Sub

Private Sub optCDBor_Click()
Load frmBorrow
frmBorrow.Show
End Sub

Private Sub optCDCol_Click()
Load frmCDMan
frmCDMan.Show
End Sub
Private Sub optFind_Click()
Load frmFind
frmFind.Show
End Sub

Private Sub Timer1_Timer()

End Sub
