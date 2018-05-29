VERSION 5.00
Begin VB.Form frmDialStatus 
   Appearance      =   0  'Flat
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5265
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5265
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      Caption         =   "&Disconnect"
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
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3120
      Top             =   0
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dialing "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmDialStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Dim i As Integer
    For i = 0 To 9
       frmDialer.Command2(i).Enabled = True
       frmDialer.Show
       frmDialer.cmdCall.Enabled = True
    Next
End Sub

Private Sub Form_Activate()
Label1.Left = 0
Label1.Left = (Me.Width - Label1.Width) / 2
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Me.Width = Me.Width
Me.Top = 3000
Me.Left = 3000
cmdCancel.Enabled = False
End Sub

Private Sub Timer1_Timer()
If frmDialer.MSComm1.CDHolding = True Then
    Label1.Caption = frmDialer.Display.Text & " Connected"
    cmdCancel.Enabled = True
Else
    Label1.Caption = "Call Disconnected"
    Unload Me
    frmDialer.Show
    frmDialer.cmdCall.Enabled = True
    frmDialer.MSComm1.PortOpen = False
End If
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = QBColor(Rnd * 12)
cmdCancel.BackColor = QBColor(Rnd * 12)
End Sub
