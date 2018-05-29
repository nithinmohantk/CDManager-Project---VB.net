VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDialer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phone Dialer"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmDialer.frx":0000
   LinkTopic       =   "frmDialer"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5655
   Begin VB.TextBox Display 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   13
      Text            =   "0"
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000040C0&
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2520
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdRedial 
      BackColor       =   &H000040C0&
      Caption         =   "Redial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdCall 
      BackColor       =   &H000040C0&
      Caption         =   "CALL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5040
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParityReplace   =   49
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   23
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1080
      Picture         =   "frmDialer.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "To dial the number, Click on number you want and click on CALL Button."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   960
      TabIndex        =   19
      Top             =   4320
      Width           =   4410
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   4215
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   735
      Left            =   840
      TabIndex        =   16
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Height          =   3495
      Left            =   1200
      TabIndex        =   15
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Height          =   3495
      Left            =   840
      TabIndex        =   14
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "frmDialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim str_author As String
  Dim i As Integer
  
Private Sub cmdCall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCall.FontBold = True
End Sub

Private Sub cmdCall_Click()
If MSComm1.PortOpen = True Then
    MSComm1.Output = "ATDT " & Display.Text & vbCr
    frmDialStatus.Show
    frmDialStatus.Label1.Caption = frmDialStatus.Label1.Caption & Display.Text
    SaveSetting App.Title, App.Path, "num", Display.Text
    cmdCall.Enabled = False
    For i = 0 To 9
        Command2(i).Enabled = False
    Next
Else
    MsgBox "Modem/Telephone line is not Connected to the Computer" & vbCrLf & "Please check whether there is any problem in modem/telephone line", vbCritical + vbOKOnly, "PORT is Not Open"
    cmdCall.Enabled = True
End If
End Sub
Private Sub cmdClear_Click()
Display.Text = ""
If Len(Display.Text) > 0 Then
cmdCall.Enabled = True
Else
cmdCall.Enabled = False
End If
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
End Sub
Private Sub Command2_Click(Index As Integer)
Display.Text = Display.Text + Command2(Index).Caption
If Len(Display.Text) > 0 Then
cmdCall.Enabled = True
Else
cmdCall.Enabled = False
End If
If Len(Display.Text) >= 18 Then
MsgBox "Invalid Number. Please re-enter.", vbOKOnly + vbInformation, "Phone Dialer"
Display.Text = ""
End If
End Sub
Private Sub cmdRedial_Click()
MSComm1.PortOpen = False
End Sub
Private Sub Display_Change()
If Len(Display.Text) = 0 Then
cmdRedial.Enabled = False
Else
cmdRedial.Enabled = True
End If
End Sub
Private Sub Form_Load()
str_author = "This program is developed by Nithin Mohan.T.K , E-Mail : nithinmohantk@gmail.com" & _
       "  for Dream Works Technologies India Ltd, Have a nice day!!!! "
Display.SelStart = 0
Display.SelLength = Len(Display.Text)
Display.SelText = Len(Display.Text)
'Display.Text = ""
On Error Resume Next
Display.Text = GetSetting(App.Title, App.Path, "num", Display.Text)
MSComm1.CommPort = 1
MSComm1.Settings = "9600,N,8,1"
MSComm1.PortOpen = True
Me.Top = 100
Me.Left = 3000
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCall.FontBold = False
End Sub

