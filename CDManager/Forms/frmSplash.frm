VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "Mci32.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   5295
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000003&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer6 
      Interval        =   3000
      Left            =   5400
      Top             =   4440
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4680
      Top             =   4440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   4440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2520
      Top             =   4440
   End
   Begin VB.Timer Timer1 
      Interval        =   8000
      Left            =   1560
      Top             =   4440
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   5880
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   873
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblLicenseTo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "LicenseTo"
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   24
      Top             =   3600
      Width           =   5415
   End
   Begin VB.Label lblReg 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "WebSite           http://www.dreamworksindia.co.nr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5010
      Width           =   4455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail :"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "nithinmohantk@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   1320
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   5760
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   4200
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   2640
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5160
      TabIndex        =   3
      Top             =   3240
      Width           =   1650
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2760
      TabIndex        =   4
      Top             =   2835
      Width           =   1200
   End
   Begin VB.Label Label12 
      BackColor       =   &H000040C0&
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   960
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Height          =   2415
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404040&
      Height          =   2415
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   7095
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1200
      TabIndex        =   2
      Top             =   3840
      Width           =   5535
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4760
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   3480
      Width           =   7095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
   i = 0
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCopyright.Caption = App.LegalCopyright
    If reg_done = True Then
        lblReg.Caption = "REGISTERED VERSION"
        lblLicenseTo.Caption = "This Product is Licensed To " & " :-  " & reg_user
    ElseIf reg_done = False Then
        lblReg.Caption = "EVALUATION VERSION"
        lblLicenseTo.Caption = "Evaluation " & 100 - try_day & " Trys Left"
    End If
'    lblCompanyProduct.Caption = App.CompanyName
'    If SysInfo1.OSVersion >= 5 Then
'        frmMain.RegAccess.hKey = HKEY_LOCAL_MACHINE
'        frmMain.RegAccess.Path = "Software\Microsoft\Windows NT\CurrentVersion"
'        frmMain.RegAccess.ValueName = "CSDVersion"
'        OSVer = frmMain.RegAccess.GetValue
'        frmMain.RegAccess.ValueName = "ProductName"
'        OS = frmMain.RegAccess.GetValue
'        frmMain.RegAccess.ValueName = "RegisteredOwner"
'        OSOwner = frmMain.RegAccess.GetValue
'    ElseIf SysInfo1.OSVersion < 5 And SysInfo1.OSVersion > 4 Then
'        frmMain.RegAccess.hKey = HKEY_LOCAL_MACHINE
'        frmMain.RegAccess.Path = "Software\Microsoft\Windows\CurrentVersion"
'        frmMain.RegAccess.ValueName = "CSDVersion"
'        OSVer = frmMain.RegAccess.GetValue
'        frmMain.RegAccess.ValueName = "ProductName"
'        OS = frmMain.RegAccess.GetValue
'        frmMain.RegAccess.ValueName = "RegisteredOwner"
'        OSOwner = frmMain.RegAccess.GetValue
'    End If
'    lblLicenseTo.Caption = "Licensed To " & " :-  " & OSOwner
'    lblPlatform.Caption = OS & " - " & OSVer
MMControl1.Command = "CLOSE"
MMControl1.FileName = App.Path + "\yahoo.wav"
MMControl1.Command = "OPEN"
MMControl1.Command = "PLAY"
Timer6.Enabled = True
lbl1.Visible = False
lbl2.Visible = False
lbl3.Visible = False
lblProductName.Visible = False
End Sub


Private Sub Timer1_Timer()
Unload Me

Load frmpassword
frmpassword.Show
End Sub

Private Sub Timer2_Timer()
lbl1.Visible = True
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
lbl2.Visible = True
Timer4.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
lbl3.Visible = True
Timer5.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
lblProductName.Visible = Not lblProductName.Visible
End Sub

Private Sub Timer6_Timer()
MMControl1.Command = "CLOSE"
MMControl1.FileName = App.Path + "\yahoo.wav"
MMControl1.Command = "OPEN"
MMControl1.Command = "PLAY"
End Sub
