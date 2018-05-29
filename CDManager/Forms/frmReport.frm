VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Manager"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6105
   Begin VB.ListBox lstCDMODE 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "frmReport.frx":0000
      Left            =   3840
      List            =   "frmReport.frx":002E
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "Select Report Mode"
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
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   2295
      Begin VB.OptionButton optBorrow 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "CD Borrow Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optCDList 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "CDRom List Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H000040C0&
      Caption         =   "&Print"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   480
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "G:\Backup\Nithin\Works\CDManager\Designers\cdlist.rpt"
      WindowTitle     =   "MyPersonal CDMangaer Report Viewer"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   7
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H000080FF&
      Height          =   1815
      Left            =   5760
      TabIndex        =   18
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H000080FF&
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Height          =   2175
      Left            =   360
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   600
      TabIndex        =   7
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      Height          =   1575
      Left            =   5280
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   5520
      TabIndex        =   17
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sf As String
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
If optCDList.Value = True Then
    CR1.ReportFileName = App.Path & "\Designers\CDList.rpt"
    If LCase(Trim(lstCDMODE.Text)) = "all(default)" Then
        sf = "{CDCol.SNO} < 100000 and {CDCol.SNO} > 0"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "AUDIO" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "DATA" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "VIDEO" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "MP3" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "LINUX" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "PC" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "GAME" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "UNIX" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "GENERAL" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "OTHERS" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "E-BOOKS" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "MAC" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "DVD" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    End If
    CR1.SelectionFormula = sf
    CR1.Connect = App.Path & "\CDMANDB.MDB"
'    CR1.Connect = "DSN=CDMAN;UID = admin;PWD="
    CR1.RetrieveDataFiles
    CR1.WindowState = crptMaximized
    CR1.Action = 1
ElseIf optBorrow.Value = True Then
    CR1.ReportFileName = App.Path & "\Designers\borrowlist.rpt"
    If LCase(Trim(lstCDMODE.Text)) = "all(default)" Then
        sf = "{BORROW.RETURNED}=0"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "AUDIO" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "DATA" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "VIDEO" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "MP3" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "LINUX" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "PC" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "GAME" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "UNIX" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "GENERAL" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "OTHERS" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "E-BOOKS" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "MAC" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    ElseIf UCase(Trim(lstCDMODE.Text)) = "DVD" Then
          sf = "{CDCol.CDTYPE}= '" & UCase(Trim(lstCDMODE.Text)) & "'"
    End If
    CR1.SelectionFormula = sf
    CR1.Connect = App.Path & "\CDMANDB.MDB"
'    CR1.Connect = "DSN=CDMAN;UID = admin;PWD="
    CR1.RetrieveDataFiles
    CR1.WindowState = crptMaximized
    CR1.Action = 1
End If

End Sub

Private Sub Form_Load()
Me.Top = 1000
Me.Left = 3000
End Sub

