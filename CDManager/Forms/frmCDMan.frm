VERSION 5.00
Begin VB.Form frmCDMan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Your CD's "
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "frmCDMan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleMode       =   0  'User
   ScaleWidth      =   7455
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "frmCDMan.frx":0442
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   37
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H000040C0&
      Caption         =   "&New"
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
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H000040C0&
      Caption         =   "&Edit"
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
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4800
      Width           =   855
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
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H000040C0&
      Caption         =   "&Save"
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
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000040C0&
      Caption         =   "&Delete"
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
      TabIndex        =   26
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000040C0&
      Caption         =   "E&xit"
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
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CDROM INFORMATION"
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
      Height          =   4575
      Left            =   1080
      TabIndex        =   20
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox cboPlatform 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCDMan.frx":16B4
         Left            =   1440
         List            =   "frmCDMan.frx":16EB
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCDMan.frx":1771
         Left            =   4800
         List            =   "frmCDMan.frx":17E1
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCDMan.frx":18BD
         Left            =   1440
         List            =   "frmCDMan.frx":18E5
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtVersion 
         Appearance      =   0  'Flat
         Height          =   285
         HideSelection   =   0   'False
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   6
         Top             =   4080
         Width           =   1575
      End
      Begin VB.ComboBox cboCDCode 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCDMan.frx":1925
         Left            =   4320
         List            =   "frmCDMan.frx":1927
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtCDQTY 
         Appearance      =   0  'Flat
         Height          =   285
         HideSelection   =   0   'False
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   765
         HideSelection   =   0   'False
         Left            =   1440
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2640
         Width           =   4455
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         Height          =   285
         HideSelection   =   0   'False
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   2
         Top             =   2160
         Width           =   4455
      End
      Begin VB.ListBox lstCDMODE 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "frmCDMan.frx":1929
         Left            =   1440
         List            =   "frmCDMan.frx":1954
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "PLATFORM    :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR            :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "MONTH        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "VERSION :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "NO.Of DISC'S :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPTION :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "DISC TITLE  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DISC IDENTIFYING CODE : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   615
         Left            =   3000
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DISC DATA   : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label txtSNO 
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SNO : -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   42
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   41
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label23 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   40
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label22 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   39
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   38
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   7080
      TabIndex        =   36
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   495
      Left            =   3120
      TabIndex        =   35
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackColor       =   &H000040C0&
      Height          =   4575
      Left            =   7080
      TabIndex        =   34
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   720
      TabIndex        =   33
      Top             =   5520
      Width           =   6735
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   720
      TabIndex        =   32
      Top             =   5280
      Width           =   6735
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080C0FF&
      Height          =   1095
      Left            =   0
      TabIndex        =   31
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   1080
      TabIndex        =   24
      Top             =   4680
      Width           =   6015
   End
   Begin VB.Label Label13 
      BackColor       =   &H000040C0&
      Height          =   135
      Left            =   720
      TabIndex        =   23
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label12 
      BackColor       =   &H000040C0&
      Height          =   4695
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H000080FF&
      Height          =   5175
      Left            =   720
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmCDMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isNew As Boolean
Private Sub DisableF()
 txtCDQTY.Enabled = False
 txtDescription.Enabled = False
 txtTitle.Enabled = False
 txtVersion.Enabled = False
 cboYear.Enabled = False
 cboMonth.Enabled = False
 cboPlatform.Enabled = False
End Sub
Private Sub EnableF()
 txtCDQTY.Enabled = True
 txtDescription.Enabled = True
 txtTitle.Enabled = True
 txtVersion.Enabled = True
 cboYear.Enabled = True
 cboMonth.Enabled = True
 cboPlatform.Enabled = True
End Sub
Private Sub ClearF()
 txtCDQTY.Text = ""
 txtDescription.Text = ""
 txtTitle.Text = ""
 txtVersion.Text = ""
 cboCDCode.Clear
 cboYear.Text = ""
 cboMonth.Text = ""
 cboPlatform.Text = ""
End Sub
Private Sub cboCDCode_click()
If isNew = False Then
Call Display
End If
End Sub

Private Sub cboCDCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If isNew = True And validate = False Then
      cboMonth.SetFocus
  ElseIf validate = True Then
      MsgBox "DISC CODE Already Exists"
      cboCDCode.SetFocus
  End If
End If
End Sub

Private Sub cboMonth_Click()
cboYear.SetFocus
End Sub
Private Sub cboMonth_GotFocus()
If cboCDCode.Text = "" Then
   MsgBox "please specify CD-ACCESS CODE "
   cboCDCode.SetFocus
End If
End Sub
Private Sub cboMonth_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cboYear.SetFocus
End If
End Sub
Private Sub cboPlatform_Click()
If Not cboPlatform.Text = "" Then
txtCDQTY.SetFocus
Else
   MsgBox "Specify the OS Platform"
   cboPlatform.SetFocus
End If
End Sub
Private Sub cboPlatform_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not cboPlatform.Text = "" Then
txtCDQTY.SetFocus
Else
   MsgBox "Specify the OS Platform"
   cboPlatform.SetFocus
End If
End If
End Sub
Private Sub cboYear_Click()
txtTitle.SetFocus
End Sub
Private Sub cboYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtTitle.SetFocus
End If
End Sub
Private Sub cmdCancel_Click()
ans = MsgBox("Are You Sure ?," & vbCrLf & "Do you want to Cancel this job?", vbQuestion + vbYesNo, "Sure????")
If ans = vbYes Then
    If isNew = True Then
        ClearF
        DisableF
        isNew = False
        cmdNew.Enabled = True
        cmdEdit.Enabled = True
    ElseIf isNew = False Then
        ClearF
        DisableF
        isNew = False
        cmdEdit.Enabled = True
        cmdNew.Enabled = True
    End If
    Call PopulateCDCode
End If
End Sub
Private Sub cmdDelete_Click()
ans = MsgBox("Are You Sure ? ", vbExclamation + vbYesNo, "Are You Sure?")
If ans = vbYes Then
    sql1 = "delete from CDCol where ACCESSCODE = '" & UCase(Trim(cboCDCode.Text)) & "' "
    conn.Execute (sql1)
    Call CommitDB
    MsgBox txtTitle.Text & "Deleted successfully"
Else
    MsgBox " Deletion Cancelled"
End If
Call ClearF
End Sub
Private Sub cmdEdit_Click()
isNew = False
Call EnableF
cmdEdit.Enabled = False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
Call ClearF
Call EnableF
If rsCDCol.State = 0 Then rsCDCol.Open
'If Not rsCDCol.EOF Then
'rsCDCol.MoveLast
txtSNO.Caption = rsCDCol.RecordCount + 1
isNew = True
cmdNew.Enabled = False
cmdEdit.Enabled = False
End Sub

Private Sub cmdSave_Click()
If isNew = True And validate = False Then
   rsCDCol.AddNew
'   rsCDCol!SNO = txtSNO.Caption
   rsCDCol!CDDESCRIPTION = Trim(txtDescription.Text)
   rsCDCol!CDTITLE = Trim(txtTitle.Text)
   rsCDCol!NOOFCDS = Val(txtCDQTY.Text)
   rsCDCol!CDTYPE = UCase(Trim(lstCDMODE.Text))
   rsCDCol!ACCESSCODE = UCase(Trim(cboCDCode.Text))
   rsCDCol!Version = Trim(txtVersion.Text)
   rsCDCol!Year = Trim(cboYear.Text)
   rsCDCol!Month = Trim(cboMonth.Text)
   rsCDCol!platform = Trim(cboPlatform.Text)
   rsCDCol.Update
   Call CommitDB
   MsgBox "New CDRom Details are ADDEDED"
   Call ClearF
   Call DisableF
   isNew = False
   cmdNew.Enabled = True
   cmdEdit.Enabled = True
   cmdNew.SetFocus
ElseIf isNew = False Then
        Call CommitDB
       If rsCDCol.EOF Or rsCDCol.BOF Then
          rsCDCol.MoveFirst
       End If
       While Not rsCDCol.EOF
          If rsCDCol!ACCESSCODE = UCase(Trim(cboCDCode.Text)) Then
              Call SaveMe
              Exit Sub
          Else
              rsCDCol.MoveNext
          End If
       Wend
       
'   sql = "update CDCol set CDTITLE = '" & Trim(txtTitle.Text) & "'," & _
'         "CDDESCRIPTION = '" & Trim(txtDescription.Text) & "'," & _
'         "NOOFCDS = '" & Val(txtCDQTY.Text) & "',VERSION = '" & Trim(txtVersion.Text) & "', " & _
'         "YEAR = '" & Trim(cboYear.Text) & "' , " & _
'         "MONTH = '" & Trim(cboMonth.Text) & "' " & _
'         "where ACCESSCODE = '" & UCase(Trim(cboCDCode.Text)) & "'; "
'   conn.Execute (sql)
    
    
End If
Call PopulateCDCode
End Sub

Private Sub Form_Load()
isNew = False
Me.Top = 20
Me.Left = 2500
Call DisableF
End Sub

Private Sub lstCDMODE_Click()
If isNew = False Then
Call PopulateCDCode
End If
End Sub

Public Sub Display()
   If rs.EOF Or rs.BOF Then
         rs.MoveFirst
     While Not rs.EOF
     If Trim(UCase(cboCDCode.Text)) = rs!ACCESSCODE Then
        txtSNO.Caption = rs!sno
        txtDescription.Text = rs!CDDESCRIPTION
        txtTitle.Text = rs!CDTITLE
        txtCDQTY.Text = rs!NOOFCDS
        txtVersion.Text = rs!Version
        cboMonth.Text = rs!Month
        cboYear.Text = rs!Year
        cboPlatform.Text = rs!platform
     End If
     rs.MoveNext
     Wend
   End If
End Sub

Public Sub PopulateCDCode()
cboCDCode.Clear
sql = " select * from CDCol where CDTYPE = '" & UCase(Trim(lstCDMODE.Text)) & "' "
Set rs = conn.Execute(sql)
If Not rs.EOF Then
   rs.MoveFirst
   While Not rs.EOF
      cboCDCode.AddItem rs!ACCESSCODE
      rs.MoveNext
   Wend
   cboCDCode.SetFocus
Else
   MsgBox "Item Doesn't exists"
End If
End Sub

Private Sub txtCDQTY_GotFocus()
If cboPlatform.Text = "" Then
    MsgBox "plz specify OS Platform"
    cboPlatform.SetFocus
End If
End Sub

Private Sub txtCDQTY_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtVersion.SetFocus
End If
End Sub

Private Sub txtDescription_GotFocus()
If Not txtTitle.Text = "" Then
    txtDescription.SetFocus
Else
     MsgBox "Title Field is Left Empty"
     txtTitle.SetFocus
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboPlatform.SetFocus
End If
End Sub

Private Sub txtTitle_GotFocus()
  If isNew = True And validate = False Then
      txtTitle.SetFocus
  ElseIf validate = True Then
      MsgBox "DISC CODE Already Exists"
      cboCDCode.SetFocus
  End If
End Sub

Public Function validate() As Boolean
Call CommitDB
If isNew = True Then
   validate = False
   If Not rsCDCol.RecordCount < 1 Then
        If rsCDCol.BOF Or rsCDCol.EOF Then
            rsCDCol.MoveFirst
        End If
        While Not rsCDCol.EOF
            If rsCDCol!ACCESSCODE = UCase(Trim(cboCDCode.Text)) Then
                validate = True
            End If
            rsCDCol.MoveNext
        Wend
    End If
End If
End Function

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Not txtTitle.Text = "" Then
    txtDescription.SetFocus
  Else
     MsgBox "Title Field is Left Empty"
     txtTitle.SetFocus
  End If
End If
End Sub


Private Sub txtVersion_GotFocus()
If txtCDQTY.Text = "" Then
    MsgBox "Enter No. of CDRoms"
    txtCDQTY.SetFocus
End If
End Sub

Private Sub txtVersion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdSave.SetFocus
End If
End Sub
Public Sub SaveMe()
       rsCDCol!ACCESSCODE = UCase(Trim(cboCDCode.Text))
       rsCDCol!CDTITLE = Trim(txtTitle.Text)
       rsCDCol!CDDESCRIPTION = Trim(txtDescription.Text)
       rsCDCol!CDTYPE = Trim(UCase(lstCDMODE.Text))
       rsCDCol!Version = Trim(txtVersion.Text)
       rsCDCol!Year = Trim(cboYear.Text)
       rsCDCol!Month = Trim(cboMonth.Text)
       rsCDCol!NOOFCDS = Val(txtCDQTY.Text)
       rsCDCol!platform = Trim(cboPlatform.Text)
       rsCDCol.Update
       Call CommitDB
       MsgBox "Updation Success"
       Call DisableF
       Call ClearF
       cmdEdit.Enabled = True
       cmdNew.Enabled = True
       Call PopulateCDCode
End Sub


