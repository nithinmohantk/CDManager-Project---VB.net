VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmFind 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEarch for CDRoms"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11085
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11085
   Begin VB.ListBox lstCDMODE 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "frmFind.frx":0442
      Left            =   2760
      List            =   "frmFind.frx":0464
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00808080&
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Search"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.ListBox lstField 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "frmFind.frx":04AB
      Left            =   5520
      List            =   "frmFind.frx":04C7
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8400
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CDCol"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   9840
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   9840
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   9840
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H000080FF&
      Height          =   1935
      Left            =   9480
      TabIndex        =   18
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H000080FF&
      Height          =   615
      Left            =   8280
      TabIndex        =   17
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackColor       =   &H000040C0&
      Height          =   1935
      Left            =   8280
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   2880
      Width           =   7095
   End
   Begin VB.Label Label9 
      BackColor       =   &H000040C0&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2640
      Width           =   7095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "** Enter the Search String"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH By :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH In :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find What?"
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
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Height          =   1935
      Left            =   1200
      TabIndex        =   13
      Top             =   720
      Width           =   7095
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Height          =   2535
      Left            =   600
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H000040C0&
      Height          =   2535
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdSearch_Click()
MSFlexGrid1.Clear
    If lstCDMODE.Text = "All(Default)" Then
       If lstField.Text = "All(Default)" Then
          sql1 = "select * from CDCol where CDTITLE like '%" & Trim(txtSearch.Text) & "%' " & _
              "or CDDESCRIPTION like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "TITLE" Then
              sql1 = "select * from CDCol where CDTITLE like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "DESCRIPTION" Then
               sql1 = "select * from CDCol where CDDESCRIPTION like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "DISCID" Then
              sql1 = "select * from CDCol where ACCESSCODE like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "MONTH" Then
              sql1 = "select * from CDCol where MONTH like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "YEAR" Then
              sql1 = "select * from CDCol where YEAR like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "VERSION" Then
              sql1 = "select * from CDCol where VERSION like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "PLATFORM" Then
              sql1 = "select * from CDCol where PLATFORM like '%" & Trim(txtSearch.Text) & "%' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       End If
    Else
        If lstField.Text = "All(Default)" Then
          sql1 = "select * from CDCol where CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' and " & _
                 "CDTITLE like '%" & Trim(txtSearch.Text) & "%' " & _
                 " order by CDType ASC" & _
                 ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "TITLE" Then
              sql1 = "select * from CDCol where CDTITLE like '%" & Trim(txtSearch.Text) & "%' " & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "DESCRIPTION" Then
               sql1 = "select * from CDCol where CDDESCRIPTION like '%" & Trim(txtSearch.Text) & "%'" & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "DISCID" Then
              sql1 = "select * from CDCol where ACCESSCODE like '%" & Trim(txtSearch.Text) & "%'" & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "MONTH" Then
              sql1 = "select * from CDCol where MONTH like '%" & Trim(txtSearch.Text) & "%'" & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "YEAR" Then
              sql1 = "select * from CDCol where YEAR like '%" & Trim(txtSearch.Text) & "%'" & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "VERSION" Then
              sql1 = "select * from CDCol where VERSION like '%" & Trim(txtSearch.Text) & "%'" & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       ElseIf UCase(Trim(lstField.Text)) = "PLATFORM" Then
              sql1 = "select * from CDCol where PLATFORM like '%" & Trim(txtSearch.Text) & "%'" & _
                    " And CDTYPE = '" & Trim(UCase(lstCDMODE.Text)) & "' order by CDType ASC" & _
              ",CDTITLE asc,ACCESSCODE asc"
       End If
              
    End If
    Set rs = conn.Execute(sql1)
    
If Not rs.EOF Then
     rs.MoveFirst
     Dim i As Integer
     i = 1
     Call loadheader
     While Not rs.EOF
      MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
      MSFlexGrid1.TextMatrix(i, 0) = Trim(rs!CDTYPE)
      MSFlexGrid1.TextMatrix(i, 1) = Trim(rs!CDTITLE)
      MSFlexGrid1.TextMatrix(i, 2) = Trim(rs!CDDESCRIPTION)
      MSFlexGrid1.TextMatrix(i, 3) = Trim(rs!ACCESSCODE)
      MSFlexGrid1.TextMatrix(i, 4) = Trim(rs!Version)
      MSFlexGrid1.TextMatrix(i, 5) = Trim(rs!NOOFCDS)
      MSFlexGrid1.TextMatrix(i, 6) = Trim(rs!Month)
      MSFlexGrid1.TextMatrix(i, 7) = Trim(rs!Year)
      i = i + 1
      
      rs.MoveNext
     Wend
Else
   MsgBox "SORRY !" & vbCrLf & "Item Not Found"
End If
End Sub

Private Sub cmdSearch_GotFocus()
If lstCDMODE.Text = "" Then
   MsgBox "Plz select CDType"
   lstCDMODE.SetFocus
End If
End Sub

Private Sub optEvery_Click()
If optEvery.Value = Checked Then
    every = True
ElseIf optEvery.Value = Unchecked Then
    every = False
End If
End Sub

Private Sub Form_Load()
lstCDMODE.Selected(0) = True
lstField.Selected(0) = True
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 2
Call loadheader
Me.Top = 50
Me.Left = 500
End Sub

Private Sub lstCDMODE_GotFocus()
If txtSearch.Text = "" Then
     MsgBox "please enter a search string"
     txtSearch.SetFocus
  Else
     lstCDMODE.SetFocus
  End If
End Sub

Private Sub lstCDMODE_KeyPress(KeyAscii As Integer)
lstField.SetFocus
End Sub
Private Sub lstField_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdSearch.SetFocus
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If txtSearch.Text = "" Then
     MsgBox "please enter a search string"
     txtSearch.SetFocus
  Else
     lstCDMODE.SetFocus
  End If
End If
End Sub

Public Sub loadheader()
MSFlexGrid1.Clear
MSFlexGrid1.Rows = 2
MSFlexGrid1.TextMatrix(0, 0) = "CATEGORY"
MSFlexGrid1.TextMatrix(0, 1) = "TITLE"
MSFlexGrid1.TextMatrix(0, 2) = "DESCRIPTION"
MSFlexGrid1.TextMatrix(0, 3) = "ACCESS CODE"
MSFlexGrid1.TextMatrix(0, 4) = "VERSION"
MSFlexGrid1.TextMatrix(0, 5) = "CD-QTY"
MSFlexGrid1.TextMatrix(0, 6) = "MONTH"
MSFlexGrid1.TextMatrix(0, 7) = "YEAR"
MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 3000
MSFlexGrid1.ColWidth(2) = 2500
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 700
MSFlexGrid1.ColWidth(6) = 670
MSFlexGrid1.ColWidth(7) = 550
End Sub
