VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A9A48D8D-D1E0-11D4-B90B-444553540000}#74.1#0"; "RegCtl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000D&
   Caption         =   "My Personal CDManager"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":08CA
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":378AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38186
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":38EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39304
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A1DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C318
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E6FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1376
      ButtonWidth     =   2725
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&CDRom Manager"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Borrow Manager"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Search 'n' Find"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CDRom &List"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Phone Dialer"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "S.O.S - &Help Me"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Get Out from Here"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   15293
            MinWidth        =   5292
            Picture         =   "frmMain.frx":3EA14
            Text            =   "Created By:  Nithin Mohan.T.K      for    © 2002 - 2005  Dream Works Technologies India Ltd"
            TextSave        =   "Created By:  Nithin Mohan.T.K      for    © 2002 - 2005  Dream Works Technologies India Ltd"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   "16/11/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            TextSave        =   "10:31 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Registry.Reg RegAccess 
      Left            =   4800
      Top             =   3240
      _ExtentX        =   979
      _ExtentY        =   450
      Hkey            =   1
      ErrorReturn     =   0
   End
   Begin VB.Menu cmdFile 
      Caption         =   "&File"
      Begin VB.Menu muFileWizard 
         Caption         =   "Run Wizard"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnusearchCDs 
         Caption         =   "Search For CDRoms"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu cmdEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu cmdManager 
      Caption         =   "&Manager"
      Begin VB.Menu mnuCDManager 
         Caption         =   "CD Manager"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuBorrowManage 
         Caption         =   "Borrow Manager"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu cmdMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuStockClear 
         Caption         =   "Clear Database"
         Shortcut        =   ^Z
      End
      Begin VB.Menu f 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileUserChangePass 
         Caption         =   "Change Password"
         Shortcut        =   +{F2}
      End
   End
   Begin VB.Menu cmdUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuUtilitiesNotepad 
         Caption         =   "Notepad"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuUtilitiesWordpad 
         Caption         =   "WordPad"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuUtilitiesCalculator 
         Caption         =   "Calculator"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuUtilitiesPaint 
         Caption         =   "Paint"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuUtilitiesWindowsExplorer 
         Caption         =   "Windows Explorer"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuUtilitiesGames 
         Caption         =   "&Games"
         Begin VB.Menu mnuUtilitiesGamesFreeCell 
            Caption         =   "FreeCell"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu mnuUtilitiesGamesMineSweeper 
            Caption         =   "MineSweeper"
            Shortcut        =   {F7}
         End
         Begin VB.Menu mnuUtilitiesGamesPinBall 
            Caption         =   "PinBall"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuUtilitiesGamesSolitaire 
            Caption         =   "Solitaire"
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu mnuUtilitiesWMP 
         Caption         =   "Windows Media Player"
      End
   End
   Begin VB.Menu cmdHelp 
      Caption         =   "&Help??"
      Begin VB.Menu cmdRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help?"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Us"
         Shortcut        =   +{F12}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i






Private Sub cmdRegister_Click()
called_by = True
Load frmRegister
frmRegister.Show
End Sub


Private Sub MDIForm_Load()
Call SaveSettings
Call LoadSettings
isList = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
SaveSettings
End
End Sub

Private Sub mnuCDManager_Click()
Load frmCDMan
frmCDMan.Show
End Sub

Private Sub mnuFileExit_Click()

ans = MsgBox("Are you really want to exit from " & App.ProductName & " v" & App.Major & "." & App.Minor & "- Build " & App.Revision, vbQuestion + vbYesNo, "Exit " & App.ProductName & " ???")
If ans = vbYes Then
    try_day = try_day + 1
    Call SaveSettings
    MsgBox "Thanks for Using " & App.ProductName & " v" & App.Major & "." & App.Minor & "- Build " & App.Revision, vbInformation + vbOKOnly
    Call disconnectDB
    End
End If
End Sub

Private Sub mnuFileUserChangePass_Click()
Load frmChgPass
frmChgPass.Show
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub
Private Sub mnusearchCDs_Click()
Load frmFind
frmFind.Show
End Sub

Private Sub mnuStockClear_Click()
If LCase(Trim(loginuser)) = "admin" Then
    Load frmConfirmPass
    frmConfirmPass.Show
Else
   MsgBox "ONLY Admin is allowed to Clear Database"
End If
End Sub

Private Sub mnuUtilitiesWordpad_Click()
Dim res As Double
res = Shell("C:\Program Files\Windows NT\Accessories\wordpad.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesNotepad_Click()
Dim res As Double
res = Shell("notepad.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesCalculator_Click()
Dim res As Double
res = Shell("calc.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesWindowsExplorer_Click()
Dim res As Double
res = Shell("explorer.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesWMP_Click()
Dim res As Double
res = Shell("C:\Program Files\Windows Media Player\wmplayer.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesPaint_Click()
Dim res As Double
res = Shell("mspaint.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesGamesFreeCell_Click()
Dim res As Double
res = Shell("freecell.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesGamesMineSweeper_Click()
Dim res As Double
res = Shell("winmine.exe", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesGamesPinBall_Click()
Dim res As Double
res = Shell("C:\Program Files\Windows NT\Pinball\PINBALL.EXE", vbNormalFocus)
End Sub
Private Sub mnuUtilitiesGamesSolitaire_Click()
Dim res As Double
res = Shell("sol.EXE", vbNormalFocus)
End Sub
Private Sub muFileWizard_Click()
Load frmWizard
frmWizard.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(1).Text = "Hello hello hello"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
       Case 1
              Load frmCDMan
              frmCDMan.Show
       Case 3
              Load frmBorrow
              frmBorrow.Show
       Case 5
              Load frmFind
              frmFind.Show
       Case 7
              Load frmReport
              frmReport.Show
       Case 9
              Load frmDialer
              frmDialer.Show
       Case 11
       Case 13
              
              ans = MsgBox("Are you really want to exit from " & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision, vbQuestion + vbYesNo, "Exit " & App.ProductName & " ???")
              If ans = vbYes Then
                 try_day = try_day + 1
                 Call SaveSettings
                 MsgBox "Thanks for Using " & App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision, vbInformation + vbOKOnly
                 Call disconnectDB
                 End
              End If
 End Select
        
End Sub
