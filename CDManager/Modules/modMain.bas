Attribute VB_Name = "modMain"
Option Explicit
Public first_reg As Boolean
Public conn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rstemp As New ADODB.Recordset
Public rsLogin As New ADODB.Recordset
Public rsCDCol As New ADODB.Recordset
Public mesg As String
Public pass_changed As Boolean
Public loginuser As String
Public loginpass As String
Public newpass1 As String
Public newpass2 As String
Public recaff As Integer
Public sql As String
Public connString As String
Public isList As Boolean
Public called_by As Boolean

Public Sub ConnectDB()
connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\CDMANDB.mdb;Mode=Read|Write|Share Deny None;Persist Security Info=False;User ID=admin"
conn.Open (connString)
conn.BeginTrans
rsCDCol.CursorLocation = adUseClient
rsCDCol.Open "[CDCol]", conn, adOpenKeyset, adLockOptimistic, adCmdTable
rsLogin.Open "[LOGIN]", conn, adOpenKeyset, adLockOptimistic, adCmdTable
End Sub
Public Sub disconnectDB()
If conn.State = 1 Then
   conn.CommitTrans
   conn.Close
End If
End Sub

Public Sub CommitDB()
conn.CommitTrans
conn.Close
Call ConnectDB
End Sub
Public Function Encrypt(ByVal strInput As String, ByVal strKey As String) As String
Dim iCount As Long
Dim lngPtr As Long
For iCount = 1 To Len(strInput)
    Mid(strInput, iCount, 1) = Chr((Asc(Mid(strInput, iCount, 1))) Xor (Asc(Mid(strKey, lngPtr + 1, 1))))
    lngPtr = ((lngPtr + 1) Mod Len(strKey))
Next iCount
Encrypt = strInput
End Function
Public Function encrypt_pass(ByVal pass As String) As String
Dim pass1(40) As String
Dim ascii(40) As String
Dim pass2 As String
Dim lenp As Integer
Dim i As Integer
lenp = Len(pass)
i = 0
While i < lenp
  i = i + 1
  pass1(i) = Mid(pass, i, 1)
  ascii(i) = Asc(pass1(i))
  ascii(i) = ascii(i) + (i + (i - 4))
  pass2 = pass2 & Chr(ascii(i))
Wend
encrypt_pass = pass2
End Function
Public Function decrypt_pass(ByVal pass As String) As String
Dim pass1(40) As String
Dim ascii(40) As Integer
Dim pass2 As String
Dim lenp As Integer
Dim i, j, k As Integer
lenp = Len(pass)
i = 1
While i <= lenp
    pass1(i) = Mid(pass, i, 1)
    ascii(i) = Asc(pass1(i))
    ascii(i) = ascii(i) - (i + (i - 4))
    pass2 = pass2 & Chr(ascii(i))
    i = i + 1
Wend
decrypt_pass = pass2
End Function


Public Sub Main()
first_reg = False
Call LoadSettings
Call ConnectDB
If Registered() = True Then
   Load frmSplash
   frmSplash.Show
Else
   called_by = False
   frmRegister.Show
End If
End Sub




