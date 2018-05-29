Attribute VB_Name = "modSub"
Option Explicit
Public Function confirm_pass(ByVal userpass As String) As Boolean
 Dim rsa As New ADODB.Recordset
 If rsa.State = 1 Then rsa.Close
 rsa.Open "login", conn, adOpenDynamic, adLockOptimistic
 rsa.Find "loginid = 'admin'"
 If decrypt_pass(rsa!Password) = userpass Then
     confirm_pass = True
 Else
     confirm_pass = False
 End If
End Function

Public Sub del_all()
    ans = MsgBox("Are you Sure to Clear Database ???", vbQuestion + vbYesNo, "Are You Sure?")
    If ans = vbYes Then
        sql = "delete from CDCol"
        conn.Execute sql
        sql = "delete from Login"
        conn.Execute sql
        sql = "delete from bdata"
        conn.Execute sql
        sql = "delete from borrow"
        conn.Execute sql, recaff
        Call CommitDB
        If recaff > 0 Then
            MsgBox "Database is Cleared Successfully ! "
        End If
     End If
End Sub

