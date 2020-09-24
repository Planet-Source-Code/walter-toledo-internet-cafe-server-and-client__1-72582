Attribute VB_Name = "Konek"
Public con As New Connection
Public rs As New Recordset
Public SQL As String

Public Sub Connect()
Dim xes, dbase
xes = "WIT"
dbase = "bill2009"
On Error GoTo HIM
If con.State = 1 Then con.Close
con.CursorLocation = adUseClient
'con.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source='" & xes & "';Initial Catalog='" & dbase & "'"
con.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=root;Data Source=WIT;Initial Catalog=bill2009"
Exit Sub
HIM:
MsgBox Err.Description, vbExclamation, frmMenu.Caption
End Sub

Public Sub UnloadAllForms()
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub


