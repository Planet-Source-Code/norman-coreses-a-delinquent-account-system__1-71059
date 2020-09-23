Attribute VB_Name = "Module1"
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741





Public Sub connection(ByRef connection As ADODB.connection, ByVal dLocation As String, ByVal pass As String)
'connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dLocation & ";Persist Security Info=False; Jet OLEDB:Database password=" & spass
 connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dLocation & ";Persist Security Info=False; Jet OLEDB:Database password=" & pass

End Sub

Public Sub recordset(ByRef recordset As ADODB.recordset, ByRef connection As ADODB.connection, ByVal ssql As String)
With recordset
    .CursorLocation = adUseClient
    .Open ssql, connection, adOpenKeyset, adLockOptimistic
End With
End Sub



Public Sub hlfocus(ByRef stext As TextBox)
With stext
    .SelStart = 0
    .SelLength = Len(stext.Text)
End With
End Sub

Sub main()
    Form19.Show 1
End Sub

Public Function sempty(ByRef stext As Variant) As Boolean
If stext.Text = "" Then
    sempty = True
    MsgBox "Please field all the field", vbExclamation, "Bank of Paracale"
    stext.SetFocus
Else
     sempty = False
End If
End Function

Public Function snumber(ByRef stext As Variant) As Boolean
If IsNumeric(stext) = False Then
snumber = True
MsgBox "Cannot accept non-numeric.", vbExclamation, "Bank of Paracale"

Else

snumber = False

End If
End Function


Public Function recfound(ByRef sRecordset As ADODB.recordset, ByVal sField As String, ByVal sfindtext As String) As Boolean
    sRecordset.Requery
    sRecordset.Find sField & "='" & sfindtext & "'"
               
If sRecordset.EOF Then
    recfound = False
Else
    recfound = True
    Username = sRecordset.Fields("UserName")
    Password = sRecordset.Fields("Password")
    End If
End Function

'****************************************************************************
'
'
'


'End Sub

Public Sub positionform(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = (Screen.Height - frm.Height) \ 2000
    LeftCorner = (Screen.Width - frm.Width) \ 2
    frm.Move LeftCorner, TopCorner
End Sub


