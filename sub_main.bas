Attribute VB_Name = "sub_main"
Private constr As String
Public con As New ADODB.Connection '���ݿ�����
Public login_bh As String '��¼�û����
Public login_xm As String '��¼�û���
Public login_pwd As String '��¼�û�����
Public model As Integer  '����ģʽ

Sub Main()
    On Error GoTo cw
    Dim rs As New ADODB.Recordset
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\data.mdb;Persist Security Info=False;Jet OLEDB:Database password=15615252266"
    con.Open constr
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from model", con, adOpenKeyset, adLockReadOnly
    If rs.Fields("acc") = 0 Then
        constr = "Provider=SQLOLEDB.1;Password=" & Trim(rs.Fields("pwd")) & ";Persist Security Info=True;User ID=" & Trim(rs.Fields("login")) & ";Initial Catalog=jfgl;Data Source=" & Trim(rs.Fields("ip"))
        If con.State = 1 Then con.Close
        con.Open constr
        frm_login.Combo1.Text = "sql2000"
        model = 0
    Else
        frm_login.Combo1.Text = "access"
        model = 1
    End If
    Load frm_login
    Exit Sub
cw:
    MsgBox Err.Description & vbCrLf & "ϵͳ���Ե���������"
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\data.mdb;Persist Security Info=False;Jet OLEDB:Database password=15615252266"
    con.Open constr
    Load frm_login
End Sub
Public Function qbh(str As String) As String
    If InStr(1, str, "-") - 1 > 0 Then
        qbh = Left(str, InStr(1, str, "-") - 1)
    Else
        qbh = str
    End If
End Function
