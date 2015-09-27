VERSION 5.00
Begin VB.Form frm_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登录"
   ClientHeight    =   3720
   ClientLeft      =   7695
   ClientTop       =   4680
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5400
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   30
      TabIndex        =   6
      Top             =   -60
      Visible         =   0   'False
      Width           =   5355
      Begin VB.CommandButton Command4 
         Caption         =   "返回(&E)"
         Height          =   375
         Left            =   3780
         TabIndex        =   16
         Top             =   2970
         Width           =   945
      End
      Begin VB.CommandButton Command3 
         Caption         =   "保存(&S)"
         Height          =   375
         Left            =   2550
         TabIndex        =   15
         Top             =   2970
         Width           =   945
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frm_login.frx":0000
         Left            =   1320
         List            =   "frm_login.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   3435
      End
      Begin VB.TextBox Text5 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   2190
         Width           =   3435
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   1560
         Width           =   3435
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   960
         Width           =   3435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         Height          =   180
         Left            =   690
         TabIndex        =   10
         Top             =   2220
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "登录用户名："
         Height          =   180
         Left            =   150
         TabIndex        =   9
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "服务器IP："
         Height          =   180
         Left            =   330
         TabIndex        =   8
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "运行模式："
         Height          =   180
         Left            =   330
         TabIndex        =   7
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1410
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   570
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户密码："
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户编号："
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As New ADODB.Recordset
Dim conl As New ADODB.Connection

Private Sub Combo1_Click()
    If conl.State = 1 Then conl.Close
    conl.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\data.mdb;Persist Security Info=False;Jet OLEDB:Database password=15615252266"
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from model", conl, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then
        Text3.Text = Trim(rs.Fields("ip"))
        Text4.Text = Trim(rs.Fields("login"))
        Text5.Text = Trim(rs.Fields("pwd"))
    End If
    If Combo1.Text = "Access" Then
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Text3.Visible = False
        Text4.Visible = False
        Text5.Visible = False
    Else
        Label4.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        Text3.Visible = True
        Text4.Visible = True
        Text5.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from login where bh='" & Trim(Text1.Text) & "' and mm='" & Trim(Text2.Text) & "'", con, adOpenDynamic, adLockOptimistic
    If rs.RecordCount > 0 Then
        login_bh = Trim(rs.Fields("bh"))
        login_pwd = Trim(rs.Fields("mm"))
        login_xm = Trim(rs.Fields("xm"))
        Load frm_main
        Me.Hide
    Else
        MsgBox "用户编号或密码不正确请检查", vbOKOnly Or vbInformation, "提示"
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    If conl.State = 1 Then conl.Close
    conl.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\data.mdb;Persist Security Info=False;Jet OLEDB:Database password=15615252266"
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from model", conl, adOpenDynamic, adLockOptimistic
    If Combo1.Text = "Access" Then
        rs.Fields("acc") = "1"
    Else
        rs.Fields("acc") = "0"
    End If
    rs.Fields("ip") = Trim(Text3.Text)
    rs.Fields("login") = Trim(Text4.Text)
    rs.Fields("pwd") = Trim(Text5.Text)
    rs.Update
    MsgBox "设置完毕，请重启程序", vbOKOnly, "提示"
    End
End Sub

Private Sub Command4_Click()
    Frame1.Visible = False
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Text1.Text = ""
    Text2.Text = ""
    If conl.State = 1 Then conl.Close
    conl.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data\data.mdb;Persist Security Info=False;Jet OLEDB:Database password=15615252266"
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from model", conl, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then
        If Trim(rs.Fields("acc")) = "1" Then
            Combo1.Text = "Access"
        Else
            Combo1.Text = "Sql2000"
        End If
        Text3.Text = Trim(rs.Fields("ip"))
        Text4.Text = Trim(rs.Fields("login"))
        Text5.Text = Trim(rs.Fields("pwd"))
    End If
    If Combo1.Text = "Access" Then
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Text3.Visible = False
        Text4.Visible = False
        Text5.Visible = False
    Else
        Label4.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        Text3.Visible = True
        Text4.Visible = True
        Text5.Visible = True
    End If
    Me.Show
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 13 Then
        Frame1.Visible = True
    ElseIf KeyCode = 13 Then
        If KeyCode = 13 Then
            SendKeys "{tab}"
        End If
    End If
    'If Trim(Text1.Text) <> "" And KeyCode = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(Text2.Text) <> "" And KeyCode = 13 Then Command1.SetFocus
End Sub
