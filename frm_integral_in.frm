VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_integral_in 
   BorderStyle     =   0  'None
   Caption         =   "员工积分录入"
   ClientHeight    =   8100
   ClientLeft      =   1125
   ClientTop       =   1965
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   5430
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   7650
      Width           =   2835
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   3330
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   7650
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "项目列表"
      Height          =   6465
      Left            =   8910
      TabIndex        =   19
      Top             =   630
      Width           =   4995
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh_list 
         Height          =   6195
         Left            =   30
         TabIndex        =   20
         Top             =   210
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   10927
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   3
         FormatString    =   "<编号       |<名称                    |<分值           "
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   3330
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   7290
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   690
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7650
      Width           =   1245
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出(&E)"
      Height          =   405
      Left            =   12630
      TabIndex        =   7
      Top             =   7620
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除(&D)"
      Height          =   405
      Left            =   11211
      TabIndex        =   6
      Top             =   7620
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   9794
      TabIndex        =   5
      Top             =   7620
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存(&S)"
      Height          =   405
      Left            =   8377
      TabIndex        =   4
      Top             =   7620
      Width           =   1245
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   6645
      Left            =   30
      TabIndex        =   8
      Top             =   540
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "<序号    |<项目编号       |<项目名称            |<加减分值    |<备注                                 "
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label13 
      Caption         =   "备注："
      Height          =   225
      Left            =   5430
      TabIndex        =   23
      Top             =   7350
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Label12"
      Height          =   180
      Left            =   4620
      TabIndex        =   22
      Top             =   7365
      Width           =   630
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "分值："
      Height          =   180
      Left            =   2760
      TabIndex        =   21
      Top             =   7725
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "员工积分录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6052
      TabIndex        =   18
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Label9"
      Height          =   180
      Left            =   13020
      TabIndex        =   17
      Top             =   7290
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "积分合计："
      Height          =   180
      Left            =   12120
      TabIndex        =   16
      Top             =   7290
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   180
      Left            =   10530
      TabIndex        =   15
      Top             =   7290
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "项目数："
      Height          =   180
      Left            =   9870
      TabIndex        =   14
      Top             =   7290
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   180
      Left            =   1980
      TabIndex        =   13
      Top             =   7732
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "项目："
      Height          =   180
      Left            =   2760
      TabIndex        =   12
      Top             =   7365
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "员工："
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   7732
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   720
      TabIndex        =   10
      Top             =   7290
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "单号："
      Height          =   180
      Left            =   150
      TabIndex        =   9
      Top             =   7290
      Width           =   540
   End
End
Attribute VB_Name = "frm_integral_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    On Error GoTo Err
    Dim i As Integer
    If Trim(msh.TextMatrix(msh.Rows - 1, 0)) <> "" Then
        con.BeginTrans
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from employee where bh='" & Trim(Text1.Text) & "'", con, adOpenDynamic, adLockOptimistic
        rs.Fields("jf") = rs.Fields("jf") + Label9.Caption
        rs.Update
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from integral_detail where 1<>1", con, adOpenDynamic, adLockOptimistic
        For i = 1 To msh.Rows - 1
            rs.AddNew
            rs.Fields("dh") = Trim(Label2.Caption)
            rs.Fields("xm_bh") = Trim(msh.TextMatrix(i, 1))
            rs.Fields("xm_mc") = Trim(msh.TextMatrix(i, 2))
            rs.Fields("bc_jf") = msh.TextMatrix(i, 3)
            rs.Fields("date_time") = Format(Now, "yyyy-MM-dd HH:mm:ss")
            rs.Fields("yg_bh") = Trim(Text1.Text)
            rs.Fields("yg_xm") = Trim(Label5.Caption)
            rs.Fields("czy_bh") = Trim(login_bh)
            rs.Fields("czy_xm") = Trim(login_xm)
            rs.Fields("memo") = Trim(Text4.Text)
            rs.Update
        Next i
        con.CommitTrans
        Command2_Click
        Exit Sub
    Else
        Text2.SetFocus
        Exit Sub
    End If
Err:
    MsgBox Err.Description, vbOKOnly, "出错了"
    con.RollbackTrans
End Sub

Private Sub Command2_Click()
    msh.Clear
    msh.FormatString = "<序号    |<项目编号       |<项目名称            |<加减分值    |<备注                                 "
    msh.Rows = 2
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select distinct dh as ds from integral_detail", con, adOpenDynamic, adLockReadOnly
    If rs.RecordCount <= 0 Then
        Label2.Caption = Format(Now, "yyyyMMdd") & "001"
    Else
        Label2.Caption = Format(Now, "yyyyMMdd") & Format(rs.RecordCount + 1, "000")
    End If
    Text1.Text = ""
    Text2.Text = ""
    Text1.Enabled = True
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Label5.Caption = ""
    Label7.Caption = "0"
    Label9.Caption = "0"
    Label12.Caption = ""
    Text3.Text = ""
    Text4.Text = ""
    Command1.Enabled = False
    If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    If msh.Rows > 2 Then
        msh.RemoveItem msh.row
    Else
        For i = 0 To msh.Cols - 1
            msh.TextMatrix(1, i) = ""
        Next i
    End If
    jfhj
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from integral_list order by xm_bh", con, adOpenKeyset, adLockReadOnly
    If rs.RecordCount <= 1 Then
        msh_list.Rows = 2
    Else
        msh_list.Rows = rs.RecordCount + 1
    End If
    For i = 1 To rs.RecordCount
        msh_list.TextMatrix(i, 0) = Trim(rs.Fields("xm_bh"))
        msh_list.TextMatrix(i, 1) = Trim(rs.Fields("xm_mc"))
        msh_list.TextMatrix(i, 2) = Trim(rs.Fields("fz"))
        If Not rs.EOF Then rs.MoveNext
    Next i
    Command2_Click
    Me.Top = frm_main.Top + (frm_main.Height - Me.Height) / 2 + 260
    Me.Left = frm_main.Left + (frm_main.Width - Me.Width) / 2 - 10
    Me.Show 1, frm_main
End Sub


Private Sub MSH_Click()
    If Text1.Enabled = True Then Text1.SetFocus
    If Text2.Enabled = True Then Text2.SetFocus
End Sub

Private Sub msh_list_Click()
    If Text1.Enabled = True Then Text1.SetFocus
    If Text2.Enabled = True Then Text2.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(Text1.Text) <> "" Then
            If rs.State = 1 Then rs.Close
            rs.CursorLocation = adUseClient
            rs.Open "select bh,xm,jf from employee where bh='" & Trim(Text1.Text) & "'", con, adOpenKeyset, adLockReadOnly
            If rs.RecordCount > 0 Then
                Label5.Caption = Trim(rs.Fields("xm"))
                Text1.Enabled = False
                Text2.Enabled = True
                'Command1.Enabled = True
                Text2.SetFocus
            Else
                MsgBox "员工不存在！", vbOKOnly Or vbInformation, "提示"
                Text1.Text = ""
                Text1.SetFocus
            End If
        End If
    End If
End Sub

'计算积分合计
Private Sub jfhj()
    Dim i As Integer
    Label9.Caption = "0"
    For i = 1 To msh.Rows - 1
        Label9.Caption = Val(Label9.Caption) + Val(msh.TextMatrix(i, 3))
    Next i
    
    Label7.Caption = msh.Rows - 1
    
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim row As Integer
        Dim i As Integer
        Dim pd As Boolean
        pd = False
        
        If Trim(Text2.Text) <> "" Then
            For i = 1 To msh_list.Rows - 1
                If Trim(Text2.Text) = msh_list.TextMatrix(i, 0) Then
                    pd = True
                    Label12.Caption = msh_list.TextMatrix(i, 1)
                    Text3.Text = msh_list.TextMatrix(i, 2)
                    Text3.Enabled = True
                    Text2.Enabled = False
                    Text3.SetFocus
                    Exit For
                End If
            Next i
            'jfhj
            If pd = False Then
                MsgBox "项目不存在！", vbOKOnly Or vbInformation, "提示"
                Text2.Text = ""
                Text2.SetFocus
            End If
        Else
            Command1.Enabled = True
            Command1.SetFocus
        End If
    End If
End Sub


Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If IsNumeric(Text3.Text) Then
            Text3.Enabled = False
            Text4.Enabled = True
            Text4.SetFocus
        Else

        End If
    End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim row As Integer
        If msh.TextMatrix(msh.Rows - 1, 0) <> "" Then
            msh.Rows = msh.Rows + 1
        End If
        row = msh.Rows - 1
        msh.TextMatrix(row, 0) = row
        msh.TextMatrix(row, 1) = Trim(Text2.Text)
        msh.TextMatrix(row, 2) = Trim(Label12.Caption)
        msh.TextMatrix(row, 3) = Trim(Text3.Text)
        msh.TextMatrix(row, 4) = Trim(Text4.Text)
        jfhj
        Text2.Text = ""
        Label12.Caption = ""
        Text3.Text = ""
        Text4.Text = ""
        Text2.Enabled = True
        Text4.Enabled = False
        Text2.SetFocus
    End If
End Sub

