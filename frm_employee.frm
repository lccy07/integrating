VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_employee 
   Caption         =   "员工信息"
   ClientHeight    =   8055
   ClientLeft      =   1200
   ClientTop       =   1815
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   13935
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4605
      Left            =   3960
      TabIndex        =   18
      Top             =   1620
      Visible         =   0   'False
      Width           =   5265
      Begin VB.TextBox Text5 
         Height          =   345
         Left            =   810
         TabIndex        =   0
         Text            =   "Text5"
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "取消(&C)"
         Height          =   405
         Left            =   3810
         TabIndex        =   8
         Top             =   4050
         Width           =   1155
      End
      Begin VB.CommandButton Command6 
         Caption         =   "保存(&S)"
         Height          =   405
         Left            =   1590
         TabIndex        =   7
         Top             =   4050
         Width           =   1155
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   810
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   3570
         Width           =   4125
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3000
         Width           =   4125
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   810
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   2430
         Width           =   4125
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   810
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1845
         Width           =   4125
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "frm_employee.frx":0000
         Left            =   810
         List            =   "frm_employee.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1290
         Width           =   4125
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   810
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   4125
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "工号："
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "积分："
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   3630
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "分组："
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3060
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "电话："
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   2490
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   780
         Width           =   540
      End
   End
   Begin VB.Frame frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   13935
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frm_employee.frx":0016
         Left            =   9510
         List            =   "frm_employee.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   307
         Width           =   1425
      End
      Begin VB.CommandButton Command5 
         Caption         =   "退出(&X)"
         Height          =   495
         Left            =   7050
         TabIndex        =   15
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "刷新(&R)"
         Height          =   495
         Left            =   5316
         TabIndex        =   14
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "修改(&E)"
         Height          =   495
         Left            =   3584
         TabIndex        =   13
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "删除(&D)"
         Height          =   495
         Left            =   1852
         TabIndex        =   12
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加(&A)"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   180
         Left            =   11580
         TabIndex        =   26
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "排序方式："
         Height          =   180
         Left            =   8520
         TabIndex        =   16
         Top             =   367
         Width           =   900
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   7335
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   12938
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "<序号    |<编号   |<姓名     |<性别   |<地址                        |<电话             |<分组             |<积分           "
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
End
Attribute VB_Name = "frm_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Command1_Click()
    Frame2.Caption = "添加员工信息"
    qc
    Frame2.Visible = True
    msh.Enabled = False
    Frame1.Enabled = False
End Sub

Private Sub Command2_Click()
    If Trim(msh.TextMatrix(msh.row, 0)) <> "" Then
        If rs.State = 1 Then rs.Close
        rs.Open "delete from employee where bh='" & Trim(msh.TextMatrix(msh.row, 1)) & "'", con, adOpenDynamic, adLockOptimistic
        If msh.Rows > 3 Then
            msh.RemoveItem msh.row
        Else
            If msh.TextMatrix(msh.row, 1) <> "" Then
            For i = 0 To msh.Cols - 1
                msh.TextMatrix(1, i) = ""
            Next i
            End If
        End If
    End If
End Sub

Private Sub Command3_Click()
    If Trim(msh.TextMatrix(msh.row, 0)) <> "" And Trim(msh.TextMatrix(msh.row, 0)) <> "合计" Then
        Frame2.Caption = "修改员工信息"
        Frame2.Visible = True
        msh.Enabled = False
        Frame1.Enabled = False
        row = msh.row
        Text5.Text = msh.TextMatrix(row, 1)
        Text1.Text = msh.TextMatrix(row, 2)
        Combo2.Text = msh.TextMatrix(row, 3)
        Text2.Text = msh.TextMatrix(row, 4)
        Text3.Text = msh.TextMatrix(row, 5)
        Combo3.Text = msh.TextMatrix(row, 6)
        Text4.Text = msh.TextMatrix(row, 7)
        Text5.Enabled = False
    End If
End Sub

Private Sub Command4_Click()
    Dim str As String
    Dim hj As Double
    
    str = "select * from employee,tb_group where employee.fz=tb_group.id order by "
    If Combo1.ListIndex = 0 Then
        str = str & "employee.bh"
    ElseIf Combo1.ListIndex = 1 Then
        str = str & "employee.fz"
    ElseIf Combo1.ListIndex = 2 Then
        str = str & "employee.jf"
    End If
    msh.Clear
    msh.FormatString = "<序号    |<编号   |<姓名     |<性别   |<地址                        |<电话             |<分组             |<积分           "
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open str, con, adOpenDynamic, adLockReadOnly
    If rs.RecordCount <= 1 Then
        msh.Rows = 2
    Else
        msh.Rows = rs.RecordCount + 1
    End If
    If rs.RecordCount > 0 Then
        Label9.Caption = "1/" & rs.RecordCount
    End If
    For i = 1 To rs.RecordCount
        msh.TextMatrix(i, 0) = i
        msh.TextMatrix(i, 1) = Trim(rs.Fields("bh"))
        msh.TextMatrix(i, 2) = Trim(rs.Fields("xm"))
        msh.TextMatrix(i, 3) = Combo2.List(rs.Fields("xb"))
        msh.TextMatrix(i, 4) = Trim(rs.Fields("dz"))
        msh.TextMatrix(i, 5) = Trim(rs.Fields("dh"))
        msh.TextMatrix(i, 6) = Trim(rs.Fields("id")) & "-" & Trim(rs.Fields("mc"))
        msh.TextMatrix(i, 7) = rs.Fields("jf")
        If Not rs.EOF Then rs.MoveNext
    Next i
    For i = 1 To msh.Rows - 1
        hj = hj + Val(msh.TextMatrix(i, 7))
    Next i
    msh.Rows = msh.Rows + 1
    msh.TextMatrix(msh.Rows - 1, 0) = "合计"
    msh.TextMatrix(msh.Rows - 1, 7) = hj
    If msh.Visible = True Then
        msh.SetFocus
    End If
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Command6_Click()
    Dim dqrow As Integer
    If Trim(Text5.Text) = "" Then
        MsgBox "编号不能为空！", vbOKOnly Or vbCritical, "警告"
        Text5.SetFocus
    ElseIf Trim(Text1.Text) = "" Then
        MsgBox "姓名不能为空！", vbOKOnly Or vbCritical, "警告"
        Text1.SetFocus
    Else
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from employee where bh='" & Trim(Text5.Text) & "'", con, adOpenDynamic, adLockOptimistic
        If rs.RecordCount <= 0 Then
            If msh.TextMatrix(msh.Rows - 2, 0) <> "" Then
                msh.Rows = msh.Rows + 1
            End If
            dqrow = msh.Rows - 2
            rs.AddNew
        Else
            For i = 1 To msh.Rows - 1
                If Trim(msh.TextMatrix(i, 1)) = Trim(Text5.Text) Then
                    dqrow = i
                    Exit For
                End If
            Next i
        End If
        rs.Fields("bh") = Trim(Text5.Text)
        rs.Fields("xm") = Trim(Text1.Text)
        rs.Fields("xb") = Combo2.ListIndex
        rs.Fields("dz") = Trim(Text2.Text)
        rs.Fields("dh") = Trim(Text3.Text)
        rs.Fields("fz") = qbh(Combo3.Text)
        rs.Update
        msh.TextMatrix(dqrow, 0) = dqrow
        msh.TextMatrix(dqrow, 1) = Trim(Text5.Text)
        msh.TextMatrix(dqrow, 2) = Trim(Text1.Text)
        msh.TextMatrix(dqrow, 3) = Combo2.List(Combo2.ListIndex)
        msh.TextMatrix(dqrow, 4) = Trim(Text2.Text)
        msh.TextMatrix(dqrow, 5) = Trim(Text3.Text)
        msh.TextMatrix(dqrow, 6) = Trim(Combo3.Text)
        msh.TextMatrix(dqrow, 7) = Trim(Text4.Text)
        
        msh.TextMatrix(msh.Rows - 1, 0) = "合计"
        
        If Text5.Enabled = True Then
            qc
            Text5.SetFocus
        Else
            Command7_Click
        End If
    End If
End Sub
Private Sub qc()
    Text5.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = "0"
    Text5.Enabled = True
End Sub
Private Sub Command7_Click()
    Frame2.Visible = False
    msh.Enabled = True
    Frame1.Enabled = True
End Sub

Private Sub Form_Load()
    Combo1.Text = Combo1.List(0)
    Combo2.Text = Combo2.List(0)
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from tb_group order by id", con, adOpenDynamic, adLockReadOnly
    Combo3.Clear
    For i = 0 To rs.RecordCount - 1
        Combo3.AddItem Trim(rs.Fields("id")) & "-" & Trim(rs.Fields("mc"))
        If Not rs.EOF Then rs.MoveNext
    Next i
    If Combo3.ListCount > 0 Then Combo3.Text = Combo3.List(0)
    Label9.Caption = ""
    Command4_Click
    Me.Show
End Sub

Private Sub Form_Resize()
    msh.Top = 0
    msh.Left = 0
    msh.Width = Me.Width
    msh.Height = Me.Height - 1200
    Frame1.Left = 0
    Frame1.Width = Me.Width
    Frame1.Height = Me.Height - msh.Height
    Frame1.Top = Me.Height - Frame1.Height - 50
End Sub

Private Sub MSH_Click()
    Label9.Caption = msh.row & "/" & msh.Rows - 1
End Sub

Private Sub msh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        MSH_Click
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text5_LostFocus()
    If Trim(Text5.Text) <> "" And Text5.Enabled = True Then
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select bh from employee where bh='" & Trim(Text5.Text) & "'"
        If rs.RecordCount > 0 Then
            MsgBox "工号不能重复！", vbOKOnly Or vbCritical, "警告"
            Text5.SetFocus
            Exit Sub
        End If
    End If
End Sub
