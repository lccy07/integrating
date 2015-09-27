VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_menger 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "管理员信息"
   ClientHeight    =   6720
   ClientLeft      =   2535
   ClientTop       =   4905
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox com_xb 
      Height          =   300
      ItemData        =   "frm_menger.frx":0000
      Left            =   6720
      List            =   "frm_menger.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txt_dh 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   3113
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出(&E)"
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmd_del 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   8640
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "保存(&S)"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txt_xm 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txt_bh 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHF 
      Height          =   6735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "<编号      |<姓名         |<性别  |<电话            |<ID      "
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lb_row 
      Caption         =   "lb_row"
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "电话："
      Height          =   180
      Left            =   6120
      TabIndex        =   11
      Top             =   3210
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "性别："
      Height          =   180
      Left            =   6120
      TabIndex        =   2
      Top             =   2250
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名："
      Height          =   180
      Left            =   6120
      TabIndex        =   1
      Top             =   1290
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编号："
      Height          =   180
      Left            =   6120
      TabIndex        =   0
      Top             =   330
      Width           =   540
   End
End
Attribute VB_Name = "frm_menger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As New ADODB.Recordset

Private Sub cmd_cancel_Click()
    txt_bh.Text = ""
    txt_xm.Text = ""
    txt_dh.Text = ""
    txt_bh.Enabled = True
    If txt_bh.Visible = True Then txt_bh.SetFocus
    cmd_del.Enabled = False
    lb_row.Caption = ""
End Sub

Private Sub cmd_del_Click()
    If txt_bh.Enabled = False Then
        If rs.State = 1 Then rs.Close
        rs.Open "delete from login where bh='" & Trim(txt_bh.Text) & "'", con, adOpenDynamic, adLockOptimistic
        If MSHF.Rows > 2 Then
            MSHF.RemoveItem lb_row.Caption
        Else
            For i = 0 To MSHF.Cols - 1
                MSHF.TextMatrix(1, i) = ""
            Next i
        End If
        cmd_cancel_Click
    End If
End Sub

Private Sub cmd_save_Click()
    Dim row As Integer
    If Trim(txt_bh.Text) = "" Then
        MsgBox "编号不能为空！", vbOKOnly Or vbCritical, "警告"
        txt_bh.SetFocus
        Exit Sub
    End If
    If Trim(txt_xm.Text) = "" Then
        MsgBox "姓名不能为空！"
        txt_xm.SetFocus
        Exit Sub
    End If
    If com_xb.Text = "男" Then
        xb = 1
    ElseIf com_xb.Text = "女" Then
        xb = 0
    End If
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from login where bh='" & Trim(txt_bh.Text) & "'", con, adOpenDynamic, adLockOptimistic
    If rs.RecordCount <= 0 Then
        rs.AddNew
        rs.Fields("bh") = Trim(txt_bh.Text)
        rs.Fields("xm") = Trim(txt_xm.Text)
        rs.Fields("mm") = Trim(txt_bh.Text)
        rs.Fields("xb") = xb
        rs.Fields("dz") = ""
        rs.Fields("dh") = Trim(txt_dh.Text)
        rs.Update
        If MSHF.TextMatrix(MSHF.Rows - 1, 0) <> "" Then
            MSHF.Rows = MSHF.Rows + 1
        End If
        row = MSHF.Rows - 1
    Else
        rs.Fields("bh") = Trim(txt_bh.Text)
        rs.Fields("xm") = Trim(txt_xm.Text)
        rs.Fields("mm") = Trim(txt_bh.Text)
        rs.Fields("xb") = xb
        rs.Fields("dz") = ""
        rs.Fields("dh") = Trim(txt_dh.Text)
        rs.Update
        For i = 1 To MSHF.Rows
            If MSHF.TextMatrix(i, 0) = Trim(txt_bh.Text) Then
                row = i
                Exit For
            End If
        Next i
    End If
    MSHF.TextMatrix(row, 0) = txt_bh.Text
    MSHF.TextMatrix(row, 1) = txt_xm.Text
    MSHF.TextMatrix(row, 2) = com_xb.Text
    MSHF.TextMatrix(row, 3) = txt_dh.Text
    cmd_cancel_Click
End Sub

Private Sub Command2_Click()

End Sub

Private Sub com_xb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_dh.SetFocus
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from login order by ID", con, adOpenDynamic, adLockOptimistic
    MSHF.ColWidth(4) = 0
    If rs.RecordCount > 0 Then
        MSHF.Rows = rs.RecordCount + 1
    Else
        MSHF.Rows = 2
    End If
    For i = 1 To rs.RecordCount
        MSHF.TextMatrix(i, 0) = Trim(rs.Fields("bh"))
        MSHF.TextMatrix(i, 1) = Trim(rs.Fields("xm"))
        If rs.Fields("xb") = "1" Then
            MSHF.TextMatrix(i, 2) = "男"
        ElseIf rs.Fields("xb") = "0" Then
            MSHF.TextMatrix(i, 2) = "女"
        End If
        MSHF.TextMatrix(i, 3) = Trim(rs.Fields("dh"))
        MSHF.TextMatrix(i, 4) = Trim(rs.Fields("id"))
        If Not rs.EOF Then rs.MoveNext
    Next i
    com_xb.Text = com_xb.List(0)
    Me.Show
End Sub

Private Sub Form_Resize()
    MSHF.Height = Me.Height
    MSHF.Top = 0
    MSHF.Left = 0
End Sub

Private Sub MSHF_DblClick()
    If MSHF.Rows >= 2 And Trim(MSHF.TextMatrix(MSHF.row, 0)) <> "" Then
        txt_bh.Text = Trim(MSHF.TextMatrix(MSHF.row, 0))
        txt_xm.Text = Trim(MSHF.TextMatrix(MSHF.row, 1))
        com_xb.Text = Trim(MSHF.TextMatrix(MSHF.row, 2))
        txt_dh.Text = Trim(MSHF.TextMatrix(MSHF.row, 3))
        cmd_del.Enabled = True
        txt_bh.Enabled = False
        lb_row.Caption = MSHF.row
    End If
End Sub

Private Sub txt_bh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txt_bh.Text) <> "" Then
            txt_xm.SetFocus
        End If
    End If
End Sub

Private Sub txt_bh_LostFocus()
    Dim pd
    If Trim(txt_bh.Text) <> "" Then
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from login where bh='" & Trim(txt_bh.Text) & "'", con, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            pd = MsgBox("编号已经存在！" & vbCrLf & "修改信息请点【是】" & vbCrLf & "否则请点【否】", vbYesNo Or vbQuestion, "询问")
            If pd = vbYes Then
                txt_bh.Enabled = False
                txt_xm.Text = Trim(rs.Fields("xm"))
                If rs.Fields("xb") = 0 Then
                    com_xb.Text = "女"
                Else
                    com_xb.Text = "男"
                End If
                txt_dh.Text = Trim(rs.Fields("dh"))
                cmd_del.Enabled = True
                txt_xm.SetFocus
                For i = 1 To MSHF.Rows
                    If MSHF.TextMatrix(i, 0) = txt_bh.Text Then
                        lb_row.Caption = i
                        Exit For
                    End If
                Next i
            ElseIf pd = vbNo Then
                txt_bh.Text = ""
                txt_bh.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txt_dh_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_save.SetFocus
    End If
End Sub

Private Sub txt_xm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txt_xm.Text) <> "" Then
            com_xb.SetFocus
        End If
    End If
End Sub
