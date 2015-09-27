VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_integral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "积分项目"
   ClientHeight    =   5760
   ClientLeft      =   3465
   ClientTop       =   2970
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   4860
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   4860
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1740
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   4860
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1020
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出(&E)"
      Height          =   495
      Left            =   6330
      TabIndex        =   6
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton cmd_del 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4350
      TabIndex        =   5
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   6330
      TabIndex        =   4
      Top             =   4380
      Width           =   1215
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "保存(&S)"
      Height          =   495
      Left            =   4350
      TabIndex        =   3
      Top             =   4380
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   5745
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   10134
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "<编号      |<名称                   |<分值    "
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编号："
      Height          =   180
      Left            =   4230
      TabIndex        =   10
      Top             =   405
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "分值："
      Height          =   180
      Left            =   4230
      TabIndex        =   9
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "名称："
      Height          =   180
      Left            =   4230
      TabIndex        =   8
      Top             =   1125
      Width           =   540
   End
End
Attribute VB_Name = "frm_integral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim row As Integer

Private Sub cmd_cancel_Click()
    Text1.Enabled = True
    cmd_del.Enabled = False
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    If Text1.Visible = True Then Text1.SetFocus
End Sub

Private Sub cmd_del_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "delete from integral_list where xm_bh='" & Trim(Text1.Text) & "'", con, adOpenDynamic, adLockOptimistic
    If MSH.Rows > 2 Then
        MSH.RemoveItem row
    Else
        For i = 0 To MSH.Cols - 1
            MSH.TextMatrix(row, i) = ""
        Next i
    End If
    cmd_cancel_Click
End Sub

Private Sub cmd_save_Click()
    If Trim(Text1.Text) = "" Then
        MsgBox "编号不能为空！", vbOKOnly Or vbCritical, "警告"
        Text1.SetFocus
    ElseIf Trim(Text2.Text) = "" Then
        MsgBox "名称不能为空！", vbOKOnly Or vbCritical, "警告"
        Text2.SetFocus
    ElseIf Not IsNumeric(Text3.Text) Then
        MsgBox "请输入一个数值！", vbOKOnly Or vbCritical, "警告"
        Text3.SetFocus
    Else
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from integral_list where xm_bh='" & Trim(Text1.Text) & "'", con, adOpenDynamic, adLockOptimistic
        If rs.RecordCount <= 0 Then
            If Trim(MSH.TextMatrix(MSH.Rows - 1, 0)) <> "" Then
                MSH.Rows = MSH.Rows + 1
            End If
            row = MSH.Rows - 1
            rs.AddNew
        Else
            For i = 1 To MSH.Rows - 1
                If Trim(MSH.TextMatrix(i, 0)) = Trim(Text1.Text) Then
                    row = i
                    Exit For
                End If
            Next i
        End If
        rs.Fields("xm_bh") = Trim(Text1.Text)
        rs.Fields("xm_mc") = Trim(Text2.Text)
        rs.Fields("fz") = Text3.Text
        rs.Update
        MSH.TextMatrix(row, 0) = Trim(Text1.Text)
        MSH.TextMatrix(row, 1) = Trim(Text2.Text)
        MSH.TextMatrix(row, 2) = Text3.Text
        cmd_cancel_Click
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    cmd_cancel_Click
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from integral_list order by xm_bh", con, adOpenDynamic, adLockReadOnly
    If rs.RecordCount <= 1 Then
        MSH.Rows = 2
    Else
        MSH.Rows = rs.RecordCount + 1
    End If
    For i = 1 To rs.RecordCount
        MSH.TextMatrix(i, 0) = Trim(rs.Fields("xm_bh"))
        MSH.TextMatrix(i, 1) = Trim(rs.Fields("xm_mc"))
        MSH.TextMatrix(i, 2) = rs.Fields("fz")
        If Not rs.EOF Then rs.MoveNext
    Next i
    Me.Show
End Sub

Private Sub MSH_DblClick()
    If Trim(MSH.TextMatrix(MSH.row, 0)) <> "" Then
        row = MSH.row
        Text1.Text = MSH.TextMatrix(row, 0)
        Text2.Text = MSH.TextMatrix(row, 1)
        Text3.Text = MSH.TextMatrix(row, 2)
        Text1.Enabled = False
        cmd_del.Enabled = True
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
