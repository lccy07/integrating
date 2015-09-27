VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_group 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "小组信息"
   ClientHeight    =   5820
   ClientLeft      =   6000
   ClientTop       =   2655
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_save 
      Caption         =   "保存(&S)"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "取消(&C)"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_del 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出(&E)"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   4980
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1395
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   4980
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   675
      Width           =   2685
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   5745
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   10134
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "<编号      |<名称                   "
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "名称："
      Height          =   180
      Left            =   4440
      TabIndex        =   8
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编号："
      Height          =   180
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "frm_group"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As New ADODB.Recordset
Private row As Integer

Private Sub cmd_cancel_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text1.Enabled = True
    cmd_del.Enabled = False
    Text1.SetFocus
End Sub

Private Sub cmd_del_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "delete from tb_group where id='" & Trim(Text1.Text) & "'", con, adOpenDynamic, adLockOptimistic
    MSH.RemoveItem row
    cmd_cancel_Click
End Sub

Private Sub cmd_save_Click()
    If Trim(Text1.Text) = "" Then
        MsgBox "小组编号不能为空！", vbOKOnly Or vbCritical, "警告"
    ElseIf Trim(Text2.Text) = "" Then
        MsgBox "小组名称不能为空！", vbOKOnly Or vbCritical, "警告"
    Else
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from tb_group where id = '" & Trim(Text1.Text) & "'", con, adOpenDynamic, adLockOptimistic
        If rs.RecordCount > 0 Then
            rs.Fields("id") = Trim(Text1.Text)
            rs.Fields("mc") = Trim(Text2.Text)
            rs.Update
            For i = 1 To MSH.Rows
                If MSH.TextMatrix(i, 0) = Trim(Text1.Text) Then
                    row = i
                    Exit For
                End If
            Next i
        Else
            rs.AddNew
            rs.Fields("id") = Trim(Text1.Text)
            rs.Fields("mc") = Trim(Text2.Text)
            rs.Update
            If MSH.TextMatrix(MSH.Rows - 1, 0) <> "" Then
                MSH.Rows = MSH.Rows + 1
            End If
            row = MSH.Rows - 1
        End If
        MSH.TextMatrix(row, 0) = Trim(Text1.Text)
        MSH.TextMatrix(row, 1) = Trim(Text2.Text)
        cmd_cancel_Click
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open "select * from tb_group order by id", con, adOpenDynamic, adLockOptimistic
    If rs.RecordCount <= 1 Then
        MSH.Rows = 2
    Else
        MSH.Rows = rs.RecordCount + 1
    End If
    For i = 1 To rs.RecordCount
        MSH.TextMatrix(i, 0) = Trim(rs.Fields("id"))
        MSH.TextMatrix(i, 1) = Trim(rs.Fields("mc"))
        If Not rs.EOF Then rs.MoveNext
    Next i
    Me.Top = 0
    Me.Left = 0
    Me.Show
End Sub

Private Sub MSH_DblClick()
    If Trim(MSH.TextMatrix(MSH.row, 0)) <> "" Then
        Text1.Enabled = False
        cmd_del.Enabled = True
        Text1.Text = MSH.TextMatrix(MSH.row, 0)
        Text2.Text = MSH.TextMatrix(MSH.row, 1)
        row = MSH.row
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Trim(Text1.Text) <> "" Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{tab}"
    End If
End Sub
