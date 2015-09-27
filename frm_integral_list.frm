VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frm_integral_list 
   Caption         =   "积分明细表"
   ClientHeight    =   8115
   ClientLeft      =   3315
   ClientTop       =   3570
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   13950
   WindowState     =   2  'Maximized
   Begin VB.Frame fr_boot 
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   7440
      Width           =   13905
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3510
         TabIndex        =   6
         Top             =   210
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd 23:59:59"
         Format          =   21430275
         CurrentDate     =   42228
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   995
         TabIndex        =   5
         Top             =   210
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd 00:00:00"
         Format          =   21430275
         CurrentDate     =   42228
      End
      Begin VB.CommandButton Command2 
         Caption         =   "退出(&E)"
         Height          =   315
         Left            =   12690
         TabIndex        =   3
         Top             =   210
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "刷新(&F)"
         Height          =   315
         Left            =   11370
         TabIndex        =   2
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "到"
         Height          =   180
         Left            =   3235
         TabIndex        =   7
         Top             =   277
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "时间：从"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   277
         Width           =   720
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   13044
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frm_integral_list.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
End
Attribute VB_Name = "frm_integral_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
    Dim hj As Double
    Dim selstr As String
    selstr = "select a.dh,a.xm_bh,a.xm_mc,a.yg_bh,a.yg_xm,a.bc_jf,a.czy_bh,a.czy_xm,a.date_time,b.fz,c.mc from integral_detail a,employee b,tb_group c where a.yg_bh=b.bh and b.fz=c.id "
    If rs.State = 1 Then rs.Close
    rs.CursorLocation = adUseClient
    If model = 1 Then
        rs.Open selstr & "and date_time between #" & Format(DTPicker1.Value, "yyyy-MM-dd HH:mm:ss") & "# and #" & Format(DTPicker2.Value, "yyyy-MM-dd HH:mm:ss") & "#", con, adOpenKeyset, adLockReadOnly
    Else
        rs.Open selstr & "and date_time between '" & Format(DTPicker1.Value, "yyyy-MM-dd HH:mm:ss") & "' and '" & Format(DTPicker2.Value, "yyyy-MM-dd HH:mm:ss") & "'", con, adOpenKeyset, adLockReadOnly
    End If
    'Debug.Print "select * from integral_detail where date_time between #" & Format(DTPicker1.Value, "yyyy-MM-dd HH:mm:ss") & "# and #" & Format(DTPicker2.Value, "yyyy-MM-dd HH:mm:ss") & "#"
    msh.Clear
    msh.FormatString = "<序号    |<单号             |<项目编号 |<项目名称    |<员工编号  |<员工姓名    |<本次积分   |<录入人编号  |<录入人姓名  |<录入日期            |<小组编号|<小组名称"
    If rs.RecordCount <= 1 Then
        msh.Rows = 2
    Else
        msh.Rows = rs.RecordCount + 1
    End If
    'msh.MergeCells = flexMergeRestrictColumns
    'msh.MergeCol(1) = True
    
    For i = 1 To rs.RecordCount
        msh.TextMatrix(i, 0) = i
        msh.TextMatrix(i, 1) = Trim(rs.Fields("dh"))
        msh.TextMatrix(i, 2) = Trim(rs.Fields("xm_bh"))
        msh.TextMatrix(i, 3) = Trim(rs.Fields("xm_mc"))
        msh.TextMatrix(i, 4) = Trim(rs.Fields("yg_bh"))
        msh.TextMatrix(i, 5) = Trim(rs.Fields("yg_xm"))
        msh.TextMatrix(i, 6) = rs.Fields("bc_jf")
        msh.TextMatrix(i, 7) = Trim(rs.Fields("czy_bh"))
        msh.TextMatrix(i, 8) = Trim(rs.Fields("czy_xm"))
        msh.TextMatrix(i, 9) = Trim(rs.Fields("date_time"))
        msh.TextMatrix(i, 10) = Trim(rs.Fields("fz"))
        msh.TextMatrix(i, 11) = Trim(rs.Fields("mc"))
        If Not rs.EOF Then rs.MoveNext
    Next i
    For i = 1 To msh.Rows - 1
        hj = hj + Val(msh.TextMatrix(i, 6))
    Next i
    msh.Rows = msh.Rows + 1
    msh.TextMatrix(msh.Rows - 1, 0) = "合计"
    msh.TextMatrix(msh.Rows - 1, 6) = hj
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Format(Now, "yyyy-MM-dd 00:00:00")
    DTPicker2.Value = Format(Now, "yyyy-MM-dd 23:59:59")
End Sub

Private Sub Form_Resize()
    msh.Top = 0
    msh.Left = 0
    msh.Width = Me.Width
    msh.Height = Me.Height - fr_boot.Height - 500
    fr_boot.Left = Me.Width - fr_boot.Width - 200
    fr_boot.Top = Me.Height - fr_boot.Height - 530
    'fr_boot.Width = Me.Width
    
End Sub
