VERSION 5.00
Begin VB.Form frm_pwd_change 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "更改密码"
   ClientHeight    =   3405
   ClientLeft      =   8145
   ClientTop       =   3225
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "退出(&E)"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存(&S)"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "校　验："
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "新密码："
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   1230
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "原密码："
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frm_pwd_change"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset

Private Sub Command1_Click()
    If Text1.Text <> login_pwd Then
        MsgBox "原密码不正确请重新输入！", vbOKOnly Or vbCritical, "警告"
        Exit Sub
    End If
    If Text2.Text <> Text3.Text Then
        MsgBox "两次输入密码不一致，请重新输入！", vbOKOnly Or vbCritical, "警告"
        Text3.Text = ""
        Text3.SetFocus
        Exit Sub
    End If
    If rs.State = 1 Then rs.Close
    rs.Open "update login set mm='" & Trim(Text2.Text) & "' where bh='" & Trim(login_bh) & "' ", con, adOpenDynamic, adLockOptimistic
    MsgBox "修改成功！", vbOKOnly Or vbInformation, "恭喜"
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Me.Show
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
