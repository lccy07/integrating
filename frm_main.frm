VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frm_main 
   BackColor       =   &H8000000C&
   Caption         =   "Ա������ϵͳ"
   ClientHeight    =   9720
   ClientLeft      =   2700
   ClientTop       =   2820
   ClientWidth     =   14250
   LinkTopic       =   "MDIForm1"
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11400
      Top             =   5040
   End
   Begin ComctlLib.StatusBar STB 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   9225
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   5292
            MinWidth        =   5292
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "��ӭʹ��Ա�����ֹ���ϵͳ"
            TextSave        =   "��ӭʹ��Ա�����ֹ���ϵͳ"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu_system 
      Caption         =   "ϵͳ"
      Begin VB.Menu menu_menger 
         Caption         =   "����Ա��Ϣ"
      End
      Begin VB.Menu menu_pwd_change 
         Caption         =   "��������"
      End
      Begin VB.Menu meun_exit 
         Caption         =   "�˳�ϵͳ"
      End
   End
   Begin VB.Menu menu_base 
      Caption         =   "������Ϣ"
      Begin VB.Menu menu_group 
         Caption         =   "С������"
      End
      Begin VB.Menu menu_employee 
         Caption         =   "Ա����Ϣ"
      End
      Begin VB.Menu menu_line1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_integral 
         Caption         =   "������Ŀ"
      End
   End
   Begin VB.Menu menu_integral_in 
      Caption         =   "����¼��"
   End
   Begin VB.Menu menu_report 
      Caption         =   "�������"
      Begin VB.Menu menu_report_integra_list 
         Caption         =   "������ϸ"
      End
      Begin VB.Menu menu_intergral_sum 
         Caption         =   "���ֻ���"
      End
   End
   Begin VB.Menu menu_about 
      Caption         =   "����"
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    If DateDiff("d", Now, "2015-10-11") > 0 And DateDiff("d", Now, "2015-10-11") < 32 Then
         STB.Panels(1).Text = "��ǰ��¼�û���" & login_xm
         STB.Panels(2).Text = "��ǰʱ�䣺" & Format(Now, "yyyy-MM-dd HH:mm:ss")
        Me.Show
    Else
        MsgBox "����ǰʹ��ϵͳΪ���ð�,����ʱ���ѵ�,����ʹ�����ϵͳ��Ӧ����ϵ��", vbOKOnly Or vbInformation, "��ʾ"
        End
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub menu_employee_Click()
    Load frm_employee
End Sub

Private Sub menu_group_Click()
    Load frm_group
End Sub

Private Sub menu_integral_Click()
    Load frm_integral
End Sub

Private Sub menu_integral_in_Click()
    On Error Resume Next
    Load frm_integral_in
End Sub

Private Sub menu_intergral_sum_Click()
    Load frm_intergral_sum
End Sub

Private Sub menu_menger_Click()
    Load frm_menger
End Sub

Private Sub menu_pwd_change_Click()
    Load frm_pwd_change
End Sub

Private Sub menu_report_integra_list_Click()
    Load frm_integral_list
End Sub

Private Sub meun_exit_Click()
    End
End Sub

Private Sub Timer1_Timer()
    STB.Panels(2).Text = "��ǰʱ�䣺" & Format(Now, "yyyy-MM-dd HH:mm:ss")
End Sub
