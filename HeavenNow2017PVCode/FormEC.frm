VERSION 5.00
Begin VB.Form FormEC 
   Caption         =   "���ӻ�ѧʵ����"
   ClientHeight    =   645
   ClientLeft      =   450
   ClientTop       =   1140
   ClientWidth     =   13425
   ControlBox      =   0   'False
   Icon            =   "FormEC.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   645
   ScaleWidth      =   13425
   Begin VB.Frame Frame3 
      Caption         =   "������"
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton Command7 
         Caption         =   "����"
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "����"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "����"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "����"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "���ݷ���"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ͼ�����"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�绯ѧ�缫"
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   2415
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "FormEC.frx":09EA
         Left            =   1320
         List            =   "FormEC.frx":09FD
         TabIndex        =   7
         Text            =   "��"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "FormEC.frx":0A14
         Left            =   120
         List            =   "FormEC.frx":0A27
         TabIndex        =   6
         Text            =   "��"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ں�"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "ȷ��"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   720
         TabIndex        =   1
         Text            =   "3"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "COM"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      Caption         =   "��Ҫ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu mnustart 
      Caption         =   "��ʼ"
      Begin VB.Menu mnunew 
         Caption         =   "�½�"
         Begin VB.Menu mnuec 
            Caption         =   "�绯ѧ"
            Begin VB.Menu mnudiandao 
               Caption         =   "�絼�ȷ���"
               Shortcut        =   ^G
            End
            Begin VB.Menu mnulizi 
               Caption         =   "���ӷ���"
               Shortcut        =   ^E
            End
         End
         Begin VB.Menu mnumore 
            Caption         =   "����"
         End
      End
      Begin VB.Menu mnuexit 
         Caption         =   "�˳�"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuset 
      Caption         =   "����"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "����"
   End
End
Attribute VB_Name = "FormEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' �����������б�������λ���κ�������ڵ�ǰ��
Private Const SWP_NOSIZE& = &H1
' ���ִ��ڴ�С
Private Const SWP_NOMOVE& = &H2
' ���ִ���λ��



Private Sub Combo1_Change()
Formlizi.Text4.Text = Combo1.Text
End Sub

Private Sub Combo1_Click()
Formlizi.Text4.Text = Combo1.Text

End Sub



Private Sub Combo2_Change()
Formlizi.Text5.Text = Combo1.Text
End Sub

Private Sub Combo2_Click()
Formlizi.Text5.Text = Combo1.Text
End Sub

Private Sub Command1_Click()
FormGC.MSComm1.CommPort = CInt(Text1.Text)
mnustart.Enabled = True
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' ��������Ϊ������ǰ
mnustart.Enabled = False

End Sub

Private Sub mnudiandao_Click()
FormGC.Show
Combo1.Enabled = True
Combo2.Enabled = True
End Sub

Private Sub mnuexit_Click()
Me.Hide
FormEDU.Show

End Sub

Private Sub mnulizi_Click()
 Formlizi.Show
 

End Sub
