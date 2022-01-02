VERSION 5.00
Begin VB.Form FormEC 
   Caption         =   "电子化学实验室"
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
      Caption         =   "工具箱"
      Height          =   615
      Left            =   5760
      TabIndex        =   8
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton Command7 
         Caption         =   "待续"
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "待续"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "待续"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "待续"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "数据分析"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "图像分析"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "电化学电极"
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
         Text            =   "右"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "FormEC.frx":0A14
         Left            =   120
         List            =   "FormEC.frx":0A27
         TabIndex        =   6
         Text            =   "左"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "串口号"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
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
      Caption         =   "重要参数"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "开始"
      Begin VB.Menu mnunew 
         Caption         =   "新建"
         Begin VB.Menu mnuec 
            Caption         =   "电化学"
            Begin VB.Menu mnudiandao 
               Caption         =   "电导度分析"
               Shortcut        =   ^G
            End
            Begin VB.Menu mnulizi 
               Caption         =   "离子分析"
               Shortcut        =   ^E
            End
         End
         Begin VB.Menu mnumore 
            Caption         =   "暂无"
         End
      End
      Begin VB.Menu mnuexit 
         Caption         =   "退出"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuset 
      Caption         =   "设置"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "关于"
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
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置



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
' 将窗口设为总在最前
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
