VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   Caption         =   "个人信息"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Text            =   "Local"
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   120
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "机主姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "本机IP"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "本机昵称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()


MsgBox ("Successful!")
Me.Hide


End Sub

Private Sub Form_Load()
Text2.Text = Winsock1.LocalIP
Text3.Text = Winsock1.LocalHostName
End Sub
