VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "DYPro Chatter 2016"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6855
   Icon            =   "Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form.frx":030A
   ScaleHeight     =   5160
   ScaleWidth      =   6855
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton C9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2880
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1440
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C7 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2880
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1440
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1440
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton C1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   4395
      TabIndex        =   18
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton C 
      BackColor       =   &H000000FF&
      Caption         =   "等待你的朋友"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Picture         =   "Form.frx":0BD4
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   6615
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "发送"
         Height          =   375
         Left            =   5640
         MouseIcon       =   "Form.frx":23FE
         Picture         =   "Form.frx":3C70
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin RichTextLib.RichTextBox Text2 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1931
         _Version        =   393217
         BackColor       =   12648384
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"Form.frx":3F7A
         MouseIcon       =   "Form.frx":4017
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin MSWinsockLib.Winsock Winsock3 
         Left            =   5640
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   6120
         Top             =   3120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   6120
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5640
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6120
         Top             =   2040
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   6120
         Top             =   1560
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6120
         Top             =   1080
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6120
         Top             =   600
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   25
         Left            =   6120
         Top             =   120
      End
      Begin RichTextLib.RichTextBox Text1 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5741
         _Version        =   393217
         BackColor       =   12648384
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"Form.frx":5E59
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSWinsockLib.Winsock WinUdpA 
      Left            =   6360
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin WMPLibCtl.WindowsMediaPlayer voice 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   5280
      Width           =   3255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   5741
      _cy             =   873
   End
   Begin VB.Menu mnust 
      Caption         =   "开始"
      Begin VB.Menu Command3 
         Caption         =   "连接"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnumin 
         Caption         =   "最小化"
         Shortcut        =   ^M
      End
      Begin VB.Menu white 
         Caption         =   "白板"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnue 
         Caption         =   "退出"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnufr 
      Caption         =   "好友"
      Begin VB.Menu mnulist 
         Caption         =   "联系人列表"
         Begin VB.Menu mnulis 
            Caption         =   "列表"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnuadd 
            Caption         =   "添加新联系人"
            Shortcut        =   {F2}
         End
      End
      Begin VB.Menu mnunei 
         Caption         =   "局域网内好友"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnus 
      Caption         =   "设置"
      Begin VB.Menu mnup 
         Caption         =   "个人信息"
         Shortcut        =   ^P
      End
      Begin VB.Menu Command2 
         Caption         =   "新端口直联"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu 
         Caption         =   "________"
      End
      Begin VB.Menu mnuf 
         Caption         =   "字体"
         Begin VB.Menu mnu10 
            Caption         =   "10"
         End
         Begin VB.Menu mnu20 
            Caption         =   "20"
         End
         Begin VB.Menu mnu25 
            Caption         =   "25"
         End
         Begin VB.Menu mnu30 
            Caption         =   "30"
         End
         Begin VB.Menu mnu40 
            Caption         =   "40"
         End
         Begin VB.Menu mnu45 
            Caption         =   "45"
         End
      End
      Begin VB.Menu mnuli 
         Caption         =   "_________"
      End
      Begin VB.Menu mnuco 
         Caption         =   "背景色"
         Begin VB.Menu mnuday 
            Caption         =   "早间模式"
            Begin VB.Menu mnuyellow 
               Caption         =   "淡黄"
            End
            Begin VB.Menu mnugreen 
               Caption         =   "淡绿"
            End
            Begin VB.Menu mnublue 
               Caption         =   "淡蓝"
            End
            Begin VB.Menu mnupur 
               Caption         =   "淡紫"
            End
            Begin VB.Menu mnured 
               Caption         =   "淡粉"
            End
            Begin VB.Menu mnuwhite 
               Caption         =   "白色"
            End
         End
         Begin VB.Menu mnunight 
            Caption         =   "夜间模式"
            Begin VB.Menu mnuheibai 
               Caption         =   "暂无"
            End
         End
      End
   End
   Begin VB.Menu mnunum 
      Caption         =   "数据"
      Begin VB.Menu mnupic 
         Caption         =   "图片文件"
      End
      Begin VB.Menu mnucopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnucut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuselecall 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnulin 
         Caption         =   "________"
      End
      Begin VB.Menu mnucr 
         Caption         =   "清空来往讯息"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuc 
         Caption         =   "撤销输入"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnusave 
         Caption         =   "保存数据"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnulook 
         Caption         =   "查看记录"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuduang 
      Caption         =   "加特技"
      Begin VB.Menu mnutop 
         Caption         =   "窗口置顶"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuxia 
         Caption         =   "变色龙"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnucmi 
         Caption         =   "暴风雨"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnucutrul 
         Caption         =   "风格化"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnupo 
         Caption         =   "警示字体"
         Begin VB.Menu Und 
            Caption         =   "下划线"
         End
         Begin VB.Menu blod 
            Caption         =   "加粗"
         End
         Begin VB.Menu italy 
            Caption         =   "意大利"
         End
         Begin VB.Menu mnu12 
            Caption         =   "1.2x"
         End
         Begin VB.Menu mnu15 
            Caption         =   "1.5x"
         End
         Begin VB.Menu mnucan 
            Caption         =   "撤销"
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuad 
         Caption         =   "警示音"
         Begin VB.Menu mnubeep 
            Caption         =   "Beep"
            Shortcut        =   {F9}
         End
      End
   End
   Begin VB.Menu mnugame 
      Caption         =   "游戏"
      Begin VB.Menu mnutic 
         Caption         =   "三子棋"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


 
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
Private Const HWND_NOTOPMOST& = -2
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置
Dim time As Integer
Dim time1 As Integer
Dim time2 As Integer
Dim time3 As Integer
Dim time4 As Integer
Dim t As Integer
Dim other As String
Dim strdata As String
Dim player As Integer
Dim C01, C02, C03, C04, C05, C06, C07, C08, C09 As Integer
Dim x1, y1, r, b, g, r1, b1, g1 As Integer
Dim CC As Long








Private Sub blod_Click()
WinUdpa.SendData "粗体"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "粗体"
End Sub






Private Sub C1_Click()
If player = 0 Then C1.Caption = "O"
If player = 1 Then C1.Caption = "X"
C01 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C1"
End Sub
Private Sub C2_Click()
If player = 0 Then C2.Caption = "O"
If player = 1 Then C2.Caption = "X"
C02 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C2"
End Sub



Private Sub C3_Click()
If player = 0 Then C3.Caption = "O"
If player = 1 Then C3.Caption = "X"
C03 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C3"
End Sub
Private Sub C4_Click()
If player = 0 Then C4.Caption = "O"
If player = 1 Then C4.Caption = "X"
C04 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C4"
End Sub
Private Sub C5_Click()
If player = 0 Then C5.Caption = "O"
If player = 1 Then C5.Caption = "X"
C05 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C5"
End Sub
Private Sub C6_Click()
If player = 0 Then C6.Caption = "O"
If player = 1 Then C6.Caption = "X"
C06 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C6"
End Sub
Private Sub C7_Click()
If player = 0 Then C7.Caption = "O"
If player = 1 Then C7.Caption = "X"
C07 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C7"
End Sub
Private Sub C8_Click()
If player = 0 Then C8.Caption = "O"
If player = 1 Then C8.Caption = "X"
C08 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C8"
End Sub
Private Sub C9_Click()
If player = 0 Then C9.Caption = "O"
If player = 1 Then C9.Caption = "X"
C09 = 1
C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
WinUdpa.SendData "C9"
End Sub

Private Sub Command1_Click()
WinUdpa.SendData Text2.Text
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & Text2.Text
Text2.Text = ""
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
mnugame.Enabled = True
WinUdpa.LocalPort = Form2.Text3.Text
WinUdpa.RemoteHost = Form2.Text1.Text
WinUdpa.RemotePort = Form2.Text2.Text
WinUdpa.Bind
Winsock1.LocalPort = Int(Form2.Text3.Text) + 94
Winsock1.RemoteHost = Form2.Text1.Text
Winsock1.RemotePort = Int(Form2.Text2.Text) + 94
Winsock1.Bind
Winsock2.LocalPort = Int(Form2.Text3.Text) + 93
Winsock2.RemoteHost = Form2.Text1.Text
Winsock2.RemotePort = Int(Form2.Text2.Text) + 93
Winsock2.Bind
Winsock3.LocalPort = Int(Form2.Text3.Text) + 92
Winsock3.RemoteHost = Form2.Text1.Text
Winsock3.RemotePort = Int(Form2.Text2.Text) + 92
Winsock3.Bind

Command3.Enabled = False
Command1.Enabled = True
mnuduang.Enabled = True
mnupic.Enabled = True
white.Enabled = True
If Form7.Text1.Text = "" Then
Form1.Caption = "DC2016 —— " & Form2.Text1.Text
Else: Form1.Caption = "DC2016 —— " & Form7.Text1.Text
End If

End Sub

Private Sub Command4_Click()
Text1.Left = 120
Text1.Width = 6255
Text2.Width = 6375
Text2.Left = 120

C1.Visible = False
C2.Visible = False
C3.Visible = False
C4.Visible = False
C5.Visible = False
C6.Visible = False
C7.Visible = False
C8.Visible = False
C9.Visible = False
C.Visible = False
Command4.Visible = False
WinUdpa.SendData "您的朋友结束三子棋"
C1.Caption = ""
C2.Caption = ""
C3.Caption = ""
C4.Caption = ""
C5.Caption = ""
C6.Caption = ""
C7.Caption = ""
C8.Caption = ""
C9.Caption = ""

End Sub

Private Sub Form_Load()
Command1.Enabled = False
Command3.Enabled = False
mnuduang.Enabled = False
voice.URL = ""
C.Enabled = False
mnugame.Enabled = False
C01 = 0
C02 = 0
C03 = 0
C04 = 0
C05 = 0
C06 = 0
C07 = 0
C08 = 0
C09 = 0
white.Enabled = False
Picture1.Visible = False
mnupic.Enabled = False
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

PopupMenu mnunum, vbPopupMenuLeftAlign

Else


Exit Sub
End If

End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

PopupMenu mnunum, vbPopupMenuLeftAlign

Else


Exit Sub
End If
End Sub

Private Sub italy_Click()
WinUdpa.SendData "意大利化"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "意大利化"
End Sub

Private Sub mnu10_Click()
Text1.Font.Size = 10
Text2.Font.Size = 10
End Sub

Private Sub mnu12_Click()
WinUdpa.SendData "1.2x"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "1.2x"
End Sub

Private Sub mnu15_Click()
WinUdpa.SendData "1.5x"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "1.5x"
End Sub

Private Sub mnu20_Click()
Text1.Font.Size = 20
Text2.Font.Size = 20
End Sub

Private Sub mnu25_Click()
Text1.Font.Size = 25
Text2.Font.Size = 25

End Sub

Private Sub mnu30_Click()
Text1.Font.Size = 30
Text2.Font.Size = 30
End Sub

Private Sub mnu40_Click()
Text1.Font.Size = 40
Text2.Font.Size = 40
End Sub

Private Sub mnu45_Click()
Text1.Font.Size = 45
Text2.Font.Size = 45
End Sub

Private Sub mnuadd_Click()
Form6.Show
End Sub

Private Sub mnubeep_Click()
WinUdpa.SendData "beep"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "beep"
End Sub

Private Sub mnublue_Click()
Form1.BackColor = &HFFC0C0
Frame1.BackColor = &HFFC0C0
Frame2.BackColor = &HFFC0C0
Text1.BackColor = &HFFC0C0
Text2.BackColor = &HFFC0C0
End Sub

Private Sub mnuc_Click()
Text2.Text = " "
End Sub

Private Sub mnucan_Click()
WinUdpa.SendData "撤销"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "撤销"
Text1.Font.Italic = False
Text1.Font.Bold = False
Text1.Font.Underline = False
Text1.Font.Size = 10
End Sub

Private Sub mnucmi_Click()
WinUdpa.SendData "暴风雨"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "暴风雨"
End Sub

Private Sub mnucr_Click()
Text1.Text = " "

End Sub

Private Sub mnucutrul_Click()
WinUdpa.SendData "风格化"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "风格化"
End Sub

Private Sub mnue_Click()
Me.Hide
Formmenu.Show



End Sub

Private Sub mnugreen_Click()
Form1.BackColor = &HC0FFC0
Frame1.BackColor = &HC0FFC0
Frame2.BackColor = &HC0FFC0
Text1.BackColor = &HC0FFC0
Text2.BackColor = &HC0FFC0
End Sub

Private Sub mnuheibai_Click()
Form1.BackColor = &H0&
Frame1.BackColor = &H0&
Frame2.BackColor = &H0&
Text1.BackColor = &H0&
Text2.BackColor = &H0&


End Sub

Private Sub mnulis_Click()
Form7.Show
End Sub

Private Sub mnulook_Click()
Form8.Show
End Sub

Private Sub mnumin_Click()
Form1.WindowState = 1
End Sub


Private Sub mnunei_Click()
Form5.Show
Form4.Show
Form4.Visible = False


End Sub

Private Sub mnup_Click()
Form3.Show

End Sub

Private Sub mnupic_Click()
CommonDialog1.Filter = "位图|*.bmp|JPG图片|*.jpg|所有文件|*.*"

     CommonDialog1.ShowOpen
Dim FileType, FiType, FileName As String

 

FileType = CommonDialog1.FileTitle

 

FiType = LCase(Right(FileType, 3))

FileName = CommonDialog1.FileName

Picture1.Visible = True
Sleep 1000
Picture1.Picture = LoadPicture(FileName)
Sleep 1000
ProgressBar1.Visible = True
ProgressBar1.Value = 0
Text1.Left = 4440
Text1.Width = 2000
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "图片文件"

WinUdpa.SendData "[图片文件]"

Sleep 1000

Dim X, Y As Integer
Dim CCC, CCCC
X = 0
Y = 0
While Y < 3375
ProgressBar1.Value = Y

X = 0
Winsock2.SendData Y
Sleep 10
While X < 4335
CCC = Picture1.Point(X, Y)

Winsock1.SendData X

Winsock3.SendData CCC

X = X + 5
Wend
Y = Y + 5
Wend
ProgressBar1.Visible = False
End Sub

Private Sub mnupur_Click()
Form1.BackColor = &HFFC0FF
Frame1.BackColor = &HFFC0FF
Frame2.BackColor = &HFFC0FF
Text1.BackColor = &HFFC0FF
Text2.BackColor = &HFFC0FF
End Sub

Private Sub mnured_Click()
Form1.BackColor = &HC0C0FF
Frame1.BackColor = &HC0C0FF
Frame2.BackColor = &HC0C0FF
Text1.BackColor = &HC0C0FF
Text2.BackColor = &HC0C0FF
End Sub

Private Sub mnusave_Click()



If Dir("\DATA") = " " Then MkDir "DATA"


Open App.Path & "\DATA\" & Date & other & ".save" For Output As #1

Close #1
Text1.SaveFile App.Path & "\DATA\" & Date & other & ".save", rtfText
    
End Sub

Private Sub mnutic_Click()
Text1.Left = 4440
Text1.Width = 2000
Text2.Left = 4440
Text2.Width = 2000
C1.Visible = True
C2.Visible = True
C3.Visible = True
C4.Visible = True
C5.Visible = True
C6.Visible = True
C7.Visible = True
C8.Visible = True
C9.Visible = True
C.Visible = True
Command4.Visible = True
WinUdpa.SendData "您的朋友开始三子棋"
C01 = 0
C02 = 0
C03 = 0
C04 = 0
C05 = 0
C06 = 0
C07 = 0
C08 = 0
C09 = 0

C1.Caption = ""
C2.Caption = ""
C3.Caption = ""
C4.Caption = ""
C5.Caption = ""
C6.Caption = ""
C7.Caption = ""
C8.Caption = ""
C9.Caption = ""
C.Caption = "等待你的朋友"

End Sub

Private Sub mnutop_Click()
WinUdpa.SendData "置顶"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "窗口置顶"
End Sub

Private Sub mnuwhite_Click()
Form1.BackColor = &HFFFFFF
Frame1.BackColor = &HFFFFFF
Frame2.BackColor = &HFFFFFF
Text1.BackColor = &HFFFFFF
Text2.BackColor = &HFFFFFF

End Sub

Private Sub mnuxia_Click()
WinUdpa.SendData "变色龙"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "变色龙"
End Sub

Private Sub mnuyellow_Click()
Form1.BackColor = &HC0FFFF
Frame1.BackColor = &HC0FFFF
Frame2.BackColor = &HC0FFFF
Text1.BackColor = &HC0FFFF
Text2.BackColor = &HC0FFFF

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

PopupMenu mnunum, vbPopupMenuLeftAlign

Else


Exit Sub
End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

PopupMenu mnunum, vbPopupMenuLeftAlign

Else


Exit Sub
End If
End Sub

Private Sub Timer1_Timer()

time = time + 1
If time > 1 Then SetWindowPos Form1.hwnd, -2, 0, 0, 0, 0, 2 Or 1
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
time1 = time1 + 1
If time1 = 1 Then
Form1.BackColor = &HC0FFC0
Frame1.BackColor = &HC0FFC0
Frame2.BackColor = &HC0FFC0
Text1.BackColor = &HC0FFC0
Text2.BackColor = &HC0FFC0
End If
If time1 = 2 Then
Form1.BackColor = &HFFC0FF
Frame1.BackColor = &HFFC0FF
Frame2.BackColor = &HFFC0FF
Text1.BackColor = &HFFC0FF
Text2.BackColor = &HFFC0FF
End If
If time1 = 3 Then
Form1.BackColor = &HFFFFFF
Frame1.BackColor = &HFFFFFF
Frame2.BackColor = &HFFFFFF
Text1.BackColor = &HFFFFFF
Text2.BackColor = &HFFFFFF
End If
If time1 = 5 Then
Form1.BackColor = &HC0FFFF
Frame1.BackColor = &HC0FFFF
Frame2.BackColor = &HC0FFFF
Text1.BackColor = &HC0FFFF
Text2.BackColor = &HC0FFFF
End If
If time1 = 6 Then
Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()
time2 = time2 + 1

If time2 = 1 Then
Form1.BackColor = &H0&
Frame1.BackColor = &H0&
Frame2.BackColor = &H0&
Text1.BackColor = &H0&
Text2.BackColor = &H0&
End If
If time2 = 2 Then
Form1.BackColor = &HFFFFFF
Frame1.BackColor = &HFFFFFF
Frame2.BackColor = &HFFFFFF
Text1.BackColor = &HFFFFFF
Text2.BackColor = &HFFFFFF
End If
If time2 = 3 Then
Form1.BackColor = &H0&
Frame1.BackColor = &H0&
Frame2.BackColor = &H0&
Text1.BackColor = &H0&
Text2.BackColor = &H0&
End If
If time2 = 4 Then
Form1.BackColor = &HFFFFFF
Frame1.BackColor = &HFFFFFF
Frame2.BackColor = &HFFFFFF
Text1.BackColor = &HFFFFFF
Text2.BackColor = &HFFFFFF
End If
If time2 = 5 Then
Form1.BackColor = &H0&
Frame1.BackColor = &H0&
Frame2.BackColor = &H0&
Text1.BackColor = &H0&
Text2.BackColor = &H0&
End If
If time2 = 6 Then
Form1.BackColor = &HFFFFFF
Frame1.BackColor = &HFFFFFF
Frame2.BackColor = &HFFFFFF
Text1.BackColor = &HFFFFFF
Text2.BackColor = &HFFFFFF
Timer2.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
time3 = time3 + 1
If time3 = 1 Then
Text1.Font.Bold = True
End If
If time3 = 2 Then
Text1.Font.Bold = False
Text1.Font.Italic = True
End If
If time3 = 3 Then
Text1.Font.Italic = False
Text1.Font.Underline = True
End If
If time3 = 5 Then
Text1.Font.Underline = False
Text1.Font.Size = Text1.Font.Size * 1.5
End If
If time3 = 6 Then
Text1.Font.Size = Text1.Font.Size / 1.5
Timer2.Enabled = False
End If

End Sub

Private Sub Timer5_Timer()
t = t + 1
If t = 1 Then voice.URL = ""
Timer5.Enabled = False

 

End Sub

Private Sub Und_Click()
WinUdpa.SendData "下划线"
Text1.Text = Text1.Text & vbCrLf & "[" & Form3.Text1.Text & " | " & Now & "]：" & "下划线"
End Sub

Private Sub white_Click()
Formwhite.Show
WinUdpa.SendData "您的朋友开始使用白板"

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData x1, vbInteger
CC = Abs(CC)
b = CC \ 65536
g = (CC Mod 65536) \ 256
r = CC Mod 256
Me.Caption = CC & "--" & r & g & b & "--" & y1
Picture1.DrawWidth = 4
Picture1.Line (x1, y1)-(x1 + 1, y1 + 1), RGB(r, g, b)
If y1 < ProgressBar1.Max Then ProgressBar1.Value = y1
If y1 >= ProgressBar1.Max Then ProgressBar1.Visible = False
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Winsock2.GetData y1, vbInteger

End Sub
Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)

Winsock3.GetData CC, vbLong

CC = Abs(CC)
b = CC \ 65536
g = (CC Mod 65536) \ 256
r = CC Mod 256
Me.Caption = CC & "--" & r & g & b & "--" & y1
Picture1.DrawWidth = 4
Picture1.Line (x1, y1)-(x1 + 1, y1 + 1), RGB(r, g, b)
If y1 < ProgressBar1.Max Then ProgressBar1.Value = y1
If y1 >= ProgressBar1.Max Then ProgressBar1.Visible = False

End Sub
Private Sub WinUdpa_DataArrival(ByVal bytesTotal As Long)
t = 0
voice.URL = App.Path & "\msg.wav"
Timer5.Enabled = True


WinUdpa.GetData strdata, vbString
If strdata = "置顶" Then
Timer1.Enabled = True
time = 0
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If

If strdata = "您的朋友开始三子棋" Then C.Enabled = True
If strdata = "您的朋友结束三子棋" Then C.Enabled = False
If strdata = "变色龙" Then
Timer2.Enabled = True
time1 = 0
End If

If strdata = "C1" Then
C01 = 1
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
If player = 0 Then C1.Caption = "X"
If player = 1 Then C1.Caption = "O"
End If

If strdata = "C2" Then
C02 = 1
If player = 0 Then C2.Caption = "X"
If player = 1 Then C2.Caption = "O"
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
End If

If strdata = "C3" Then
C03 = 1
If player = 0 Then C3.Caption = "X"
If player = 1 Then C3.Caption = "O"
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
End If

If strdata = "C4" Then
C04 = 1
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
If player = 0 Then C4.Caption = "X"
If player = 1 Then C4.Caption = "O"
End If

If strdata = "C5" Then
C05 = 1
If player = 0 Then C5.Caption = "X"
If player = 1 Then C5.Caption = "O"
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
End If

If strdata = "C6" Then
C06 = 1
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
If player = 0 Then C6.Caption = "X"
If player = 1 Then C6.Caption = "O"
End If

If strdata = "C7" Then
C07 = 1
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
If player = 0 Then C7.Caption = "X"
If player = 1 Then C7.Caption = "O"
End If

If strdata = "C8" Then
C08 = 1
If player = 0 Then C8.Caption = "X"
If player = 1 Then C8.Caption = "O"
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
End If

If strdata = "C9" Then
C09 = 1
If player = 0 Then C9.Caption = "X"
If player = 1 Then C9.Caption = "O"
If Not C01 = 1 Then C1.Enabled = True
If Not C02 = 1 Then C2.Enabled = True
If Not C03 = 1 Then C3.Enabled = True
If Not C04 = 1 Then C4.Enabled = True
If Not C05 = 1 Then C5.Enabled = True
If Not C06 = 1 Then C6.Enabled = True
If Not C07 = 1 Then C7.Enabled = True
If Not C08 = 1 Then C8.Enabled = True
If Not C09 = 1 Then C9.Enabled = True
End If
 If (C1.Caption + C2.Caption + C3.Caption = "OOO") Or (C1.Caption + C5.Caption + C9.Caption = "OOO") Or (C1.Caption + C4.Caption + C7.Caption = "OOO") Or (C4.Caption + C5.Caption + C6.Caption = "OOO") Or (C7.Caption + C8.Caption + C9.Caption = "OOO") Or (C5.Caption + C2.Caption + C8.Caption = "OOO") Or (C3.Caption + C6.Caption + C9.Caption = "OOO") Or (C3.Caption + C5.Caption + C7.Caption = "OOO") Then
  If player = 1 Then
  MsgBox "你赢了!"
  WinUdpa.SendData ""
  C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
C.Caption = "退出并重新进入以重新开始"
End If
If player = 0 Then
  MsgBox "你输了!"
  WinUdpa.SendData ""
  C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
C.Caption = "退出并重新进入以重新开始"
End If
End If
If (C1.Caption + C2.Caption + C3.Caption = "XXX") Or (C1.Caption + C5.Caption + C9.Caption = "XXX") Or (C1.Caption + C4.Caption + C7.Caption = "XXX") Or (C4.Caption + C5.Caption + C6.Caption = "XXX") Or (C7.Caption + C8.Caption + C9.Caption = "XXX") Or (C5.Caption + C2.Caption + C8.Caption = "XXX") Or (C3.Caption + C6.Caption + C9.Caption = "XXX") Or (C3.Caption + C5.Caption + C7.Caption = "XXX") Then
  If player = 0 Then
  MsgBox "你赢了!"
  WinUdpa.SendData ""
  C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
C.Caption = "退出并重新进入以重新开始"
End If
If player = 1 Then
  MsgBox "你输了!"
  C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
C.Caption = "退出并重新进入以重新开始"
End If
End If
If C01 + C02 + C03 + C04 + C5 + C06 + C07 + C08 + C09 = 9 And Not ((C1.Caption + C2.Caption + C3.Caption = "XXX") Or (C1.Caption + C5.Caption + C9.Caption = "XXX") Or (C1.Caption + C4.Caption + C7.Caption = "XXX") Or (C4.Caption + C5.Caption + C6.Caption = "XXX") Or (C7.Caption + C8.Caption + C9.Caption = "XXX") Or (C5.Caption + C2.Caption + C8.Caption = "XXX") Or (C3.Caption + C6.Caption + C9.Caption = "XXX") Or (C3.Caption + C5.Caption + C7.Caption = "XXX")) And Not ((C1.Caption + C2.Caption + C3.Caption = "OOO") Or (C1.Caption + C5.Caption + C9.Caption = "OOO") Or (C1.Caption + C4.Caption + C7.Caption = "OOO") Or (C4.Caption + C5.Caption + C6.Caption = "OOO") Or (C7.Caption + C8.Caption + C9.Caption = "OOO") Or (C5.Caption + C2.Caption + C8.Caption = "OOO") Or (C3.Caption + C6.Caption + C9.Caption = "OOO") Or (C3.Caption + C5.Caption + C7.Caption = "OOO")) Then

If Not C1.Caption + C2.Caption + C3.Caption + C4.Caption + C5.Caption + C6.Caption = "" Then
MsgBox "和棋"
  WinUdpa.SendData ""
  C1.Enabled = False
C2.Enabled = False
C3.Enabled = False
C4.Enabled = False
C5.Enabled = False
C6.Enabled = False
C7.Enabled = False
C8.Enabled = False
C9.Enabled = False
C.Caption = "退出并重新进入以重新开始"
End If
End If

If strdata = "[图片文件]" Then
Picture1.Visible = True
Text1.Left = 4440
Text1.Width = 2000
ProgressBar1.Visible = True
x1 = 0
y1 = 0
Picture1.Cls

End If





  
 
 

If strdata = "暴风雨" Then
Timer3.Enabled = True
time2 = 0
End If
If strdata = "风格化" Then
Timer4.Enabled = True
time3 = 0
End If
If strdata = "beep" Then
MsgBox ("请注意！")
End If
If strdata = "意大利化" Then
Text1.Font.Italic = True
End If

If strdata = "粗体" Then
Text1.Font.Bold = True
End If
If strdata = "下划线" Then
Text1.Font.Underline = True
End If

If strdata = "1.2x" Then
Text1.Font.Size = Text1.Font.Size * 1.2
End If
 If strdata = "1.5x" Then
 Text1.Font.Size = Text1.Font.Size * 1.5
 End If
 If strdata = "撤销" Then
Text1.Font.Italic = False
Text1.Font.Bold = False
Text1.Font.Underline = False
Text1.Font.Size = 10
End If
 



If Form7.Text1.Text = "" Then
Text1.Text = Text1.Text & vbCrLf & "[" & Form2.Text1.Text & " | " & Now & "]：" & strdata
other = Form2.Text1.Text
Else: Text1.Text = Text1.Text & vbCrLf & "[" & Form7.Text1.Text & " | " & Now & "]：" & strdata
other = Form7.Text1.Text
End If
End Sub

 Private Sub C_Click()
Dim a As Integer

Randomize
a = Int(Rnd * (1 - 0 + 1))
C.Caption = a
C.Enabled = False

WinUdpa.SendData C.Caption
C.FontSize = 12


C.Caption = a & vbCrLf & "耐心等待对方响应"
 While Not strdata = "0" And Not strdata = "1"
 WinUdpa.GetData strdata, vbString
 
 Wend

If strdata = "0" And Not Int(strdata) = a Then
C.Caption = "后手"
C.Enabled = False
C.FontSize = 20
player = 1
End If

If strdata = "1" And Not Int(strdata) = a Then
C.Caption = "先手"
C.Enabled = False
C.FontSize = 20
player = 0
C1.Enabled = True
C2.Enabled = True
C3.Enabled = True
C4.Enabled = True
C5.Enabled = True
C6.Enabled = True
C7.Enabled = True
C8.Enabled = True
C9.Enabled = True
End If

 If Int(strdata) = a Then
C.Caption = "重来"
C.Enabled = True
C.FontSize = 20
strdata = ""
End If

End Sub

Private Sub mnuCopy_Click()

Clipboard.Clear

Clipboard.SetText Text2.SelText

 

End Sub

 



 

Private Sub mnuCut_Click()

Clipboard.Clear

Clipboard.SetText Text2.SelText

 
 

Text2.SelText = ""

End Sub



 

Private Sub mnuSelectAll_Click()

 

Text2.SelStart = 0

Text2.SelLength = Len(Text2.Text)

End Sub

 



 

Private Sub mnuPaste_Click()



Text2.SelText = Clipboard.GetText

 

End Sub

