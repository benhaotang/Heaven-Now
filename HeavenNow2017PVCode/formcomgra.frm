VERSION 5.00
Begin VB.Form formcomgra 
   Caption         =   "串口绘图器"
   ClientHeight    =   5640
   ClientLeft      =   4395
   ClientTop       =   3990
   ClientWidth     =   8595
   LinkTopic       =   "Form9"
   Picture         =   "formcomgra.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   615
      Left            =   9960
      TabIndex        =   8
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "测量"
      Height          =   615
      Left            =   8640
      Picture         =   "formcomgra.frx":1302
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   4050
      Left            =   8640
      Pattern         =   "*.bmp"
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         FillColor       =   &H00FFFFFF&
         Height          =   5175
         Left            =   120
         Picture         =   "formcomgra.frx":2604
         ScaleHeight     =   5115
         ScaleWidth      =   8235
         TabIndex        =   1
         ToolTipText     =   "双击以清除实验数据"
         Top             =   240
         Width           =   8295
         Begin VB.Timer Timer1 
            Interval        =   10
            Left            =   240
            Top             =   120
         End
         Begin VB.Label time 
            BackColor       =   &H8000000D&
            Caption         =   "0"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1320
            TabIndex        =   4
            Top             =   0
            Width           =   495
         End
         Begin VB.Line Line15 
            X1              =   0
            X2              =   8280
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line14 
            X1              =   0
            X2              =   8280
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line13 
            X1              =   0
            X2              =   8280
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line12 
            X1              =   0
            X2              =   8280
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line11 
            X1              =   0
            X2              =   8280
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line10 
            X1              =   0
            X2              =   8280
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   8160
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   8280
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Line Line7 
            X1              =   0
            X2              =   8280
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   8280
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   8160
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   8280
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   8280
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label YI 
            BackColor       =   &H80000009&
            Caption         =   "10"
            Height          =   255
            Left            =   7800
            TabIndex        =   3
            Top             =   4920
            Width           =   495
         End
         Begin VB.Label YM 
            BackColor       =   &H80000009&
            Caption         =   "10"
            Height          =   375
            Left            =   7800
            TabIndex        =   2
            Top             =   0
            Width           =   495
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   0
            X2              =   8280
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000D&
            BorderWidth     =   2
            X1              =   1320
            X2              =   1320
            Y1              =   0
            Y2              =   5040
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Pervious"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu y 
      Caption         =   "Y坐标"
   End
   Begin VB.Menu T 
      Caption         =   "T长度"
   End
   Begin VB.Menu Command3 
      Caption         =   "高级菜单"
   End
End
Attribute VB_Name = "formcomgra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim an As Integer
Dim tt As Integer
Dim tn As Integer
Dim val As Double
Dim wi As Integer
Dim n As Double
Dim m As Double


Private Sub Command1_Click()
FormCom.MSComm1.PortOpen = True
Timer1.Enabled = True
FormA.Command1.Enabled = False
FormA.Command2.Enabled = True
File1.Enabled = False
Command2.Enabled = True
Command1.Enabled = False

End Sub

Private Sub Command2_Click()
FormCom.MSComm1.PortOpen = False
File1.Enabled = True
Timer1.Enabled = False
FormA.Command1.Enabled = True
FormA.Command2.Enabled = False
Command2.Enabled = False
Command1.Enabled = True


End Sub

Private Sub Command3_Click()
Timer2.Enabled = True
End Sub

Private Sub File1_DblClick()
Picture1.Cls
Picture1.Picture = LoadPicture(App.Path & "\" & File1.FileName)

End Sub

Private Sub Form_Load()
Picture1.Scale (0, 10)-(1000, -10)
tt = 1
tn = 0
End Sub


Private Sub Picture1_dblClick()
Picture1.Cls
Label2.Caption = 0

End Sub

Private Sub T_Click()
tt = Int(Int(InputBox("T length", "Heaven Now 串口通讯：Tchange")) / 10)
If tt < 1 Then
MsgBox ("输入错误-Heaven Now串口通讯:输入数值过小，请重试！")
tt = 1
End If

End Sub

Private Sub Timer1_Timer()
tn = tn + 1
If tn = tt Then
an = an + 1
If FormCom.MSComm1.InBufferCount > 0 Then
cr = Split(Trim(FormCom.MSComm1.Input), ".")
If UBound(cr) - LBound(cr) + 1 = 1 Then val = cr(0)
If UBound(cr) - LBound(cr) + 1 > 1 Then val = cr(0) & "." & cr(1)

End If

Picture1.Line (an - 1000 * Int(an / 1000), val)-(an - 1000 * Int(an / 1000), 0), RGB(255, 0, 0)
Line1.x1 = (an - 1000 * Int(an / 1000)) + 5
Line1.x2 = (an - 1000 * Int(an / 1000)) + 5
time.Left = (an - 1000 * Int(an / 1000)) + 5
time.Caption = an * tt


If an Mod 1000 = 1 Then
Picture1.AutoRedraw = True
SavePicture Picture1.Image, App.Path & "\" & Int(an / 1000) & ".BMP"
Picture1.Cls
End If
If an = 1000 Then Timer2.Enabled = True

tn = 0
End If

End Sub

Private Sub Timer2_Timer()
If wi < 10 Or wi = 10 Then
Me.Width = Me.Width + 200
wi = wi + 1


End If

If wi > 11 And wi < 23 Then
Me.Width = Me.Width - 200
wi = wi + 1
End If
If wi = 11 Then
Timer2.Enabled = False
wi = 12
Command3.Caption = "隐藏高级"


End If
If wi = 23 Then
Timer2.Enabled = False
wi = 0
Command3.Caption = "高级菜单"

End If
End Sub

Private Sub y_Click()
n = InputBox("YMax", "Heaven Now 串口通讯：Ychange")
YM.Caption = n
m = InputBox("YMin", "Heaven Now 串口通讯：Ychange")
YI.Caption = m
Picture1.Scale (0, n)-(1000, m)
End Sub
