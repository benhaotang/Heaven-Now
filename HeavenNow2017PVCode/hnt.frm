VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormG 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Heaven Now Graffiti "
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "hnt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   15240
   Begin VB.CommandButton Command12 
      Caption         =   "字"
      Height          =   255
      Left            =   10680
      TabIndex        =   38
      Top             =   9720
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "圆"
      Height          =   255
      Left            =   10320
      TabIndex        =   35
      Top             =   9720
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "实矩"
      Height          =   255
      Left            =   9840
      TabIndex        =   31
      Top             =   9720
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "空矩"
      Height          =   255
      Left            =   9360
      TabIndex        =   30
      Top             =   9720
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "直线"
      Height          =   255
      Left            =   8880
      TabIndex        =   26
      Top             =   9720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   7320
      TabIndex        =   25
      Text            =   "请先输入图片地址"
      Top             =   9960
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "再按加载图片"
      Height          =   345
      Left            =   7320
      TabIndex        =   24
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "蓝"
      Height          =   255
      Left            =   8280
      TabIndex        =   23
      Top             =   9720
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "黑"
      Height          =   255
      Left            =   7920
      TabIndex        =   22
      Top             =   9720
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "白"
      Height          =   255
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   21
      Top             =   9720
      Width           =   375
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   10200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   14040
      TabIndex        =   15
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "截图并保存"
      Height          =   495
      Left            =   11760
      TabIndex        =   11
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      Height          =   495
      Index           =   0
      Left            =   11760
      TabIndex        =   10
      Top             =   10200
      Width           =   1335
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   10200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   10200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   10200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   9720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   3
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label18 
      Caption         =   "控制栏目"
      Height          =   255
      Left            =   11400
      TabIndex        =   37
      Top             =   9480
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "特殊工具"
      Height          =   255
      Left            =   9000
      TabIndex        =   36
      Top             =   9480
      Width           =   735
   End
   Begin VB.Line Line11 
      X1              =   11040
      X2              =   11040
      Y1              =   9600
      Y2              =   10800
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      X1              =   11160
      X2              =   11640
      Y1              =   10200
      Y2              =   10680
   End
   Begin VB.Label Label16 
      Caption         =   "先双击确定基准（粉红），再单击确定目标（淡黄）。"
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      Top             =   10200
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Y2"
      Height          =   255
      Left            =   10680
      TabIndex        =   33
      Top             =   9960
      Width           =   375
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   32
      Top             =   9960
      Width           =   495
   End
   Begin VB.Line Line10 
      X1              =   9720
      X2              =   10200
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Y1"
      Height          =   255
      Left            =   9360
      TabIndex        =   29
      Top             =   9960
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "X1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   28
      Top             =   9960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      Height          =   495
      Left            =   11160
      Shape           =   5  'Rounded Square
      Top             =   10200
      Width           =   495
   End
   Begin VB.Label Label7 
      Height          =   255
      Index           =   1
      Left            =   11040
      TabIndex        =   27
      Top             =   9960
      Width           =   255
   End
   Begin VB.Label Label7 
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   13
      Top             =   9960
      Width           =   255
   End
   Begin VB.Line Line9 
      X1              =   7320
      X2              =   7320
      Y1              =   9600
      Y2              =   10800
   End
   Begin VB.Label Label11 
      Caption         =   "背景"
      Height          =   255
      Left            =   7440
      TabIndex        =   20
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "粗细"
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "样式"
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   9840
      Width           =   375
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   13560
      X2              =   13560
      Y1              =   9840
      Y2              =   10200
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   13440
      TabIndex        =   14
      Top             =   9720
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   615
      Left            =   13080
      Shape           =   3  'Circle
      Top             =   9960
      Width           =   975
   End
   Begin VB.Line Line7 
      BorderWidth     =   5
      X1              =   11160
      X2              =   11520
      Y1              =   9720
      Y2              =   10080
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   11400
      TabIndex        =   12
      Top             =   9600
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   375
      Left            =   11160
      Top             =   9720
      Width           =   375
   End
   Begin VB.Line Line5 
      X1              =   8760
      X2              =   15240
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Label Label3 
      Caption         =   "画笔风格"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   9480
      Width           =   735
   End
   Begin VB.Line Line4 
      X1              =   8880
      X2              =   8880
      Y1              =   9600
      Y2              =   10800
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   8880
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Label Label2 
      Caption         =   "颜色设定"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   9480
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   5880
      X2              =   5880
      Y1              =   9600
      Y2              =   10800
   End
   Begin VB.Line Line1 
      X1              =   -240
      X2              =   6000
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Label Label1 
      Caption         =   "Blue(0~255)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   5
      Top             =   9720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Red(0~255)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Green(0~255)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   9720
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Made by DYPro Mike TANG,copyright DYPro Program Studio(2012~2014) Thanks for using Heaven Now"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   9240
      Width           =   8535
   End
   Begin VB.Label Label5 
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   9600
      Width           =   15255
   End
End
Attribute VB_Name = "FormG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long
Dim b As Long
Dim C As Long
Dim d As Long



Private Sub Chart3D1_Click()
Chart3D1.LoadURL
End Sub

Private Sub Command10_Click()
Line (a, b)-(C, d), RGB(Slider1, Slider3, Slider2), BF
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "请确保已经选好了地址，再执行！警告！否则程序崩溃！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command11_Click()
Circle (a, b), Sqr(Abs(C - a) * Abs(C - a) + Abs(b - d) * Abs(b - d)), RGB(Slider1, Slider3, Slider2)
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "警告！画过大的园程序会崩溃！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "文字的添加只有专业版才有哦，快激活专业版吧"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "使用左键拖动选取，使用右键重新拾取。警告！一旦你退出截图模式，以前所画的都会清空，请慎重使用！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command4_Click()
FormG.BackColor = RGB(255, 255, 255)
Label9.BackColor = RGB(255, 255, 255)
Label9.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Command5_Click()
FormG.BackColor = RGB(0, 0, 0)
Label9.BackColor = RGB(0, 0, 0)
Label9.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "警告！更换背景会清空原来所做的画，图片输错地址会使程序崩溃！谨慎使用！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command6_Click()
FormG.BackColor = RGB(0, 0, 255)
Label9.BackColor = RGB(0, 0, 225)
Label9.ForeColor = RGB(255, 255, 255)
End Sub



Private Sub Command7_Click()
FormG.Picture = LoadPicture(Text1.Text)
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "警告！更换背景会清空原来所做的画，图片输错地址会使程序崩溃！谨慎使用！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command8_Click()
Line (a, b)-(C, d), RGB(Slider1, Slider3, Slider2)
End Sub




Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "请确保已经选好了地址，再执行！警告！否则程序崩溃！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Command9_Click()
Line (a, b)-(C, d), RGB(Slider1, Slider3, Slider2), B
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "请确保已经选好了地址，再执行！警告！否则程序崩溃！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub EnterpriseLogonControl1_GotFocus()

End Sub

Private Sub Form_Click()
C = FormG.CurrentX
d = FormG.CurrentY
Label14.Caption = C
Label15.Caption = d
End Sub

Private Sub Form_dblClick()
a = FormG.CurrentX
b = FormG.CurrentY
Label12.Caption = a
Label13.Caption = b
End Sub

Private Sub Form_Load()
Slider5.Max = 50
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormG.CurrentX = X
FormG.CurrentY = Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 And Slider4 = 1 Then FormG.Line -(X, Y), RGB(Slider1, Slider3, Slider2)
If Button = 1 And Slider4 = 2 Then FormG.Line -(X, Y), RGB(Slider1, Slider3, Slider2), B
If Button = 1 And Slider4 = 3 Then FormG.Line -(X, Y), RGB(Slider1, Slider3, Slider2), BF
FormG.DrawWidth = Slider5.Value
Label9.Caption = "Made by DYPro Mike TANG,copyright DYPro Program Studio(2012~2016) Thanks for using Heaven Now"

End Sub
Private Sub Command2_Click()
Label9.Caption = "使用左键拖动选取，使用右键重新拾取。警告！一旦你退出截图模式，以前所画的都会清空，请慎重使用！"
Label9.ForeColor = RGB(255, 0, 0)
Shell "SnippingTool.exe"

End Sub
Private Sub Command1_Click(Index As Integer)
FormG.Cls
End Sub

Private Sub Command3_Click()
FormG.Visible = False
Formmenu.Visible = True
End Sub

Private Sub Label16_Click()
Label9.Caption = "请确保已经选好了地址，再执行！警告！否则程序崩溃！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub MSChart1_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

Private Sub Slider5_Click()
Label9.Caption = "注意！使用较粗的画笔需要放慢速度画，不然会产生断点！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Slider5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "注意！使用较粗的画笔需要放慢画速，不然会产生断点！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "警告！更换背景会清空原来所做的画，图片输错地址会使程序崩溃！谨慎使用！"
Label9.ForeColor = RGB(255, 0, 0)
End Sub
