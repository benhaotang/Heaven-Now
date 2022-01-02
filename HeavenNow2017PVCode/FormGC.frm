VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormGC 
   Caption         =   "溶液分析仪--电导度"
   ClientHeight    =   6675
   ClientLeft      =   2505
   ClientTop       =   3795
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10905
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Max             =   5
      SelStart        =   3
      Value           =   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "STOP"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6615
      Left            =   7080
      TabIndex        =   9
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   1
      Cols            =   4
   End
   Begin VB.Timer Timer2 
      Interval        =   25
      Left            =   1800
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   6795
      TabIndex        =   1
      ToolTipText     =   "双击以清除实验数据"
      Top             =   2880
      Width           =   6855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1200
      Top             =   2520
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "电导曲线"
      Height          =   3855
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   7095
   End
   Begin VB.Label Label9 
      Caption         =   "y轴比例/10E-n"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "V"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "西门子"
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
      Left            =   5880
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "欧姆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "*25微秒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Menu mnustart 
      Caption         =   "开始"
      Begin VB.Menu mnucutpic 
         Caption         =   "截图"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusave 
         Caption         =   "保存并处理"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuset 
      Caption         =   "设置"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "FormGC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSComm1.PortOpen = True
Timer1.Enabled = True
Slider1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
MSComm1.PortOpen = False
Slider1.Enabled = True
End Sub

Private Sub Form_Load()

Dim y1, y2 As Double
Dim x1, x2 As Integer

Picture1.Scale (0, 0.0015)-(1000, -0.0001)

MSFlexGrid1.TextMatrix(0, 0) = "时间/us"
MSFlexGrid1.TextMatrix(0, 1) = "电压/V"
MSFlexGrid1.TextMatrix(0, 2) = "电阻/Ω"
MSFlexGrid1.TextMatrix(0, 3) = "电导/S"



End Sub

Private Sub mnuabout_Click()
frmAboutGC.Show

End Sub

Private Sub mnucutpic_Click()
Shell "SnippingTool.exe"

End Sub

Private Sub mnusave_Click()
Picture1.AutoRedraw = True
SavePicture Picture1.Image, App.Path & "\recent.BMP"
FormG.Show
FormG.Picture = LoadPicture(App.Path & "\recent.BMP")
End Sub

Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
Case comEvCD
Case comEvCTS
Case comEvDSR
Case comEvRing

Case comEvReceive
Label1.Caption = Trim(MSComm1.Input)
Case comEvSend

End Select

End Sub

Private Sub Picture1_dblClick()
Picture1.Cls
Label2.Caption = 0

End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Slider1.Value = 0 Then Picture1.Scale (0, 1.5)-(1000, -0.0001)
If Slider1.Value = 1 Then Picture1.Scale (0, 0.15)-(1000, -0.0001)
If Slider1.Value = 2 Then Picture1.Scale (0, 0.015)-(1000, -0.0001)
If Slider1.Value = 3 Then Picture1.Scale (0, 0.0015)-(1000, -0.0001)
If Slider1.Value = 4 Then Picture1.Scale (0, 0.00015)-(1000, -0.0000001)
If Slider1.Value = 5 Then Picture1.Scale (0, 0.000015)-(1000, -0.0000001)


End Sub

Private Sub Timer1_Timer()
  x2 = Label2.Caption + 1
If MSComm1.InBufferCount > 0 Then
 Label1.Caption = CDbl(Trim(MSComm1.Input)) + 0
 If Label1.Caption = "" Then
 Label1.Caption = "0"
 End If
 
  y2 = (10000 * Label1.Caption / (5 - Label1.Caption + 0.01) + 0.01)

  Label6.Caption = 1 / y2
  End If
 
  If y2 = 100 Then
  Picture1.Line (Label2.Caption, y1)-(x2, 0), RGB(255, 0, 0)
 Else
  Picture1.Line (Label2.Caption, y1)-(x2, Label6.Caption), RGB(255, 0, 0)
 End If

 Label2.Caption = x2
 Label3.Caption = y2
 x1 = x2
 y1 = y2
  MSFlexGrid1.Rows = x2 + 1
 MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows * 25

 MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = Label1.Caption
 If Label6.Caption = 0 Then
 MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = 0
 Else
  MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = 1 / Label6.Caption
  End If
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = Label6.Caption

Dim a As Integer

Dim b As Integer

If x2 Mod 1000 = 1 Then
Picture1.Cls

a = 1000 * Int(x2 / 1000)
b = a + 1000
If Slider1.Value = 0 Then Picture1.Scale (a, 1.5)-(b, -0.0001)
If Slider1.Value = 1 Then Picture1.Scale (a, 0.15)-(b, -0.0001)
If Slider1.Value = 2 Then Picture1.Scale (a, 0.015)-(b, -0.0001)
If Slider1.Value = 3 Then Picture1.Scale (a, 0.0015)-(b, -0.0001)
If Slider1.Value = 4 Then Picture1.Scale (a, 0.00015)-(b, -0.0000001)
If Slider1.Value = 5 Then Picture1.Scale (a, 0.000015)-(b, -0.0000001)
End If
End Sub
