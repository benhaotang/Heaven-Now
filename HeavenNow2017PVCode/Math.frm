VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Math 
   Caption         =   "Form9"
   ClientHeight    =   4425
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6930
   LinkTopic       =   "Form9"
   ScaleHeight     =   4425
   ScaleWidth      =   6930
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RSH 
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Math.frx":0000
   End
   Begin RichTextLib.RichTextBox LSH 
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Math.frx":008F
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSScriptControlCtl.ScriptControl msscript1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Text            =   "x^2"
      Top             =   4080
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox calc 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Math.frx":011E
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      DrawWidth       =   2
      Height          =   3735
      Left            =   360
      ScaleHeight     =   3675
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   3480
         Y1              =   0
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   0
         X2              =   6240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lb 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "y="
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnustart 
      Caption         =   "开始"
      Begin VB.Menu windows 
         Caption         =   "窗口调节"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnusave 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu gaoji 
      Caption         =   "高级"
      Begin VB.Menu wei 
         Caption         =   "未整理方程输入"
         Shortcut        =   ^I
      End
      Begin VB.Menu dong 
         Caption         =   "动态方程"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu guanyu 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim y As Double
Dim temp As Double
Dim i As Double
Dim col As Integer
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim en As Integer
Dim biaodashi As String
Dim todo As Integer
Dim st As Integer
Dim xm As Double
Dim YM As Double
Dim ste As String
Dim xmin As Double
Dim ymin As Double
Dim xmax As Double
Dim ymax As Double
Dim txmin As Double
Dim tymin As Double
Dim txmax As Double
Dim tymax As Double





Private Sub Command1_Click()

ProgressBar1.Visible = True
en = 0
Load lb(col)
lb(col).BackStyle = 0
Set lb(col).Container = Picture1
calc.Text = Text1.Text
calc.Find "X"
calc.Text = Replace(calc.Text, "X", "x")
Text1.Text = calc.Text

r = 255 - 50 * (col Mod 5)
g = 50 * (col Mod 5)
b = 255 - 10 * (col Mod 25)
If r < 0 Then r = 255
If g > 255 Then g = 0
If b < 0 Then b = 255
For i = xmin - 0.01 To xmax Step 0.01 * (xmax - xmin) / 20
On Error Resume Next
ProgressBar1.Value = (i + xmax + 0.01) / (xmax - xmin) * 101
If y < ymax And y > ymin And en = 0 Then
lb(col).Left = i
lb(col).Top = y
lb(col).Visible = True
lb(col).Width = 2 * (xmax - xmin) / 20
lb(col).Height = 0.5 * (ymax - ymin) / 10
lb(col).Caption = Text1.Text
lb(col).ForeColor = RGB(r, g, b)
en = 1
End If

On Error GoTo 999
calc.Text = Text1.Text
temp = y
Label1.Caption = i
calc.Find "x"
calc.Text = Replace(calc.Text, "x", "(" & i & ")")
On Error Resume Next
y = msscript1.Eval(calc.Text)
Label2.Caption = y
On Error Resume Next
Picture1.Line (i - 0.01, temp)-(i, y), RGB(r, g, b)

999:
Err.Clear


Next i
col = col + 1
ProgressBar1.Visible = False

End Sub

Private Sub dong_Click()
todo = 0
ProgressBar1.Visible = True
biaodashi = InputBox("请输入表达式(使用a作为可变变量)：y=", "Heaven Now 数学终结者：动态表达式", todo = 1)
ste = InputBox("步长？范围？（请输入'step,Amin,Amax'）", "Heaven Now 数学终结者：动态表达式", todo = 1)

biao = Split(ste, ",")
If UBound(biao) - LBound(biao) < 2 Then
MsgBox ("输入错误！")
todo = 1
End If
If todo = 0 Then
st = CDbl(biao(0))
xm = CDbl(biao(1))
YM = CDbl(biao(2))
en = 0
Load lb(col)
lb(col).BackStyle = 0
Set lb(col).Container = Picture1
calc.Text = biaodashi
calc.Find "X"
calc.Text = Replace(calc.Text, "X", "x")
calc.Find "A"
calc.Text = Replace(calc.Text, "A", "a")

biaodashi = calc.Text

r = 255 - 50 * (col Mod 5)
g = 50 * (col Mod 5)
b = 255 - 10 * (col Mod 25)
If r < 0 Then r = 255
If g > 255 Then g = 0
If b < 0 Then b = 255
For j = xm To YM Step st
For i = xmin - 0.01 To xmax Step 0.01 * (xmax - xmin) / 20
ProgressBar1.Value = (i + xmax + 0.01) / ((xmax - xmin) * CInt((YM - xm) / st)) * 101
If y < 5 And y > -5 And en = 0 Then
lb(col).Left = i
lb(col).Top = y
lb(col).Visible = True
lb(col).Width = 2 * (xmax - xmin) / 20
lb(col).Height = 0.5 * (ymax - ymin) / 10
lb(col).Caption = biaodashi
lb(col).ForeColor = RGB(r, g, b)
en = 1
End If

On Error GoTo 999
calc.Text = biaodashi
temp = y
Label1.Caption = i
calc.Find "x"
calc.Text = Replace(calc.Text, "x", "(" & i & ")")
calc.Find "a"
calc.Text = Replace(calc.Text, "a", "(" & j & ")")

On Error Resume Next
y = msscript1.Eval(calc.Text)
Label2.Caption = y
On Error Resume Next
Picture1.Line (i - 0.01, temp)-(i, y), RGB(r, g, b)

999:
Err.Clear


Next i
Picture1.DrawWidth = Picture1.DrawWidth + 1
Next j

col = col + 1
End If
Picture1.DrawWidth = 2
ProgressBar1.Visible = False
End Sub

Private Sub exit_Click()
Me.Hide
FormEDU.Show

End Sub

Private Sub Form_Load()
col = 1
Picture1.Width = 2.6 * Math.Width
Picture1.Height = 2 * Math.Height
Text1.Top = Picture1.Top + 0.56 * Picture1.Width + 50
Text1.Width = 0.6 * Picture1.Width
Command1.Left = Text1.Left + Text1.Width
Command1.Top = Text1.Top
Command1.Width = 0.4 * Picture1.Width
ProgressBar1.Width = Command1.Width
ProgressBar1.Top = Command1.Top
ProgressBar1.Left = Command1.Left
ProgressBar1.Visible = False
Picture1.Scale (-10, 5)-(10, -5)
Picture1.Line (1000, 0)-(-1000, 0), RGB(0, 0, 0)
Picture1.Line (0, 1000)-(0, -1000), RGB(0, 0, 0)
Line2.x1 = -10
Line2.x2 = 10
Line1.y1 = -5
Line1.y2 = 5
Line1.Visible = False
Line2.Visible = False
Label3.Top = Text1.Top

Math.Caption = "Heaven Now 数学终结者"
xmin = -10
xmax = 10
ymin = -5
ymax = 5


End Sub


Private Sub guanyu_Click()
mathabout.Show
End Sub

Private Sub mnusave_Click()
CommonDialog1.Filter = "BMP Files|*.bmp"


CommonDialog1.ShowSave
Picture1.AutoRedraw = True
SavePicture Picture1.Image, CommonDialog1.FileName
End Sub

Private Sub Picture1_DblClick()
Picture1.Cls
For i = 0 To lb.UBound Step 1
lb(i).Visible = False
Next i
Picture1.Line (1000, 0)-(-1000, 0), RGB(0, 0, 0)
Picture1.Line (0, 1000)-(0, -1000), RGB(0, 0, 0)

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Line1.Visible = True
Line2.Visible = True
Picture1.ToolTipText = X & "," & y
Line1.x1 = X
Line1.x2 = X
Line2.y1 = y
Line2.y2 = y

End Sub

Private Sub wei_Click()
ProgressBar1.Visible = True
todo = 0
en = 0
biaodashi = InputBox("请输入表达式(如x*y=1/x)：", "Heaven Now 数学终结者：未整理表达式", todo = 1)
biao = Split(biaodashi, "=")
If UBound(biao) - LBound(biao) = 0 Then
MsgBox ("输入错误01：等号未找到！")
todo = 1
End If
If todo = 0 Then
r = 255 - 50 * (col Mod 5)
g = 50 * (col Mod 5)
b = 255 - 10 * (col Mod 25)
If r < 0 Then r = 255
If g > 255 Then g = 0
If b < 0 Then b = 255
LSH.Text = biao(0)
RSH.Text = biao(1)
LSH.Find "Y"
LSH.Text = Replace(LSH.Text, "Y", "y")
LSH.Find "X"
LSH.Text = Replace(LSH.Text, "X", "x")
RSH.Find "Y"
RSH.Text = Replace(RSH.Text, "Y", "y")
RSH.Find "X"
RSH.Text = Replace(RSH.Text, "X", "x")
Load lb(col)
lb(col).Caption = LSH.Text & "=" & RSH.Text
lb(col).ForeColor = RGB(r, g, b)

biao = Split(lb(col).Caption, "=")
LSH.Text = biao(0)
RSH.Text = biao(1)
For xi = xmin To xmax Step 0.05 * (xmax - xmin) / 20
 For yt = ymin To ymax Step 0.05 * (ymax - ymin) / 10
 biao = Split(lb(col).Caption, "=")
LSH.Text = biao(0)
RSH.Text = biao(1)
 LSH.Find "y"
LSH.Text = Replace(LSH.Text, "y", yt)
LSH.Find "x"
LSH.Text = Replace(LSH.Text, "x", xi)
RSH.Find "y"
RSH.Text = Replace(RSH.Text, "y", yt)
RSH.Find "x"
RSH.Text = Replace(RSH.Text, "x", xi)
On Error Resume Next
lhs = Round(msscript1.Eval(LSH.Text), 4)
On Error Resume Next
rhs = Round(msscript1.Eval(RSH.Text), 4)

If Abs(lhs - rhs) < 0.035 Then
On Error Resume Next
Picture1.Line (xi - 0.05, yt)-(xi, yt), RGB(r, g, b)
If en = 0 Then
lb(col).Left = xi
lb(col).Top = yt
lb(col).Visible = True
en = 1
End If
End If

Next yt
On Error Resume Next
ProgressBar1.Value = (xi + xmax) / (xmax - xmin) * 101
Next xi
col = col + 1
lb(col).Width = 4 * (xmax - xmin) / 20
lb(col).Height = 0.5 * (ymax - ymin) / 10
End If
ProgressBar1.Visible = False

End Sub



Private Sub windows_Click()
todo = 0
txmin = xmin
txmax = xmax
tymin = ymin
tymax = ymax
biaodashi = InputBox("请输入视窗范围：(Xmin,Ymin,Xmax,Ymax)", "Heaven Now 数学终结者：视窗调节", todo = 1)
biao = Split(biaodashi, ",")
If UBound(biao) - LBound(biao) < 3 Then
MsgBox ("输入错误！")
todo = 1
End If
If todo = 0 Then
xmin = biao(0)
xmax = biao(2)
ymin = biao(1)
ymax = biao(3)
Picture1.AutoRedraw = True
SavePicture Picture1.Image, App.Path & "\recenttemp.BMP"
Picture1.Cls

Picture1.Scale (xmin, ymax)-(xmax, ymin)
Picture1.Line (xmax, 0)-(xmin, 0), RGB(0, 0, 0)
Picture1.Line (0, ymax)-(0, ymin), RGB(0, 0, 0)
Picture1.PaintPicture LoadPicture(App.Path & "\recenttemp.BMP"), txmin, tymax, txmax - txmin, tymax - tymin


End If

End Sub
