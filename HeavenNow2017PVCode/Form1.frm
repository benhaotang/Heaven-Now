VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormC 
   Caption         =   "Heaven Now（天堂即刻办公）2016 个人中文版 "
   ClientHeight    =   5850
   ClientLeft      =   2505
   ClientTop       =   3120
   ClientWidth     =   10320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   5850
   ScaleWidth      =   10320
   Begin VB.Timer Timer3 
      Interval        =   40
      Left            =   1080
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2016"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "请等待。。。。。。。"
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   5280
      Width           =   1935
   End
End
Attribute VB_Name = "FormC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim n As Integer

Dim t  As Integer



Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置


Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为总在最前
Me.Width = 0
Me.Height = 0

i = 0
'Dim Y() As Byte

 '  Open App.Path & "\hn.swf" For Binary Access Write As #1
  '  Put #1, , Y
  ' Close #1
'  Do Until App.Path & "\hn.swf" <> ""
'      DoEvents
'   Loop
 ' ShockwaveFlash1.Movie = App.Path & "\hn.swf"

End Sub

Private Sub Timer1_Timer()


i = i + 25

ProgressBar1.Value = i
If i >= 100 Then
Me.Hide
Formmenu.Show

Timer1.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()

t = t + 1

Me.Width = Me.Width + 970
Me.Height = Me.Height + 585

If t = 10 Then
Timer3.Enabled = False
Timer1.Enabled = True
Timer2.Enabled = True
End If

End Sub
