VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "状态 - 加载中..."
   ClientHeight    =   210
   ClientLeft      =   1770
   ClientTop       =   4485
   ClientWidth     =   1905
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   210
   ScaleWidth      =   1905
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   120
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "正在加载中...请稍候"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()

i = i + 1

If i >= 5 Then Form5.Hide
Timer1.Enabled = False
Form4.Visible = True
Form5.Hide


End Sub

