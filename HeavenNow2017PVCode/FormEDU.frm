VERSION 5.00
Begin VB.Form FormEDU 
   Caption         =   "选择"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "FormEDU.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2085
   ScaleWidth      =   3975
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "单片机串口通讯"
      Height          =   1455
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Math Buster"
      Height          =   1455
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "电子化学实验室"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FormEDU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FormEC.Show
Me.Hide

End Sub

Private Sub Command2_Click()
Math.Show
Me.Hide

End Sub

Private Sub Command3_Click()
FormCom.Show
Me.Hide

End Sub

Private Sub Command4_Click()
Formmenu.Show
Me.Hide

End Sub
