VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "链接"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2445
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存设置"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "个人端口"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "对方端口"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
MsgBox ("Successful!")
Me.Hide
Form1.Command3.Enabled = True
    
End Sub


