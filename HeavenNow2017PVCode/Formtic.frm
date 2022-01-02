VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Formtic 
   Caption         =   "三子棋"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form9"
   ScaleHeight     =   4830
   ScaleWidth      =   5715
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock WinUdpA 
      Left            =   4800
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton C 
      BackColor       =   &H000000FF&
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
      Left            =   4440
      Picture         =   "Formtic.frx":0000
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton C3 
      Enabled         =   0   'False
      Height          =   1695
      Index           =   7
      Left            =   2760
      TabIndex        =   8
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton C4 
      Enabled         =   0   'False
      Height          =   1695
      Index           =   6
      Left            =   -120
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton C5 
      Enabled         =   0   'False
      Height          =   1695
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton C6 
      Enabled         =   0   'False
      Height          =   1695
      Index           =   4
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton C7 
      Enabled         =   0   'False
      Height          =   1455
      Index           =   3
      Left            =   -120
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton C8 
      Enabled         =   0   'False
      Height          =   1455
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton C9 
      Enabled         =   0   'False
      Height          =   1455
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton C2 
      Enabled         =   0   'False
      Height          =   1695
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton C1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   4200
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line5 
      X1              =   4200
      X2              =   4200
      Y1              =   0
      Y2              =   4800
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   2760
      X2              =   2760
      Y1              =   0
      Y2              =   4800
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1320
      Y1              =   0
      Y2              =   4800
   End
End
Attribute VB_Name = "Formtic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub C_Click()
Randomize
a = Int(Rnd * (100 - 1 + 1)) + 1
C.Caption = a
C.Enabled = False

WinUdpA.SendData C.Caption
C.FontSize = 12


C.Caption = a & vbCrLf & "耐心等待对方响应"
Dim strdata As String
WinUdpA.GetData strdata, vbString
While strdata = ""
WinUdpA.GetData strdata, vbString
Wend

If Int(strdata) > a Then
C.Caption = "后手"
C.Enabled = False
End If

If Int(strdata) < a Then
C.Caption = "先手"
C.Enabled = False
End If

If Int(strdata) = a Then C.Caption = "重来"

C.FontSize = 20

End Sub

Private Sub Form_Load()

WinUdpA.LocalPort = Form2.Text3.Text
WinUdpA.RemoteHost = Form2.Text1.Text
WinUdpA.RemotePort = Form2.Text2.Text
Form1.WinUdpA.Close
WinUdpA.Bind
End Sub

Private Sub WinUdpA_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
WinUdpA.GetData strdata, vbString
End Sub

