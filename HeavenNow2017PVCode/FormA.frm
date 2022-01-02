VERSION 5.00
Begin VB.Form FormA 
   Caption         =   "Adjust"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   1215
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "串口号"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   720
         TabIndex        =   6
         Text            =   "9600"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "暂停"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   720
         TabIndex        =   2
         Text            =   "3"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "开始"
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "波特率"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "COM"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "FormA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If FormCom.MSComm1.PortOpen = True Then FormCom.MSComm1.PortOpen = False
FormCom.MSComm1.Settings = Text3.Text & ",n,8,1"
FormCom.MSComm1.CommPort = Text1.Text
FormCom.MSComm1.PortOpen = True
FormCom.Timer1.Enabled = True
FormCom.Text2.Enabled = True
Me.Hide
FormCom.mnuhuitu.Enabled = True
Command2.Enabled = True
Command1.Enabled = False
formcomgra.Command1.Enabled = False
formcomgra.Command2.Enabled = True
End Sub

Private Sub Command2_Click()
FormCom.MSComm1.PortOpen = False
FormCom.Timer1.Enabled = False
FormCom.Text2.Enabled = False
FormCom.mnuhuitu.Enabled = False
Command2.Enabled = False
Command1.Enabled = True
formcomgra.Command1.Enabled = True
formcomgra.Command2.Enabled = False

End Sub

