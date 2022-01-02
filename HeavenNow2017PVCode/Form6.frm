VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form6 
   Caption         =   "添加..."
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   1860
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   4935
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form6.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   0
      Picture         =   "Form6.frx":009D
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "端口号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "对方IP"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "备注姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RichTextBox1.Text = Text1.Text & vbCrLf & Text2.Text & vbCrLf & Text3.Text & vbCrLf & "-" & vbCrLf & "-" & vbCrLf & "-"
Text4.Text = "\F\" & Text1.Text & ".log"

If Dir("\F") = " " Then MkDir "F"
If Dir(Text4.Text) = "" Then
Open App.Path & Text4.Text For Output As #1
Close #1
End If
RichTextBox1.SaveFile App.Path & Text4.Text, rtfText

MsgBox ("Successful!")
Me.Hide

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RichTextBox1.Text = Text1.Text & vbCrLf & Text2.Text & vbCrLf & Text3.Text
Text4.Text = "\F\" & Text1.Text & ".log"

End Sub

Private Sub Command2_Click()
Form6.Hide


End Sub

Private Sub Command3_Click()

End Sub

