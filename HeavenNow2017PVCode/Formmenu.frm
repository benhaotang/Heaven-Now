VERSION 5.00
Begin VB.Form Formmenu 
   Caption         =   "Heaven Now 2016 ��ʼ�˵�"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   Icon            =   "Formmenu.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Formmenu.frx":1872
   ScaleHeight     =   5595
   ScaleWidth      =   10290
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   4080
      Top             =   240
   End
   Begin VB.CommandButton Command9 
      Caption         =   "����"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ѧϰ"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   8280
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�˽�רҵ�棬���ʰ棬������ҳ"
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "���ڱ����"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ʹͿѻ��ø��ʺ���"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ϳѻ"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�߼���̴���(PRO)"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ۺ��ļ�����"
      Height          =   375
      Left            =   1080
      Picture         =   "Formmenu.frx":67D9
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2016"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   7800
      X2              =   7800
      Y1              =   4680
      Y2              =   5160
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7680
      TabIndex        =   10
      Top             =   4560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   735
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   675
      Left            =   480
      Picture         =   "Formmenu.frx":6BCC
      Top             =   4200
      Width           =   675
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   600
      Picture         =   "Formmenu.frx":83F6
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   480
      Picture         =   "Formmenu.frx":A228
      Top             =   2280
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   675
      Left            =   480
      Picture         =   "Formmenu.frx":BA52
      Top             =   1440
      Width           =   675
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label3 
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Formmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer

Private Sub Command1_Click()
Formcentral.Show
Formmenu.Visible = False
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "���������Դ���΢��Word��doc����΢��Excel��csv�����ı��ĵ���txt����RTF�ĵ���rtf�����ļ������򵥵Ĳ��������򵥵Ľ��棬���ḻ�Ľ��棬ʹ��İ칫֮�ø����" & (Chr(13) & Chr(10)) & "�������ߣ���������"
End Sub

Private Sub Command2_Click()
Formweb2.Show
Formmenu.Visible = False
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "���������Խ��м򵥵ı���Ͳ��ԡ�֧��BAT��HTML����Ҫ����רҵ�棬��֪�����ڼ���רҵ�棬�������İ�ť" & (Chr(13) & Chr(10)) & "�������ߣ���������"
End Sub

Private Sub Command3_Click()
FormG.Show
Formmenu.Visible = False
End Sub
Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "���������Խ��и���Ϳѻ��ͼƬ�޸ġ�����PS���������գ��ʺϳ�ѧ��ʹ��!���ܽ���ȫ������򵥻�ͼ��Ҫ��" & (Chr(13) & Chr(10)) & "�������ߣ���������"
End Sub

Private Sub Command4_Click()
Formweb.Show
Formmenu.Visible = False
End Sub
Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "ʹ��JN��DyproCuriousSight  DY����ͼƬ�������ʹʹ�ù��̸����ʣ�����"
End Sub

Private Sub Command5_Click()
frmAbout.Show
End Sub

Private Sub Command6_Click()
Formweb2.Show
Formmenu.Visible = False
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "������ҳ��http://dyprodd.oicp.net/static/mt/" & (Chr(13) & Chr(10)) & "�����ҳ��http://dyprodd.oicp.net/static/mt/hn.html" & (Chr(13) & Chr(10)) & "�������ߣ���������"
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Command8_Click()
FormEDU.Show
Me.Hide

End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "������һЩ�߼��ĵ��ӻ�ѧϰ�˿ڡ������ѧϰ�°빦����" & (Chr(13) & Chr(10)) & "�������ߣ���������"
End Sub

Private Sub Command9_Click()
Me.Hide
Form1.Show

End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "����UPDЭ�飬ͬһ������������ֻ�赯ָ֮�С�" & (Chr(13) & Chr(10)) & "�������ߣ���������"
End Sub

Private Sub Form_Load()
Me.Height = 0

End Sub

Private Sub Timer1_Timer()

t = t + 1


Me.Height = Me.Height + 560

If t = 10 Then

Timer1.Enabled = False

End If

End Sub
