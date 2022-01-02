VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Heaven Now CSV Table"
   ClientHeight    =   9630
   ClientLeft      =   -135
   ClientTop       =   750
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "����"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "hncsv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   15240
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4455
      Left            =   4440
      TabIndex        =   2
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"hncsv.frx":1872
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Picture         =   "hncsv.frx":190F
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   $"hncsv.frx":1D8B
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   10335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�"
      Begin VB.Menu mnuopen 
         Caption         =   "��"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnunew 
         Caption         =   "�½�"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnusave 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuexit 
         Caption         =   "�˳�"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuadd 
      Caption         =   "����"
      Begin VB.Menu mnutable 
         Caption         =   "���"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭"
      Begin VB.Menu mnuundo 
         Caption         =   "����"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnucopy 
         Caption         =   "����"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucut 
         Caption         =   "����"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuselecall 
         Caption         =   "ȫѡ"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "����"
      Begin VB.Menu mnufind 
         Caption         =   "����"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuexchange 
         Caption         =   "�滻"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnufindon 
         Caption         =   "��������"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuwindows 
      Caption         =   "�����������ã����ɱ��棩"
      Begin VB.Menu mnuword 
         Caption         =   "����"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnusize 
         Caption         =   "��С"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim sFind As String


 

Dim FileType, FiType As String

 


 

Private Sub Form_Load()



 

Me.Height = 6000

Me.Width = 9000

End Sub

 

Private Sub Form_Resize()

On Error Resume Next
RichTextBox1.Top = 500

RichTextBox1.Left = 20

RichTextBox1.Height = ScaleHeight - 40

RichTextBox1.Width = ScaleWidth - 40

End Sub

 


 



'Private Sub mnuexchange_Click()
'sFind = InputBox("������Ҫ�滻���֡��ʣ�", "Heaven Now CSV Table���滻����I", sFind)
'w = InputBox("�������滻���֡��ʣ�", "Heaven Now CSV Table���滻II", sFind)
'RichTextBox1.Find sFind
'RichTextBox1.Text = Replace(RichTextBox1.Text, sFind, w)
'End Sub

Private Sub mnuNew_Click()

 

RichTextBox1.Text = ""


 

FileName = "Heaven Now CSV Table��δ����"

Me.Caption = FileName

 

End Sub

 

 

Private Sub mnuOpen_Click()

CommonDialog1.Filter = "��񣨶��ŷָ�����*csv��|*.csv|*.txt|�����ļ���*.*��|*.*"

CommonDialog1.ShowOpen

RichTextBox1.Text = ""



 

FileName = CommonDialog1.FileName

 

RichTextBox1.LoadFile FileName

 

Me.Caption = "Heaven Now CSV Table��" & FileName

 

End Sub



 

Private Sub mnuSave_Click()

 

CommonDialog1.Filter = "��񣨶��ŷָ�����*csv��|*.csv|*.txt|�����ļ���*.*��|*.*"


CommonDialog1.ShowSave

 

FileType = CommonDialog1.FileTitle

 

FiType = LCase(Right(FileType, 3))

FileName = CommonDialog1.FileName

 

Select Case FiType

Case "txt"

RichTextBox1.SaveFile FileName, rtfText

Case "csv"

RichTextBox1.SaveFile FileName, csvRTF

Case "*.*"

RichTextBox1.SaveFile FileName

 

End Select

Me.Caption = "Heaven Now CSV Table��" & FileName

End Sub

 

 

Private Sub mnuExit_Click()

End

End Sub

 

 

Private Sub mnuCopy_Click()

Clipboard.Clear

Clipboard.SetText RichTextBox1.SelText

 

End Sub

 



 

Private Sub mnuCut_Click()

Clipboard.Clear

Clipboard.SetText
RichTextBox1.SelText

 

RichTextBox1.SelText = ""

End Sub



 

Private Sub mnuSelectAll_Click()

 

RichTextBox1.SelStart = 0

RichTextBox1.SelLength = Len(RichTextBox1.Text)

End Sub

 



 

Private Sub mnuPaste_Click()

 

RichTextBox1.SelText = Clipboard.GetText

 

End Sub

 

 

Private Sub mnuFind_Click()

 sFind = InputBox("������Ҫ���ҵ��֡��ʣ�", "Heaven Now CSV Table����������", sFind)

RichTextBox1.Find sFind

End Sub

 



 

Private Sub mnuFindOn_Click()

 

RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength + 1

 

RichTextBox1.Find sFind, , Len(RichTextBox1)

 

End Sub

 



 

Private Sub mnusize_Click()
w = InputBox("��������Ҫ�Ĵ�С��", "Heaven Now TXT Word�������С")
RichTextBox1.Font.Size = w
End Sub



Private Sub mnutable_Click()
n = InputBox("������������", "Heaven Now TXT Word���������")
n1 = InputBox("������������", "Heaven Now TXT Word���������")
Dim a As Integer

For a = 1 To n1

shit = shit & ","

Next a
For a = 1 To n

RichTextBox1.Text = RichTextBox1.Text & shit & vbCrLf

Next a

End Sub

Private Sub mnuword_Click()
w = InputBox("��������Ҫ�����壨ȫ�ƣ���", "Heaven Now CSV Table������")
RichTextBox1.Font = w
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then

PopupMenu mnuEdit, vbPopupMenuLeftAlign
PopupMenu mnuwindows, vbPopupMenuLeftAlign

Else

Exit Sub

End If

End Sub


Private Sub RichTextBox1_KeyUp(KeyCode As Integer, Shift As Integer)

 

If KeyCode = vbKeySpace Then

RichTextBox1.SelFontName = CommonDialog1.FontName

 

End If

End Sub

 



