VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formcentral 
   Caption         =   "Heaven Now �ۺϴ���"
   ClientHeight    =   9630
   ClientLeft      =   255
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
   Icon            =   "hntxt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   15240
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   10800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16960
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"hntxt.frx":1E32
      MouseIcon       =   "hntxt.frx":1ECF
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
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭"
      Begin VB.Menu mnuundo 
         Caption         =   "������PRO��"
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
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuselecall 
         Caption         =   "ȫѡ"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuinput 
      Caption         =   "����"
      Begin VB.Menu mnugr 
         Caption         =   "Ϳѻ"
      End
      Begin VB.Menu mnupic 
         Caption         =   "ͼƬ��PRO�����˰���ø��Ʋ���"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnutable 
         Caption         =   "csv���"
         Shortcut        =   ^T
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
      Caption         =   "������������"
      Begin VB.Menu mnuword 
         Caption         =   "ȫ������"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnusize 
         Caption         =   "ȫ�ִ�С"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnufont 
         Caption         =   "ѡ����������"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "����"
      Begin VB.Menu mnuhlp 
         Caption         =   "ʹ������"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuabout 
         Caption         =   "��������"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "Formcentral"
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
RichTextBox1.Top = 0

RichTextBox1.Left = 20

RichTextBox1.Height = ScaleHeight - 40

RichTextBox1.Width = ScaleWidth - 40

End Sub

 


 



Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuexchange_Click()

sFind = InputBox("������Ҫ�滻���֡��ʣ�", "Heaven Now �ۺϴ����滻����I", sFind)
w = InputBox("�������滻���֡��ʣ�", "Heaven Now �ۺϴ����滻II", sFind)
RichTextBox1.Find sFind
RichTextBox1.Text = Replace(RichTextBox1.Text, sFind, w)
End Sub

Private Sub mnufont_Click()
 On Error Resume Next
 With CommonDialog1
 .FontName = RichTextBox1.SelFontName
 .FontSize = RichTextBox1.SelFontSize
 .Color = RichTextBox1.SelColor
 .FontBold = RichTextBox1.SelBold
 .FontItalic = RichTextBox1.SelItalic
 .FontUnderline = RichTextBox1.SelUnderline
 .FontStrikethru = RichTextBox1.SelStrikeThru
 End With
 CommonDialog1.Flags = cdlCFEffects Or cdlCFForceFontExist Or cdlCFScreenFonts
 CommonDialog1.ShowFont
 With RichTextBox1
 .SelFontName = CommonDialog1.FontName
 .SelFontSize = CommonDialog1.FontSize
 .SelColor = CommonDialog1.Color
 .SelBold = CommonDialog1.FontBold
 .SelItalic = CommonDialog1.FontItalic
 .SelUnderline = CommonDialog1.FontUnderline
 .SelStrikeThru = CommonDialog1.FontStrikethru
 End With
 RichTextBox1.SetFocus
End Sub

Private Sub mnugr_Click()
FormG.Show

End Sub

Private Sub mnuhlp_Click()
MsgBox ("�û�����:" & (Chr(13) & Chr(10)) & "1.���ͼƬ���ƣ������֧���ڼ�����������ʽ��֧�ָ�ʽ��ͼƬֱ�Ӹ��ƽ����������������Զ����ǵ��ڴ�С��" & (Chr(13) & Chr(10)) & "�������ް�������")
End Sub

Private Sub mnuNew_Click()

 

RichTextBox1.Text = ""


 

FileName = "Heaven Now �ۺϴ���δ����"

Me.Caption = FileName

 

End Sub

 

 

Private Sub mnuOpen_Click()

CommonDialog1.Filter = "Miscorsoft Word��*doc������֧���������������doc��|*.doc|Miscorsoft Excel CSV ���ŷָ���*doc����ֻ�ܴ��ı���|*.csv|Miscorsoft�ı��ĵ���*txt��|*.txt|RTF�ļ���*rtf��|*.rtf|�����ļ���*.*��|*.*"

CommonDialog1.ShowOpen

RichTextBox1.Text = ""



 

FileName = CommonDialog1.FileName

 

RichTextBox1.LoadFile FileName

 

Me.Caption = "Heaven Now �ۺϴ���" & FileName

 

End Sub



 





Private Sub mnuSave_Click()

 

CommonDialog1.Filter = "Miscorsoft Word��*doc�����������֧�ֱ��������doc��|*.doc|Miscorsoft Excel CSV ���ŷָ���*doc����ֻ�ܴ��ı���|*.csv|Miscorsoft�ı��ĵ���*txt��|*.txt|RTF�ļ���*rtf��|*.rtf|�����ļ���*.*��|*.*"


CommonDialog1.ShowSave

 

FileType = CommonDialog1.FileTitle

 

FiType = LCase(Right(FileType, 3))

FileName = CommonDialog1.FileName

 

Select Case FiType

Case "txt"

RichTextBox1.SaveFile FileName, rtfText

Case "doc"

RichTextBox1.SaveFile FileName, doc
Case "csv"
RichTextBox1.SaveFile FileName, csv

Case "rtf"

RichTextBox1.SaveFile FileName, rtfRTF

Case "*.*"

RichTextBox1.SaveFile FileName

 

End Select

Me.Caption = "Heaven Now �ۺϴ���" & FileName

End Sub

 

 

Private Sub mnuexit_Click()
Formcentral.Visible = False
Formmenu.Visible = True
End Sub

 

 

Private Sub mnuCopy_Click()

Clipboard.Clear

Clipboard.SetText RichTextBox1.SelText

 

End Sub

 



 

Private Sub mnuCut_Click()

Clipboard.Clear

Clipboard.SetText RichTextBox1.SelText

 
 

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

 sFind = InputBox("������Ҫ���ҵ��֡��ʣ�", "Heaven Now �ۺϴ���:����", sFind)

RichTextBox1.Find sFind

End Sub

 



 

Private Sub mnuFindOn_Click()

 

RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength + 1

 

RichTextBox1.Find sFind, , Len(RichTextBox1)

 

End Sub

 



 

Private Sub mnusize_Click()
w = InputBox("��������Ҫ�Ĵ�С��", "Heaven Now �ۺϴ��������С")
RichTextBox1.Font.Size = w
End Sub



Private Sub mnutable_Click()

n = InputBox("������������", "Heaven Now �ۺϴ����������")
n1 = InputBox("������������", "Heaven Now �ۺϴ����������")
Dim a As Integer

For a = 1 To n1

shit = shit & ","

Next a
For a = 1 To n
RichTextBox1.Text = RichTextBox1.Text & shit & vbCrLf

Next a

End Sub

Private Sub mnuword_Click()
w = InputBox("��������Ҫ�����壨ȫ�ƣ���", "Heaven Now �ۺϴ�������")
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



