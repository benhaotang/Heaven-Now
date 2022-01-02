VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formcentral 
   Caption         =   "Heaven Now 综合处理"
   ClientHeight    =   9630
   ClientLeft      =   255
   ClientTop       =   750
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "宋体"
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
      Caption         =   "文件"
      Begin VB.Menu mnuopen 
         Caption         =   "打开"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnunew 
         Caption         =   "新建"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnusave 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuexit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑"
      Begin VB.Menu mnuundo 
         Caption         =   "撤销（PRO）"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnucopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnucut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuselecall 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuinput 
      Caption         =   "插入"
      Begin VB.Menu mnugr 
         Caption         =   "涂鸦"
      End
      Begin VB.Menu mnupic 
         Caption         =   "图片（PRO）个人版可用复制操作"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnutable 
         Caption         =   "csv表格"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "搜索"
      Begin VB.Menu mnufind 
         Caption         =   "查找"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuexchange 
         Caption         =   "替换"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnufindon 
         Caption         =   "继续查找"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuwindows 
      Caption         =   "窗口字体设置"
      Begin VB.Menu mnuword 
         Caption         =   "全局字体"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnusize 
         Caption         =   "全局大小"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnufont 
         Caption         =   "选定文字字体"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "帮助"
      Begin VB.Menu mnuhlp 
         Caption         =   "使用疑难"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuabout 
         Caption         =   "关于我们"
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

sFind = InputBox("请输入要替换的字、词：", "Heaven Now 综合处理：替换内容I", sFind)
w = InputBox("请输入替换的字、词：", "Heaven Now 综合处理：替换II", sFind)
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
MsgBox ("用户疑难:" & (Chr(13) & Chr(10)) & "1.软件图片复制：本软件支持在剪贴板的正规格式或支持格式的图片直接复制进本软件，本软件可以对他们调节大小。" & (Chr(13) & Chr(10)) & "其他暂无帮助……")
End Sub

Private Sub mnuNew_Click()

 

RichTextBox1.Text = ""


 

FileName = "Heaven Now 综合处理：未命名"

Me.Caption = FileName

 

End Sub

 

 

Private Sub mnuOpen_Click()

CommonDialog1.Filter = "Miscorsoft Word（*doc）（不支持其他软件制作的doc）|*.doc|Miscorsoft Excel CSV 逗号分隔（*doc）（只能纯文本）|*.csv|Miscorsoft文本文档（*txt）|*.txt|RTF文件（*rtf）|*.rtf|所有文件（*.*）|*.*"

CommonDialog1.ShowOpen

RichTextBox1.Text = ""



 

FileName = CommonDialog1.FileName

 

RichTextBox1.LoadFile FileName

 

Me.Caption = "Heaven Now 综合处理：" & FileName

 

End Sub



 





Private Sub mnuSave_Click()

 

CommonDialog1.Filter = "Miscorsoft Word（*doc）（其他软件支持本软件做的doc）|*.doc|Miscorsoft Excel CSV 逗号分隔（*doc）（只能纯文本）|*.csv|Miscorsoft文本文档（*txt）|*.txt|RTF文件（*rtf）|*.rtf|所有文件（*.*）|*.*"


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

Me.Caption = "Heaven Now 综合处理：" & FileName

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

 sFind = InputBox("请输入要查找的字、词：", "Heaven Now 综合处理:查找", sFind)

RichTextBox1.Find sFind

End Sub

 



 

Private Sub mnuFindOn_Click()

 

RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength + 1

 

RichTextBox1.Find sFind, , Len(RichTextBox1)

 

End Sub

 



 

Private Sub mnusize_Click()
w = InputBox("请输入所要的大小：", "Heaven Now 综合处理：字体大小")
RichTextBox1.Font.Size = w
End Sub



Private Sub mnutable_Click()

n = InputBox("请输入行数：", "Heaven Now 综合处理：添加行数")
n1 = InputBox("请输入列数：", "Heaven Now 综合处理：添加列数")
Dim a As Integer

For a = 1 To n1

shit = shit & ","

Next a
For a = 1 To n
RichTextBox1.Text = RichTextBox1.Text & shit & vbCrLf

Next a

End Sub

Private Sub mnuword_Click()
w = InputBox("请输入所要的字体（全称）：", "Heaven Now 综合处理：字体")
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



