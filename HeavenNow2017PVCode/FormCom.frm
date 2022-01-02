VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FormCom 
   Caption         =   "串口通讯"
   ClientHeight    =   5415
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4695
   LinkTopic       =   "Form9"
   ScaleHeight     =   5415
   ScaleWidth      =   4695
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5106
      _Version        =   393217
      DisableNoScroll =   -1  'True
      TextRTF         =   $"FormCom.frx":0000
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "双击发送"
      Top             =   3120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   0   'False
      DisableNoScroll =   -1  'True
      TextRTF         =   $"FormCom.frx":009D
   End
   Begin VB.Menu mnustart 
      Caption         =   "开始"
      Begin VB.Menu mnuset 
         Caption         =   "设置"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuexit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnutext 
      Caption         =   "文本"
      Begin VB.Menu mnuget 
         Caption         =   "送达"
         Begin VB.Menu mnuSelectAll1 
            Caption         =   "全选"
         End
         Begin VB.Menu mnucopy1 
            Caption         =   "复制"
         End
         Begin VB.Menu mnuPaste1 
            Caption         =   "粘贴"
         End
         Begin VB.Menu mnufind1 
            Caption         =   "查找"
         End
         Begin VB.Menu mnufindon1 
            Caption         =   "查找下一个"
         End
      End
      Begin VB.Menu mnusend 
         Caption         =   "发送"
         Begin VB.Menu mnuSelectAll2 
            Caption         =   "全选"
         End
         Begin VB.Menu mnucopy2 
            Caption         =   "复制"
         End
         Begin VB.Menu mnupaste2 
            Caption         =   "粘贴"
         End
         Begin VB.Menu mnufind2 
            Caption         =   "查找"
         End
         Begin VB.Menu mnufindon2 
            Caption         =   "查找下一个"
         End
      End
   End
   Begin VB.Menu mnuhuitu 
      Caption         =   "绘图"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FormCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sFind As String

Private Sub Form_Load()
Text1.Top = 0

Text1.Left = 20

Text1.Height = ScaleHeight / 4 * 3

Text1.Width = ScaleWidth - 40
Text2.Top = ScaleHeight / 4 * 3

Text2.Left = 20

Text2.Height = ScaleHeight / 4

Text2.Width = ScaleWidth - 40
Text2.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Text1.Top = 0

Text1.Left = 20

Text1.Height = ScaleHeight / 4 * 3

Text1.Width = ScaleWidth - 40
Text2.Top = ScaleHeight / 4 * 3

Text2.Left = 20

Text2.Height = ScaleHeight / 4

Text2.Width = ScaleWidth - 40
End Sub

Private Sub list1_click()

End Sub

Private Sub mnuCopy1_Click()

Clipboard.Clear

Clipboard.SetText Text1.SelText

 

End Sub

 



 

Private Sub mnuCut1_Click()

Clipboard.Clear

Clipboard.SetText Text1.SelText

 
 

RichTextBox1.SelText = ""

End Sub



 

Private Sub mnupause_Click()

End Sub

Private Sub mnufindon_Click()

End Sub

Private Sub mnuexit_Click()
Me.Hide
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
FormEDU.Show

End Sub

Private Sub mnuhuitu_Click()
formcomgra.Show

End Sub

Private Sub mnuSelectAll1_Click()

 

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

 



 

Private Sub mnuPaste1_Click()

 

Text1.SelText = Clipboard.GetText

 

End Sub

 

 

Private Sub mnuFind1_Click()

 sFind = InputBox("请输入要查找的字、词：", "Heaven Now串口通讯:查找", sFind)

Text1.Find sFind

End Sub

 



 

Private Sub mnuFindOn1_Click()

 

Text1.SelStart = Text1.SelStart + Text1.SelLength + 1

 

Text1.Find sFind, , Len(Text1)

 

End Sub
Private Sub mnuCopy2_Click()

Clipboard.Clear

Clipboard.SetText Text2.SelText

 

End Sub

 



 

Private Sub mnuCut2_Click()

Clipboard.Clear

Clipboard.SetText Text2.SelText

 
 

RichTextBox2.SelText = ""

End Sub



 





Private Sub mnuSelectAll2_Click()

 

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

 



 

Private Sub mnuPaste2_Click()

 

Text2.SelText = Clipboard.GetText

 

End Sub

 

 

Private Sub mnuFind2_Click()

 sFind = InputBox("请输入要查找的字、词：", "Heaven Now 串口通讯:查找", sFind)

Text2.Find sFind

End Sub

 



 

Private Sub mnuFindOn2_Click()

 

Text2.SelStart = Text2.SelStart + Text2.SelLength + 2

 

Text2.Find sFind, , Len(Text2)

 

End Sub



Private Sub mnuset_Click()
FormA.Show
End Sub

Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
Case comEvCD
Case comEvCTS
Case comEvDSR
Case comEvRing

Case comEvReceive
Text1.Text = Text1.Text & vbCrLf & "& Now &" & ":" & Trim(MSComm1.Input)
Case comEvSend
Text1.Text = Text2.Text & vbCrLf & "You:" & Trim(MSComm1.Output)
Text2.Text = ""
End Select
End Sub

Private Sub Text2_DblClick()
Text1.Text = Text1.Text & vbCrLf & "You:" & Text2.Text
MSComm1.Output = Text2.Text
Text2.Text = ""
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Text1.Top = 0

Text1.Left = 20

Text1.Height = ScaleHeight / 4 * 3

Text1.Width = ScaleWidth - 40
Text2.Top = ScaleHeight / 4 * 3

Text2.Left = 20

Text2.Height = ScaleHeight / 4

Text2.Width = ScaleWidth - 40
End Sub
Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Text1.Top = 0

Text1.Left = 20

Text1.Height = ScaleHeight / 4 * 3

Text1.Width = ScaleWidth - 40
Text2.Top = ScaleHeight / 4 * 3

Text2.Left = 20

Text2.Height = ScaleHeight / 4

Text2.Width = ScaleWidth - 40
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If Button = 2 Then

PopupMenu mnuget, vbPopupMenuLeftAlign


Else

Exit Sub

End If

End Sub
Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If Button = 2 Then

PopupMenu mnusend, vbPopupMenuLeftAlign


Else

Exit Sub

End If

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Timer1_Timer()
If MSComm1.InBufferCount > 0 Then
Text1.Text = Text1.Text & vbCrLf & "MCU:" & Trim(MSComm1.Input)
 
End If
End Sub

