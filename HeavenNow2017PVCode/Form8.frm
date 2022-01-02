VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form8 
   Caption         =   "记录"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "字号"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   4695
      Begin ComctlLib.Slider Slider1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   327682
         Min             =   10
         Max             =   50
         SelStart        =   10
         Value           =   10
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "查看记录"
      Height          =   3135
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4895
         _Version        =   393217
         BackColor       =   8438015
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"Form8.frx":1E32
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "记录列表"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   15
         Left            =   0
         TabIndex        =   4
         Top             =   3120
         Width           =   1935
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2730
         Left            =   120
         Pattern         =   "*.save"
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_Click()
RichTextBox1.LoadFile App.Path & "\DATA\" & File1.FileName
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\DATA"
End Sub




Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RichTextBox1.Font.Size = Slider1.Value
End Sub
