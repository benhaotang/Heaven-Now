VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form7 
   Caption         =   "联系人"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   2145
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   3975
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"Form7.frx":0000
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "选中人信息"
      Height          =   2175
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Height          =   270
         Left            =   3360
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Height          =   270
         Left            =   3360
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Height          =   270
         Left            =   3360
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "更新"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Height          =   270
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   960
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "社交账号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         Caption         =   "电话"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "地址"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "端口"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "IP"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "昵称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "联系人列表"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   2175
         Left            =   1560
         TabIndex        =   2
         Top             =   0
         Width           =   15
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   1830
         Left            =   120
         Pattern         =   "*.log"
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "昵称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.Text1.Text = Text2.Text
Form2.Text2.Text = Text3.Text
MsgBox ("Successful!")
Me.Hide
Form2.Show
    
End Sub

Private Sub Command2_Click()
Form7.Hide
End Sub

Private Sub Command3_Click()
RichTextBox1.Text = Text1.Text & vbCrLf & Text2.Text & vbCrLf & Text3.Text & vbCrLf & Text5.Text & vbCrLf & Text6.Text & vbCrLf & Text7.Text
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

Private Sub File1_Click()
Dim a As String
Dim i  As Integer
Dim n As Integer

i = 0
n = 1
Open App.Path & "\F\" & File1.FileName For Input As #1
Do Until i = Val(n)
Line Input #1, a
i = i + 1
Loop

Close
Text1.Text = a

Dim a1 As String
Dim i1  As Integer
Dim n1 As Integer

i1 = 0
n1 = 2
Open App.Path & "\F\" & File1.FileName For Input As #1
Do Until i1 = Val(n1)
Line Input #1, a1
i1 = i1 + 1
Loop

Close
Text2.Text = a1
 
 Dim a2 As String
Dim i2  As Integer
Dim n2 As Integer

i2 = 0
n2 = 3
Open App.Path & "\F\" & File1.FileName For Input As #1
Do Until i2 = Val(n2)
Line Input #1, a2
i2 = i2 + 1
Loop
Close
Text3.Text = a2
Dim b As String
Dim r As Integer
Dim m As Integer

r = 0
m = 4
Open App.Path & "\F\" & File1.FileName For Input As #1
Do Until r = Val(m)
Line Input #1, b
r = r + 1
Loop
Close
Text5.Text = b
 
 Dim a4 As String
Dim i4  As Integer
Dim n4 As Integer

i4 = 0
n4 = 5
Open App.Path & "\F\" & File1.FileName For Input As #1
Do Until i4 = Val(n4)
Line Input #1, a4
i4 = i4 + 1
Loop

Close
Text6.Text = a4
 Dim a5 As String
Dim i5  As Integer
Dim n5 As Integer

i5 = 0
n5 = 6
Open App.Path & "\F\" & File1.FileName For Input As #1
Do Until i5 = Val(n5)
Line Input #1, a5
i5 = i5 + 1
Loop

Close
Text7.Text = a5

End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\F"
End Sub


Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RichTextBox1.Text = Text1.Text & vbCrLf & Text2.Text & vbCrLf & Text3.Text
Text4.Text = "\F\" & Text1.Text & ".log"

End Sub

