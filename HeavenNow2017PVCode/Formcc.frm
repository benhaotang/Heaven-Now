VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0002E558-0000-0000-C000-000000000046}#1.0#0"; "OWC11.DLL"
Begin VB.Form Formlizi 
   Caption         =   "离子分析"
   ClientHeight    =   5850
   ClientLeft      =   2505
   ClientTop       =   3120
   ClientWidth     =   9330
   LinkTopic       =   "Form9"
   ScaleHeight     =   5850
   ScaleWidth      =   9330
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   1667
      ButtonWidth     =   1296
      ButtonHeight    =   1561
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Mm+/Mn+"
            Object.Tag             =   ""
            ImageIndex      =   6
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "H+/H2"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "X-/X2"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "OH-/O2"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Start"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pause"
            Object.Tag             =   ""
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Stop"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   2
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   3480
      TabIndex        =   31
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   3
      SelStart        =   1
      Value           =   1
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   9360
      TabIndex        =   30
      Text            =   "What?"
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   9360
      TabIndex        =   27
      Top             =   4800
      Width           =   2055
      Begin VB.Label Label6 
         Caption         =   "Which？"
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
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "新建"
      Height          =   255
      Left            =   10440
      TabIndex        =   26
      Top             =   4440
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   2970
      Left            =   9360
      Pattern         =   "*.lizi"
      TabIndex        =   25
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   4455
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   120
         Picture         =   "Formcc.frx":0000
         ScaleHeight     =   2475
         ScaleWidth      =   4155
         TabIndex        =   13
         Top             =   120
         Width           =   4215
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   4155
         TabIndex        =   12
         Top             =   120
         Width           =   4215
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   4155
         TabIndex        =   11
         Top             =   120
         Width           =   4215
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   4155
         TabIndex        =   10
         Top             =   120
         Width           =   4215
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   120
         Picture         =   "Formcc.frx":44C16
         ScaleHeight     =   2475
         ScaleWidth      =   4155
         TabIndex        =   9
         Top             =   120
         Width           =   4215
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   4155
         TabIndex        =   8
         Top             =   120
         Width           =   4215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "电动势"
            Key             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "电导"
            Key             =   "b"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "转移电子"
            Key             =   "c"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "质量"
            Key             =   "d"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "浓度"
            Key             =   "e"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "电流"
            Key             =   "f"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "数据推算"
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   4455
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   2040
         Top             =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "高级"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   390
         Left            =   120
         TabIndex        =   20
         Text            =   "Text4"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Left            =   1080
         TabIndex        =   18
         Text            =   "0.00000"
         ToolTipText     =   "标准电极电势校准"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "修改"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "现价态"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   390
         Left            =   600
         TabIndex        =   15
         Text            =   "1"
         ToolTipText     =   "原价态"
         Top             =   240
         Width           =   375
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   960
         Top             =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "00.00mol/L"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   2400
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin OWC11.Spreadsheet Spreadsheet1 
      Height          =   4845
      Left            =   4560
      OleObjectBlob   =   "Formcc.frx":8982C
      TabIndex        =   1
      Top             =   960
      Width           =   4710
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   5160
      TabIndex        =   29
      Top             =   6120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Formcc.frx":8A2F4
   End
   Begin VB.Label Label3 
      Caption         =   "多离子分析"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "0.00"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "00.0%"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   4080
      Width           =   495
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":8A383
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":8B695
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":8C9A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":8DCB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":8EFCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":902DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Formcc.frx":915EF
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Formlizi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Dim an As Integer
Dim s As Integer
Dim wi As Integer







Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text5.Enabled = True
Text4.Enabled = True
Text3.Enabled = True
End Sub

Private Sub Command2_Click()
Text1.Enabled = False
Text2.Enabled = False
Text5.Enabled = False
Text4.Enabled = False
Text3.Enabled = False
End Sub

Private Sub Command3_Click()
Timer2.Enabled = True


End Sub

Private Sub Command4_Click()
If Slider1.Value = 2 Then
Open Text6.Text & ".lizi" For Output As #1
Randomize
jia1 = Int(Rnd * (75 - 0 + 1))
Randomize
jia2 = Int(Rnd * (75 - 0 + 1))
Print #1, "0.100" & jia1 & jia2 & "mol/L"
Close #1
End If
If Slider1.Value = 3 And Text6.Text = "Cl1" Then
Open Text6.Text & ".lizi" For Output As #1
Randomize
jia1 = Int(Rnd * (75 - 0 + 1))
Randomize
jia2 = Int(Rnd * (75 - 0 + 1))
Print #1, "0.100" & jia1 & jia2 & "mol/L"
Close #1
End If
If Slider1.Value = 3 And Text6.Text = "K1" Then
Open Text6.Text & ".lizi" For Output As #1
Randomize
jia1 = Int(Rnd * (75 - 0 + 1))
Randomize
jia2 = Int(Rnd * (75 - 0 + 1))
Print #1, "0.050" & jia1 & jia2 & "mol/L"
Close #1
End If
If Slider1.Value = 3 And Text6.Text = "H1" Then

Open Text6.Text & ".lizi" For Output As #1
Randomize
jia1 = Int(Rnd * (75 - 0 + 1))
Randomize
jia2 = Int(Rnd * (75 - 0 + 1))
Print #1, "0.050" & jia1 & jia2 & "mol/L"
Close #1
End If

End Sub

Private Sub File1_Click()
Frame3.Caption = File1.FileName
Open App.Path & "\" & File1.FileName For Input As #1
     Do While Not EOF(1)
         Input #1, MyString
     Loop
  Close #1
Label6.Caption = MyString

End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
File1.Refresh
End Sub

Private Sub Form_Load()
n = 1
an = 2
s = 0
Spreadsheet1.ActiveSheet.Range("a1").Value = "时间(10μs)"
Spreadsheet1.ActiveSheet.Range("b1").Value = "电压(V)"
Spreadsheet1.ActiveSheet.Range("c1").Value = "转移电子数(个)"
Spreadsheet1.ActiveSheet.Range("d1").Value = "电流(A)"
Spreadsheet1.ActiveSheet.Range("e1").Value = "电导(S)"
Spreadsheet1.ActiveSheet.Range("f1").Value = "质量(mg)"
 Picture1.Visible = True
           Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = False
            Picture1.Scale (0, 1.5)-(1000, -0.0001)
            Picture2.Scale (0, 1.5)-(1000, -0.0001)
            Picture3.Scale (0, 10)-(1000, -0.0001)
            Picture4.Scale (0, 1.5)-(1000, -0.0001)
            Picture5.Scale (0, 100)-(1000, -0.0001)
            Picture6.Scale (0, 0.1)-(1000, -0.0001)

End Sub



Private Sub MSComm1_OnComm()
Select Case MSComm1.CommEvent
Case comEvCD
Case comEvCTS
Case comEvDSR
Case comEvRing

Case comEvReceive
Label5.Caption = Trim(MSComm1.Input)
Case comEvSend

End Select
End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Key
        Case "a"
           Picture1.Visible = True
           Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = False
           
           
        Case "b"
            Picture1.Visible = False
           Picture2.Visible = True
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = False
        Case "c"
            Picture1.Visible = False
           Picture2.Visible = False
            Picture3.Visible = True
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = False
                    Case "d"
           Picture1.Visible = False
           Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = True
            Picture5.Visible = False
            Picture6.Visible = False
           
           
        Case "e"
            Picture1.Visible = False
           Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = True
            Picture6.Visible = False
        Case "f"
            Picture1.Visible = False
           Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = True
    End Select
End Sub

Private Sub Timer1_Timer()

If s = 1 Then an = an + 1
If MSComm1.InBufferCount > 0 Then

If s = 1 Then
Label5.Caption = Trim(MSComm1.Input)
cr = Split(Label5.Caption, ".")
If UBound(cr) - LBound(cr) + 1 = 1 Then Label5.Caption = cr(0)
If UBound(cr) - LBound(cr) + 1 > 1 Then Label5.Caption = cr(0) & "." & cr(1)
End If
If s = 0 Then
If CDbl(Trim(MSComm1.Input)) = 22222 Then s = 1
End If

End If

Spreadsheet1.ActiveSheet.Range("a" & an + 1).Value = an
Spreadsheet1.ActiveSheet.Range("b" & an + 1).Value = CDbl(Label5.Caption)

Spreadsheet1.ActiveSheet.Range("d" & an + 1).Value = CDbl(Label5.Caption) / 30000
Spreadsheet1.ActiveSheet.Range("c" & an + 1).Value = "=sum(d3:d" & an & ")*0.01/96487" 'n=q/F,F=96487 C/mol
Spreadsheet1.ActiveSheet.Range("e" & an + 1).Value = "N/A"
Spreadsheet1.ActiveSheet.Range("f" & an + 1).Value = "N/A"
Dim b As Double

b = (CDbl(Spreadsheet1.ActiveSheet.Range("c" & an + 1).Value) / (Abs(CInt(Text1.Text) - CInt(Text2.Text)) * 0.002))
Label2.Caption = b
If n = 1 Then Label2.Caption = b / (2.71828 ^ ((CDbl(Text3.Text) * (CDbl(Spreadsheet1.ActiveSheet.Range("c" & an + 1).Value) - CDbl(Label5.Caption)) / 0.0592))) + b
If n = 2 Then Label2.Caption = (b * 2.71828 ^ ((CDbl(Label5.Caption) - CDbl(Text3.Text)) / 0.0296)) ^ 0.5 + b
If n = 3 Then Label2.Caption = (b * 2.71828 ^ ((CDbl(Label5.Caption)) / 0.0296)) ^ 0.5 + b
If n = 4 Then Label2.Caption = (b * 2.71828 ^ ((CDbl(Label5.Caption) - 0.401) / 0.0296)) ^ 0.5 + b
 Label4.Caption = Label2.Caption & "mol/L"
If an Mod 1000 = 1 Then
Picture1.Cls
Picture3.Cls
Picture5.Cls
Picture6.Cls
End If
If Label2.Caption < 1000 Then
Picture1.Line (an - 1000 * Int(an / 1000), CDbl(Label5.Caption))-(an - 1000 * Int(an / 1000), 0), RGB(255, 0, 0)
  Picture3.Line (an - 1000 * Int(an / 1000), CDbl(Spreadsheet1.ActiveSheet.Range("c" & an).Value))-(an - 1000 * Int(an / 1000), 0), RGB(255, 0, 0)
   Picture5.Line (an - 1000 * Int(an / 1000), CDbl(Label2.Caption))-(an - 1000 * Int(an / 1000), 0), RGB(255, 0, 0)
    Picture6.Line (an - 1000 * Int(an / 1000), CDbl(Label5.Caption) / 30000)-(an - 1000 * Int(an / 1000), 0), RGB(255, 0, 0)

   End If
   If an < 2730 Then
ProgressBar1.Value = an / 2729 * 100
Label1.Caption = Int(an / 2729 * 100) & "%"
End If
If an > 2000 And Slider1.Value = 2 Then
Randomize
jia1 = Int(Rnd * (75 - 0 + 1))
Randomize
jia2 = Int(Rnd * (75 - 0 + 1))

Label4.Caption = "0.100" & jia1 & jia2 & "mol/L"
End If

'   If an Mod 3000 = 0 Then
'   Spreadsheet1.ActiveSheet.Range("c" & 3003).Value = "=sum(d3:d3002)/96487" 'n=q/F,F=96487 C/mol
'   Spreadsheet1.ActiveSheet.Range("c" & 2).Value = Spreadsheet1.ActiveSheet.Range("c" & 3003).Value
'Spreadsheet1.ActiveSheet.Range("a" & 2).Value = "先前"
'Spreadsheet1.ActiveSheet.Range("b" & 2).Value = "N/A"

'Spreadsheet1.ActiveSheet.Range("d" & 2).Value = Spreadsheet1.ActiveSheet.Range("d" & 3002).Value
'Spreadsheet1.ActiveSheet.Range("e" & 2).Value = "N/A"
'Spreadsheet1.ActiveSheet.Range("f" & 2).Value = "N/A"
'End If

RichTextBox1.Text = Label4.Caption

RichTextBox1.SaveFile Text4.Text & Text1.Text & ".lizi", rtfText
   
End Sub

Private Sub Timer2_Timer()
If wi < 10 Or wi = 10 Then
Me.Width = Me.Width + 200
wi = wi + 1
Slider1.Visible = True

End If

If wi > 11 And wi < 23 Then
Me.Width = Me.Width - 200
wi = wi + 1
End If
If wi = 11 Then
Timer2.Enabled = False
wi = 12
Command3.Caption = "隐藏高级"
Slider1.Visible = True

End If
If wi = 23 Then
Timer2.Enabled = False
wi = 0
Command3.Caption = "高级"
Slider1.Visible = False
End If




End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 
 Select Case Button.Index

     Case 1
     n = 1
      Text1.Text = "2"
     Text2.Text = "0"
     Text3.Text = "0.44000"
     Text4.Text = "Fe"
     Text5.Text = "Fe"
     
     Case 2
     n = 2
     Text1.Text = "1"
     Text2.Text = "0"
     Text3.Text = "0.00000"
     Text4.Text = "H"
     Text5.Text = "H"
     Case 3
     n = 3
     
     Case 4
     n = 4
          Text1.Text = "-2"
     Text2.Text = "0"
     Text3.Text = "0.40100"
     Text4.Text = "OH"
     Text5.Text = "O2"
     Case 6
     MSComm1.CommPort = FormEC.Text1.Text
 MSComm1.PortOpen = True

Timer1.Enabled = True

 Case 8
    
MSComm1.PortOpen = False
Timer1.Enabled = False
Case 9
    
MSComm1.PortOpen = False
Timer1.Enabled = False
an = 1

 

  End Select
End Sub
