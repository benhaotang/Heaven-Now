VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   ClientHeight    =   8595
   ClientLeft      =   2820
   ClientTop       =   2010
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8280
      Width           =   10815
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      ExtentX         =   19050
      ExtentY         =   14631
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10785
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10785
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   2670
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":02E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":05C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Formmenu.Show
Formmenu.Visible = True
frmBrowser.Visible = False
End Sub

Private Sub Form_Load()
brwWebBrowser.Navigate ("http://dyprodd.oicp.net/static/jn/products.htm")
frmBrowser.Caption = "使用 JN 的 DyproCuriousSight DY奇视图片浏览器 做得更好"
End Sub
