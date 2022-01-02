VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Formweb2 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12330
   Icon            =   "Formweb2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12330
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8400
      Width           =   12375
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      ExtentX         =   21828
      ExtentY         =   14843
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
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
End
Attribute VB_Name = "Formweb2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Command1_Click()
Formmenu.Show
Formmenu.Visible = True
Formweb2.Visible = False
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate ("http://dyprodd.oicp.net/static/mt/hn.html")
Formweb2.Caption = "Heaven Now 软件主页"
End Sub

