VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Formweb 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   Icon            =   "Formweb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10980
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7200
      Width           =   11055
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      ExtentX         =   19500
      ExtentY         =   12726
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
Attribute VB_Name = "Formweb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Command1_Click()
Formmenu.Show
Formmenu.Visible = True
Formweb.Visible = False
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate ("http://dyprodd.oicp.net/static/jn/products.htm")
Formweb.Caption = "使用 JN 的 DyproCuriousSight DY奇视图片浏览器 做得更好"
End Sub

