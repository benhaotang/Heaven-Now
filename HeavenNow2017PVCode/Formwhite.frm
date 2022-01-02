VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Formwhite 
   BackColor       =   &H8000000B&
   Caption         =   "°×°å"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   1402
      ButtonWidth     =   1085
      ButtonHeight    =   1244
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   840
      Width           =   6255
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock4 
         Left            =   1320
         Top             =   4440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock Winsock3 
         Left            =   3240
         Top             =   3600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   3240
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   5040
         Top             =   2520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSWinsockLib.Winsock WinUdpa 
         Left            =   5040
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3480
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   34
         ImageHeight     =   41
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":10FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":21F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":32EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":43E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":54E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":65DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":76D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Formwhite.frx":87D0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Formwhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x2 As Single
Dim y2 As Single
Dim x3 As Single
Dim y3 As Single
Dim x4 As Single
Dim y4 As Single
Dim r, g, b As Integer
Dim n, n1 As Integer




Private Sub Form_Load()
n = 1
WinUdpa.LocalPort = Int(Form2.Text3.Text) + 99
WinUdpa.RemoteHost = Form2.Text1.Text
WinUdpa.RemotePort = Int(Form2.Text2.Text) + 99
WinUdpa.Bind
Winsock1.LocalPort = Int(Form2.Text3.Text) + 98
Winsock1.RemoteHost = Form2.Text1.Text
Winsock1.RemotePort = Int(Form2.Text2.Text) + 98
Winsock1.Bind
Winsock2.LocalPort = Int(Form2.Text3.Text) + 97
Winsock2.RemoteHost = Form2.Text1.Text
Winsock2.RemotePort = Int(Form2.Text2.Text) + 97
Winsock2.Bind
Winsock3.LocalPort = Int(Form2.Text3.Text) + 96
Winsock3.RemoteHost = Form2.Text1.Text
Winsock3.RemotePort = Int(Form2.Text2.Text) + 96
Winsock3.Bind
Winsock4.LocalPort = Int(Form2.Text3.Text) + 95
Winsock4.RemoteHost = Form2.Text1.Text
Winsock4.RemotePort = Int(Form2.Text2.Text) + 95
Winsock4.Bind
Picture1.AutoRedraw = True
Picture1.CurrentX = x1
Picture1.CurrentY = y1
r = 255
     g = 0
     b = 0
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
If Button = 1 Then
Winsock1.SendData X
Winsock3.SendData Y
Picture1.Line (x2, y2)-(X, Y), RGB(r, g, b)
End If
x2 = X
y2 = Y
WinUdpa.SendData x2
Winsock2.SendData y2
Picture1.Refresh


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index

     Case 1

     r = 255
     g = 0
     b = 0
     Winsock4.SendData 1

     Case 2

     r = 0
     g = 255
     b = 0
     Winsock4.SendData 2
     Case 3
     r = 0
     g = 0
     b = 255
     Winsock4.SendData 3

     Case 4
     r = 255
     g = 255
     b = 0
     Winsock4.SendData 4

     Case 5
     r = 0
     g = 0
     b = 0
     Winsock4.SendData 5
     
Case 6
     r = 255
     g = 255
     b = 255
     Winsock4.SendData 6
     Case 8
     Picture1.Cls
     Winsock4.SendData 7
     Case 9
     CommonDialog1.Filter = "Î»Í¼|*.bmp"

     CommonDialog1.ShowSave

 

FileType = CommonDialog1.FileTitle

 

FiType = LCase(Right(FileType, 3))

FileName = CommonDialog1.FileName
SavePicture Picture1.Image, FileName

     Case 10
     Me.Hide
     Form1.WinUdpa.SendData "ÄúµÄÅóÓÑÒÑ½áÊø»­°å"
     
     
  End Select
End Sub



Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData x4, vbSingle
End Sub



Private Sub WinUdpa_DataArrival(ByVal bytesTotal As Long)
WinUdpa.GetData x3, vbSingle
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Winsock2.GetData y3, vbSingle
End Sub
Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
Winsock3.GetData y4, vbSingle
If n = 1 Then Picture1.Line (x3, y3)-(x4, y4), RGB(255, 0, 0)
If n = 2 Then Picture1.Line (x3, y3)-(x4, y4), RGB(0, 255, 0)
If n = 3 Then Picture1.Line (x3, y3)-(x4, y4), RGB(0, 0, 255)
If n = 4 Then Picture1.Line (x3, y3)-(x4, y4), RGB(255, 255, 0)
If n = 5 Then Picture1.Line (x3, y3)-(x4, y4), RGB(0, 0, 0)
If n = 6 Then Picture1.Line (x3, y3)-(x4, y4), RGB(255, 255, 255)

End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)
n1 = n

Winsock4.GetData n, vbInteger
If n = 7 Then
Picture1.Cls
n = n1
End If


End Sub

