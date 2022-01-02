VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form4 
   Caption         =   "局域网内用户"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   4680
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "登录用户"
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin MSWinsockLib.Winsock dns 
         Left            =   1560
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "确定"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Text            =   "IP"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Text            =   "主机名"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3090
         ItemData        =   "Form4.frx":0000
         Left            =   120
         List            =   "Form4.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const resource_connected As Long = &H1&
Private Const resource_globalnet As Long = &H2&
Private Const resource_remembered As Long = &H3&


Private Const resourcedisplaytype_directory& = &H9
Private Const resourcedisplaytype_domain& = &H1
Private Const resourcedisplaytype_file& = &H4
Private Const resourcedisplaytype_generic& = &H0
Private Const resourcedisplaytype_group& = &H5
Private Const resourcedisplaytype_network& = &H6
Private Const resourcedisplaytype_root& = &H7
Private Const resourcedisplaytype_server& = &H2
Private Const resourcedisplaytype_share& = &H3
Private Const resourcedisplaytype_shareadmin& = &H8
Private Const resourcetype_any As Long = &H0&
Private Const resourcetype_disk As Long = &H1&
Private Const resourcetype_print As Long = &H2&
Private Const resourcetype_unknown As Long = &HFFFF&
Private Const resourceusage_all As Long = &H0&
Private Const resourceusage_connectable As Long = &H1&
Private Const resourceusage_container As Long = &H2&
Private Const resourceusage_reserved As Long = &H80000000
Private Const no_error = 0
Private Const error_more_data = 234 'l // dderror
Private Const resource_enum_all As Long = &HFFFF

Private Type Netresource
       dwScope As Long
       dwType As Long
       dwdisplaytype As Long
       dwUsage As Long
       plocalname As Long
       premotename As Long
       pcomment As Long
       pprovider As Long
End Type

Private Type Netresource_BUf
        dwScope As Long
        dwType As Long
        dwdisplaytype As Long
        dwUsage As Long
        slocalname As String
        sremotename As String
        scomment As String
        sprovider As String
End Type

Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Netresource, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function ValidateRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Sub copymem Lib "kernel32" Alias "rtlmovememory" (lpto As Any, lpfrom As Any, ByVal llen As Long)
Private Declare Sub copymembyptr Lib "kernel32" Alias "rtlmovememory" (ByVal lpto As Long, ByVal lpfrom As Long, ByVal llen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpya" (ByVal lpstring1 As String, ByVal lpstring2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlena" (ByVal lpstring As Any) As Long

Private Sub Command1_Click()
MsgBox ("Successful!")
Me.Hide
Form2.Show
Form2.Text1.Text = Text2.Text

End Sub

Private Sub Command2_Click()
Form4.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
Const MAX_RESOURCES = 256
Const NOT_A_CONTAINER = -1
Dim bFirstTime As Boolean
Dim lReturn As Long
Dim hEnum As Long
Dim lCount As Long
Dim lMin As Long
Dim lLength As Long
Dim l As Long
Dim lBufferSize As Long
Dim lLastIndex As Long
Dim uNetApi(0 To MAX_RESOURCES) As Netresource
Dim uNet() As Netresource_BUf
bFirstTime = True

Do
  If bFirstTime Then
     lReturn = WNetOpenEnum(resource_globalnet, resourcetype_any, resourceusage_all, uNetApi(0), hEnum)
     bFirstTime = False
     Else
         If uNet(lLastIndex).dwUsage And resourceusage_container Then
            lReturn = WNetOpenEnum(resource_globalnet, resourcetype_any, resourceusage_all, uNetApi(lLastIndex), hEnum)
            Else
               lReturn = NOT_A_CONTAINER
               hEnum = 0
         End If
     lLastIndex = lLastIndex + 1
  End If
  
  If lReturn = no_error Then
     lCount = resource_enum_all
     Do
       lBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
       lReturn = WNetEnumResource(hEnum, lCount, uNetApi(0), lBufferSize)
      If lCount > 0 Then
         ReDim Preserve uNet(0 To lMin + lCount - 1) As Netresource_BUf   '以前是netresourece
         For l = 0 To lCount - 1
             uNet(lMin + l).dwScope = uNetApi(l).dwScope
             uNet(lMin + l).dwType = uNetApi(l).dwType
             uNet(lMin + l).dwdisplaytype = uNetApi(l).dwdisplaytype
             uNet(lMin + l).dwUsage = uNetApi(l).dwUsage
             If uNetApi(l).plocalname Then
                lLength = lstrlen(uNetApi(l).plocalname)
                uNet(lMin + l).slocalname = Space$(lLength)
                copymem ByVal uNet(lMin + l).slocalname, ByVal uNetApi(l).plocalname, lLength
             End If
             
             If uNetApi(l).premotename Then
                lLength = lstrlen(uNetApi(l).premotename)
                uNet(lMin + l).sremotename = Space$(lLength)
                copymem ByVal uNet(lMin + l).sremotename, ByVal uNetApi(l).premotename, lLength
             End If
        Next l
      End If
      lMin = lMin + lCount
    Loop While lReturn = error_more_data
  End If
  If hEnum Then l = WNetCloseEnum(hEnum)
Loop While lLastIndex < lMin
If UBound(uNet) > 0 Then
   For l = 0 To UBound(uNet)
       If uNet(l).dwdisplaytype = resourcedisplaytype_server Then List1.AddItem uNet(l).sremotename
   Next l
End If
End Sub

Private Sub list1_click()

Text2.Text = List1.Text


End Sub
