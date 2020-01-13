VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl Use_ctrl 
   BackStyle       =   0  '透明
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
   ClipBehavior    =   0  '无
   ScaleHeight     =   1815
   ScaleWidth      =   2370
   Begin VB.Timer Timer1 
      Left            =   1215
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   30
      Picture         =   "Use_ctrl.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   30
      Width           =   540
   End
   Begin MSCommLib.MSComm MSCom 
      Left            =   150
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Use_ctrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim M_Fh As Single
Dim M_WY As Single
Dim M_dFH As Single
Dim M_dWY As Single
Dim M_dBX As Single
Dim M_ComPort As Integer
Dim M_Enable As Boolean
Dim M_Data_Count As Long
Event onTest(FuHe As Single, WeiYi As Single)
Public Property Get ComPort() As Integer
    ComPort = M_ComPort
End Property
Public Property Let ComPort(N_ComPort As Integer)
    M_ComPort = N_ComPort
    PropertyChanged "comPort"
End Property
'Public Property Get FH() As Integer
'  dFH = M_disFH
'End Property
'Public Property Let FH(N_FH As Integer)
'  M_disFH = N_FH
'  PropertyChanged "FH"
'End Property
'Public Property Get dWY() As Integer
'  dWY = M_dWY
'End Property
'Public Property Let dWY(n_dWY As Integer)
'   dWY = n_dWY
'  PropertyChanged "dWY"
'End Property
Public Property Get Enabled() As Boolean
    Enabled = M_Enable
End Property
Public Property Let Enabled(N_Enable As Boolean)
    If N_Enable = True Then
        Call InitCom(True)
        M_Enable = N_Enable
    Else
        M_Data_Count = 0
        Call InitCom(False)
        M_Enable = N_Enable
    End If
    PropertyChanged "Enabled"
End Property
Private Sub InitCom(T As Boolean)
    If T = True Then
        MSCom.CommPort = M_ComPort
        MSCom.PortOpen = True
        Timer1.Interval = 100
    Else
        MSCom.PortOpen = False
        Timer1.Interval = 0
    End If
End Sub

Private Sub Timer1_Timer()
    DoEvents
    If M_Enable = True Then
        YL_data
    End If
End Sub
Private Sub UserControl_InitProperties()
    M_ComPort = 1
    m_Resu = 100
    M_Enable = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
M_ComPort = PropBag.ReadProperty("comPort", 1)
'M_disFH = PropBag.ReadProperty("disFH", 1)
'M_disWY = PropBag.ReadProperty("disWY", 1)
'M_disBX = PropBag.ReadProperty("disBX", 1)
'm_Resu = PropBag.ReadProperty("Resu", 100)
M_Enable = PropBag.ReadProperty("Enabled", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("comPort", M_ComPort, 1)
    'Call PropBag.WriteProperty("disFH", M_disFH, 1)
    'Call PropBag.WriteProperty("disWY", M_disWY, 1)
    'Call PropBag.WriteProperty("disBX", M_disBX, 1)
    'Call PropBag.WriteProperty("Resu", m_Resu, 100)
    Call PropBag.WriteProperty("Enabled", M_Enable, False)
End Sub

Private Sub YL_data() '数据模拟
    Dim i As Long
    M_Data_Count = M_Data_Count + 1
    If Data_Count > 50 Then
        M_Fh = 1.3 * M_Data_Count * Rnd(Format(Time, "ss"))
        M_WY = M_Data_Count * 0.2
    Else
        M_Fh = 1.3 * M_Data_Count
        M_WY = M_Data_Count * 0.2
    End If
    Data_OutPut
End Sub
Private Sub Data_OutPut()
    MSCom.Output = "A" & M_Fh & "B" & M_WY
    Return_data
End Sub
Private Sub Return_data()
    Dim P_Fh As Single
    Dim P_Wy As Single
    Dim P_Str_Data As String
    Dim PP As String
    P_Str_Data = MSCom.Input
    If Trim(P_Str_Data) = "" Then Exit Sub
    For i = 1 To Len(P_Str_Data)
        If Mid(P_Str_Data, i, 1) = "A" Then
            P_Fh = Val(Mid(P_Str_Data, i + 1, Len(P_Str_Data)))
        End If
        If Mid(P_Str_Data, i, 1) = "B" Then
            P_Wy = Val(Mid(P_Str_Data, i + 1, Len(P_Str_Data)))
        End If
    Next
    RaiseEvent onTest(P_Fh, P_Wy)
End Sub
