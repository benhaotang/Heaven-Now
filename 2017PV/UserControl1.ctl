VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00FF8080&
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   ScaleHeight     =   1815
   ScaleWidth      =   1980
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1320
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   600
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Name1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label dis 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   15
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ȱʡ����ֵ:
Const m_def_nam = ""
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'���Ա���:
Dim m_nam As String
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'�¼�����:
Event Click() 'MappingInfo=Label3,Label3,-1,Click
Attribute Click.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ťʱ������"
'Event Click()
Event DblClick()
Attribute DblClick.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ť���ٴΰ��²��ͷ���갴ťʱ������"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "���û���ӵ�н���Ķ����ϰ��������ʱ������"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "���û����º��ͷ� ANSI ��ʱ������"
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "���û���ӵ�н���Ķ������ͷż�ʱ������"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseDown.VB_Description = "���û���ӵ�н���Ķ����ϰ�����갴ťʱ������"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseMove.VB_Description = "���û��ƶ����ʱ������"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseUp.VB_Description = "���û���ӵ�н���Ķ������ͷ���귢����"
Dim time, ti As Integer



'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "ָ�� Label �� Shape �ı�����ʽ��͸���Ļ��ǲ�͸���ġ�"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "����/���ö���ı߿���ʽ��"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "ǿ����ȫ�ػ�һ������"
     
End Sub

Private Sub Label1_Change()
Picture1.Picture = LoadPicture(Label1.Caption)

End Sub

Private Sub Name_Click()

End Sub



Private Sub Timer1_Timer()
ti = ti + 1
On Error Resume Next
If ti Mod (4 * time) > time And ti Mod (4 * time) < 2 * time Then
dis.Visible = True
Picture1.Visible = False
dis.Height = dis.Height + 0.5 * 375
End If
On Error Resume Next
If ti Mod (4 * time) > 2 * time And ti Mod (4 * time) < 3 * time Then
dis.Visible = True
Picture1.Visible = False
dis.Height = dis.Height - 0.5 * 375
End If
On Error Resume Next
If ti Mod (4 * time) > 3 * time Then
Picture1.Visible = True
End If
If ti Mod (4 * time) = 0 Or ti > 4 * time Then
dis.Visible = False
ti = 0
End If

End Sub

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
   
ti = 0
If Label1.Caption <> "" Then Image.Picture = LoadPicture(Label1.Caption)
Timer1.Enabled = True

    m_nam = m_def_nam
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    dis.Caption = PropBag.ReadProperty("disc", "")
    Name1.Caption = PropBag.ReadProperty("nam", "")
    Label1.Caption = PropBag.ReadProperty("imageu", "")
    m_nam = PropBag.ReadProperty("nam", m_def_nam)
    Label2.BorderStyle = PropBag.ReadProperty("dura", 0)
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("imageu", Picture, Nothing)
    Call PropBag.WriteProperty("disc", dis.Caption, "")
    Call PropBag.WriteProperty("nam", Name1.Caption, "")
End Sub
'
''ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
''MappingInfo=Image,Image,-1,Picture
'Public Property Get imageu() As Picture
'    Set imageu = Image.Picture
'End Property
   
'


    
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=dis,dis,-1,Caption
Public Property Get disc() As String
Attribute disc.VB_Description = "����/���ö���ı������л�ͼ��������ı���"
    disc = dis.Caption
End Property

Public Property Let disc(ByVal New_disc As String)
    dis.Caption() = New_disc
    PropertyChanged "disc"
End Property
'
''ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
''MappingInfo=Name,Name,-1,Caption
'Public Property Get nam() As String
'    nam = Name1.Caption
'End Property
'
'Public Property Let nam(ByVal New_nam As String)
'    Name1.Caption() = New_nam
'    PropertyChanged "nam"
'End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get imageu() As String
Attribute imageu.VB_Description = "����/���ö���ı������л�ͼ��������ı���"
    imageu = Label1.Caption
End Property

Public Property Let imageu(ByVal New_imageu As String)
    Label1.Caption() = New_imageu
    PropertyChanged "imageu"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=13,0,0,
Public Property Get nam() As String
Attribute nam.VB_Description = "����/���ö���ı������л�ͼ��������ı���"
    nam = m_nam
End Property

Public Property Let nam(ByVal New_nam As String)
    Name1.Caption = New_nam
    PropertyChanged "nam"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=Label2,Label2,-1,BorderStyle
Public Property Get dura() As Integer
Attribute dura.VB_Description = "����/���ö���ı߿���ʽ��"
    dura = Label2.BorderStyle
End Property

Public Property Let dura(ByVal New_dura As Integer)
    time = New_dura
    PropertyChanged "dura"
End Property

Private Sub Label3_Click()
    RaiseEvent Click
End Sub

