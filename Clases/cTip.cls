VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API
'#################

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

'UDT ENUM
'#################


Enum e_TipoIcono
    ISinIcono = 0
    ivbInformation = 1
    ivbExclamation = 2
    ivbCritical = 3
End Enum

Enum e_EstiloIcono
    EstiloClasico
    EstiloBalloon
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

'Var
'###################

Private m_Handle As Long
Private m_Handle2 As Long
Private t_ToolTip As TOOLINFO

Private m_Icono As e_TipoIcono
Private m_EstiloToolTip As e_EstiloIcono
Private m_Texto As String
Private m_Duracion As Long
Private m_ColorFondo As Long
Private m_Titulo As String
Private m_ColorTexto As Long


'Sub y Func
'#################

Public Function Iniciar(ByVal h As Long) As Boolean
   Dim hSt As Long
   
   If m_Handle <> 0 Then DestroyWindow m_Handle
   m_Handle2 = h
   
   hSt = &H1 Or &H2
   
   If m_EstiloToolTip = EstiloBalloon Then hSt = hSt Or &H40
    
   m_Handle = CreateWindowEx(0&, "tooltips_class32", vbNullString, hSt, _
      &H80000000, &H80000000, &H80000000, &H80000000, 0&, 0&, App.hInstance, 0&)
               
   With t_ToolTip

      .lFlags = &H10 Or &H1
      .hwnd = m_Handle2
      .lId = m_Handle2
      .hInstance = App.hInstance
      .lSize = Len(t_ToolTip)
   End With
   
   SendMessage m_Handle, (&H400 + 4), 0&, t_ToolTip
   
   If m_Titulo <> "" Or m_Icono <> ISinIcono Then
      SendMessage m_Handle, (&H400 + 32), CLng(m_Icono), ByVal m_Titulo
   End If

   If m_ColorTexto <> Empty Then
      SendMessage m_Handle, (&H400 + 20), m_ColorTexto, 0&
   End If

   If m_ColorFondo <> Empty Then
      SendMessage m_Handle, (&H400 + 19), m_ColorFondo, 0&
   End If
   
   SendMessageLong m_Handle, (&H400 + 3), 2, m_Duracion
   SendMessageLong m_Handle, (&H400 + 3), 3, 400
End Function

Private Sub Class_Initialize()
   InitCommonControls
   m_Duracion = 5000
End Sub

Private Sub Class_Terminate()
   Terminar
End Sub

Public Sub Terminar()
   If m_Handle <> 0 Then DestroyWindow m_Handle
End Sub



'Prop
'#############

Public Property Let Icono(ByVal Valor As e_TipoIcono)
   m_Icono = Valor
   If m_Handle <> 0 And m_Titulo <> Empty And m_Icono <> ISinIcono Then
      SendMessage m_Handle, (&H400 + 32), CLng(m_Icono), ByVal m_Titulo
   End If
End Property

Public Property Get Icono() As e_TipoIcono
   Icono = m_Icono
End Property

Public Property Let ColorFuente(ByVal Valor As Long)
   m_ColorTexto = Valor
   If m_Handle <> 0 Then SendMessage m_Handle, (&H400 + 20), m_ColorTexto, 0&
End Property

Public Property Get ColorFuente() As Long
   ColorFuente = m_ColorTexto
End Property

Public Property Let Titulo(ByVal Valor As String)
   m_Titulo = Valor
   If m_Handle <> 0 And m_Titulo <> Empty And m_Icono <> ISinIcono Then
      SendMessage m_Handle, (&H400 + 32), CLng(m_Icono), ByVal m_Titulo
   End If
End Property

Public Property Get Titulo() As String
   Titulo = t_ToolTip.lpStr
End Property

Public Property Let ColorFondo(ByVal Valor As Long)
   m_ColorFondo = Valor
   If m_Handle <> 0 Then
      SendMessage m_Handle, (&H400 + 19), m_ColorFondo, 0&
   End If
End Property

Public Property Get ColorFondo() As Long
   ColorFondo = m_ColorFondo
End Property

Public Property Let Texto(ByVal Valor As String)
   m_Texto = Valor
   t_ToolTip.lpStr = Valor
   If m_Handle <> 0 Then
      SendMessage m_Handle, (&H400 + 12), 0&, t_ToolTip
   End If
End Property

Public Property Get Texto() As String
   Texto = m_Texto
End Property

Public Property Get Duracion() As Long
   Duracion = m_Duracion
End Property

Public Property Let Duracion(ByVal lData As Long)
   m_Duracion = lData
End Property

Public Property Let Estilo(ByVal Valor As e_EstiloIcono)
   m_EstiloToolTip = Valor
End Property

Public Property Get Estilo() As e_EstiloIcono
   Estilo = m_EstiloToolTip
End Property


