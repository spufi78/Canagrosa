VERSION 5.00
Begin VB.UserControl ControlPanelXP 
   BackColor       =   &H00CCCCFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00CC3300&
   MouseIcon       =   "ControlPanelXP.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   4320
      MouseIcon       =   "ControlPanelXP.ctx":0152
      MousePointer    =   99  'Custom
      Picture         =   "ControlPanelXP.ctx":02A4
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   0
      Top             =   105
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   180
      Left            =   4320
      Picture         =   "ControlPanelXP.ctx":064E
      Top             =   2745
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   4320
      Picture         =   "ControlPanelXP.ctx":09F8
      Top             =   3015
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "ControlPanelXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
'Nombre:            ControlPanelXP
'Autor:             Leandro Ascierto
'Fecha:             31/08/08
'Revición:          1
'Descripcion:       Contenedor de controles
'-------------------------------------------------------------------
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Private Const CLR_INVALID = -1

Private Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type


Dim Blue As Integer
Dim Green As Integer
Dim Red As Integer
Dim i As Integer
Dim m_Height As Long

Dim m_Open As Boolean
Dim m_HeaderColor As Long
Dim m_Caption As String
Dim m_TextColor As Long
Dim m_Icon As StdPicture
Dim m_PictureBack As StdPicture
Dim m_CanExpand As Boolean
Dim m_Brightness As Integer
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Expand(State As Boolean)

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property


Public Property Get Brightness() As Integer
    Brightness = m_Brightness
End Property

Public Property Let Brightness(ByVal New_Brightness As Integer)
    m_Brightness = IIf(New_Brightness > 10, 10, Abs(New_Brightness))
    PropertyChanged "Brightness"
    RedrawControl
End Property



Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    RedrawControl
End Property

Public Property Get PictureBack() As Picture
    Set PictureBack = m_PictureBack
End Property

Public Property Set PictureBack(ByVal New_PictureBack As Picture)
    Set m_PictureBack = New_PictureBack
    PropertyChanged "PictureBack"
    RedrawControl
End Property

Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property
Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    PropertyChanged "Icon"
    RedrawControl
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    m_TextColor = New_TextColor
    PropertyChanged "TextColor"
    RedrawControl
End Property

Public Property Get HeaderColor() As OLE_COLOR
    HeaderColor = m_HeaderColor
End Property

Public Property Let HeaderColor(ByVal New_HeaderColor As OLE_COLOR)
    m_HeaderColor = New_HeaderColor
    PropertyChanged "HeaderColor"
    RedrawControl
End Property

Public Property Get PanelOpen() As Boolean
    PanelOpen = m_Open
End Property

Public Property Let PanelOpen(ByVal New_PanelOpen As Boolean)
    m_Open = New_PanelOpen
    PropertyChanged "PanelOpen"
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    RedrawControl
End Property

Public Property Get CanExpand() As Boolean
    CanExpand = m_CanExpand
End Property

Public Property Let CanExpand(ByVal New_CanExpand As Boolean)
    m_CanExpand = New_CanExpand
    PropertyChanged "CanExpand"
    Picture1.Visible = m_CanExpand
End Property

Private Sub UserControl_InitProperties()
m_Caption = Ambient.DisplayName
m_TextColor = vbBlue
m_HeaderColor = vbRed
UserControl.BackColor = &HCCCCFF
m_Open = True
m_CanExpand = True
m_Brightness = 4
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub


Public Property Get SizeHeight() As Single
Attribute SizeHeight.VB_Description = "Establece el Alto en tiempo de ejecución"
Attribute SizeHeight.VB_MemberFlags = "100c"
    SizeHeight = m_Height
End Property

Public Property Let SizeHeight(ByVal New_SizeHeight As Single)
m_Height = New_SizeHeight
PropertyChanged "Height"
UserControl_Resize
RedrawControl
End Property


Private Sub UserControl_Paint()
RedrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    Set m_PictureBack = PropBag.ReadProperty("PictureBack", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HCCCCFF)
    m_TextColor = PropBag.ReadProperty("TextColor", vbBlue)
    m_HeaderColor = PropBag.ReadProperty("HeaderColor", vbRed)
    m_Open = PropBag.ReadProperty("PanelOpen", True)
    m_CanExpand = PropBag.ReadProperty("CanExpand", True)
    m_Height = PropBag.ReadProperty("Height", 1000)
    m_Brightness = PropBag.ReadProperty("Brightness", 4)
    Picture1.Visible = m_CanExpand
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("PictureBack", m_PictureBack, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HCCCCFF)
    Call PropBag.WriteProperty("TextColor", m_TextColor, vbBlue)
    Call PropBag.WriteProperty("HeaderColor", m_HeaderColor, vbRed)
    Call PropBag.WriteProperty("PanelOpen", m_Open, True)
    Call PropBag.WriteProperty("CanExpand", m_CanExpand, True)
    Call PropBag.WriteProperty("Height", m_Height, 1000)
    Call PropBag.WriteProperty("Brightness", m_Brightness, 4)
End Sub



Private Sub RedrawControl()
    'Dim Rgn As Long
    Dim LightColor As Integer
    Dim DarkColor As Integer
    Dim mBrush As Long
    Dim r As RECT
    

    
    
    'UserControl.AutoRedraw = True
    'UserControl.Cls

    If Not PictureBack Is Nothing Then
        mBrush = CreatePatternBrush(m_PictureBack)
        SetRect r, 0, 24, UserControl.ScaleWidth, UserControl.ScaleHeight
        FillRect UserControl.hdc, r, mBrush
        DeleteObject mBrush
    End If



    
    GetRgb TranslateColor(m_HeaderColor)
    
    For i = 24 To 1 Step -1
        LightColor = (40 - i) * m_Brightness
        UserControl.Line (0, i)-(UserControl.Width, i), RGB(Red + LightColor, Green + LightColor, Blue + LightColor)
    Next
    
    
    UserControl.Line (0, 24)-(UserControl.Width, 24), TransColorRGB(UserControl.BackColor, -150)
    UserControl.Line (0, 25)-(UserControl.Width, 25), TransColorRGB(UserControl.BackColor, -100)
    UserControl.Line (0, 26)-(UserControl.Width, 26), TransColorRGB(UserControl.BackColor, -50)
    
    UserControl.ForeColor = TransColorRGB(m_HeaderColor, -50)
    RoundRect UserControl.hdc, 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 9, 9
    
    If Not m_Icon Is Nothing Then
        UserControl.PaintPicture m_Icon, 6, 4, 16, 16
        UserControl.CurrentX = 26: UserControl.CurrentY = 6
    Else
        UserControl.CurrentX = 8: UserControl.CurrentY = 6
    End If
    
    UserControl.ForeColor = m_TextColor
    UserControl.Print m_Caption
    
    'UserControl.AutoRedraw = False
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function


Private Function TransColorRGB(color As Long, Intencity As Long) As Long

GetRgb TranslateColor(color)

If Intencity < 0 Then
    Red = IIf(Red + Intencity >= 0, Red + Intencity, 0)
    Green = IIf(Green + Intencity >= 0, Green + Intencity, 0)
    Blue = IIf(Blue + Intencity >= 0, Blue + Intencity, 0)
Else
    Red = IIf(Red + Intencity <= 255, Red + Intencity, 255)
    Green = IIf(Green + Intencity <= 255, Green + Intencity, 255)
    Blue = IIf(Blue + Intencity <= 255, Blue + Intencity, 255)
End If

TransColorRGB = RGB(Red, Green, Blue)
End Function


Private Sub GetRgb(color As Long)
    Blue = ((color And &HFF0000) / 65536)
    Green = ((color And &HFF00FF00) / 256&)
    Red = color Mod 256
End Sub

Private Sub Picture1_Click()
m_Open = IIf(m_Open, False, True)
UserControl_Resize
RedrawControl
RaiseEvent Expand(m_Open)

End Sub

Private Sub UserControl_Resize()
Dim Rgn As Long

If Not Ambient.UserMode Then
    m_Height = UserControl.Height
    PropertyChanged "Height"
Else
    If m_Open = False Then
        UserControl.Height = 390
        Picture1 = Image3.Picture
    Else
        UserControl.Width = UserControl.Width
        UserControl.Height = m_Height
        Picture1 = Image2.Picture
    End If
End If

    Rgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 9, 9)
    SetWindowRgn UserControl.Hwnd, Rgn, True
    DeleteObject Rgn


Picture1.Left = UserControl.ScaleWidth - 24
'RedrawControl
End Sub

Private Sub UserControl_Show()
Dim Rgn As Long
Rgn = CreateRoundRectRgn(0, 0, 13, 13, 10, 10)
SetWindowRgn Picture1.Hwnd, Rgn, True

UserControl_Resize



End Sub
