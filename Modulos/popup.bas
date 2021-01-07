Attribute VB_Name = "popup"
Public Declare Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
            

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst _
    As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 _
    As Long, ByVal un2 As Long) As Long
    

Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_COPYRETURNORG = &H4
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1

Private Const ILC_COLOR = &H0
Private Const ILC_MASK = &H1
Public Const ILC_COLOR4 = &H4
Public Const ILC_COLOR8 = &H8
Public Const ILC_COLOR16 = &H10
Public Const ILC_COLOR24 = &H18
Public Const ILC_COLOR32 = &H20
Public Const ILD_NORMAL = 0
Public Enum POP
    IDCLOSE = 2
End Enum
Function LoadBitmap(Path As String) As Long
    LoadBitmap = LoadImage(App.hInstance, Path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
End Function

Public Sub MensajePop(popup As XtremeSuiteControls.PopupControl, mensaje As Long, titulo As String, texto As String)
    Dim Item As PopupControlItem
    
    popup.RemoveAllItems
    popup.Icons.RemoveAll
    
    Set Item = popup.AddItem(2, 6, 170, 19, "Mensajes de usuario...")
    Item.Hyperlink = False
    
    ' Titulo
    Set Item = popup.AddItem(0, 27, 360, 0, titulo)
    Item.TextAlignment = DT_CENTER
    Item.Font.Bold = True
    Item.CalculateHeight
    Item.Hyperlink = False
'    Item.CalculateWidth

    ' Texto
    Set Item = popup.AddItem(10, 50, 360, 100, texto)
    Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    Item.CalculateHeight
    Item.Hyperlink = False
'    Item.ID = IDSITE
    
    Set Item = popup.AddItem(351, 6, 364, 19, "")
    Item.SetIcons LoadBitmap(ReadINI(App.Path + "\config.ini", "logo", "recursos") & "cerrar.bmp"), 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
    Item.Caption = mensaje
    Item.Hyperlink = False
    Item.ID = POP.IDCLOSE
'    Set Item = Popup.AddItem(7, 6, 20, 19, "")
'    Item.SetIcons frmMenu.botones.ListImages(1).Picture, 0, xtpPopupItemIconNormal
'    Set Item = popup.AddItem(115, 102, 160, 120, "")
'    Item.SetIcons LoadBitmap(ReadINI(App.Path + "\config.ini", "logo", "recursos") & "logo.jpg"), 0, xtpPopupItemIconNormal
    
    popup.VisualTheme = xtpPopupThemeMSN
    popup.SetSize 370, 400

End Sub

Public Sub popupCreacion(mensaje As Long, titulo As String, texto As String)
    Dim POP As XtremeSuiteControls.PopupControl
    Set POP = frmMenu.PopupControl
    MensajePop POP, mensaje, titulo, texto
    POP.Animation = xtpPopupAnimationSlide
'    POP.ShowDelay = False
    POP.ShowDelay = 5000
    POP.Show
End Sub

