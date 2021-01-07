Attribute VB_Name = "centrarMdi"

Option Explicit

Public Declare Function _
    SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, _
        ByVal uParam As Long, ByRef lpvParam As Any, _
        ByVal fuWinIni As Long) As Long
Private Type RECT
    Izq As Long
    Sup As Long
    Der As Long
    Inf As Long
End Type

Public Const SPI_GETWORKAREA = 48

Public Sub CentrarForma(frm As Form)
Dim R As RECT
Dim lRes As Long
Dim lAncho As Long
Dim lLargo As Long
    With frm
        If .WindowState = vbMinimized Or .WindowState = vbMaximized Then Exit Sub
    End With
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, R, 0)
    If lRes Then
        With R
            .Izq = Screen.TwipsPerPixelX * .Izq
            .Sup = Screen.TwipsPerPixelY * .Sup
            .Der = Screen.TwipsPerPixelX * .Der
            .Inf = Screen.TwipsPerPixelY * .Inf
            lAncho = .Der - .Izq
            lLargo = .Inf - .Sup
            frm.Move .Izq + (lAncho - frm.Width) \ 2, .Sup + (lLargo - frm.Height) \ 2
        End With
    End If
End Sub

