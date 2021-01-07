VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim aPos As Integer 'Mantiene el valor global de la posición del scroll
Dim WithEvents ScrollBar As VB.VScrollBar 'Variable para crear el Control y utilizar sus procedimientos/eventos
Attribute ScrollBar.VB_VarHelpID = -1
Dim Cargado As Boolean

Sub DibujarScrollBar()
    If Not Cargado Then
        Set ScrollBar = Screen.ActiveForm.Controls.Add("VB.VScrollBar", _
                        "ScrollBar1", Screen.ActiveForm) 'Ejemplo extraido de la página de la microsoft
        ScrollBar.Max = Screen.ActiveForm.Height
        Ajustar
        ScrollBar.Visible = True
        ScrollBar.SmallChange = 50
        ScrollBar.LargeChange = 100
        Cargado = True
    End If
    Ajustar
End Sub

Private Sub Ajustar()
Dim ControlVerticalMaximo As Integer
Dim cControl As Control
Dim AltoStatus As Integer

On Error Resume Next
    For Each cControl In Screen.ActiveForm.Controls
        If cControl.Top > ControlVerticalMaximo Then
            ControlVerticalMaximo = cControl.Top + cControl.Height 'Obtengo la posición vertical del control que se encuentre mas abajo
        End If
        
        If TypeName(cControl) = "StatusBar" Then 'Si tengo un control estatus también se evalúa su altura
            AltoStatus = cControl.Height
        End If
    Next
    
    If ScrollBar.Height > ControlVerticalMaximo Then
        ScrollBar.Max = ScrollBar.Height - ControlVerticalMaximo + 50
    Else
        ScrollBar.Max = ControlVerticalMaximo - ScrollBar.Height + 50
    End If
    
    If ControlVerticalMaximo <= Screen.ActiveForm.ScaleHeight Then
        ScrollBar.Enabled = False
    Else
        ScrollBar.Enabled = True
    End If
    
    ScrollBar.Height = Screen.ActiveForm.ScaleHeight - AltoStatus
    ScrollBar.Top = 0
    ScrollBar.Left = Screen.ActiveForm.ScaleWidth - ScrollBar.Width
    aPos = 0
    ScrollBar.value = 0
End Sub

Private Sub ScrollBar_Change()
    cambiar
End Sub

Private Sub ScrollBar_Scroll()
    cambiar
End Sub

Private Sub cambiar()
Dim a As Control
Dim Pos As Integer

On Error Resume Next
    
    Pos = aPos - ScrollBar.value
    
    For Each a In Screen.ActiveForm.Controls
        If Not TypeName(a) = "VScrollBar" Then
            If a.Container = "" Then
                a.Top = a.Top + Pos 'Muevo los controles dependiendo de la posición del scroll
            End If
        End If
    Next
    'Debug.Print "Min: " & ScrollBar.min
    'Debug.Print "Max: " & ScrollBar.Max
    
    aPos = ScrollBar.value
End Sub
