VERSION 5.00
Begin VB.Form frmEtiquetasLocalizador 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Lector de código de barras"
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   Icon            =   "frmEtiquetasLocalizador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   2145
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   2850
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   2295
      Picture         =   "frmEtiquetasLocalizador.frx":030A
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "F1-Código de Barras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   2175
   End
End
Attribute VB_Name = "frmEtiquetasLocalizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
txtcodigo.Text = ""
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 400
    Me.Top = frmMenu.SmartMenuXP1.Height + 150
    
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Hide
        buscar_codigo
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    ElseIf KeyAscii = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub buscar_codigo()
    If txtcodigo <> "" Then
        Select Case UCase(Left(txtcodigo, 1))
        Case "M"
            gmuestra = CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
            frmVerMuestra.Show 1
        Case "F"
            gdoc = CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
            frmListadoDocPago.Show
        Case "R"
            frmREX_Gestion.Show
            frmREX_Gestion.CARGAR_CODIGO CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
        Case "P"
            frmRPR_Gestion.Show
            frmRPR_Gestion.CARGAR_CODIGO CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
        Case "E" ' Equipo
            consulta_equipo CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
        Case "C" ' Calibracion
            consulta_calibracion CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
        Case "V" ' Verificación
            consulta_verificacion CLng(Mid(txtcodigo, 2, Len(txtcodigo) - 1))
        Case Else
            MsgBox "No localizo el código de barras.", vbCritical, App.Title
            txtcodigo = ""
            txtcodigo.SetFocus
        End Select
    End If
    txtcodigo = ""
End Sub
Private Sub consulta_equipo(ID As Long)
    Dim objfrm As New frmEquipoEdicion
    Dim lngid As Long
    Dim objEquipo As New clsEquipos
    lngid = ID
    
    Call objEquipo.Carga(lngid)
    
    Set objfrm.EQUIPO = objEquipo
    
    If objEquipo.getALTA_BAJA = 1 Then
        objfrm.TipoEdicion = visualizar
    Else
        objfrm.TipoEdicion = edicion
    End If
    
    objfrm.Show vbModal
    
    Unload objfrm
    Set objfrm = Nothing
End Sub
Private Sub consulta_calibracion(ID As Long)
    Dim objfrm  As New frmEquipoEdicionCalibracion
    Dim strId As String
    Dim intEstado As Integer
    Dim mvarobjEquipo As New clsEquipos
    Dim oEC As New clsEquipoCalibracion
    oEC.Carga ID
    mvarobjEquipo.Carga oEC.getEQUIPO_ID
    
    strId = ID
    intEstado = oEC.getESTADO
    
    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = strId
        If intEstado = 0 Or intEstado = 3 Then
            .TipoEdicion = enumTipoEdicion.edicion
        Else
            .TipoEdicion = enumTipoEdicion.visualizar
        End If
                
        .Show vbModal
    End With
    
    Unload objfrm
    Set objfrm = Nothing
End Sub
Private Sub consulta_verificacion(ID As Long)
    Dim objfrm  As New frmEquipoEdicionVerificacion
    Dim strId As String
    Dim intEstado As Integer
    Dim mvarobjEquipo As New clsEquipos
    Dim oEV As New clsEquipoVerificacion
    oEV.Carga ID
    mvarobjEquipo.Carga oEV.getEQUIPO_ID
    
    strId = ID
    intEstado = oEV.getESTADO
    
    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = strId
        
        If intEstado = 0 Then
            .TipoEdicion = edicion ' si no está cerrado
        Else
            .TipoEdicion = visualizar
        End If
        
        
        .Show vbModal
        
    End With
    Unload objfrm
    Set objfrm = Nothing

End Sub
