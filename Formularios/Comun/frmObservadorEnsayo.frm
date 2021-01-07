VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmObservadorEnsayo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indicar Observador en Ensayo"
   ClientHeight    =   8340
   ClientLeft      =   4335
   ClientTop       =   4290
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmObservadorEnsayo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCualificacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Cualificación"
      Height          =   885
      Left            =   1620
      Picture         =   "frmObservadorEnsayo.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7335
      Width           =   1500
   End
   Begin VB.Frame frmDuplicado 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ensayo duplicado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   90
      TabIndex        =   22
      Top             =   6165
      Width           =   7170
      Begin pryCombo.miCombo cmbResultado1 
         Height          =   315
         Left            =   1215
         TabIndex        =   23
         Top             =   270
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
      End
      Begin pryCombo.miCombo cmbResultado2 
         Height          =   315
         Left            =   1215
         TabIndex        =   25
         Top             =   675
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado 2º"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   26
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado 1º"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   915
      End
   End
   Begin VB.Frame frmDatos 
      BackColor       =   &H00C0C0C0&
      Height          =   4605
      Left            =   90
      TabIndex        =   9
      Top             =   1395
      Width           =   7170
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2790
         TabIndex        =   21
         Top             =   855
         Width           =   4200
      End
      Begin VB.CheckBox chkEs_Recualificacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Marque si se trata de una RECUALIFICACIÓN"
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   2385
         Width           =   6135
      End
      Begin VB.CheckBox chkEs_duplicado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "El ensayo se realiza por duplicado, formador y observador aportan un resultado y se evalua."
         Height          =   285
         Left            =   180
         TabIndex        =   19
         Top             =   2025
         Width           =   6765
      End
      Begin VB.OptionButton opObservador 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observador"
         Height          =   285
         Index           =   1
         Left            =   4995
         TabIndex        =   18
         Top             =   1575
         Width           =   1230
      End
      Begin VB.OptionButton opObservador 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formador"
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   17
         Top             =   1575
         Width           =   1230
      End
      Begin VB.TextBox txtobservaciones 
         Height          =   1320
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3195
         Width           =   6855
      End
      Begin pryCombo.miCombo cmbobservador 
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   1170
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
      End
      Begin pryCombo.miCombo cmbEnsayo 
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   450
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿La persona que observa actua como?"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   16
         Top             =   1620
         Width           =   2790
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones sobre el resultado de la formación"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   13
         Top             =   2925
         Width           =   3480
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Quién realiza el ensayo?"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   12
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¿Quién es el Observador?"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   11
         Top             =   900
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver P.N.T."
      Height          =   885
      Left            =   90
      Picture         =   "frmObservadorEnsayo.frx":26E4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7335
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7320
      Width           =   1185
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4365
      Left            =   450
      TabIndex        =   6
      Top             =   3555
      Visible         =   0   'False
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   7699
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin pryCombo.miCombo cmbPNT 
      Height          =   330
      Left            =   90
      TabIndex        =   27
      Top             =   900
      Visible         =   0   'False
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   582
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Evidencias de la cualificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   450
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   630
      Width           =   390
   End
   Begin VB.Label lblCodPnt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   900
      Width           =   7095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCodTipoMuestraEnsayo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   270
      Width           =   7275
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Muestra/Ensayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   1830
   End
End
Attribute VB_Name = "frmObservadorEnsayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MUESTRA_ID As Long
Public TIPO_DETERMINACION_ENSAYO_ID As Long
Public DETERMINACION_ENSAYO_ID  As Long
'M0961-I
Public SELLANTE_ID As Long
Public ENSAYO As Long
'M0961-F
Public ES_CONTROL_EFICACIA As Boolean
'MANTIS-807-I
Public FORMULARIO_ORIGEN As Integer
'MANTIS-807-F
Public MUESTRA_CERRADA As Boolean
Public TIPO_OBSERVACION_ID As Integer

Private DOCUMENTO_ID As Long
'Private ACTOR_ES_FORMADOR As Integer

Private Sub cmdCualificacion_Click()
    Dim oEC As New clsEmpleados_cualificaciones_m
   On Error GoTo cmdCualificacion_Click_Error

    If oEC.CargaPorMuestra(MUESTRA_ID) = False Then
        MsgBox "La muestra indicada no se encuentra en ningúna cualificación.", vbExclamation, App.Title
    Else
        Dim oC As New clsEmpleados_cualificaciones
        oC.Carga oEC.getCUALIFICACION_ID
        With frmEmpleados_Cualificaciones_Nueva
            .EMPLEADO_ID = oC.getEMPLEADO_ID
            .ID_CUALIFICACION = oEC.getCUALIFICACION_ID
            .Show 1
        End With
    End If
    Set oEC = Nothing

   On Error GoTo 0
   Exit Sub

cmdCualificacion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCualificacion_Click of Formulario frmObservadorEnsayo"
End Sub

' Para el tipo,
'    Si es Control de Eficacia, será el tipo de ensayo
'    Si es Determinacion, será el tipo de Determinacion (PNT)

'Private strCad As String
'Private rs As ADODB.RecordSet
'Private mvarobjUsuariosCualificados As New clsMc_cualificaciones_pnt

'Private mvarblnActualizar As Boolean
'Private mvarblnCargando_Inicial As Boolean
Private Sub cmdPNT_Click()
   On Error GoTo cmdPNT_Click_Error

     If cmbPNT.visible = True Then
        DOCUMENTO_ID = cmbPNT.getPK_SALIDA
     End If
     If DOCUMENTO_ID <> 0 Then
         Dim oPNT As New clsCa_documentos
         oPNT.mostrar DOCUMENTO_ID, True
         Set oPNT = Nothing
     Else
         MsgBox "No tiene PNT Vínculado.", vbExclamation, App.Title
     End If

   On Error GoTo 0
   Exit Sub

cmdPNT_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPNT_Click of Formulario frmObservadorEnsayo"
End Sub

Private Function comprobar_datos() As Boolean
    Dim strMsg  As String

    strMsg = ""
    comprobar_datos = False

    If cmbEnsayo.getPK_SALIDA <= 0 Then
        strMsg = strMsg & vbCrLf & " - Debe señalar el usuario que realiza el Ensayo."
    End If
    
    If cmbobservador.getPK_SALIDA <= 0 Then
        strMsg = strMsg & vbCrLf & " - Debe señalar el Observador del Ensayo."
    End If
    
    If cmbPNT.visible = True And cmbPNT.getTEXTO = "" Then
        strMsg = strMsg & vbCrLf & " - Debe indicar el PNT."
    End If
'    If DOCUMENTO_ID <= 0 Then
'        strMsg = strMsg & vbCrLf & " - El Tipo de Ensayo o Tipo de Determinación no tiene vinculado un Documento de Calidad. No es posible averiguar personal cualificado para este ensayo."
'    End If
    
    
    If frmDuplicado.visible = True Then
        If cmbResultado1.getTEXTO = "" Then
            strMsg = strMsg & vbCrLf & " - Debe indicar quien registra el Resultado 1º."
        End If
'SOLOCITADO POR LORENA : NO VALIDAR EL SEGUNDO RESULTADO
'        If cmbResultado2.getTEXTO = "" Then
'            strMsg = strMsg & vbCrLf & " - Debe indicar quien registra el Resultado 2º."
'        End If
    End If

    If Trim(strMsg) <> "" Then
        MsgBox "Se han encontrado los siguientes Errores: " & strMsg, vbInformation, "Observador del Ensayo"
        Exit Function
    End If

    comprobar_datos = True

End Function

'Private Sub determinar_rol_usuarios()
'
'    ' determina el rol que juega cada uno
'
'    Set rs = mvarobjUsuariosCualificados.Listado_por_pnt(DOCUMENTO_ID)
'
'    If rs.RecordCount = 0 Then
'        ' cuando no existe un pnt concreto, no se puede evaluar quien es formador y quien no
'        ' se toma por defecto que el actor está en formacion, y no se carga ningún usurario en combo
'
'        'llenar_combo cmbUsuarios, New clsUsuarios, 0, frmUsuarios, " AND ID_EMPLEADO=0"
'
'        cmbUsuarios.desactivar
'
'        lblUsuarioActual.Caption = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
'        cmdOk.Enabled = False
'
'        lblEnFormacion.Caption = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
'
'        lblFormador.ForeColor = vbGrayText
'        lblFormador.Caption = "(No existen usuarios cualificados para este Documento de Calidad)"
'        chkEsrecualificacion.Visible = False
'        Exit Sub
'    End If
'
'
'    ' revisa si el usuario actual es formador
'
'    rs.Filter = "USUARIO_ID = '" & CStr(USUARIO.getID_EMPLEADO) & "'"
'
'    If rs.RecordCount = 0 Then
'        ' no ha encontrado al usuario actor como formador
'        ACTOR_ES_FORMADOR = 0
'    Else
'        ACTOR_ES_FORMADOR = 1
'    End If
'
'    chkEsrecualificacion.Visible = False
'
'
'    rs.Filter = ""
'
'    mostrar_info_roles
'
'    lblUsuarioActual.Caption = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
'
'
'
'End Sub

'Private Sub mostrar_info_roles()
'
'    If ACTOR_ES_FORMADOR = 1 Then
'
'        ' El que realiza el ensayo es el FORMADOR
'
'        lblFormador.ForeColor = vbBlack
'        lblFormador.Caption = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
'
'        lblEnFormacion.Caption = "(Aún no especificado)"
'
'        llenar_combo cmbUsuarios, New clsUsuarios, 0, frmUsuarios, " AND ID_EMPLEADO <> " & CStr(USUARIO.getID_EMPLEADO)
'
'    Else
'        ' El que realiza el ensayo es el que está en formación
'
'        lblEnFormacion.ForeColor = vbBlack
'        lblEnFormacion.Caption = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
'
'        lblFormador.Caption = "(Aún no especificado)"
'
'        ' llena el combo con los cualificados para un pnt concreto
'        mvarobjUsuariosCualificados.llenar_combo_usuarios_cualificados cmbUsuarios, USUARIO.getID_EMPLEADO, True
'    End If
'End Sub


Private Sub opciones_edicion()

'    If MUESTRA_CERRADA Then
'        frmDatos.Enabled = False
'        cmdok.Enabled = False
'    End If


End Sub

Private Sub presentar_datos(oMo As clsMuestras_observadores)

    With oMo
        ES_CONTROL_EFICACIA = .getES_CONTROL_EFICACIA
        DOCUMENTO_ID = .getDOCUMENTO_ID
        MUESTRA_ID = .getMUESTRA_ID
        TIPO_DETERMINACION_ENSAYO_ID = .getTIPO_DETERMINACION_ENSAYO_ID
        DETERMINACION_ENSAYO_ID = .getDETERMINACION_ENSAYO_ID
        'M0961-I
        'SELLANTE_ID = .getSELLANTE_ID
        'ENSAYO = .getENSAYO
        'M0961-F
        
        If .getES_CONTROL_EFICACIA = 1 Then
            Dim oCE As New clsCe_tipos_ensayos
            oCE.Carga TIPO_DETERMINACION_ENSAYO_ID
            Dim oTA As New clsTipos_analisis
            oTA.CARGAR oCE.getTIPO_ANALISIS_ID
            lblCodTipoMuestraEnsayo.Caption = oTA.getNOMBRE & ": " & oCE.getNOMBRE
            Set oCE = Nothing
        Else
            'M0961-I
            'Dim oTD As New clsTipos_determinacion
            'oTD.CargarTipoDeterminacion .getTIPO_DETERMINACION_ENSAYO_ID
            'lblCodTipoMuestraEnsayo.Caption = oTD.getNOMBRE
            'Set oTD = Nothing
            
            Select Case FORMULARIO_ORIGEN
            Case 2 'formulario de Sellantes

                    Dim oSEnsayos   As New clsSellantes_ensayos
                    Dim oSeTipo     As New clsSellantes

                    oSEnsayos.Carga Me.SELLANTE_ID, Me.ENSAYO
                    oSeTipo.Carga (Me.SELLANTE_ID)
                    
                    lblCodTipoMuestraEnsayo.Caption = "Ensayo: " & oSEnsayos.getENSAYO & "  / Sellante: " & oSeTipo.getPROCESO
                    
                    Set oSEnsayos = Nothing
                    Set oSeTipo = Nothing
            Case 5 ' Plasma
                presentar_datos_plasma
                cmbPNT.MostrarElemento .getDOCUMENTO_ID
                    
            Case Else 'otros
            
                Dim oTD As New clsTipos_determinacion
                oTD.CargarTipoDeterminacion .getTIPO_DETERMINACION_ENSAYO_ID
                lblCodTipoMuestraEnsayo.Caption = oTD.getNOMBRE
'                DOCUMENTO_ID = oTD.getPNT_VINCULADO
                Set oTD = Nothing
'                Dim oPNT As New clsCa_documentos
'                If DOCUMENTO_ID = 0 Then
'                    lblCodPnt.Caption = "N/A"
'                ElseIf oPNT.Carga(DOCUMENTO_ID) Then
'                    lblCodPnt.Caption = "[" & oPNT.getCODIGO & "] " & oPNT.getNOMBRE
'                Else
'                    lblCodPnt.Caption = "N/A"
'                End If
                
            End Select
            'M0961-F
        End If
        If .getDOCUMENTO_ID <> 0 Then
            Dim oCA As New clsCa_documentos
            oCA.Carga .getDOCUMENTO_ID
            lblCodPnt.Caption = oCA.getNOMBRE
            Set oCA = Nothing
        End If
        chkEs_Recualificacion.Value = IIf(.getES_RECUALIFICACION = 1, vbChecked, vbUnchecked)
        chkEs_duplicado.Value = IIf(.getES_DUPLICADO = 1, vbChecked, vbUnchecked)
        
        cmbEnsayo.MostrarElemento .getUSUARIO_ID_REALIZACION
        cmbobservador.MostrarElemento .getUSUARIO_ID_OBSERVADOR
        
        cmbResultado1.MostrarElemento .getUSUARIO_RESULTADO1
        cmbResultado2.MostrarElemento .getUSUARIO_RESULTADO2
'        If .getREALIZACION_ES_FORMADOR = 1 Then
'            opEnsayo(0).value = True
'        Else
'            opEnsayo(1).value = True
'        End If
                
        If .getOBSERVADOR_ES_FORMADOR = 1 Then
            opObservador(0).Value = True
        Else
            opObservador(1).Value = True
        End If
        
        txtObservaciones = .getOBSERVACIONES
        
    End With

    Set oMo = Nothing
End Sub

Private Sub presentar_datos_ce()

    
    Dim oTipo_CE As New clsCe_tipos_ensayos
    Dim oTA As New clsTipos_analisis
    Dim oPNT As New clsCa_documentos
    
    oTipo_CE.Carga TIPO_DETERMINACION_ENSAYO_ID
    oTA.CARGAR oTipo_CE.getTIPO_ANALISIS_ID
    lblCodTipoMuestraEnsayo.Caption = oTA.getNOMBRE & ": " & oTipo_CE.getNOMBRE
    DOCUMENTO_ID = oTipo_CE.getPNT_VINCULADO
    
    If DOCUMENTO_ID = 0 Then
        lblCodPnt.Caption = "N/A"
    ElseIf oPNT.Carga(DOCUMENTO_ID) Then
        lblCodPnt.Caption = "[" & oPNT.getCODIGO & "] " & oPNT.getNOMBRE
    Else
        lblCodPnt.Caption = "N/A"
    End If
    
    Set oTipo_CE = Nothing
    Set oPNT = Nothing
    
End Sub
Private Sub presentar_datos_plasma()
       
   On Error GoTo presentar_datos_plasma_Error

    lblCodTipoMuestraEnsayo.Caption = "PLASMA"
    lblCodPnt.Caption = "N/A"
    lblCodPnt.visible = False
    
    cmbPNT.visible = True
    ' Cargar combo pnts plasma
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, " ID_DOCUMENTO IN (2527,2662,2663,2664,2665,2666) "

    

   On Error GoTo 0
   Exit Sub

presentar_datos_plasma_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure presentar_datos_plasma of Formulario frmObservadorEnsayo"
    
End Sub
Private Sub presentar_datos_determinacion()

    Dim oDeter As New clsDeterminaciones
    Dim oTipo_Deter As New clsTipos_determinacion
    Dim oPNT As New clsCa_documentos

    oDeter.CargarDeterminacion DETERMINACION_ENSAYO_ID
    oTipo_Deter.CargarTipoDeterminacion oDeter.getTIPO_DETERMINACION_ID
    
    TIPO_DETERMINACION_ENSAYO_ID = oDeter.getTIPO_DETERMINACION_ID
    DOCUMENTO_ID = oTipo_Deter.getPNT_VINCULADO
    lblCodTipoMuestraEnsayo.Caption = oTipo_Deter.getNOMBRE
    
    If DOCUMENTO_ID = 0 Then
        lblCodPnt.Caption = "N/A"
    ElseIf oPNT.Carga(DOCUMENTO_ID) Then
        lblCodPnt.Caption = "[" & oPNT.getCODIGO & "] " & oPNT.getNOMBRE
    Else
        lblCodPnt.Caption = "N/A"
    End If
    Set oTipo_Deter = Nothing
    Set oDeter = Nothing
    Set oPNT = Nothing
    
End Sub
'Private Sub cmbUsuarios_change()
''If mvarblnCargando_Inicial Then Exit Sub
'
'Dim blnCualificado As Boolean
'
'    If ACTOR_ES_FORMADOR = 1 Then
'        chkEsrecualificacion.Visible = mvarobjUsuariosCualificados.comprobar_usuario_es_cualificado(cmbUsuarios.getPK_SALIDA, DOCUMENTO_ID)
'
'        lblEnFormacion.Caption = cmbUsuarios.getTEXTO
'    Else
'        lblFormador.Caption = cmbUsuarios.getTEXTO
'    End If
'
'End Sub

'Private Sub cmdAvisoRecualificacion_Click()
'Dim strCad As String
'
'
'strCad = "ATENCIÓN: " & vbCrLf
'strCad = strCad & "Tanto el usuario que realiza el ensayo como el Observador, se encuentran cualificados para el Documento Pertinente." & vbCrLf
'strCad = strCad & "Se observa a su vez que no consta como RECUALIFICACIÓN, por ello, si decide marcarlo ahora, los roles de 'Formador'" & vbCrLf
'strCad = strCad & "y 'En Formación' de los usuarios implicados, no variará cuando seleccione la casilla la primera ocasión." & vbCrLf
'
'MsgBox strCad, vbCritical, "Aviso Importante Observador"
'
'End Sub
'M0961-I
'MANTIS-807-I
'Private Sub presentar_datos_sellantes()
'
'    Dim oSE As New clsSellantes_resultados
'    Dim oSeTipo As New clsSellantes
'    Dim oTipo_Deter As New clsTipos_determinacion
'    Dim oPNT As New clsCa_documentos
'
'    oSE.Carga MUESTRA_ID
'    oTipo_Deter.CargarTipoDeterminacion oSE.getTIPO_DETERMINACION_ID
'
'    DOCUMENTO_ID = oTipo_Deter.getPNT_VINCULADO
'    TIPO_DETERMINACION_ENSAYO_ID = oSE.getTIPO_DETERMINACION_ID
'
'    If TIPO_DETERMINACION_ENSAYO_ID = 0 Then
'       oSeTipo.Carga (oSE.getSELLANTE_ID)
'
'       lblCodTipoMuestraEnsayo.Caption = "Sellante: " & oSeTipo.getPROCESO
'       lblCodTipoMuestraEnsayo.AutoSize = True
'    Else
'        lblCodTipoMuestraEnsayo.Caption = oTipo_Deter.getNOMBRE
'    End If
'
'
'    If DOCUMENTO_ID = 0 Then
'
'        lblCodPnt.Caption = "N/A"
'        bloquear_controles
'
'    ElseIf oPNT.Carga(DOCUMENTO_ID) Then
'
'        lblCodPnt.Caption = "[" & oPNT.getCODIGO & "] " & oPNT.getNOMBRE
'        cmdPNT.Enabled = True
'
'    Else
'
'        lblCodPnt.Caption = "N/A"
'        bloquear_controles
'
'    End If
'
'    Set oTipo_Deter = Nothing
'    Set oSE = Nothing
'    Set oPNT = Nothing
'
'End Sub
Private Sub presentar_datos_sellantes()
    
    Dim oSe         As New clsSellantes_resultados
    Dim oSEnsayos   As New clsSellantes_ensayos
    Dim oSeTipo     As New clsSellantes
    Dim oTipo_Deter As New clsTipos_determinacion
    Dim oPNT        As New clsCa_documentos

    
    oSe.Carga MUESTRA_ID
    oSEnsayos.Carga Me.SELLANTE_ID, Me.ENSAYO
    oTipo_Deter.CargarTipoDeterminacion oSe.getTIPO_DETERMINACION_ID
    
    DOCUMENTO_ID = oTipo_Deter.getPNT_VINCULADO
    TIPO_DETERMINACION_ENSAYO_ID = oSe.getTIPO_DETERMINACION_ID

    oSeTipo.Carga (oSe.getSELLANTE_ID)

    lblCodTipoMuestraEnsayo.Caption = "Ensayo: " & oSEnsayos.getENSAYO & "  / Sellante: " & oSeTipo.getPROCESO
    lblCodTipoMuestraEnsayo.AutoSize = True
    
    If DOCUMENTO_ID = 0 Then

         lblCodPnt.Caption = "Norma: " & oSEnsayos.getNORMA_CRITERIO
         bloquear_controles
 
     ElseIf oPNT.Carga(DOCUMENTO_ID) Then
 
         lblCodPnt.Caption = "[" & oPNT.getCODIGO & "] " & oPNT.getNOMBRE
         cmdPNT.Enabled = True
 
     Else
 
         lblCodPnt.Caption = "N/A"
         bloquear_controles
 
    End If
    
    Set oSeTipo = Nothing
    Set oTipo_Deter = Nothing
    Set oSEnsayos = Nothing
    Set oSe = Nothing
    Set oPNT = Nothing

End Sub
'M0961-F

Private Sub bloquear_controles()
     cmdPNT.Enabled = False
     cmdok.Enabled = False
     cmbResultado1.desactivar
     cmbResultado2.desactivar
     txtObservaciones.Enabled = False
     chkEs_duplicado.Enabled = False
     chkEs_Recualificacion.Enabled = False
     cmbEnsayo.desactivar
     cmbobservador.desactivar
End Sub
'MANTIS-807-F



Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    
   On Error GoTo cmdok_Click_Error

    If Not comprobar_datos Then Exit Sub

    Dim oMo As New clsMuestras_observadores
    
'M0961-I
    If SELLANTE_ID = 0 Then
'M0961-F
        With oMo
            .Eliminar MUESTRA_ID, DETERMINACION_ENSAYO_ID
            
            .setMUESTRA_ID = MUESTRA_ID
            .setDETERMINACION_ENSAYO_ID = DETERMINACION_ENSAYO_ID
            .setTIPO_DETERMINACION_ENSAYO_ID = TIPO_DETERMINACION_ENSAYO_ID
'M0961-I
            .setSELLANTE_ID = SELLANTE_ID
            .setENSAYO = ENSAYO
'M0961-F
            If cmbPNT.visible = True Then
                .setDOCUMENTO_ID = cmbPNT.getPK_SALIDA
            Else
                .setDOCUMENTO_ID = DOCUMENTO_ID
            End If
            .setFECHA = Format(Date, "yyyy-mm-dd")
            .setES_CONTROL_EFICACIA = IIf(ES_CONTROL_EFICACIA, 1, 0)
            
            .setUSUARIO_ID_REALIZACION = cmbEnsayo.getPK_SALIDA
            .setUSUARIO_ID_OBSERVADOR = cmbobservador.getPK_SALIDA
    '        .setREALIZACION_ES_FORMADOR = IIf(opEnsayo(0).value = True, 1, 0)
            .setOBSERVADOR_ES_FORMADOR = IIf(opObservador(0).Value = True, 1, 0)
            
            .setUSUARIO_RESULTADO1 = cmbResultado1.getPK_SALIDA
            .setUSUARIO_RESULTADO2 = cmbResultado2.getPK_SALIDA
            
            .setES_DUPLICADO = IIf(chkEs_duplicado.Value = vbChecked, 1, 0)
            .setES_RECUALIFICACION = IIf(chkEs_Recualificacion.Value = vbChecked, 1, 0)
            .setOBSERVACIONES = txtObservaciones
            oMo.Insertar
        
         End With

'M0961-I
    Else
        With oMo
            .EliminarSellante MUESTRA_ID, SELLANTE_ID, ENSAYO
            
            .setMUESTRA_ID = MUESTRA_ID
            .setDETERMINACION_ENSAYO_ID = DETERMINACION_ENSAYO_ID
            .setTIPO_DETERMINACION_ENSAYO_ID = TIPO_DETERMINACION_ENSAYO_I
            .setSELLANTE_ID = SELLANTE_ID
            .setENSAYO = ENSAYO
            If cmbPNT.visible = True Then
                .setDOCUMENTO_ID = cmbPNT.getPK_SALIDA
            Else
                .setDOCUMENTO_ID = DOCUMENTO_ID
            End If
            .setFECHA = Format(Date, "yyyy-mm-dd")
            .setES_CONTROL_EFICACIA = IIf(ES_CONTROL_EFICACIA, 1, 0)
            
            .setUSUARIO_ID_REALIZACION = cmbEnsayo.getPK_SALIDA
            .setUSUARIO_ID_OBSERVADOR = cmbobservador.getPK_SALIDA
            .setOBSERVADOR_ES_FORMADOR = IIf(opObservador(0).Value = True, 1, 0)
            
            .setUSUARIO_RESULTADO1 = cmbResultado1.getPK_SALIDA
            .setUSUARIO_RESULTADO2 = cmbResultado2.getPK_SALIDA
            
            .setES_DUPLICADO = IIf(chkEs_duplicado.Value = vbChecked, 1, 0)
            .setES_RECUALIFICACION = IIf(chkEs_Recualificacion.Value = vbChecked, 1, 0)
            .setOBSERVACIONES = txtObservaciones
            oMo.Insertar
        End With
    End If
'M0961-F
    Set oMo = Nothing
    MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmObservadorEnsayo"

End Sub

Private Sub chkEsrecualificacion_Click()
'    If mvarblnCargando_Inicial Then Exit Sub
    
'    Dim strCad As String
'    Dim lngIdObservador As Long
'
'
'    If ACTOR_ES_FORMADOR = 0 Then
'        ACTOR_ES_FORMADOR = 1
'    Else
'        ACTOR_ES_FORMADOR = 0
'    End If
'
'    If cmdAvisoRecualificacion.Visible Then
'        chkEsrecualificacion.ForeColor = vbBlack
'        chkEsrecualificacion.Tag = ""
'        cmdAvisoRecualificacion.Visible = False
'        Exit Sub
'    End If
'
'
'    lngIdObservador = cmbUsuarios.getPK_SALIDA
'    cmbUsuarios.Limpiar
'
''    mvarblnCargando_Inicial = True
'
'    If ACTOR_ES_FORMADOR = 0 Then
'        mvarobjUsuariosCualificados.llenar_combo_usuarios_cualificados cmbUsuarios, USUARIO.getID_EMPLEADO, True
'    Else
'        llenar_combo cmbUsuarios, New clsUsuarios, 0, frmUsuarios, " AND ID_EMPLEADO <> " & CStr(USUARIO.getID_EMPLEADO)
'    End If
'    cmbUsuarios.cargar_datos
'    cmbUsuarios.MostrarElemento lngIdObservador
''    mvarblnCargando_Inicial = False
'
'    ' Cambia los roles
'    strCad = lblEnFormacion.Caption
'    lblEnFormacion.Caption = lblFormador.Caption
'    lblFormador.Caption = strCad

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    
    llenar_combo cmbEnsayo, New clsUsuarios, 0, frmUsuarios, " OR ANULADO <> 0 "
    llenar_combo cmbobservador, New clsUsuarios, 0, frmUsuarios, " OR ANULADO <> 0 "
    llenar_combo cmbResultado1, New clsUsuarios, 0, frmUsuarios, " OR ANULADO <> 0 "
    llenar_combo cmbResultado2, New clsUsuarios, 0, frmUsuarios, " OR ANULADO <> 0 "
    
    Dim oMo As New clsMuestras_observadores
    
    If MUESTRA_ID <> 0 Then
        Dim oM As New clsMuestra
        oM.CargaMuestra MUESTRA_ID
        If oM.getANALISIS_DUPLICADO = 1 Then
            frmDuplicado.visible = True
        Else
            frmDuplicado.visible = False
        End If
    End If
'M0961-I
'   If Not oMo.Carga(MUESTRA_ID, DETERMINACION_ENSAYO_ID) Then
    If SELLANTE_ID = 0 Then
       Carga = oMo.Carga(MUESTRA_ID, DETERMINACION_ENSAYO_ID)
    Else
       Carga = oMo.CargaSellante(MUESTRA_ID, SELLANTE_ID, ENSAYO)
    End If

    If Not Carga Then
'M0961-F
        If MUESTRA_CERRADA Then
            MsgBox "ATENCIÓN: La Muestra está CERRADA y no se tienen datos sobre su Observación", vbInformation, "Observador"
        End If
        
'        mvarblnActualizar = False
        
        If ES_CONTROL_EFICACIA Then
            presentar_datos_ce
        Else
            'MANTIS-807-I
            'presentar_datos_determinacion
            Select Case FORMULARIO_ORIGEN
            Case 2 'formulario de Sellantes
                presentar_datos_sellantes
            Case 5 ' Plasma
                presentar_datos_plasma
            Case Else 'otros
                presentar_datos_determinacion
            End Select
            
            'MANTIS-807-F
            
        End If
'        determinar_rol_usuarios
    Else
        presentar_datos oMo
    End If
'    opciones_edicion
    Set oMo = Nothing
'    mvarblnCargando_Inicial = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then cmdcancel_Click
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "Muestra", 2500, lvwColumnLeft
        .Add , , "Fecha", 2500, lvwColumnLeft
    End With
End Sub
