VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoEdicionMtoFechasEdicion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Mantenimiento de Equipos"
   ClientHeight    =   8130
   ClientLeft      =   2160
   ClientTop       =   2445
   ClientWidth     =   11970
   ClipControls    =   0   'False
   Icon            =   "frmEquipoEdicionMtoFechasEdicion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkRapidoTodos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conforme Rapido Todos Eq. Pte."
      Height          =   870
      Left            =   1110
      Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7230
      Width           =   1500
   End
   Begin VB.CommandButton cmdOkRapido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conforme Rapido"
      Height          =   870
      Left            =   30
      Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":12B4
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7230
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7230
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Mantenimiento"
      Height          =   5265
      Left            =   0
      TabIndex        =   14
      Top             =   1920
      Width           =   11925
      Begin VB.ListBox lstAcciones 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2280
         ItemData        =   "frmEquipoEdicionMtoFechasEdicion.frx":1B7E
         Left            =   1830
         List            =   "frmEquipoEdicionMtoFechasEdicion.frx":1B91
         Style           =   1  'Checkbox
         TabIndex        =   35
         Top             =   2010
         Width           =   6525
      End
      Begin VB.TextBox txtCertificado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   33
         Top             =   1635
         Width           =   6540
      End
      Begin VB.CommandButton cmdMostrarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   9150
         Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":1BD1
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ver norma"
         Top             =   1620
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8370
         Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":1E26
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Buscar documento"
         Top             =   1620
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8775
         Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":2097
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Eliminar documento"
         Top             =   1620
         Width           =   360
      End
      Begin VB.CommandButton cmdEscanearCert 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   8775
         Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":222B
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Escanear documento"
         Top             =   2295
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   1830
         MaxLength       =   1024
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   4320
         Width           =   10005
      End
      Begin MSComCtl2.DTPicker txtFechaActual 
         Height          =   300
         Left            =   1830
         TabIndex        =   20
         Top             =   1305
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   40197
      End
      Begin MSDataListLib.DataCombo cmbProcedimiento 
         Height          =   315
         Left            =   10290
         TabIndex        =   27
         Top             =   1650
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   315
         Left            =   1830
         TabIndex        =   28
         Top             =   930
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   556
      End
      Begin pryCombo.miCombo cmbProtocolo 
         Height          =   315
         Left            =   1830
         TabIndex        =   37
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   556
      End
      Begin VB.Frame fraEstadoIntervencion 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   8970
         TabIndex        =   23
         Top             =   2580
         Width           =   2355
         Begin VB.OptionButton optEstadMto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cerrado No Conforme"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton optEstadMto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Realizado"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   25
            Top             =   390
            Width           =   1605
         End
         Begin VB.OptionButton optEstadMto 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Previsto"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   24
            Top             =   60
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin pryCombo.miCombo cmbPlan 
         Height          =   315
         Left            =   1830
         TabIndex        =   42
         Top             =   270
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   556
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plan de Mantenimiento"
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   41
         Top             =   330
         Width           =   1620
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Señale las acciones realizadas"
         Height          =   390
         Index           =   6
         Left            =   150
         TabIndex        =   36
         Top             =   2070
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documento Adjunto"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   34
         Top             =   1695
         Width           =   1410
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   4590
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   17
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   17
         Left            =   150
         TabIndex        =   16
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Mantenimiento"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   1380
         Width           =   1605
      End
   End
   Begin VB.Frame frmDatosEquipo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Equipo"
      Height          =   1305
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   11925
      Begin VB.TextBox txtProtocolo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   540
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   210
         Width           =   4995
      End
      Begin VB.TextBox txtModelo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   570
         Width           =   4995
      End
      Begin VB.TextBox txtNSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   210
         Width           =   4695
      End
      Begin VB.TextBox txtFamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   930
         Width           =   4995
      End
      Begin VB.TextBox txtPlan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Plan de Mantenimiento AAAAAAAAAAAAAAAAAAAAAA"
         Top             =   540
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.TextBox txtPeriodicidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   7170
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   930
         Width           =   4695
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Protocolo"
         Height          =   195
         Index           =   9
         Left            =   6690
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modelo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   3
         Left            =   6330
         TabIndex        =   9
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   990
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   8
         Left            =   6270
         TabIndex        =   7
         Top             =   1020
         Width           =   870
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7230
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de Equipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   13
      Top             =   45
      Width           =   2640
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11340
      Picture         =   "frmEquipoEdicionMtoFechasEdicion.frx":25E5
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar Fechas para el Mantenimiento de Equipos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   315
      Width           =   3585
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   11955
   End
End
Attribute VB_Name = "frmEquipoEdicionMtoFechasEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarblnMostrarCierre As Boolean
Private mvarenuTipoEdicion As enumTipoEdicion
Private mvarobjMantenimiento As clsEquipoMantenimiento

Private mvarobjPlan As New clsPlanMantenimiento
Private mvarobjEquipo As clsEquipos

Private mvarcolMantenimientos  As clsGenericCollection
Private mvarlngID_MANTENIMIENTO As Long
Private mvarblnResultado As Boolean
Private mvarblnVieneDeCuaderno As Boolean
Private mvarlngIdEvento As Long
Private mvardtmFechaPrevista As Date
Private mvarlngidEquipo As Long

Private mvarblnOkRapido As Boolean
Public Property Get idEquipo() As Long

    idEquipo = mvarlngidEquipo

End Property

Public Property Let idEquipo(ByVal lngidEquipo As Long)

    mvarlngidEquipo = lngidEquipo

End Property


Private Sub PresentarDatos()
'Dim objItem As clsEquipos_planes_Acciones

    Set mvarobjPlan = New clsPlanMantenimiento
    
    With mvarobjEquipo
        txtNombre.Text = .getNOMBRE
        txtNSerie.Text = .getSERIE
        If Not .getFAMILIA Is Nothing Then
            txtFamilia.Text = .getFAMILIA.getNOMBRE
        End If
        txtModelo.Text = .getMODELO
    End With
    
    If mvarenuTipoEdicion = Alta Then
        Me.Caption = "Alta de Mantenimiento de Equipo"
        Set mvarobjMantenimiento = New clsEquipoMantenimiento
        Call mvarobjPlan.carga(mvarobjEquipo.getPLAN_MANTENIMIENTO_ID)
        ' muestra el plan por defecto para el equipo
        cmbPlan.MostrarElemento mvarobjEquipo.getPLAN_MANTENIMIENTO_ID
        cmbResponsable.MostrarElemento mvarobjEquipo.getMANTENEDOR_ID
        cmbProtocolo.MostrarElemento mvarobjPlan.getPROTOCOLO_ID
        txtFechaActual.Value = Now
    Else
        If mvarenuTipoEdicion = EDICION Then
            Me.Caption = "Edición de Mantenimiento de Equipo"
        Else
            Me.Caption = "Visualización de Mantenimiento de Equipo (Solo Lectura)"
        End If
        
        Set mvarobjMantenimiento = New clsEquipoMantenimiento
        'Call mvarobjPlan.Carga(mvarobjEquipo.getPLAN_MANTENIMIENTO_ID)
        mvarobjMantenimiento.carga mvarlngID_MANTENIMIENTO
        Call mvarobjPlan.carga(mvarobjMantenimiento.getPLANMTO_ID)
        cmbPlan.MostrarElemento mvarobjMantenimiento.getPLANMTO_ID
        txtObservaciones.Text = mvarobjMantenimiento.getOBSERVACIONES
        
        cmbProtocolo.MostrarElemento mvarobjMantenimiento.getPROCEDIMIENTO_ID
        cmbResponsable.MostrarElemento mvarobjMantenimiento.getMANTENEDOR_ID
        optEstadMto(mvarobjMantenimiento.getESTADO).Value = True
        
        If IsDate(mvarobjMantenimiento.getFECHA_ACTUAL) Then
            txtFechaActual.Value = CDate(mvarobjMantenimiento.getFECHA_ACTUAL)
        Else
            txtFechaActual.Value = Now
        End If
        txtCertificado = mvarobjMantenimiento.getRUTA_CERTIFICADO
            
    End If
    
    'txtPlan.Text = mvarobjPlan.getNOMBRE
    txtperiodicidad.Text = mvarobjPlan.getFRECUENCIA
'    If Not mvarobjMantenimiento.CERTIFICADO Is Nothing Then
'        If Trim(mvarobjMantenimiento.CERTIFICADO.getRUTA_TEMPORAL) <> "" Then
'            txtCertificado.Text = mvarobjMantenimiento.CERTIFICADO.getNOMBRE_ARCHIVO_TEMP
'        Else
'            txtCertificado.Text = mvarobjMantenimiento.CERTIFICADO.getNOMBRE_ARCHIVO
'        End If
'    End If
    
    lstAcciones.Clear
    
    Call PresentarDatos_Acciones
    
    
    
        
End Sub

Private Sub RecogerDatos()

    With mvarobjMantenimiento
        .setPLANMTO_ID = cmbPlan.getPK_SALIDA
        .setPLAN_MANTENIMIENTO = cmbPlan.getTEXTO
        
        .setFECHA_ACTUAL = Format(txtFechaActual.Value)
        If mvarblnOkRapido Then
            .setESTADO = 1
        Else
            .setESTADO = IIf(optEstadMto(0).Value, 0, 1)
        End If
        .setEQUIPO_ID = mvarobjEquipo.getID_EQUIPO
        
        .setPROCEDIMIENTO_ID = cmbProtocolo.getPK_SALIDA
        
        .setPROCEDIMIENTO = cmbProtocolo.getTEXTO
        .setMANTENEDOR_ID = cmbResponsable.getPK_SALIDA
        .setRESPONSABLE = cmbResponsable.getTEXTO
        .setOBSERVACIONES = txtObservaciones.Text
        If mvarblnOkRapido Then
            .setACCIONES = recoger_acciones_todas()
        Else
            .setACCIONES = recoger_acciones()
        End If
        .setRUTA_CERTIFICADO = txtCertificado
    End With

End Sub

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Get id_mantenimiento() As Long

    id_mantenimiento = mvarlngID_MANTENIMIENTO

End Property

Public Property Let id_mantenimiento(ByVal lngID_MANTENIMIENTO As Long)

    mvarlngID_MANTENIMIENTO = lngID_MANTENIMIENTO

End Property


Private Sub cmbPlan_change()
Set mvarobjPlan = New clsPlanMantenimiento

mvarobjPlan.carga cmbPlan.getPK_SALIDA
Call PresentarDatos_Acciones
End Sub

Private Sub cmdAdjuntarCertificado_Click()
On Error GoTo cmdAdjuntarCertificado_Click_Error
    
    If mvarenuTipoEdicion = Alta Then
        MsgBox "Guarde primero el Mantenimiento para poder asignar documentos.", vbCritical, App.Title
        Exit Sub
    End If
    
    cd.ShowOpen
    
    If Trim(cd.FileName) = "" Then Exit Sub
    
'    mvarobjMantenimiento.CERTIFICADO.setRUTA_TEMPORAL = cd.FileName
'    mvarobjMantenimiento.CERTIFICADO.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtCertificado.Text = cd.FileTitle
    
    Dim oD As New clsDocumentacion
    Dim salida As String
    salida = oD.SubirEquipo(mvarlngidEquipo, 2, CLng(mvarlngID_MANTENIMIENTO), 0, cd.FileName, cd.FileTitle)
    If salida <> "" Then
        MsgBox "Se ha producido un error al subir el documento : " & salida, vbCritical, App.Title
    Else
        txtCertificado.Text = cd.FileTitle
    End If
    
    

On Error GoTo 0
    Exit Sub
cmdAdjuntarCertificado_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarCertificado_Click of Formulario frmEquipoEdicionMtoFechasEdicion"
End Sub

Private Sub cmdEliminarCertificado_Click()
    Dim oD As New clsDocumentacion
    If oD.EliminarEquipo(mvarobjEquipo.getID_EQUIPO, 2, CLng(mvarobjMantenimiento.getID_MANTENIMIENTO), 0) = "" Then
       txtCertificado.Text = ""
    End If
    Set oD = Nothing

'txtCertificado.Text = ""
'mvarobjMantenimiento.CERTIFICADO.setID_AUX = enumIdAux.ID_AUX_ELIMINADO

End Sub

'Private Sub cmdEscanearCert_Click()
'Dim strArchivo As String
'
'    strArchivo = EscanearATemp
'
'    If Trim(strArchivo) = "" Then Exit Sub
'
'    mvarobjMantenimiento.CERTIFICADO.setRUTA_TEMPORAL = strArchivo
'    mvarobjMantenimiento.CERTIFICADO.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtCertificado.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
'
'End Sub


Private Sub cmdMostrarCertificado_Click()
    
    Dim oD As New clsDocumentacion
   On Error GoTo cmdMostrarCertificado_Click_Error

    oD.CargarEquipo mvarobjEquipo.getID_EQUIPO, 2, CLng(mvarobjMantenimiento.getID_MANTENIMIENTO), 0, True
    Set oD = Nothing
    
'    Dim objAI As New clsArchivoAdjunto
'    Dim destino As String, r As Double
'
'    Set objAI = mvarobjMantenimiento.CERTIFICADO
'
'    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
'        destino = objAI.getRUTA_TEMPORAL
'    ElseIf (objAI.getRUTA <> "") Then
'        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\MTO\" & mvarobjMantenimiento.getID_MANTENIMIENTO & "\CERT\" & objAI.getNOMBRE_ARCHIVO
'    End If
'
'    On Error GoTo fallo
'
'    If Dir(destino, vbArchive) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If

   On Error GoTo 0
   Exit Sub

cmdMostrarCertificado_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrarCertificado_Click of Formulario frmEquipoEdicionMtoFechasEdicion"

End Sub

Private Sub cmdOkRapido_Click()
    mvarblnOkRapido = True

    If Not ComprobarDatos Then Exit Sub

    Call RecogerDatos
    
    mvarblnOkRapido = False
    
    'If Not mvarblnVieneDeCuaderno Then
        ' no viene de cuaderno
        If mvarenuTipoEdicion = Alta Then
            mvarobjMantenimiento.setEQUIPO_ID = mvarlngidEquipo
            mvarobjMantenimiento.Insertar
            'mvarobjEquipo.Mantenimientos.Add mvarobjMantenimiento
        Else
            'Call mvarobjEquipo.Mantenimientos.Replace(CStr(mvarobjMantenimiento.getID_MANTENIMIENTO), mvarobjMantenimiento)
            Call mvarobjMantenimiento.Modificar(mvarlngID_MANTENIMIENTO)
        End If
    'Else
        ' actualiza el dato directamente en la base de datos
        'Call mvarobjMantenimiento.Modificar(mvarlngIdEvento)
    'End If
'M1051    If Not mvarblnVieneDeCuaderno Then
'M1051        Call mvarobjEquipo.Carga_Mantenimiento
'M1051    End If
    
    mvarblnResultado = True
    Me.Hide

End Sub

Private Sub cmdOkRapidoTodos_Click()
    mvarblnOkRapido = True

    If Not ComprobarDatos Then Exit Sub

    RecogerDatos
    
    mvarblnOkRapido = False
    
    If Not mvarobjEquipo.ok_rapido_equipos_pendientes_diario(mvarobjEquipo.getID_EQUIPO, txtFechaActual.Value, cmbResponsable.getPK_SALIDA) Then
        MsgBox "No se han encontrado Mantenimientos Previstos para este equipo en esta fecha o anteriores.", vbInformation, "Realizado Rápido"
    End If
    mvarblnResultado = True
    Me.Hide

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjMantenimiento = Nothing

    Set mvarobjEquipo = Nothing
End Sub


Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()

    If Not ComprobarDatos Then Exit Sub

    Call RecogerDatos
    
    If mvarenuTipoEdicion = Alta Then
        mvarobjMantenimiento.setEQUIPO_ID = mvarlngidEquipo
        mvarobjMantenimiento.Insertar
'JGM        If mvarobjMantenimiento.getESTADO = 0 Then
'JGM            ' solo revisa las operaciones pendientes si queda en previsto.
'JGM            ' porque al cerrarlo, ya no es previsto, entonces no sirve revisar nada
'JGM            mvarobjMantenimiento.revisar_mantenimiento_pendiente
'JGM        End If
    Else
        Call mvarobjMantenimiento.Modificar(mvarlngID_MANTENIMIENTO)
        
    End If
    
    mvarblnResultado = True
    Me.Hide

End Sub

Private Sub Form_Load()

log Me.Name
cargar_botones Me


'Call cargar_combos

If mvarblnVieneDeCuaderno Then
    Set mvarobjEquipo = New clsEquipos
    mvarobjEquipo.carga mvarlngidEquipo
    mvarenuTipoEdicion = EDICION
    mvarlngID_MANTENIMIENTO = mvarlngIdEvento
End If

Call cargar_combos

cmdOkRapido.visible = mvarblnVieneDeCuaderno
cmdOkRapidoTodos.visible = mvarblnVieneDeCuaderno
mvarblnOkRapido = mvarblnVieneDeCuaderno

mvarlngidEquipo = mvarobjEquipo.getID_EQUIPO

Call OpcionesEdicion

Call PresentarDatos

fraEstadoIntervencion.visible = mvarblnMostrarCierre Or mvarblnVieneDeCuaderno

End Sub


Private Sub cargar_combos()
'Dim oCA_Doc As New clsCa_documentos
        
    llenar_combo cmbResponsable, New clsUsuarios, 0, Me, ""
    llenar_combo cmbProtocolo, New clsCa_documentos, 0, frmCA_Documento, ""
    llenar_combo cmbPlan, New clsPlanMantenimiento, 0, frmEquipoPlanMtoEdicion, " id_plan_mto in (select plan_mantenimiento_id from eq_planes_mantenimiento_equipos where equipo_id = " & CStr(mvarobjEquipo.getID_EQUIPO) & ") or id_plan_mto = " & CStr(mvarobjEquipo.getPLAN_MANTENIMIENTO_ID)
    
    
End Sub

Public Property Get MostrarCierre() As Boolean

    MostrarCierre = mvarblnMostrarCierre

End Property

Public Property Let MostrarCierre(ByVal blnMostrarCierre As Boolean)

    mvarblnMostrarCierre = blnMostrarCierre

End Property

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

Public Property Get Mantenimiento() As clsEquipoMantenimiento

    Set Mantenimiento = mvarobjMantenimiento

End Property

Public Property Set Mantenimiento(objMantenimiento As clsEquipoMantenimiento)

    Set mvarobjMantenimiento = objMantenimiento

End Property

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Private Sub OpcionesEdicion()

    If mvarenuTipoEdicion = visualizar Then
        lstAcciones.Enabled = False
        
        If USUARIO.getID_EMPLEADO <> 69 And USUARIO.getID_EMPLEADO <> 31 Then
            cmdok.Enabled = False
            txtFechaActual.Enabled = False
        End If
        'cmbProcedimiento.Enabled = False
        cmbProtocolo.desactivar
        cmbResponsable.desactivar
        cmbPlan.desactivar
        txtObservaciones.Locked = True
        txtCertificado.Locked = True
            cmdMostrarCertificado.Left = cmdAdjuntarCertificado.Left
            cmdAdjuntarCertificado.visible = False
            cmdEscanearCert.visible = False
            cmdEliminarCertificado.visible = False
        fraEstadoIntervencion.Enabled = False
        cmdOkRapido.visible = False
        cmdOkRapidoTodos.visible = False
    End If

End Sub

Private Function ComprobarDatos() As Boolean

Dim strCad As String
strCad = ""

    ComprobarDatos = False

    'If getDataComboSel(cmbProcedimiento) <= 0 Then
    '    strCad = strCad & vbCrLf & " - Debe señalar el Procedimiento para el Mantenimiento de Equipo."
    '    'MsgBox "Debe señalar el Procedimiento para el Mantenimiento de Equipo", vbInformation, "Editar Mantenimiento de Equipo"
    'End If
    
    If cmbPlan.getPK_SALIDA = 0 Then
        strCad = strCad & vbCrLf & " - Debe señalar a que plan pertenece el Mantenimiento de Equipo."
        'MsgBox "Debe señalar el Procedimiento para el Mantenimiento de Equipo", vbInformation, "Editar Mantenimiento de Equipo"
    End If
    
    If cmbProtocolo.getPK_SALIDA = 0 Then
        strCad = strCad & vbCrLf & " - Debe señalar el Procedimiento (Protocolo) para el Mantenimiento de Equipo."
        'MsgBox "Debe señalar el Procedimiento para el Mantenimiento de Equipo", vbInformation, "Editar Mantenimiento de Equipo"
    End If
    
    If txtFechaActual.Value = CDate("01/01/1900") Then
        strCad = strCad & vbCrLf & " - Debe señalar una fecha válida para el Mantenimiento de Equipo."
        'MsgBox "Debe señalar una fecha válida para el Mantenimiento de Equipo", vbInformation, "Editar Mantenimiento de Equipo"
    ElseIf txtFechaActual.Value < CDate(Format(Now, "dd/mm/yyyy")) And optEstadMto(0).Value = False Then
        'M00496-I
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
            strCad = strCad & vbCrLf & " - Debe señalar una fecha válida posterior a o igual a hoy."
        End If
        'M00496-F
    Else
        ' cuando la fecha está correcta, comprueba que no se ponga una fecha anterior a la fecha de próximo mantenimiento
        If Not mvarblnOkRapido Then
            Dim trs As ADODB.Recordset
            Set trs = mvarobjEquipo.devolver_fecha_prox_mantenimiento
            On Error Resume Next
            trs.MoveFirst
            
            If Err.Number = 0 Then
'V220513-I
'               If trs!id_mantenimiento <> 0 And trs!id_mantenimiento <> mvarlngID_MANTENIMIENTO And CLng(trs("PLANMTO_ID")) = cmbPlan.getPK_SALIDA Then
                If optEstadMto(1).Value = True And trs!id_mantenimiento <> 0 And trs!id_mantenimiento <> mvarlngID_MANTENIMIENTO And CLng(trs("PLANMTO_ID")) = cmbPlan.getPK_SALIDA Then
'V220513-F
'                    If CDate(trs!fecha_actual) >= (txtFechaActual.value) Then
                    If CDate(trs!fecha_actual) <= (txtFechaActual.Value) Then
                        strCad = strCad & vbCrLf & " - Existe un mantenimiento previsto para la fecha " & Format(trs!fecha_actual, "dd/mm/yyyy") & ". La fecha del mantenimiento debe ser posterior"
                    End If
                End If
            End If
            
        End If
    End If
    
    If cmbResponsable.getPK_SALIDA <= 0 Then
        strCad = strCad & vbCrLf & " - Debe señalar un Responsable para el Mantenimiento de Equipo."
        'MsgBox "Debe señalar un Responsable para el Mantenimiento de Equipo", vbInformation, "Editar Mantenimiento de Equipo"
    End If

    If Trim(strCad) <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strCad, vbInformation, "Editar Mantenimiento de Equipo"
        Exit Function
    End If
    
    ComprobarDatos = True


End Function


Public Property Get VieneDeCuaderno() As Boolean

    VieneDeCuaderno = mvarblnVieneDeCuaderno

End Property

Public Property Let VieneDeCuaderno(ByVal blnVieneDeCuaderno As Boolean)

    mvarblnVieneDeCuaderno = blnVieneDeCuaderno

End Property

Public Property Get IdEvento() As Long

    IdEvento = mvarlngIdEvento

End Property

Public Property Let IdEvento(ByVal lngIdEvento As Long)

    mvarlngIdEvento = lngIdEvento

End Property

Public Property Get FechaPrevista() As Date

    FechaPrevista = mvardtmFechaPrevista

End Property

Public Property Let FechaPrevista(ByVal dtmFechaPrevista As Date)

    mvardtmFechaPrevista = dtmFechaPrevista

End Property

Private Sub PresentarDatos_Acciones()
Dim arrAcciones() As String
Dim x As Long, total As Long
Dim rs As ADODB.Recordset
Dim blnEncontrado As Boolean

If mvarobjMantenimiento Is Nothing Then Exit Sub

If Trim(mvarobjMantenimiento.getACCIONES) <> "" Then
    
    arrAcciones = Split(mvarobjMantenimiento.getACCIONES, ";")
    total = UBound(arrAcciones)
Else
    total = -1
End If

If mvarlngID_MANTENIMIENTO <> 0 Then
    Set rs = mvarobjMantenimiento.cargar_acciones_realizadas(mvarlngID_MANTENIMIENTO, False)
    
    If Not rs Is Nothing Then
        
        lstAcciones.Clear
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            While Not rs.EOF
                lstAcciones.AddItem rs("NOMBRE")
                lstAcciones.ItemData(lstAcciones.ListCount - 1) = rs("ID_ACCION")
                blnEncontrado = False
                    If total >= 0 Then
                        For x = 0 To total
                            If arrAcciones(x) = rs("ID_ACCION") Then
                                blnEncontrado = True
                                Exit For
                            End If
                        Next x
                    End If
                    lstAcciones.Selected(lstAcciones.ListCount - 1) = blnEncontrado
                rs.MoveNext
            Wend
        End If
    End If
Else
    Set rs = mvarobjPlan.devolver_acciones(mvarobjPlan.getID_PLAN_MTO)
    
    If Not rs Is Nothing Then
        lstAcciones.Clear
        
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
                
            While Not rs.EOF
                lstAcciones.AddItem rs("NOMBRE")
                lstAcciones.ItemData(lstAcciones.ListCount - 1) = rs("ID_ACCION")
                blnEncontrado = False
                    'If total >= 0 Then
                    '    For x = 0 To total
                    '        If arrAcciones(x) = rs("ID_ACCION") Then
                    '            blnEncontrado = True
                    '            Exit For
                    '        End If
                    '    Next x
                    'End If
                    'lstAcciones.Selected(lstAcciones.ListCount - 1) = blnEncontrado
                rs.MoveNext
            Wend
        End If
    End If
End If
End Sub

Private Function recoger_acciones() As String

    Dim x As Integer
    Dim acc As String
    
    acc = ""
    
    For x = 0 To lstAcciones.ListCount - 1
        If lstAcciones.Selected(x) Then
            acc = acc & lstAcciones.ItemData(x) & ";"
        End If
    Next x
    
    If Trim(acc) <> "" Then
        acc = Left(acc, Len(acc) - 1)
    End If
    
    recoger_acciones = acc

End Function

Private Function recoger_acciones_todas() As String

    Dim x As Integer
    Dim acc As String
    
    acc = ""
    
    For x = 0 To lstAcciones.ListCount - 1
        'If lstAcciones.Selected(x) Then
            acc = acc & lstAcciones.ItemData(x) & ";"
        'End If
    Next x
    
    If Trim(acc) <> "" Then
        acc = Left(acc, Len(acc) - 1)
    End If
    
    recoger_acciones_todas = acc

End Function


