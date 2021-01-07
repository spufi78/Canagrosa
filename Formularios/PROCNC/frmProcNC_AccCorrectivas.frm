VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcNC_AccCorrectivas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acción Correctiva/Preventiva"
   ClientHeight    =   7260
   ClientLeft      =   6945
   ClientTop       =   2040
   ClientWidth     =   8400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   37
      Top             =   675
      Width           =   8250
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ACCION PREVENTIVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4095
         TabIndex        =   39
         Top             =   315
         Width           =   2670
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ACCION CORRECTIVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   38
         Top             =   315
         Width           =   2670
      End
   End
   Begin VB.CommandButton cmdEnviarAvisoFinalizada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informar Fin Acción"
      Height          =   870
      Left            =   3930
      Picture         =   "frmProcNC_AccCorrectivas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6360
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdCerrarYGenerarPNC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar Y Generar P.N.C."
      Height          =   870
      Left            =   1980
      Picture         =   "frmProcNC_AccCorrectivas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6360
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdCerrarExito 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar Acc. Correctiva"
      Height          =   870
      Left            =   30
      Picture         =   "frmProcNC_AccCorrectivas.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6360
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdTramitar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitar Acc. Correctiva"
      Height          =   870
      Left            =   30
      Picture         =   "frmProcNC_AccCorrectivas.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6360
      Width           =   1905
   End
   Begin TabDlg.SSTab tabAccCorrectiva 
      Height          =   4755
      Left            =   0
      TabIndex        =   4
      Top             =   1485
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8387
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Datos Accion "
      TabPicture(0)   =   "frmProcNC_AccCorrectivas.frx":2328
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabel(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabel(9)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPlazoDefinido"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDescripcion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFechaPuestaEnMarcha"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbResponsableImplantacion"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTitulo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Resolución Accion Correctiva"
      TabPicture(1)   =   "frmProcNC_AccCorrectivas.frx":2344
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabel(8)"
      Tab(1).Control(1)=   "txtFechaResolucion"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(5)=   "Frame4"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtTitulo 
         Height          =   315
         Left            =   60
         MaxLength       =   150
         TabIndex        =   31
         Top             =   600
         Width           =   8145
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -74880
         TabIndex        =   25
         Top             =   1770
         Width           =   8205
         Begin VB.OptionButton optEvidencias_si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   27
            Top             =   270
            Width           =   705
         End
         Begin VB.OptionButton optEvidencias_no 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   26
            Top             =   270
            Width           =   645
         End
         Begin VB.Label lblLabel 
            Caption         =   $"frmProcNC_AccCorrectivas.frx":2360
            Height          =   555
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   90
            Width           =   5415
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   -74880
         TabIndex        =   21
         Top             =   1410
         Width           =   8205
         Begin VB.OptionButton optComunicado_a_departamentos_no 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   23
            Top             =   90
            Width           =   645
         End
         Begin VB.OptionButton optComunicado_a_departamentos_si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   22
            Top             =   90
            Width           =   705
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "¿Se ha comunicado la modificación a todos los departamentos?"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   90
            Width           =   4515
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   -74880
         TabIndex        =   17
         Top             =   930
         Width           =   8205
         Begin VB.OptionButton optEfectiva_si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   19
            Top             =   120
            Width           =   705
         End
         Begin VB.OptionButton optEfectiva_no 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   18
            Top             =   120
            Width           =   645
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "¿Son Efectivas?"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   1170
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   -74880
         TabIndex        =   13
         Top             =   450
         Width           =   8205
         Begin VB.OptionButton optAccionCorrectivaEnPlazo_No 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   16
            Top             =   120
            Width           =   645
         End
         Begin VB.OptionButton optAccionCorrectivaEnPlazo_Si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   15
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "¿La Acción Correctiva ha sido puesta en marcha en plazo?"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   4185
         End
      End
      Begin VB.ComboBox cmbResponsableImplantacion 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3465
         Width           =   5565
      End
      Begin MSComCtl2.DTPicker txtFechaPuestaEnMarcha 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52101121
         CurrentDate     =   40156
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   2220
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1200
         Width           =   8145
      End
      Begin MSComCtl2.DTPicker txtPlazoDefinido 
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   4230
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52101121
         CurrentDate     =   40156
      End
      Begin MSComCtl2.DTPicker txtFechaResolucion 
         Height          =   315
         Left            =   -72450
         TabIndex        =   29
         Top             =   3060
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   52101121
         CurrentDate     =   40156
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Título de la Acción"
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   32
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Resolución Incidencia"
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   30
         Top             =   3090
         Width           =   2070
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Responsable de la Implantación"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   3480
         Width           =   2265
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Plazo Definido para la Resolución"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   4260
         Width           =   2385
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Puesta En Marcha"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   3870
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Descripción de la Acción "
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acción Correctiva/Preventiva"
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
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   3030
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Describa los datos de la Acción Correctiva, rellenando los siguientes campos."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   345
      Width           =   5460
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   7830
      Picture         =   "frmProcNC_AccCorrectivas.frx":2408
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   10995
   End
End
Attribute VB_Name = "frmProcNC_AccCorrectivas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarblnResultado As Boolean
Private mvarobjAccionCorrectiva As clsProcNcAccionCorrectora
Private mvarenumTipoEdicion As enumTipoEdicion

Private mvarobjPnc As clsProcNc

Private Sub OpcionesEdicion()
On Error GoTo OpcionesEdicion_Error
    
If mvarenumTipoEdicion = visualizar Then
   'txtTitulo.Enabled = False
   'txtDescripcion.Enabled = False
   txtTitulo.Locked = True
   txtDescripcion.Locked = True
   cmdTramitar.Enabled = False
   cmdCerrarExito.Enabled = False
   cmdCerrarYGenerarPNC.Enabled = False
   cmbResponsableImplantacion.Enabled = False
   txtFechaPuestaEnMarcha.Enabled = False
   txtPlazoDefinido.Enabled = False
   txtFechaResolucion.Enabled = False
   
   optAccionCorrectivaEnPlazo_No.Enabled = False
   optAccionCorrectivaEnPlazo_Si.Enabled = False
   optComunicado_a_departamentos_no.Enabled = False
   optComunicado_a_departamentos_si.Enabled = False
   optEfectiva_no.Enabled = False
   optEfectiva_si.Enabled = False
   optEvidencias_no.Enabled = False
   optEvidencias_si.Enabled = False
   
    cmdok.Enabled = False
End If
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.OpcionesEdicion"
    Exit Sub
OpcionesEdicion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.OpcionesEdicion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure OpcionesEdicion of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Public Property Let PK(ByVal dato As Long)
On Error GoTo PK_Error
    
Set mvarobjAccionCorrectiva = New clsProcNcAccionCorrectora
Set mvarobjPnc = New clsProcNc



Call mvarobjAccionCorrectiva.Carga(dato)
Call mvarobjPnc.Carga(mvarobjAccionCorrectiva.getPROCNC_ID)

If USUARIO.getID_EMPLEADO = mvarobjAccionCorrectiva.getRESPONSABLE_ID And mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
    cmdEnviarAvisoFinalizada.Visible = True
End If

If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.Calidad) = vbChecked Then
    mvarenumTipoEdicion = enumTipoEdicion.EDICION
Else
    mvarenumTipoEdicion = enumTipoEdicion.visualizar
    cmdok.Visible = False
    cmdCerrarExito.Visible = False
    cmdCerrarYGenerarPNC.Visible = False
    cmdTramitar.Visible = False
    
    cmbResponsableImplantacion.Enabled = False
    txtFechaPuestaEnMarcha.Enabled = False
    txtFechaResolucion.Enabled = False
    txtPlazoDefinido.Enabled = False
    txtTitulo.Locked = True
    txtDescripcion.Locked = True
End If

Call mvarobjAccionCorrectiva.Carga(dato)
Call mvarobjPnc.Carga(mvarobjAccionCorrectiva.getPROCNC_ID)

Call Form_Load
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.PK"
    Exit Property
PK_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.PK"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PK of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Property


Public Property Set Pnc(ByRef obj As clsProcNc)
On Error GoTo Pnc_Error
    
Set mvarobjPnc = obj
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.Pnc"
    Exit Property
Pnc_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.Pnc"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Pnc of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Property

Private Sub PresentarDatos()
Dim objUsuarios As New clsUsuarios
    
On Error GoTo PresentarDatos_Error
    
With mvarobjAccionCorrectiva
    opTipo(.getTIPO_ID).value = True

    txtTitulo.Text = .getTITULO
    txtDescripcion.Text = .getDESCRIPCION
    
    txtFechaPuestaEnMarcha.value = .getFECHA_PUESTA_EN_MARCHA
    
    txtPlazoDefinido.value = .getFECHA_PREVISTA
    For x = 0 To cmbResponsableImplantacion.ListCount - 1
        If cmbResponsableImplantacion.ItemData(x) = .getRESPONSABLE_ID Then
            cmbResponsableImplantacion.ListIndex = x
            Exit For
        End If
    Next x
    
    
    txtFechaResolucion.value = .getFECHA_RESOLUCION
    
    If .getRESOLUCION_EFECTIVA > -1 Then
        optEfectiva_si.value = (.getRESOLUCION_EFECTIVA = 1)
    End If
    
    If .getRESOLUCION_COMUNICADO_MODIFICACIONES > -1 Then
        optComunicado_a_departamentos_si.value = (.getRESOLUCION_COMUNICADO_MODIFICACIONES = 1)
    End If
    
    If .getRESOLUCION_EVIDENCIAS > -1 Then
        optEvidencias_si.value = (.getRESOLUCION_EVIDENCIAS = 1)
    End If
    
    If .getRESOLUCION_PUESTA_MARCHA_EN_PLAZO > -1 Then
        optAccionCorrectivaEnPlazo_Si.value = (.getRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 1)
    End If


End With
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.PresentarDatos"
    Exit Sub
PresentarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.PresentarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Public Property Get TipoEdicion() As enumTipoEdicion

   On Error GoTo TipoEdicion_Error

    TipoEdicion = mvarenumTipoEdicion

   On Error GoTo 0
   Exit Property

TipoEdicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TipoEdicion of Formulario frmProcNC_AccCorrectivas"

End Property

Public Property Let TipoEdicion(ByVal enumTipoEdicion As enumTipoEdicion)

   On Error GoTo TipoEdicion_Error

    mvarenumTipoEdicion = enumTipoEdicion

   On Error GoTo 0
   Exit Property

TipoEdicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TipoEdicion of Formulario frmProcNC_AccCorrectivas"

End Property


Public Property Get AccionCorrectiva() As clsProcNcAccionCorrectora

    Set AccionCorrectiva = mvarobjAccionCorrectiva

End Property

Public Property Set AccionCorrectiva(ByRef VALOR As clsProcNcAccionCorrectora)

    Set mvarobjAccionCorrectiva = VALOR

End Property

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub cmdCerrarExito_Click()
Dim strFecha As String, strTitulo As String

On Error GoTo cmdCerrarExito_Click_Error
    
strFecha = Format(txtFechaResolucion.value, "dd/mm/yyyy")
strTitulo = txtTitulo.Text

If Not ComprobarDatos("Cerrar Acción Correctiva", True) Then Exit Sub

If MsgBox("Ha decidido cerrar la Acción Correctora " & strTitulo & " en la siguiente Fecha: " & strFecha, vbInformation + vbYesNo, "Cerrar Accion Correctiva") = vbNo Then Exit Sub

mvarobjAccionCorrectiva.setESTADO_ID = C_PROCNC_ESTADOS.CERRADA
mvarobjAccionCorrectiva.setESTADO = "P.N.C. Cerrada"

Call GuardarDatos

mvarobjAccionCorrectiva.EnviarMensaje_AccCerradaOk


mvarblnResultado = True
Me.Hide

    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdCerrarExito_Click"
    Exit Sub
cmdCerrarExito_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdCerrarExito_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdCerrarExito_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdEnviarAvisoFinalizada_Click()


    Dim strMensaje As String
    
    
On Error GoTo cmdEnviarAvisoFinalizada_Click_Error
    
    strMensaje = InputBox("Para dar por finalizada la Accion Correctiva, escriba un mensaje si desea indicar alguna observacion de los responsables de calidad.", "Mensaje de Finalizacion de Tarea")
    
    If Trim(strMensaje) = "" Then
        If MsgBox("No indica observaciones. ¿Desea enviar un aviso de finalización de tareas de todas formas?", vbInformation + vbYesNo, "Informar Acción Finalizada") = vbNo Then _
            Exit Sub
    End If

    mvarobjAccionCorrectiva.EnviarMensaje_AccFinalizadaPorResponsable strMensaje
        
    Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdEnviarAvisoFinalizada_Click"
    Exit Sub
cmdEnviarAvisoFinalizada_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdEnviarAvisoFinalizada_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdEnviarAvisoFinalizada_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdOk_Click()

On Error GoTo cmdok_Click_Error
    
If Not ComprobarDatos() Then Exit Sub

Call GuardarDatos


mvarblnResultado = True
Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdok_Click"
    Exit Sub
cmdok_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdok_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdsalir_Click()

On Error GoTo cmdSalir_Click_Error
    
mvarblnResultado = False

Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdSalir_Click"
    Exit Sub
cmdSalir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdSalir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdSalir_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub cmdTramitar_Click()
Dim strRESPONSABLE As String

On Error GoTo cmdTramitar_Click_Error
    
If Not ComprobarDatos("Tramitar Acción Correctiva") Then Exit Sub

strRESPONSABLE = cmbResponsableImplantacion.List(cmbResponsableImplantacion.ListIndex)

If MsgBox("Va a Tramitar la Acción Correctiva. El Responsable de su implantación será: " & strRESPONSABLE & vbCrLf & " ¿Desea Continuar?", vbYesNo, "Tramitar Accion Correctiva") = vbNo Then Exit Sub

mvarobjAccionCorrectiva.setESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION
mvarobjAccionCorrectiva.setESTADO = "En Tramitación"

Call GuardarDatos

mvarobjAccionCorrectiva.EnviarMensaje_AccTramitada

mvarblnResultado = True
Me.Hide

    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdTramitar_Click"
    Exit Sub
cmdTramitar_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdTramitar_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdTramitar_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdCerrarYGenerarPNC_Click()

Dim strFecha As String, strTitulo As String

On Error GoTo cmdCerrarYGenerarPNC_Click_Error
    
strFecha = Format(txtFechaResolucion.value, "dd/mm/yyyy")
strTitulo = txtTitulo.Text

If Not ComprobarDatos("Cerrar Acc. Correctiva Y Generar Incidencia", True) Then Exit Sub

If MsgBox("Ha decidido cerrar la Acción Correctora " & strTitulo & " y generar un nueva Incidencia en la siguiente Fecha: " & strFecha, vbInformation + vbYesNo, "Cerrar Accion Correctiva") = vbNo Then Exit Sub

mvarobjAccionCorrectiva.setESTADO_ID = C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL
mvarobjAccionCorrectiva.setESTADO = "Cerrada (Genera Nueva Incidencia)"

Call GuardarDatos

mvarblnResultado = True
Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdCerrarYGenerarPNC_Click"
    Exit Sub
cmdCerrarYGenerarPNC_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdCerrarYGenerarPNC_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdCerrarYGenerarPNC_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub Form_Load()
    
Dim objUsuarios As New clsUsuarios
On Error GoTo Form_Load_Error
    
    log (Me.Name)
    cargar_botones Me
    
    Cargar_ComboBox cmbResponsableImplantacion, objUsuarios
    
    If mvarenumTipoEdicion = Alta Then
        
        Set mvarobjAccionCorrectiva = New clsProcNcAccionCorrectora
        mvarobjAccionCorrectiva.setID_AUX = -2
        mvarobjAccionCorrectiva.setESTADO_ID = C_PROCNC_ESTADOS.ABIERTA
        mvarobjAccionCorrectiva.setESTADO = "Abierta"
        tabAccCorrectiva.TabVisible(1) = False
        txtFechaPuestaEnMarcha.value = Now
        txtPlazoDefinido.value = Now
        cmdTramitar.Visible = False
        Exit Sub
    End If
    
    If mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
        tabAccCorrectiva.TabVisible(1) = False
        'cmdTramitar.Visible = True
    End If
    
    If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
            cmdTramitar.Visible = False
            cmdCerrarExito.Visible = False
            cmdCerrarYGenerarPNC.Visible = False
    Else
        If mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
            cmdTramitar.Visible = False
            cmdCerrarExito.Visible = True
            cmdCerrarYGenerarPNC.Visible = True
        End If
        
        If mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA Or mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL Then
            cmdTramitar.Visible = False
            cmdCerrarExito.Visible = False
            cmdCerrarYGenerarPNC.Visible = False
        End If
    End If
    
    ' PresentaDatos
    Call PresentarDatos
    
    Call OpcionesEdicion
    
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Function ComprobarDatos(Optional ByVal prmtitulo As String = "Guardar Acción Correctiva", Optional ByVal prmCerrando As Boolean = False) As Boolean
Dim blnErr As Boolean, blnErrCerrando As Boolean
Dim strMensaje As String
On Error GoTo ComprobarDatos_Error
    
blnErr = False
blnErrCerrando = False

strMensaje = "Se han detectado los siguientes errores: "

If opTipo(0).value = False And opTipo(1).value = False Then
    strMensaje = strMensaje & vbCrLf & " - Indique si se trata de una Accioón Correctiva/Preventiva"
    blnErr = True
End If
    

If Trim(txtTitulo.Text) = "" Then
    strMensaje = strMensaje & vbCrLf & " - Ha de establecer un título para la Acción Correctiva"
    blnErr = True
    'Exit Function
End If


If Trim(txtDescripcion.Text) = "" Then
    strMensaje = strMensaje & vbCrLf & " - Ha de establecer un texto descriptivo de la acción correctiva"
    blnErr = True
    'Exit Function
End If


If CLng(txtFechaPuestaEnMarcha.value) > CLng(txtPlazoDefinido.value) Then
    strMensaje = strMensaje & vbCrLf & " - La fecha de Puesta en Marcha no podrá ser posterior a la Fecha Prevista de Resolución."
    blnErr = True
    'Exit Function
End If

If cmbResponsableImplantacion.ListIndex < 0 Then
    strMensaje = strMensaje & vbCrLf & " - Debe señalar un Responsable para la Implantación de la Acción Correctiva."
    blnErr = True
    'Exit Function
ElseIf cmbResponsableImplantacion.ItemData(cmbResponsableImplantacion.ListIndex) = 0 Then
    strMensaje = strMensaje & vbCrLf & " - Debe señalar un Responsable para la Implantación de la Acción Correctiva."
    blnErr = True
    'Exit Function
End If

If prmCerrando Then
    If Not (optAccionCorrectivaEnPlazo_No.value Or optAccionCorrectivaEnPlazo_Si.value) Then
        blnErrCerrando = True
    End If
    
    If Not (optComunicado_a_departamentos_no.value Or optComunicado_a_departamentos_si.value) Then
        blnErrCerrando = True
    End If
    If Not (optEfectiva_no.value Or optEfectiva_si.value) Then
        blnErrCerrando = True
    End If
    If Not (optEvidencias_no.value Or optEvidencias_si.value) Then
        blnErrCerrando = True
    End If
    
    If blnErrCerrando Then
        strMensaje = strMensaje & vbCrLf & " - Debe señalar Todas las respuestas en la resolución de la Acción Correctiva."
        blnErr = True
    End If
    
End If


If blnErr Then
    MsgBox strMensaje, vbInformation, prmtitulo
    ComprobarDatos = False
Else
    ComprobarDatos = True
End If




    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.ComprobarDatos"
    Exit Function
ComprobarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.ComprobarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Private Sub GuardarDatos()
On Error GoTo GuardarDatos_Error
    
With mvarobjAccionCorrectiva
    .setTIPO_ID = 0
    If opTipo(1).value = True Then
        .setTIPO_ID = 1
    End If
    .setTITULO = txtTitulo.Text
    .setPROCNC_ID = mvarobjPnc.getID_PROCNC
    .setDESCRIPCION = txtDescripcion.Text
    .setFECHA_PUESTA_EN_MARCHA = txtFechaPuestaEnMarcha.value
    .setFECHA_PREVISTA = txtPlazoDefinido.value
    .setRESPONSABLE_ID = cmbResponsableImplantacion.ItemData(cmbResponsableImplantacion.ListIndex)
    .setRESPONSABLE = cmbResponsableImplantacion.List(cmbResponsableImplantacion.ListIndex)
    .setFECHA_RESOLUCION = txtFechaResolucion.value
    
    
    If optAccionCorrectivaEnPlazo_Si.value Then
        .setRESOLUCION_EFECTIVA = 1
    ElseIf optAccionCorrectivaEnPlazo_No.value Then
        .setRESOLUCION_EFECTIVA = 0
    Else
        .setRESOLUCION_EFECTIVA = -1
    End If
    
    If optComunicado_a_departamentos_si.value Then
        .setRESOLUCION_COMUNICADO_MODIFICACIONES = 1
    ElseIf optComunicado_a_departamentos_no.value Then
        .setRESOLUCION_COMUNICADO_MODIFICACIONES = 0
    Else
        .setRESOLUCION_COMUNICADO_MODIFICACIONES = -1
    End If
    
    If optEvidencias_si.value Then
        .setRESOLUCION_EVIDENCIAS = 1
    ElseIf optEvidencias_no.value Then
        .setRESOLUCION_EVIDENCIAS = 0
    Else
        .setRESOLUCION_EVIDENCIAS = -1
    End If

    If optAccionCorrectivaEnPlazo_Si.value Then
        .setRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 1
    ElseIf optAccionCorrectivaEnPlazo_No.value Then
        .setRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 0
    Else
        .setRESOLUCION_PUESTA_MARCHA_EN_PLAZO = -1
    End If
    
    If .getID_ACCION <> 0 Then
        .Modificar
    Else
        .Insertar
    End If
    
End With
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.GuardarDatos"
    Exit Sub
GuardarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.GuardarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatos of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
