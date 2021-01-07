VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcNCEdicion_AccionCorrectiva 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acción Correctiva / Preventiva"
   ClientHeight    =   7320
   ClientLeft      =   3855
   ClientTop       =   2010
   ClientWidth     =   8640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTipo 
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
      Top             =   585
      Width           =   8520
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
         TabIndex        =   39
         Top             =   315
         Width           =   2670
      End
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
      Picture         =   "frmProcNCEdicion_AccionCorrectiva.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6390
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdCerrarYGenerarPNC 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar Y Generar P.N.C."
      Height          =   870
      Left            =   1980
      Picture         =   "frmProcNCEdicion_AccionCorrectiva.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6390
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdCerrarExito 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar Acción"
      Height          =   870
      Left            =   30
      Picture         =   "frmProcNCEdicion_AccionCorrectiva.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6390
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdTramitar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitar Acción"
      Height          =   870
      Left            =   30
      Picture         =   "frmProcNCEdicion_AccionCorrectiva.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6390
      Width           =   1905
   End
   Begin TabDlg.SSTab tabAccCorrectiva 
      Height          =   4890
      Left            =   45
      TabIndex        =   4
      Top             =   1410
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8625
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Datos Accion"
      TabPicture(0)   =   "frmProcNCEdicion_AccionCorrectiva.frx":2328
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
      Tab(0).Control(8)=   "txtTitulo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmbResponsableImplantacion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Resolución Accion"
      TabPicture(1)   =   "frmProcNCEdicion_AccionCorrectiva.frx":2344
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabel(8)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtFechaResolucion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraResolucion(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraResolucion(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraResolucion(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraResolucion(3)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin pryCombo.miCombo cmbResponsableImplantacion 
         Height          =   315
         Left            =   2550
         TabIndex        =   36
         Top             =   3645
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   556
      End
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   30
         MaxLength       =   150
         TabIndex        =   30
         Top             =   690
         Width           =   8205
      End
      Begin VB.Frame fraResolucion 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   3
         Left            =   -74880
         TabIndex        =   24
         Top             =   1770
         Width           =   8205
         Begin VB.OptionButton optEvidencias_si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   26
            Top             =   270
            Width           =   705
         End
         Begin VB.OptionButton optEvidencias_no 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   25
            Top             =   270
            Width           =   645
         End
         Begin VB.Label lblLabel 
            Caption         =   $"frmProcNCEdicion_AccionCorrectiva.frx":2360
            Height          =   555
            Index           =   7
            Left            =   120
            TabIndex        =   27
            Top             =   90
            Width           =   5415
         End
      End
      Begin VB.Frame fraResolucion 
         BorderStyle     =   0  'None
         Height          =   400
         Index           =   2
         Left            =   -74880
         TabIndex        =   20
         Top             =   1410
         Width           =   8205
         Begin VB.OptionButton optComunicado_a_departamentos_no 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   22
            Top             =   90
            Width           =   645
         End
         Begin VB.OptionButton optComunicado_a_departamentos_si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   21
            Top             =   90
            Width           =   705
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "¿Se ha comunicado la modificación a todos los departamentos?"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   23
            Top             =   90
            Width           =   4515
         End
      End
      Begin VB.Frame fraResolucion 
         BorderStyle     =   0  'None
         Height          =   400
         Index           =   1
         Left            =   -74880
         TabIndex        =   16
         Top             =   930
         Width           =   8205
         Begin VB.OptionButton optEfectiva_si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   18
            Top             =   120
            Width           =   705
         End
         Begin VB.OptionButton optEfectiva_no 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   17
            Top             =   120
            Width           =   645
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "¿Son Efectivas?"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1170
         End
      End
      Begin VB.Frame fraResolucion 
         BorderStyle     =   0  'None
         Height          =   400
         Index           =   0
         Left            =   -74880
         TabIndex        =   12
         Top             =   450
         Width           =   8205
         Begin VB.OptionButton optAccionCorrectivaEnPlazo_No 
            Caption         =   "No"
            Height          =   195
            Left            =   7200
            TabIndex        =   15
            Top             =   120
            Width           =   645
         End
         Begin VB.OptionButton optAccionCorrectivaEnPlazo_Si 
            Caption         =   "Sí"
            Height          =   195
            Left            =   5940
            TabIndex        =   14
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "¿La Acción ha sido puesta en marcha en plazo?"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   3420
         End
      End
      Begin MSComCtl2.DTPicker txtFechaPuestaEnMarcha 
         Height          =   315
         Left            =   2550
         TabIndex        =   7
         Top             =   4035
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51838977
         CurrentDate     =   40156
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   2280
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1275
         Width           =   8145
      End
      Begin MSComCtl2.DTPicker txtPlazoDefinido 
         Height          =   315
         Left            =   2550
         TabIndex        =   9
         Top             =   4425
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51838977
         CurrentDate     =   40156
      End
      Begin MSComCtl2.DTPicker txtFechaResolucion 
         Height          =   315
         Left            =   -72450
         TabIndex        =   28
         Top             =   3060
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51838977
         CurrentDate     =   40156
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Título de la Acción"
         Height          =   195
         Index           =   9
         Left            =   30
         TabIndex        =   31
         Top             =   450
         Width           =   1350
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Resolución Incidencia"
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   29
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
         Top             =   3705
         Width           =   2265
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Plazo para la Resolución"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   4485
         Width           =   1755
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Puesta En Marcha"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   4095
         Width           =   1815
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Descripción de la Acción"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   6
         Top             =   1050
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6390
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6390
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acción Correctiva / Preventiva"
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
      Width           =   3150
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Describa los datos de la Acción Correctiva / Preventiva, rellenando los siguientes campos."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   6390
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8010
      Picture         =   "frmProcNCEdicion_AccionCorrectiva.frx":23FD
      Top             =   60
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   8730
   End
End
Attribute VB_Name = "frmProcNCEdicion_AccionCorrectiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public PK_PNC As Long

Private mvarobjAccionCorrectiva As New clsProcNcAccionCorrectora
Private mvarenuEstado_pnc As C_PROCNC_ESTADOS
Private mvarenuNivelAcceso As C_PROCNC_NIVELES_ACCESO

Private Sub OpcionesEdicion()
    On Error GoTo OpcionesEdicion_Error
    
    If PK = 0 Then
        
        ' cuando se está dando de alta, lo unico que muestra es la primera pestaña, da igual todo lo demás
        tabAccCorrectiva.TabVisible(1) = False
        
        ' Y no muestra ningun boton
        cmdTramitar.Visible = False
        cmdCerrarExito.Visible = False
        cmdCerrarYGenerarPNC.Visible = False
        
        Exit Sub
    End If
    
    
    ' cuando es edicion, depende del estado del pnc, del estado de la acción, y de quien seas
    
    '1º, por estaod de la accion
    If mvarenuEstado_pnc = PTE_PLAN_ACCIONES_CORRECTIVAS Then
        ' cuando es pte de acciones correctivas, solo el jefe de equipo y calidad, pueden modificar. pero...
        cmdTramitar.Visible = False
        cmdCerrarExito.Visible = False
        cmdCerrarYGenerarPNC.Visible = False
        cmdEnviarAvisoFinalizada.Visible = False
        tabAccCorrectiva.TabVisible(1) = False
        ' si no es autorizado, solo lectura
        If mvarenuNivelAcceso <> JEFE_EQUIPO_INVESTIGACION And mvarenuNivelAcceso <> JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO And mvarenuNivelAcceso <> ACCESO_TOTAL Then
            txtTitulo.Enabled = False
            txtdescripcion.Enabled = False
            cmbResponsableImplantacion.desactivar
            txtFechaPuestaEnMarcha.Enabled = False
            txtPlazoDefinido.Enabled = False
            cmdok.Enabled = False
            frmTipo.Enabled = False
        End If
    ElseIf mvarenuEstado_pnc = pte_cierre Or CERRADA_PARCIAL_EVAL Then
        ' Solo pueden modificar los que tengan acceso total
                
        ' para todos los demás, depende del estado de la propia accion correctiva
        If mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
            ' cuando está abierta, en este estado de pnc, solo calidad puede tramitar
            cmdCerrarExito.Visible = False
            cmdEnviarAvisoFinalizada.Visible = False
            cmdCerrarYGenerarPNC.Visible = False
            
            If mvarenuNivelAcceso = ACCESO_TOTAL Then
                cmdTramitar.Visible = True
            Else
                ' no siendo acceso  total, lo desactiva todo
                cmdTramitar.Visible = False
                cmdok.Enabled = False
                frmTipo.Enabled = False
                txtTitulo.Enabled = False
                txtdescripcion.Enabled = False
                cmbResponsableImplantacion.desactivar
                txtFechaPuestaEnMarcha.Enabled = False
                txtPlazoDefinido.Enabled = False
            End If
        ElseIf mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
            ' no se puede volver a tramitar
            cmdTramitar.Visible = False
            
            'Acceso total para calidad
            If mvarenuNivelAcceso = ACCESO_TOTAL Then
                cmdEnviarAvisoFinalizada.Visible = False
                cmdCerrarExito.Visible = True
                cmdCerrarYGenerarPNC.Visible = True
                
            ElseIf USUARIO.getID_EMPLEADO = mvarobjAccionCorrectiva.getRESPONSABLE_ID Then
                tabAccCorrectiva.TabVisible(1) = False ' no deja ver la segunda accion
                cmdEnviarAvisoFinalizada.Visible = True
                cmdCerrarExito.Visible = False
                cmdCerrarYGenerarPNC.Visible = False
                cmbResponsableImplantacion.desactivar
                txtFechaPuestaEnMarcha.Enabled = False
                txtPlazoDefinido.Enabled = False
            Else
                ' Si no es responsable de calidad o de la accion correctiva, solo lectura.
                cmdEnviarAvisoFinalizada.Visible = False
                tabAccCorrectiva.TabVisible(1) = False ' no deja ver la segunda accion
                    ' pero desactiva todo lo que contiene
                    fraResolucion(0).Enabled = False
                    fraResolucion(1).Enabled = False
                    fraResolucion(2).Enabled = False
                    fraResolucion(3).Enabled = False
                    txtFechaResolucion.Enabled = False
                
                cmdCerrarExito.Visible = False
                cmdCerrarYGenerarPNC.Visible = False
                
                
                txtTitulo.Enabled = False
                txtdescripcion.Enabled = False
                cmbResponsableImplantacion.desactivar
                txtFechaPuestaEnMarcha.Enabled = False
                txtPlazoDefinido.Enabled = False
                cmdok.Enabled = False
                frmTipo.Enabled = False
            End If
            ' aviso a finalizar para los implicados
        Else ' LA ACCION CORRECTIVA ESTÁ CERRADA
            
            'LO DESACTIVA TODO
            
            cmdEnviarAvisoFinalizada.Visible = False
            
            tabAccCorrectiva.TabVisible(1) = True ' no deja ver la segunda accion
                ' pero desactiva todo lo que contiene
                fraResolucion(0).Enabled = False
                fraResolucion(1).Enabled = False
                fraResolucion(2).Enabled = False
                fraResolucion(3).Enabled = False
                txtFechaResolucion.Enabled = False
            cmdCerrarExito.Visible = False
            cmdCerrarYGenerarPNC.Visible = False
            cmdTramitar.Visible = False
            txtTitulo.Enabled = False
            txtdescripcion.Enabled = False
            cmbResponsableImplantacion.desactivar
            txtFechaPuestaEnMarcha.Enabled = False
            txtPlazoDefinido.Enabled = False
            cmdok.Enabled = False
            frmTipo.Enabled = False
        End If
    Else ' al estar cerrada total, solo deja visualizarla
        cmdEnviarAvisoFinalizada.Visible = False
        tabAccCorrectiva.TabVisible(1) = True ' no deja ver la segunda accion
            ' pero desactiva todo lo que contiene
            fraResolucion(0).Enabled = False
            fraResolucion(1).Enabled = False
            fraResolucion(2).Enabled = False
            fraResolucion(3).Enabled = False
            txtFechaResolucion.Enabled = False
        cmdCerrarExito.Visible = False
        cmdCerrarYGenerarPNC.Visible = False
        
        txtTitulo.Enabled = False
        txtdescripcion.Enabled = False
        cmbResponsableImplantacion.desactivar
        txtFechaPuestaEnMarcha.Enabled = False
        txtPlazoDefinido.Enabled = False
        cmdok.Enabled = False
        frmTipo.Enabled = False
    End If
    
    
    Exit Sub
    
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.OpcionesEdicion"
    Exit Sub
OpcionesEdicion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.OpcionesEdicion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure OpcionesEdicion of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub PresentarDatos()
   
On Error GoTo PresentarDatos_Error
    
If PK = 0 Then
    mvarobjAccionCorrectiva.setID_AUX = -2
    mvarobjAccionCorrectiva.setESTADO_ID = C_PROCNC_ESTADOS.ABIERTA
    mvarobjAccionCorrectiva.setESTADO = "Abierta"
    
    txtFechaPuestaEnMarcha.Value = Now
    txtPlazoDefinido.Value = Now
    txtFechaResolucion.Value = Now
    Exit Sub
End If
    

With mvarobjAccionCorrectiva
    opTipo(.getTIPO_ID).Value = True
    
    txtTitulo.Text = .getTITULO
    txtdescripcion.Text = .getDESCRIPCION
    txtFechaPuestaEnMarcha.Value = .getFECHA_PUESTA_EN_MARCHA
    txtPlazoDefinido.Value = .getFECHA_PREVISTA
    
    cmbResponsableImplantacion.MostrarElemento .getRESPONSABLE_ID
        
    txtFechaResolucion.Value = .getFECHA_RESOLUCION
    If .getRESOLUCION_EFECTIVA <> -1 Then optEfectiva_si.Value = (.getRESOLUCION_EFECTIVA = 1)
    If .getRESOLUCION_EFECTIVA <> -1 Then optEfectiva_no.Value = (.getRESOLUCION_EFECTIVA = 0)
    
    If .getRESOLUCION_COMUNICADO_MODIFICACIONES <> -1 Then optComunicado_a_departamentos_si.Value = (.getRESOLUCION_COMUNICADO_MODIFICACIONES = 1)
    If .getRESOLUCION_COMUNICADO_MODIFICACIONES <> -1 Then optComunicado_a_departamentos_no.Value = (.getRESOLUCION_COMUNICADO_MODIFICACIONES = 0)
    
    If .getRESOLUCION_EVIDENCIAS <> -1 Then optEvidencias_si.Value = (.getRESOLUCION_EVIDENCIAS = 1)
    If .getRESOLUCION_EVIDENCIAS <> -1 Then optEvidencias_no.Value = (.getRESOLUCION_EVIDENCIAS = 0)
    
    If .getRESOLUCION_PUESTA_MARCHA_EN_PLAZO <> -1 Then optAccionCorrectivaEnPlazo_Si.Value = (.getRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 1)
    If .getRESOLUCION_PUESTA_MARCHA_EN_PLAZO <> -1 Then optAccionCorrectivaEnPlazo_No.Value = (.getRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 0)
    
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
        
    strFecha = Format(txtFechaResolucion.Value, "dd/mm/yyyy")
    strTitulo = txtTitulo.Text
    
    If Not ComprobarDatos("Cerrar Acción", True) Then Exit Sub
    If MsgBox("Ha decidido cerrar la Acción : " & strTitulo & " en la siguiente Fecha: " & strFecha & " ¿Esta seguro?", vbQuestion + vbYesNo, "Cerrar Accion") = vbNo Then Exit Sub
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
    
    strMensaje = InputBox("Para dar por finalizada la Accion, escriba un mensaje si desea indicar alguna observacion de los responsables de calidad.", "Mensaje de Finalizacion de Tarea")
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

Private Sub cmdok_Click()
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

Private Sub cmdcancel_Click()
On Error GoTo cmdcancel_Click_Error
    mvarblnResultado = False
    Me.Hide
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdSalir_Click"
    Exit Sub
cmdcancel_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AccCorrectivas.cmdSalir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdcancel_Click of Formulario frmProcNC_AccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub cmdTramitar_Click()
    Dim strRESPONSABLE As String
On Error GoTo cmdTramitar_Click_Error
    If Not ComprobarDatos("Tramitar Acción") Then Exit Sub
    strRESPONSABLE = cmbResponsableImplantacion.getTEXTO
    If MsgBox("Va a Tramitar la Acción. El Responsable de su implantación será: " & strRESPONSABLE & vbCrLf & " ¿Desea Continuar?", vbQuestion + vbYesNo, "Tramitar Accion") = vbNo Then Exit Sub
    mvarobjAccionCorrectiva.setESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION
    mvarobjAccionCorrectiva.setESTADO = "En Tramitación"
    Call GuardarDatos
'JGM    mvarobjAccionCorrectiva.EnviarMensaje_AccTramitada
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
    strFecha = Format(txtFechaResolucion.Value, "dd/mm/yyyy")
    strTitulo = txtTitulo.Text
    If Not ComprobarDatos("Cerrar Acc. Y Generar Incidencia", True) Then Exit Sub
    If MsgBox("Ha decidido cerrar la Acción : " & strTitulo & " y generar un nueva Incidencia en la siguiente Fecha: " & strFecha, vbInformation + vbYesNo, "Cerrar Accion") = vbNo Then Exit Sub

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
    llenar_combo cmbResponsableImplantacion, New clsUsuarios, 0, frmUsuarios, ""
        
    If PK <> 0 Then
        mvarobjAccionCorrectiva.Carga PK
        PK_PNC = mvarobjAccionCorrectiva.getPROCNC_ID
    End If
        
    If mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
        tabAccCorrectiva.TabVisible(1) = False
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

Private Function ComprobarDatos(Optional ByVal prmtitulo As String = "Guardar Acción", Optional ByVal prmCerrando As Boolean = False) As Boolean
Dim blnErr As Boolean, blnErrCerrando As Boolean
Dim strMensaje As String
On Error GoTo ComprobarDatos_Error
    
blnErr = False
blnErrCerrando = False

strMensaje = "Se han detectado los siguientes errores: "

If opTipo(0).Value = False And opTipo(1).Value = False Then
    strMensaje = strMensaje & vbCrLf & " - Indique si se trata de una Accioón Correctiva/Preventiva"
    blnErr = True
End If

If Trim(txtTitulo.Text) = "" Then
    strMensaje = strMensaje & vbCrLf & " - Ha de establecer un título para la Acción"
    blnErr = True
    'Exit Function
End If


If Trim(txtdescripcion.Text) = "" Then
    strMensaje = strMensaje & vbCrLf & " - Ha de establecer un texto descriptivo de la acción"
    blnErr = True
    'Exit Function
End If


If CLng(txtFechaPuestaEnMarcha.Value) > CLng(txtPlazoDefinido.Value) Then
    strMensaje = strMensaje & vbCrLf & " - La fecha de Puesta en Marcha no podrá ser posterior a la Fecha Prevista de Resolución."
    blnErr = True
    'Exit Function
End If

If cmbResponsableImplantacion.getPK_SALIDA <= 0 Then
    strMensaje = strMensaje & vbCrLf & " - Debe señalar un Responsable para la Implantación de la Acción."
    blnErr = True
    'Exit Function
'ElseIf cmbResponsableImplantacion.ItemData(cmbResponsableImplantacion.ListIndex) = 0 Then
'    strMensaje = strMensaje & vbCrLf & " - Debe señalar un Responsable para la Implantación de la Acción Correctiva."
'    blnErr = True
'    'Exit Function
End If

    If prmCerrando Then
        If Not (optAccionCorrectivaEnPlazo_No.Value Or optAccionCorrectivaEnPlazo_Si.Value) Then
            blnErrCerrando = True
        End If
        
        If Not (optComunicado_a_departamentos_no.Value Or optComunicado_a_departamentos_si.Value) Then
            blnErrCerrando = True
        End If
        If Not (optEfectiva_no.Value Or optEfectiva_si.Value) Then
            blnErrCerrando = True
        End If
        If Not (optEvidencias_no.Value Or optEvidencias_si.Value) Then
            blnErrCerrando = True
        End If
        
        If blnErrCerrando Then
            strMensaje = strMensaje & vbCrLf & " - Debe señalar Todas las respuestas en la resolución de la Acción."
            blnErr = True
        End If
        
    End If
    
    
    If blnErr Then
        MsgBox strMensaje, vbExclamation, prmtitulo
        ComprobarDatos = False
    Else
        ' Comprueba si ha modificado el usuario Responsable de la implantacion
        ' Siempre que la Acc. Correctiva esté en Tramitacion
        If mvarobjAccionCorrectiva.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
            If cmbResponsableImplantacion.getPK_SALIDA <> mvarobjAccionCorrectiva.getRESPONSABLE_ID Then
                MsgBox "Se ha modificado el Responsable de implantación de la Accion. Se le enviará un aviso al nuevo responsable, y otro mensaje indicando el cambio al anterior responsable.", vbInformation, "Guardar Accion Correctiva"
            End If
        End If
    
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
    If opTipo(1).Value = True Then
        .setTIPO_ID = 1
    End If
    
    .setTITULO = txtTitulo.Text
    .setPROCNC_ID = PK_PNC
    .setDESCRIPCION = txtdescripcion.Text
    .setFECHA_PUESTA_EN_MARCHA = txtFechaPuestaEnMarcha.Value
    .setFECHA_PREVISTA = txtPlazoDefinido.Value
    .setRESPONSABLE_ID = cmbResponsableImplantacion.getPK_SALIDA
    .setRESPONSABLE = cmbResponsableImplantacion.getTEXTO
    .setFECHA_RESOLUCION = txtFechaResolucion.Value
    
    If optAccionCorrectivaEnPlazo_Si.Value Then
        .setRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 1
    ElseIf optAccionCorrectivaEnPlazo_No.Value Then
        .setRESOLUCION_PUESTA_MARCHA_EN_PLAZO = 0
    Else
        .setRESOLUCION_PUESTA_MARCHA_EN_PLAZO = -1
    End If
    
    If optEfectiva_si.Value Then
        .setRESOLUCION_EFECTIVA = 1
    ElseIf optEfectiva_no.Value Then
        .setRESOLUCION_EFECTIVA = 0
    Else
        .setRESOLUCION_EFECTIVA = -1
    End If
    
    If optComunicado_a_departamentos_si.Value Then
        .setRESOLUCION_COMUNICADO_MODIFICACIONES = 1
    ElseIf optComunicado_a_departamentos_no.Value Then
        .setRESOLUCION_COMUNICADO_MODIFICACIONES = 0
    Else
        .setRESOLUCION_COMUNICADO_MODIFICACIONES = -1
    End If
    
    If optEvidencias_si.Value Then
        .setRESOLUCION_EVIDENCIAS = 1
    ElseIf optEvidencias_no.Value Then
        .setRESOLUCION_EVIDENCIAS = 0
    Else
        .setRESOLUCION_EVIDENCIAS = -1
    End If
    
    If PK = 0 Then
        .Insertar
    Else
        .Modificar
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

Public Property Get estado_pnc() As C_PROCNC_ESTADOS

    estado_pnc = mvarenuEstado_pnc

End Property

Public Property Let estado_pnc(ByVal enuestado_pnc As C_PROCNC_ESTADOS)

    mvarenuEstado_pnc = enuestado_pnc

End Property

Public Property Get NivelAcceso() As C_PROCNC_NIVELES_ACCESO

    NivelAcceso = mvarenuNivelAcceso

End Property

Public Property Let NivelAcceso(ByVal enuNivelAcceso As C_PROCNC_NIVELES_ACCESO)

    mvarenuNivelAcceso = enuNivelAcceso

End Property
