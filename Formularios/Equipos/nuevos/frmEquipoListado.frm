VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoListado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Equipos de Medición y Ensayo"
   ClientHeight    =   10080
   ClientLeft      =   2940
   ClientTop       =   1545
   ClientWidth     =   16755
   Icon            =   "frmEquipoListado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   16755
   Begin VB.CommandButton cmdRecepcionMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alta Muestra Calibración"
      Height          =   870
      Left            =   8685
      Picture         =   "frmEquipoListado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Generar etiqueta"
      Top             =   9165
      Width           =   2130
   End
   Begin VB.CommandButton cmdmail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enviar Mail"
      Height          =   870
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   9165
      Width           =   1050
   End
   Begin VB.CommandButton cmdAsignacionRapida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asig. Rápida"
      Height          =   870
      Left            =   7590
      Picture         =   "frmEquipoListado.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Listados"
      Top             =   9165
      Width           =   1050
   End
   Begin VB.CommandButton cmdVerAvisos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Avisos"
      Height          =   540
      Left            =   12750
      Picture         =   "frmEquipoListado.frx":2406
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Listados"
      Top             =   9210
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   13440
      Top             =   9270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoListado.frx":27DB
            Key             =   "E0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoListado.frx":2B1E
            Key             =   "E1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoListado.frx":2D53
            Key             =   "E2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoListado.frx":2EF7
            Key             =   "E3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoListado.frx":3199
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   600
      Left            =   14085
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Generar etiqueta"
      Top             =   9180
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdListado_Seleccion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listados"
      Height          =   870
      Left            =   6525
      Picture         =   "frmEquipoListado.frx":99FB
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Listados"
      Top             =   9165
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   45
      TabIndex        =   32
      Top             =   630
      Width           =   16650
      Begin VB.CheckBox chkMTL 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   15105
         TabIndex        =   71
         Top             =   1845
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox chkCP 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   16140
         TabIndex        =   70
         Top             =   1845
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkNadcap 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         Height          =   315
         Left            =   13365
         TabIndex        =   69
         Top             =   1800
         Width           =   1065
      End
      Begin VB.CheckBox chkENAC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         Height          =   315
         Left            =   13365
         TabIndex        =   68
         Top             =   1485
         Width           =   795
      End
      Begin VB.CheckBox chkCalExterna 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externa"
         Enabled         =   0   'False
         Height          =   285
         Left            =   12555
         TabIndex        =   64
         Top             =   135
         Width           =   885
      End
      Begin VB.CheckBox chkCalInterna 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interna"
         Enabled         =   0   'False
         Height          =   285
         Left            =   11700
         TabIndex        =   63
         Top             =   135
         Width           =   885
      End
      Begin VB.CheckBox chkFechaCalibracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Calibración"
         Height          =   315
         Left            =   90
         TabIndex        =   62
         Top             =   2340
         Width           =   1695
      End
      Begin VB.CheckBox chkSoloCanagrosa 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Sólo Equipos de Canagrosa"
         Height          =   315
         Left            =   10170
         TabIndex        =   58
         Top             =   2160
         Width           =   2865
      End
      Begin VB.CheckBox chkSoloClientes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Sólo Equipos de Clientes"
         Height          =   315
         Left            =   10170
         TabIndex        =   56
         Top             =   1890
         Width           =   2865
      End
      Begin VB.TextBox txtNumCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7740
         TabIndex        =   54
         Top             =   1980
         Width           =   2310
      End
      Begin VB.CheckBox chkPrioritario 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prioritarios"
         Height          =   315
         Left            =   10170
         TabIndex        =   51
         Top             =   1620
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.CheckBox chkFS 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Fuera Servicio"
         Height          =   315
         Left            =   10170
         TabIndex        =   16
         Top             =   1305
         Width           =   2100
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   1125
         TabIndex        =   10
         Top             =   1260
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   609
      End
      Begin VB.CheckBox chkEqDeBaja 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Eq. de Baja"
         Height          =   315
         Left            =   10170
         TabIndex        =   15
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox txtNSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1125
         TabIndex        =   4
         Top             =   540
         Width           =   1410
      End
      Begin VB.CheckBox chkConManteminiento 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Con Mantenimiento"
         Height          =   315
         Left            =   10170
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkConVerificacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Con Verificación"
         Height          =   345
         Left            =   10170
         TabIndex        =   13
         Top             =   400
         Width           =   1545
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3915
         TabIndex        =   1
         Top             =   180
         Width           =   2340
      End
      Begin VB.TextBox txtNEqipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1125
         TabIndex        =   0
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox txtFabricante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1125
         TabIndex        =   7
         Top             =   900
         Width           =   1410
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   13815
         TabIndex        =   29
         Top             =   135
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo cmbFamilia 
         Height          =   315
         Left            =   3915
         TabIndex        =   8
         Top             =   900
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEstadoVer 
         Height          =   315
         Left            =   7740
         TabIndex        =   9
         Top             =   1260
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbLocalizacion 
         Height          =   315
         Left            =   3915
         TabIndex        =   5
         Top             =   540
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEstadoCal 
         Height          =   315
         Left            =   7740
         TabIndex        =   6
         Top             =   900
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEstadoMto 
         Height          =   315
         Left            =   7740
         TabIndex        =   11
         Top             =   1620
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbTipoEquipo 
         Height          =   315
         Left            =   7740
         TabIndex        =   3
         Top             =   540
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbAccesorios 
         Height          =   315
         Left            =   13410
         TabIndex        =   46
         Top             =   2160
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   345
         Left            =   1125
         TabIndex        =   48
         Top             =   1620
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Filtro"
         Height          =   960
         Left            =   15435
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   675
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   345
         Left            =   1125
         TabIndex        =   52
         Top             =   1980
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbEstado 
         Height          =   315
         Left            =   7740
         TabIndex        =   2
         Top             =   180
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1800
         TabIndex        =   59
         Top             =   2340
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51642369
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3735
         TabIndex        =   60
         Top             =   2340
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51642369
         CurrentDate     =   38002
      End
      Begin VB.CheckBox chkConCalibracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Con Calibración"
         Height          =   285
         Left            =   10170
         TabIndex        =   12
         Top             =   150
         Width           =   1605
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Height          =   315
         Left            =   7740
         TabIndex        =   66
         Top             =   2340
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "- MTL"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   38
         Left            =   14580
         TabIndex        =   73
         Top             =   1845
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "- CP"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   39
         Left            =   15705
         TabIndex        =   72
         Top             =   1845
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Left            =   6345
         TabIndex        =   67
         Top             =   2430
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   3195
         TabIndex        =   61
         Top             =   2385
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado Equipo"
         Height          =   195
         Left            =   6300
         TabIndex        =   57
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Num.Equip.Cliente"
         Height          =   240
         Index           =   5
         Left            =   6345
         TabIndex        =   55
         Top             =   2055
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   75
         TabIndex        =   53
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Left            =   75
         TabIndex        =   49
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   75
         TabIndex        =   47
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Equipo"
         Height          =   195
         Left            =   6300
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado Mto."
         Height          =   195
         Left            =   6330
         TabIndex        =   43
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricante"
         Height          =   240
         Index           =   4
         Left            =   75
         TabIndex        =   41
         Top             =   975
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado Verificación"
         Height          =   195
         Left            =   6330
         TabIndex        =   40
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado Calibración"
         Height          =   240
         Left            =   6330
         TabIndex        =   39
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Localización"
         Height          =   240
         Left            =   2580
         TabIndex        =   38
         Top             =   615
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Area Metrológica"
         Height          =   240
         Index           =   3
         Left            =   2580
         TabIndex        =   37
         Top             =   975
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   240
         Index           =   2
         Left            =   2580
         TabIndex        =   36
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Equipo"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   255
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Num. Serie"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   34
         Top             =   615
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Left            =   11835
         TabIndex        =   33
         Top             =   990
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   870
      Left            =   11925
      Picture         =   "frmEquipoListado.frx":A2C5
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Ver ficha del equipo"
      Top             =   9180
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Duplicar equipo"
      Top             =   9165
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   15660
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9165
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Añadir equipo"
      Top             =   9165
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Modificar equipo"
      Top             =   9165
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Eliminar equipo"
      Top             =   9165
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5685
      Left            =   60
      TabIndex        =   28
      Top             =   3435
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   10028
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
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
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar"
      Height          =   870
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exportar datos a impresora o excel"
      Top             =   9165
      Width           =   1050
   End
   Begin VB.Label lblNota 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NOTA: Los equipos señalados en Rojo se encuentran FUERA DE SERVICIO"
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   13695
      TabIndex        =   42
      Top             =   9210
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ventana de gestión de Equipos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   31
      Top             =   330
      Width           =   2220
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   16155
      Picture         =   "frmEquipoListado.frx":AB8F
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Equipos de Medición y Ensayo"
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
      TabIndex        =   30
      Top             =   30
      Width           =   4410
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   -45
      Top             =   0
      Width           =   17175
   End
End
Attribute VB_Name = "frmEquipoListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public informe As String
Public criterio As String

Private bln_no_actualizar As Boolean
Private mvarCRITERIO_LISTADO As String


Private Sub chkCP_Click()
    cargar_lista
End Sub

Private Sub chkENAC_Click()
    cargar_lista
End Sub

Private Sub chkMTL_Click()
    cargar_lista
End Sub

Private Sub chkNADCAP_Click()
    lblCampos(38).visible = chkNadcap.Value
    lblCampos(39).visible = chkNadcap.Value
    chkMTL.visible = chkNadcap.Value
    chkCP.visible = chkNadcap.Value
    If chkNadcap.Value = Unchecked Then
        chkMTL.Value = Unchecked
        chkCP.Value = Unchecked
    End If
    cargar_lista
End Sub

Private Sub cmbCentro_Change()
    cargar_lista
End Sub
Private Sub cmdRecepcionMuestra_Click()
   On Error GoTo cmdRecepcionMuestra_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oEquipo As New clsEquipos
    If oEquipo.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
        oEquipo.PonerEquipoEnM1 lista.ListItems(lista.selectedItem.Index).Text
        With frmRecepcion
            .cmbClientes.MostrarElemento oEquipo.getCLIENTE_ID
            .cmbTM.MostrarElemento TIPOS_MUESTRAS.CALIBRACION_EXTERIOR
            Dim rs As ADODB.Recordset
            c = "select id_tipo_analisis from tipos_analisis where nombre = '" & oEquipo.getNOMBRE & "'"
            Set rs = datos_bd(c)
            If rs.RecordCount > 0 Then
                .cmbDatos(2).BoundText = rs(0)
            Else
                .cmbDatos(2).Text = ""
            End If
            .Text1(1) = oEquipo.getNOMBRE & " Id.:" & oEquipo.getNUMERO_EQUIPO_CLIENTE
            .cmbDatos(0).BoundText = USUARIO.getID_EMPLEADO
            .cmbCentro.BoundText = CENTROS.CENTRO_SEVILLA
            .cmbDatos(4).BoundText = 1
            .chkOpcion(4).Value = Checked
            .chkFechaSolicitudNA.Value = Checked
            .cmbCalibracionId.visible = True
            .lblcalibracion.visible = True
            .cargarCalibraciones lista.ListItems(lista.selectedItem.Index).Text
            .Show
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdRecepcionMuestra_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRecepcionMuestra_Click of Formulario frmEquipoListado"
End Sub

Private Sub chkAV_Click()
    cargar_lista
End Sub

Private Sub chkCalExterna_Click()
    cargar_lista
End Sub

Private Sub chkCalInterna_Click()
    cargar_lista
End Sub

Private Sub chkFechaCalibracion_Click()
    If chkFechaCalibracion.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
    cargar_lista
End Sub

Private Sub chkFS_Click()
    cargar_lista
End Sub

Private Sub chkPrioritario_Click()
    cargar_lista
End Sub

Private Sub chkRPR_Click()
    cargar_lista
End Sub

Private Sub chkSoloCanagrosa_Click()
    chkSoloClientes.Value = Unchecked
    If chkSoloCanagrosa.Value = Checked Then
        chkSoloClientes.Enabled = False
    Else
        chkSoloClientes.Enabled = True
    End If
    cargar_lista
End Sub

Private Sub chkSoloClientes_Click()
    chkSoloCanagrosa.Value = Unchecked
    If chkSoloClientes.Value = Checked Then
        chkSoloCanagrosa.Enabled = False
    Else
        chkSoloCanagrosa.Enabled = True
    End If
    cargar_lista
End Sub

Private Sub chkSPR_Click()
    cargar_lista
End Sub

Private Sub cmbAccesorios_Change()
    cargar_lista
End Sub

Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmbEstado_Change()
    'M1124-I
    cargar_lista
    'M1124-F
End Sub

Private Sub cmbEstadoCal_Change()
    cargar_lista
End Sub
Private Sub cmbEstadoMto_Change()
    cargar_lista
End Sub
Private Sub cmbEstadoVer_Change()
    cargar_lista
End Sub
Private Sub cmbfamilia_Change()
    cargar_lista
End Sub
Private Sub cmbLocalizacion_Change()
    cargar_lista
End Sub
Private Sub cmbProveedor_change()
    cargar_lista
End Sub

Private Sub cmbResponsable_Change()
    cargar_lista
End Sub

Private Sub cmbTipoEquipo_Change()
    cargar_lista
End Sub
Private Sub cmdAsignacionRapida_Click()
    Dim objfrm As New frmEquiposAsignacionRapida
    objfrm.Show 1
    Set objfrm = Nothing
End Sub
Private Sub cmdLimpiar_Click()
    bln_no_actualizar = True
    chkConCalibracion.Value = vbUnchecked
    chkConManteminiento.Value = vbUnchecked
    chkConVerificacion.Value = vbUnchecked
    chkEqDeBaja.Value = vbUnchecked
    chkFS.Value = vbUnchecked
    chkPrioritario.Value = vbUnchecked
    
    chkEnac.Value = vbUnchecked
    chkNadcap.Value = vbUnchecked
    
    txtNEqipo.Text = ""
    txtNombre.Text = ""
    txtDescripcion.Text = ""
    txtNSerie.Text = ""
    txtFabricante.Text = ""
    'M1124-I
    cmbEstado.BoundText = ""
    cmbCentro.BoundText = ""
    'M1124-F
    cmbAccesorios.BoundText = ""
    cmbLocalizacion.BoundText = ""
    cmbFamilia.BoundText = ""
    cmbTipoEquipo.BoundText = ""
    cmbEstadoCal.BoundText = ""
    cmbEstadoMto.BoundText = ""
    cmbEstadoVer.BoundText = ""
    cmbResponsable.limpiar
    cmbProveedor.limpiar
    
    bln_no_actualizar = False

    'M1050-I
    cmbClientes.limpiar
    txtNumCliente = ""
    chkSoloClientes.Value = vbUnchecked
    'M1050-F
    cargar_lista
    
End Sub

Private Sub cmdListado_Seleccion_Click()
    informe = ""
    criterio = ""
    frmEquipos_Listados_Seleccion.Show 1
    informe = frmEquipos_Listados_Seleccion.informe
    criterio = frmEquipos_Listados_Seleccion.criterio
    
    If informe <> "" Then
        With frmReport
            .iniciar
            .informe = informe
            .criterio = criterio
            .imprimir = False
            .generar
            .visible = True
        End With
    End If
    
    Unload frmEquipos_Listados_Seleccion
End Sub
Private Sub cmdAnadir_Click()
    Dim objfrm As New frmEquipoEdicion
    Dim lngid As Long
    Dim objEquipo As New clsEquipos
    
    On Error GoTo cmdAnadir_Click_Error
    
    Set objfrm.EQUIPO = objEquipo
    objfrm.TipoEdicion = Alta
    
    objfrm.Show vbModal
    
    If objfrm.RESULTADO Then
        Call cargar_lista
    End If
    
    Unload objfrm
    Set objfrm = Nothing
    
    On Error GoTo 0
        Exit Sub
cmdAnadir_Click_Error:
        'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmEquipoListado"
End Sub

Private Sub cmdduplicar_Click()
    Dim strId As String
    Dim eq As clsEquipos
    
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
   
    strId = lista.ListItems(lista.selectedItem.Index)
    
    Set eq = New clsEquipos
    
    If MsgBox("Va a duplicar el equipo " & strId & " . ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        If eq.duplicar_equipo(strId) Then
        Me.MousePointer = 0
            MsgBox "Equipo duplicado correctamente. El nuevo duplicado posee el Nº " & strId, vbInformation, "Duplicar Equipo"
            cargar_lista
        Else
        Me.MousePointer = 0
            MsgBox "No fue posible duplicar el Equipo.", vbExclamation, "Duplicar Equipo"
        End If
     End If
    Exit Sub
    
fallo:
        Me.MousePointer = 0
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el equipo : " & lista.ListItems(lista.selectedItem.Index).SubItems(2), vbQuestion + vbYesNo, App.Title) = vbYes Then
            ' Analisis
            Dim oEquipo As New clsEquipos
            oEquipo.Eliminar lista.ListItems(lista.selectedItem.Index)
            cargar_lista
        End If
    Else
        MsgBox "Debe seleccionar el equipo que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdFicha_Click()
    If lista.ListItems.Count > 0 Then
        Dim oEquipo As New clsEquipos
        
        Call oEquipo.ImprimirFichaEquipo(lista.ListItems(lista.selectedItem.Index))
        Set oEquipo = Nothing
    Else
        MsgBox "Debe seleccionar el equipo cuya ficha desea ver.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdetiqueta_Click()
    Dim i As Long
    Dim strEquipos As String
    Dim booAlgunoSeleccionado As Boolean
    
On Error GoTo trataError

        Dim objReport As New frmReport
        
        With objReport
            Firmas.copiar_firma_responsable_tecnico
            .iniciar
            .informe = "Equipos\rptEquipos_Etiqueta"
            strEquipos = "{equipos.ID_EQUIPO} in ["
            booAlgunoSeleccionado = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked Then
                    strEquipos = strEquipos & CLng(lista.ListItems(i)) & ".00,"
                    booAlgunoSeleccionado = True
                End If
            Next i
            If booAlgunoSeleccionado Then
                strEquipos = Left(strEquipos, Len(strEquipos) - 1) & "]"
                .criterio = strEquipos
                .imprimir = False
                .generar
                .visible = True
            Else
                MsgBox "Debe marcar los equipos para los que desea generar etiqueta.", vbOKOnly + vbInformation, App.Title
            End If
        End With
        Unload objReport
        Set objReport = Nothing
        
        log ("Final impresion de etiquetas de equipos")

    
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir las etiquetas.", vbCritical, Err.Description
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea exportar a excel?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            generar_excel_listado
        Else
            With frmReport
                .iniciar
                .informe = "Equipos\rptEquipos_Listado"
                .criterio = mvarCRITERIO_LISTADO
                .imprimir = False
                .generar
                .visible = True
            End With
        End If
    Else
        MsgBox "Debe seleccionar el equipo cuya ficha desea imprimir.", vbOKOnly + vbInformation, App.Title
    End If
End Sub
Private Sub generar_excel_listado()
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo generar_excel_listado_Error

    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Add
    Set XLS = XLW.Worksheets(1)
    Me.MousePointer = 11
    XLW.Worksheets(3).Delete
    XLW.Worksheets(2).Delete
    XLW.Worksheets(1).Name = "Listado de Equipos"
    XLS.Range("1:1").HorizontalAlignment = xlCenter
    XLS.Range("1:1").VerticalAlignment = xlCenter
    XLS.Range("1:1").RowHeight = 30
    XLS.Range("1:1").WrapText = True
    'Cabecera
    XLS.Cells(1, 1) = "NºEquipo"
    XLS.Cells(1, 2) = "Estado"
    XLS.Cells(1, 3) = "NºEquipo Cliente"
    XLS.Cells(1, 4) = "Cliente"
    XLS.Cells(1, 5) = "Nombre"
    XLS.Cells(1, 6) = "NºSerie"
    XLS.Cells(1, 7) = "Prox.Calibración"
    XLS.Cells(1, 8) = "Prox.Verificación"
    XLS.Cells(1, 9) = "Prox.Mantenimiento"
    XLS.Cells(1, 10) = "Localización"
    XLS.Cells(1, 11) = "Responsable"
    XLS.Cells(1, 12) = "Calibrador"
    XLS.Cells(1, 13) = "Centro"
    
    i = 2
    ' Datos
    For i = 1 To lista.ListItems.Count
        XLS.Cells(i + 1, 1) = lista.ListItems(i).Text
        XLS.Cells(i + 1, 2) = lista.ListItems(i).SubItems(1)
        XLS.Cells(i + 1, 3) = lista.ListItems(i).SubItems(2)
        XLS.Cells(i + 1, 4) = lista.ListItems(i).SubItems(17) ' Cliente
        XLS.Cells(i + 1, 5) = lista.ListItems(i).SubItems(3)
        XLS.Cells(i + 1, 6) = lista.ListItems(i).SubItems(4)
        XLS.Cells(i + 1, 7) = Format(lista.ListItems(i).SubItems(6), "yyyy-mm-dd")
        XLS.Cells(i + 1, 8) = Format(lista.ListItems(i).SubItems(8), "yyyy-mm-dd")
        XLS.Cells(i + 1, 9) = Format(lista.ListItems(i).SubItems(10), "yyyy-mm-dd")
        XLS.Cells(i + 1, 10) = lista.ListItems(i).SubItems(11)
        XLS.Cells(i + 1, 11) = lista.ListItems(i).SubItems(12)
        XLS.Cells(i + 1, 12) = lista.ListItems(i).SubItems(16) ' Calibrador
        XLS.Cells(i + 1, 13) = lista.ListItems(i).SubItems(18) ' Centro
    Next
    For i = 1 To 13
        XLS.Columns(i).AutoFit
    Next
    XLS.Range("2:" & lista.ListItems.Count + 1).HorizontalAlignment = xlLeft
    
    Me.MousePointer = 0
    XLA.visible = True
'    Set XLS = Nothing
'    Set XLW = Nothing
'    Set XLA = Nothing
   On Error GoTo 0
   Exit Sub

generar_excel_listado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_excel_listado of Formulario frmEquipoListado"
    
End Sub


Private Sub cmdmail_Click()
   On Error GoTo cmdmail_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Me.MousePointer = 11
    ' Generar PDF con el listado
    Dim Listado As String
    On Error Resume Next
    Listado = App.Path & "\Listado de Equipos.pdf"
    Kill Listado
    With frmReport
        .iniciar
        .informe = "Equipos\rptEquipos_Listado"
        .criterio = mvarCRITERIO_LISTADO
        .imprimir = False
        .pdf = Listado
        .generar
        .visible = False
    End With
    Set frm = Nothing
    ' Enviar correo
    If Dir(Listado) Then
        Dim ref As String
        Dim des As String
        ref = "Listado de Equipos"
        If cmbResponsable.getTEXTO <> "" Then
            Dim oUsuario As New clsUsuarios
            oUsuario.CARGAR cmbResponsable.getPK_SALIDA
            des = oUsuario.getEMAIL
        Else
            des = "Indique el destinatario"
        End If
        genera_correo des, ref, "", Listado, Me.hdc
        Me.MousePointer = 0
    Else
        Me.MousePointer = 0
        MsgBox "Error al generar el listado.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdmail_Click_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmail_Click of Formulario frmEquipoListado"

End Sub

Private Sub cmdModificar_Click()

On Error GoTo cmdModificar_Click_Error

    Dim objfrm As New frmEquipoEdicion
    Dim lngid As Long
    Dim objEquipo As New clsEquipos
    
    lngid = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    If lngid <= 0 Then Exit Sub
    
    Call objEquipo.Carga(lngid)
    
    Set objfrm.EQUIPO = objEquipo
    
    If objEquipo.getALTA_BAJA = 1 Then
        objfrm.TipoEdicion = visualizar
    Else
        objfrm.TipoEdicion = EDICION
    End If
    
    objfrm.Show vbModal
    
    cargar_linea_lista lngid, lista.selectedItem.Index
    
    Unload objfrm
    Set objfrm = Nothing
    
    On Error GoTo 0
        Exit Sub
cmdModificar_Click_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmEquipoListado"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdVerAvisos_Click()
    cmbEstadoCal.BoundText = "1"
    cmbEstadoVer.BoundText = "1"
    cmbEstadoMto.BoundText = "1"
    
    cargar_lista
End Sub
Private Sub chkConCalibracion_Click()
    If chkConCalibracion.Value = Checked Then
        chkCalInterna.Enabled = True
        chkCalExterna.Enabled = True
    Else
        chkCalInterna.Enabled = False
        chkCalExterna.Enabled = False
    End If
    Call cargar_lista
End Sub
Private Sub chkConManteminiento_Click()
    Call cargar_lista
End Sub
Private Sub chkConVerificacion_Click()
    Call cargar_lista
End Sub
Private Sub chkEqDeBaja_Click()
    Call cargar_lista
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 50
    Me.Left = 50
    
    bln_no_actualizar = False
    
    cargar_botones Me

    Call cargar_combos
    cabecera
        
    fdesde = Date - 31
    fhasta = Date + 31
'    cargar_lista
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    'M1124-I
    cargar_combo cmbEstado, New clsEq_Estados
    cargar_combo cmbCentro, New clsCentros
    'M1124-F
    
    oDeco.cargar_combo cmbFamilia, DECODIFICADORA.EQ_FAMILIAS
    oDeco.cargar_combo cmbLocalizacion, DECODIFICADORA.EQ_SITUACIONES
        
    oDeco.cargar_combo cmbTipoEquipo, DECODIFICADORA.EQ_TIPOS_EQUIPO
    oDeco.cargar_combo cmbAccesorios, DECODIFICADORA.EQ_SINO
    
    llenar_combo cmbResponsable, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    cargar_estad_cal_ver cmbEstadoCal
    cargar_estad_cal_ver cmbEstadoVer
    cargar_estad_cal_ver cmbEstadoMto
    
    
    Dim consulta As String
    consulta = "SELECT DISTINCT A.ID_PROVEEDOR,A.NOMBRE " & _
               "  FROM PROVEEDORES A,EQUIPOS B " & _
               " Where A.ID_PROVEEDOR = B.PROVEEDOR_ID "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
                   
        With cmbProveedor
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setQUERY = consulta
            .setTABLA = "PROVEEDORES"
            .setDESCRIPCION = "Proveedores"
            .setPK = "ID_PROVEEDOR"
            .setCAMPO = "A.NOMBRE"
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmProveedores_Detalle
        End With
    End If
End Sub
Private Sub cargar_lista()
    
   On Error GoTo cargar_lista_Error

    If bln_no_actualizar Then Exit Sub

    Dim rs As ADODB.Recordset
    
    Dim oEq As New clsEquipos
    Dim intEstado As Integer
    Dim objLitem As ListItem, objSI As ListSubItem
'    Dim dtmFechaServidor As Date
    Dim blnFueraServicio As Boolean
'    Dim str_es_acc As String, str_tipo_eq As String
'    Dim rs_c As ADODB.Recordset, rs_v As ADODB.Recordset, rs_m As ADODB.Recordset
'    Dim oOP As New clsEquiposOperacionesPendientes
    Dim fecha_cvm As String
    lista.ListItems.Clear
    
'    Set rs_c = oOP.Listado_calibraciones_pendientes()
'    Set rs_v = oOP.Listado_verificaciones_pendientes()
'    Set rs_m = oOP.Listado_mantenimientos_pendientes()
        
        
    mvarCRITERIO_LISTADO = ""
        
    str_es_acc = CStr(getDataComboSel(cmbAccesorios))
    If str_es_acc = "-1" Or str_es_acc = "2" Then
        str_es_acc = ""
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " And {equipos.ES_ACCESORIO} = " & str_es_acc
    End If
    
    str_tipo_eq = CStr(getDataComboSel(cmbTipoEquipo))
    If str_tipo_eq = "-1" Then
        str_tipo_eq = ""
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " And {equipos.TIPO_EQUIPO_ID} = " & str_tipo_eq
    End If
    Dim proveedor As Long
    If cmbProveedor.getTEXTO = "" Then
        proveedor = 0
    Else
        proveedor = cmbProveedor.getPK_SALIDA
    End If
    Dim responsable As Long
    If cmbResponsable.getTEXTO = "" Then
        responsable = 0
    Else
        responsable = cmbResponsable.getPK_SALIDA
    End If
    'M1050-I
    Dim cliente As Long
    If cmbClientes.getTEXTO = "" Then
        cliente = 0
    Else
        cliente = cmbClientes.getPK_SALIDA
    End If
    Set rs = oEq.Listado(Trim(txtDescripcion.Text), Trim(txtNSerie.Text), Trim(txtNEqipo.Text), Trim(txtNombre.Text), cmbFamilia.BoundText, cmbLocalizacion.BoundText, txtFabricante.Text, chkEqDeBaja.Value, chkConCalibracion.Value, chkConVerificacion.Value, chkConManteminiento.Value, cmbEstadoCal.BoundText, cmbEstadoVer.BoundText, cmbEstadoMto.BoundText, str_es_acc, str_tipo_eq, mvarCRITERIO_LISTADO, proveedor, chkFS.Value, responsable, chkPrioritario.Value, cliente, txtNumCliente, chkSoloClientes.Value, chkSoloCanagrosa.Value, cmbEstado.BoundText, chkCalInterna.Value, chkCalExterna.Value, cmbCentro.BoundText, chkEnac.Value, chkNadcap.Value, chkMTL.Value, chkCP.Value)
    Dim cont As Long
    cont = 0
    If rs.RecordCount <> 0 Then
'        dtmFechaServidor = Now
        Do
                cont = cont + 1
                Set objLitem = lista.ListItems.Add(, , Format(rs("ID_EQUIPO"), "00000"))
                blnFueraServicio = (CInt(rs("FUERA_SERVICIO")) = 1)
                With objLitem
                    If blnFueraServicio Then
                        .bold = True
                        .ForeColor = RGB(255, 0, 0)
                    End If
                    Set objSI = .ListSubItems.Add(, , rs("ESTADO_ID"))
                    If blnFueraServicio Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    If IsNull(rs("NUMERO_EQUIPO_CLIENTE")) Then
                        Set objSI = .ListSubItems.Add(, , "")
                    Else
                        Set objSI = .ListSubItems.Add(, , CStr(rs("NUMERO_EQUIPO_CLIENTE")))
                    End If
                    If blnFueraServicio Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    Set objSI = .ListSubItems.Add(, , rs("NOMBRE"))
                    If blnFueraServicio Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    Set objSI = .ListSubItems.Add(, , rs("SERIE"))
                    If blnFueraServicio Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    If rs("ESTADO_ID") = "B" Or rs("ESTADO_ID") = "F/S" Or rs("ESTADO_ID") = "CAU" Or rs("ESTADO_ID") = "E" Or rs("ESTADO_ID") = "I" Or rs("ESTADO_ID") = "R" Then
                        .ListSubItems.Add , , ""
                        .ListSubItems.Add , , ""
                        .ListSubItems.Add , , ""
                        .ListSubItems.Add , , ""
                        .ListSubItems.Add , , ""
                        .ListSubItems.Add , , ""
                    Else
                        'CALIBRACION
                        .ListSubItems.Add , , "", CInt(rs("CAL_ESTADO"))
                        If Not IsNull(rs("CAL_FECHA_PREVISTA")) Then
                            Set objSI = .ListSubItems.Add(, , Format(rs("CAL_FECHA_PREVISTA"), "dd/mm/yyyy"))
                            If blnFueraServicio Then
                                objSI.bold = True
                                objSI.ForeColor = RGB(255, 0, 0)
                            End If
                        Else
                            Set objSI = .ListSubItems.Add(, , "")
                        End If
                        'VERIFICACION
                        .ListSubItems.Add , , "", CInt(rs("VER_ESTADO"))
                        If Not IsNull(rs("VER_FECHA_PREVISTA")) Then
                            Set objSI = .ListSubItems.Add(, , Format(rs("VER_FECHA_PREVISTA"), "dd/mm/yyyy"))
                            If blnFueraServicio Then
                                objSI.bold = True
                                objSI.ForeColor = RGB(255, 0, 0)
                            End If
                        Else
                            Set objSI = .ListSubItems.Add(, , "")
                        End If
                        'MANTENIMIENTO
                        .ListSubItems.Add , , "", CInt(rs("MAN_ESTADO"))
                        If Not IsNull(rs("MAN_FECHA_PREVISTA")) Then
                            Set objSI = .ListSubItems.Add(, , Format(rs("MAN_FECHA_PREVISTA"), "dd/mm/yyyy"))
                            If blnFueraServicio Then
                                objSI.bold = True
                                objSI.ForeColor = RGB(255, 0, 0)
                            End If
                        Else
                            Set objSI = .ListSubItems.Add(, , "")
                        End If
                    
                    End If
                    Set objSI = .ListSubItems.Add(, , rs("NOMBRE_LOCALIZACION"))
                    If blnFueraServicio Then
                        objSI.bold = True
                        objSI.ForeColor = RGB(255, 0, 0)
                    End If
                    If Not IsNull(rs("RESPONSABLE")) Then
                        Set objSI = .ListSubItems.Add(, , rs("RESPONSABLE"))
                        If blnFueraServicio Then
                            objSI.bold = True
                            objSI.ForeColor = RGB(255, 0, 0)
                        End If
                    End If
                    .ListSubItems.Add , , IIf(CInt(rs("AV")), "X", "")
                    .ListSubItems.Add , , IIf(CInt(rs("SPR")), "X", "")
                    .ListSubItems.Add , , IIf(CInt(rs("RPR")), "X", "")
                    
                    .ListSubItems.Add , , rs("CALIBRADOR") ' Calibrador
                    .ListSubItems.Add , , IIf(Not IsNull(rs("CLIENTE")), rs("CLIENTE"), "") ' cliente
                    .ListSubItems.Add , , IIf(Not IsNull(rs("CENTRO")), rs("CENTRO"), "") ' CENTRO
                End With
'                If CInt(rs("PRIORITARIO")) = 1 Then
'                    lista.ListItems(lista.ListItems.Count).SmallIcon = 5
'                End If
                If chkFechaCalibracion.Value = Checked Then
                    If Trim(lista.ListItems(lista.ListItems.Count).SubItems(6)) <> "" Then
                        If CDate(lista.ListItems(lista.ListItems.Count).SubItems(6)) < CDate(fdesde) Or CDate(lista.ListItems(lista.ListItems.Count).SubItems(6)) > CDate(fhasta) Then
                            lista.ListItems.Remove lista.ListItems.Count
                            cont = cont - 1
                        End If
                    Else
                        lista.ListItems.Remove lista.ListItems.Count
                        cont = cont - 1
                    End If
                End If
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    lblsubtitulo = "Ventana de gestión de Equipos. Número de equipos mostrados : " & cont
    Set oEq = Nothing

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmEquipoListado"
End Sub

Public Sub cargar_linea_lista(ID As Long, linea As Long)
    Dim rs As ADODB.Recordset
    Dim oEq As New clsEquipos
    Dim intEstado As Integer
    Dim objLitem As ListItem, objSI As ListSubItem
    Dim blnFueraServicio As Boolean
    Dim fecha_cvm As String
    Dim i As Integer
    
    Set rs = oEq.Listado_equipo_unico(ID)
    rs.MoveFirst
            
    Set objLitem = lista.ListItems(linea)
    blnFueraServicio = (CInt(rs("FUERA_SERVICIO")) = 1)
    With objLitem
        If blnFueraServicio Then
        End If
        .ListSubItems(1).Text = rs("ESTADO_ID")
        .ListSubItems(2).Text = rs("NUMERO_EQUIPO_CLIENTE")
        .ListSubItems(3).Text = rs("NOMBRE")
        .ListSubItems(4).Text = rs("SERIE")
        If rs("ESTADO_ID") = "B" Or rs("ESTADO_ID") = "F/S" Then
            .ListSubItems(5).ReportIcon = vbNull
            .ListSubItems(6) = ""
            .ListSubItems(7).ReportIcon = vbNull
            .ListSubItems(8) = ""
            .ListSubItems(9).ReportIcon = vbNull
            .ListSubItems(10) = ""
        Else
            If Not IsNull(rs("CAL_FECHA_PREVISTA")) Then
                .ListSubItems(6) = rs("CAL_FECHA_PREVISTA")
            Else
                .ListSubItems(6) = ""
            End If
            .ListSubItems(7).ReportIcon = CInt(rs("VER_ESTADO"))
            If Not IsNull(rs("VER_FECHA_PREVISTA")) Then
                .ListSubItems(8).Text = rs("VER_FECHA_PREVISTA")
            Else
                .ListSubItems(8).Text = ""
            End If
            .ListSubItems(9).ReportIcon = CInt(rs("MAN_ESTADO"))
            If Not IsNull(rs("MAN_FECHA_PREVISTA")) Then
                .ListSubItems(10).Text = rs("MAN_FECHA_PREVISTA")
            Else
                .ListSubItems(10).Text = ""
            End If
        End If
        .ListSubItems(11).Text = rs("NOMBRE_LOCALIZACION")
        
        If Not IsNull(rs("RESPONSABLE")) Then
            .ListSubItems(12).Text = rs("RESPONSABLE")
        Else
            .ListSubItems(12).Text = ""
        End If
        
        .ListSubItems(13).Text = IIf(CInt(rs("AV")), "X", "")
        .ListSubItems(14).Text = IIf(CInt(rs("SPR")), "X", "")
        .ListSubItems(15).Text = IIf(CInt(rs("RPR")), "X", "")
'        .ListSubItems(16).Text = IIf(Not IsNull(rs("CALIBRADOR")), rs("CALIBRADOR"), "")  ' cliente
        .ListSubItems(17).Text = IIf(Not IsNull(rs("CLIENTE")), rs("CLIENTE"), "")  ' cliente
        
        .ListSubItems(18).Text = rs("CENTRO")
        If blnFueraServicio Then
            .bold = True
            .ForeColor = RGB(255, 0, 0)
            For i = 1 To lista.ColumnHeaders.Count - 1
                .ListSubItems(i).bold = True
                .ListSubItems(i).ForeColor = RGB(255, 0, 0)
            Next
        Else
            .bold = False
            .ForeColor = vbBlack
            For i = 1 To lista.ColumnHeaders.Count - 1
                .ListSubItems(i).bold = False
                .ListSubItems(i).ForeColor = vbBlack
            Next
        End If
    End With
    If CInt(rs("PRIORITARIO")) = 1 Then
        lista.ListItems(linea).SmallIcon = 5
    End If
    lista_Click
    
    Set oEq = Nothing
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub
Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºEquipo", 850, lvwColumnLeft
        .Add , , "Estado", 700, lvwColumnCenter
        .Add , , "NºCliente", 1200, lvwColumnCenter
        .Add , , "Nombre Equipo", 2750, lvwColumnLeft
        .Add , , "NºSerie", 1400, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
        .Add , , "Prox. Calib.", 1050, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
        .Add , , "Prox. Verif.", 1050, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
        .Add , , "Prox. Mto.", 1050, lvwColumnCenter
        .Add , , "Localiz.", 1000, lvwColumnLeft
        .Add , , "Responsable", 1350, lvwColumnLeft
        .Add , , "AV", 0, lvwColumnCenter
        .Add , , "SPr", 0, lvwColumnCenter
        .Add , , "RPr", 0, lvwColumnCenter
        .Add , , "Calibrador", 1500, lvwColumnLeft
        .Add , , "Cliente", 0, lvwColumnLeft
        .Add , , "Centro", 1500, lvwColumnLeft
    End With
End Sub
'Private Function devolverTipoIconoEstado(ByVal id_eq As Long, ByVal fuera_uso As Boolean, ByRef rs As ADODB.Recordset, ByRef fecha As String) As String
'    fecha = ""
'    rs.Filter = "equipo_id = " & CStr(id_eq)
'
'    If rs.RecordCount = 0 Or fuera_uso Then
'        devolverTipoIconoEstado = "nada"
'        Exit Function
'    End If
'
'    rs.Sort = "fecha_prevista asc"
'    rs.MoveFirst
'
'    fecha = Format(rs!FECHA_PREVISTA, "dd/mm/yyyy")
'
'    Select Case CInt(rs!ESTADO)
'        Case 0
'            devolverTipoIconoEstado = "estado_pendiente"
'        Case 1
'            devolverTipoIconoEstado = "estado_preaviso"
'        Case 2
'            devolverTipoIconoEstado = "estado_ok"
'    End Select
'
'End Function


'Private Function devolverTipoIconoEstado_old(ByVal fecha_actual As Date, ByVal fecha_preaviso As Date, FECHA_PREVISTA As String) As String
'
'    Dim f_prev As Date
'
'    If fecha_preaviso = CDate("1900-01-01") Then
'        devolverTipoIconoEstado_old = "nada"
'        Exit Function
'    End If
'
'    If Not IsDate(FECHA_PREVISTA) Then
'        devolverTipoIconoEstado_old = "nada"
'        Exit Function
'    End If
'
'    f_prev = CDate(FECHA_PREVISTA)
'
'    If fecha_actual < fecha_preaviso Then
'        ' aun no ha llegado a la fecha de preaviso
'        devolverTipoIconoEstado_old = "estado_ok"
'        Exit Function
'    End If
'
'    If fecha_actual < f_prev Then
'        ' aun no ha llegado a la fecha de preaviso
'        devolverTipoIconoEstado_old = "estado_preaviso"
'        Exit Function
'    End If
'
'    If f_prev <= fecha_actual Then
'        ' aun no ha llegado a la fecha de preaviso
'        devolverTipoIconoEstado_old = "estado_pendiente"
'        Exit Function
'    End If
'
'    ' si no se cumple los supuestos anteriores, no devuelve nada para el icono
'    devolverTipoIconoEstado_old = "nada"
'
'
'End Function
'Private Function devolverEstadoCalVer(ByVal fecha_actual As Date, ByVal fecha_preaviso As Date, FECHA_PREVISTA As String) As Integer
'
'    Dim f_prev As Date
'
'    If fecha_preaviso = CDate("1900-01-01") Then
'        devolverEstadoCalVer = -1
'        Exit Function
'    End If
'
'    If Not IsDate(FECHA_PREVISTA) Then
'        devolverEstadoCalVer = -1
'        Exit Function
'    End If
'
'    f_prev = CDate(FECHA_PREVISTA)
'
'    If fecha_actual < fecha_preaviso Then
'        ' aun no ha llegado a la fecha de preaviso
'        devolverEstadoCalVer = 0
'        Exit Function
'    End If
'
'    If fecha_actual < f_prev Then
'        ' aun no ha llegado a la fecha de preaviso
'        devolverEstadoCalVer = 1
'        Exit Function
'    End If
'
'    If f_prev <= fecha_actual Then
'        ' aun no ha llegado a la fecha de preaviso
'        devolverEstadoCalVer = 2
'        Exit Function
'    End If
'
'    ' si no se cumple los supuestos anteriores, no devuelve nada para el icono
'    devolverEstadoCalVer = -1
'
'
'End Function

Private Sub txtDescripcion_Change()
    If txtDescripcion <> "" Then
        cargar_lista
    Else
        lista.ListItems.Clear
    End If
End Sub

Private Sub txtFabricante_Change()
    If txtFabricante <> "" Then
        cargar_lista
    Else
        lista.ListItems.Clear
    End If
End Sub

Private Sub txtNEqipo_Change()
    If txtNEqipo.Text <> "" Then
        cargar_lista
    Else
        lista.ListItems.Clear
    End If
End Sub
Private Sub txtNEqipo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_SoloNumerico(txtNEqipo, KeyAscii)
End Sub
Private Sub txtnombre_Change()
    If txtNombre <> "" Then
        cargar_lista
    Else
        lista.ListItems.Clear
    End If
End Sub

Private Sub txtNSerie_Change()
    If txtNSerie <> "" Then
        cargar_lista
    Else
        lista.ListItems.Clear
    End If
End Sub
'Private Function devolverFechaCorrecta(ByVal fecha As String) As String
'    devolverFechaCorrecta = ""
'
'    If Not IsDate(fecha) Then
'        Exit Function
'    End If
'
'    If CDate(fecha) = CDate("1900-01-01") Then
'        Exit Function
'    End If
'    devolverFechaCorrecta = Format(fecha, "dd/mm/yyyy")
'End Function
Private Sub cargar_estad_cal_ver(ByRef cmb As DataCombo)

    Dim rs As ADODB.Recordset
    Dim sql As String
    
    sql = " SELECT -1 AS ID_ESTADO, '' as descripcion "
    sql = sql & " union "
    sql = sql & " SELECT 0 AS ID_ESTADO, 'CORRECTO' as descripcion "
    sql = sql & " union "
    sql = sql & " SELECT 1 AS ID_ESTADO, 'EN PREAVISO' as descripcion "
    sql = sql & " union "
    sql = sql & " SELECT 2 AS ID_ESTADO, 'PENDIENTE' as descripcion "
    
    Set rs = datos_bd(sql)

    With cmb
        Set .RowSource = rs
        .ListField = rs(1).Name
        .BoundColumn = rs(0).Name
    End With

    Set rs = Nothing
    
End Sub

Private Sub txtNumCliente_Change()
    cargar_lista
End Sub
