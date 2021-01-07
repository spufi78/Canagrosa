VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmListadoDocPago 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de documentos de pago"
   ClientHeight    =   11685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15690
   Icon            =   "frmListadoDocPago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11685
   ScaleWidth      =   15690
   Begin VB.Frame frmDatosEspeciales 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Pedido a la factura seleccionada"
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
      Height          =   1335
      Left            =   2925
      TabIndex        =   102
      Top             =   5940
      Visible         =   0   'False
      Width           =   9135
      Begin pryCombo.miCombo cmdPedidosAsginar 
         Height          =   330
         Left            =   945
         TabIndex        =   103
         Top             =   540
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdAsignarPedido 
         Height          =   795
         Left            =   7605
         TabIndex        =   104
         Top             =   315
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "Informar Pedido"
         Appearance      =   5
         Picture         =   "frmListadoDocPago.frx":030A
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   135
         TabIndex        =   105
         Top             =   630
         Width           =   690
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "GENERANDO FIRMA ELECTRÓNICA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1680
      Left            =   4860
      TabIndex        =   95
      Top             =   6075
      Visible         =   0   'False
      Width           =   4920
      Begin VB.CheckBox chkDeterminaciones 
         Caption         =   "Por Determinaciones"
         Height          =   195
         Left            =   225
         TabIndex        =   100
         Top             =   1125
         Width           =   2220
      End
      Begin VB.CommandButton cmdAlbaranCrear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear Albaran"
         Height          =   510
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   990
         Width           =   1140
      End
      Begin VB.CommandButton cmdAlbaranCerrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   510
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Firma Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   270
         TabIndex        =   97
         Top             =   450
         Width           =   4380
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "CREAR ALBARAN DESDE FACTURA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   96
         Top             =   0
         Width           =   4915
      End
   End
   Begin VB.CommandButton cmdSal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   555
      Left            =   14130
      Picture         =   "frmListadoDocPago.frx":6B6C
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   11025
      Width           =   1470
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   7020
      TabIndex        =   66
      Top             =   9990
      Width           =   3480
      Begin VB.CommandButton cmdIberia 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Iberia"
         Enabled         =   0   'False
         Height          =   645
         Left            =   2160
         Picture         =   "frmListadoDocPago.frx":D3BE
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   900
         Width           =   1020
      End
      Begin VB.CommandButton cmdImprimir2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1125
         Picture         =   "frmListadoDocPago.frx":13C10
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   225
         Width           =   1020
      End
      Begin VB.CommandButton cmdCartaPago 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Carta Pago"
         Enabled         =   0   'False
         Height          =   645
         Left            =   90
         Picture         =   "frmListadoDocPago.frx":1A462
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   225
         Width           =   1020
      End
      Begin VB.CommandButton cmdmail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enviar Mail"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1125
         Picture         =   "frmListadoDocPago.frx":20CB4
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   900
         Width           =   1020
      End
      Begin VB.CheckBox chkFirma 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Firma Digital"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2205
         TabIndex        =   67
         Top             =   450
         Width           =   1230
      End
      Begin VB.CommandButton cmdListado2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Listado"
         Enabled         =   0   'False
         Height          =   645
         Left            =   90
         Picture         =   "frmListadoDocPago.frx":27506
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   900
         Width           =   1020
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Oculto"
      Height          =   870
      Left            =   6030
      TabIndex        =   58
      Top             =   6345
      Visible         =   0   'False
      Width           =   7125
      Begin VB.CommandButton cmdRecibo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recibo"
         Height          =   600
         Left            =   6030
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdFactorizadas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Factorizadas"
         Enabled         =   0   'False
         Height          =   600
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   180
         Width           =   1020
      End
      Begin VB.CheckBox chkfactorizar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factorizada en fecha"
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   60
         Top             =   315
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3690
         TabIndex        =   59
         Top             =   315
         Width           =   780
      End
      Begin MSComCtl2.DTPicker ffactorizada 
         Height          =   330
         Left            =   2025
         TabIndex        =   61
         Top             =   315
         Width           =   1545
         _ExtentX        =   2725
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
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   8421631
         Format          =   16515073
         CurrentDate     =   38002
      End
   End
   Begin VB.Frame frmFirma 
      Caption         =   "GENERANDO FIRMA ELECTRÓNICA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1365
      Left            =   4860
      TabIndex        =   55
      Top             =   4995
      Visible         =   0   'False
      Width           =   4920
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "GENERANDO FIRMA ELECTRÓNICA..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   4915
      End
      Begin VB.Label lblFirma2 
         Alignment       =   2  'Center
         Caption         =   "Firma Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   270
         TabIndex        =   57
         Top             =   900
         Width           =   4380
      End
      Begin VB.Label lblFirma 
         Alignment       =   2  'Center
         Caption         =   "Firma Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   270
         TabIndex        =   56
         Top             =   450
         Width           =   4380
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   10530
      TabIndex        =   45
      Top             =   9990
      Width           =   2310
      Begin VB.CommandButton cmdInformarPedido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "F5-Informar Pedido"
         Height          =   645
         Left            =   1125
         Picture         =   "frmListadoDocPago.frx":2DD58
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   900
         Width           =   1110
      End
      Begin VB.CommandButton cmdCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         Height          =   645
         Left            =   1125
         Picture         =   "frmListadoDocPago.frx":345AA
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton cmdVerPedido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Pedido"
         Height          =   645
         Left            =   90
         Picture         =   "frmListadoDocPago.frx":3ADFC
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   225
         Width           =   1020
      End
      Begin VB.CommandButton cmdPedidosCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ped. Cliente"
         Height          =   645
         Left            =   90
         Picture         =   "frmListadoDocPago.frx":4164E
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   900
         Width           =   1020
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   45
      TabIndex        =   17
      Top             =   9990
      Width           =   6960
      Begin VB.CommandButton cmdAim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos ADS"
         Height          =   645
         Left            =   5715
         Picture         =   "frmListadoDocPago.frx":47EA0
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   900
         Width           =   1110
      End
      Begin VB.CommandButton cmdDesglose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desglose Contable"
         Height          =   645
         Left            =   5715
         Picture         =   "frmListadoDocPago.frx":482F3
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton cmdAlbaran 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear Albaran"
         Height          =   645
         Left            =   4590
         Picture         =   "frmListadoDocPago.frx":4EB45
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   900
         Width           =   1110
      End
      Begin VB.CommandButton cmdDuplicar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Duplicar"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1215
         Picture         =   "frmListadoDocPago.frx":55397
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   900
         Width           =   1110
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   645
         Left            =   4590
         Picture         =   "frmListadoDocPago.frx":5BBE9
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton cmdDatos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Otros Datos"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1215
         Picture         =   "frmListadoDocPago.frx":6243B
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton cmdConceptos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Conptos  F. Muestras"
         Enabled         =   0   'False
         Height          =   645
         Left            =   3465
         Picture         =   "frmListadoDocPago.frx":68C8D
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   900
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdDescobrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descobrar"
         Enabled         =   0   'False
         Height          =   645
         Left            =   75
         Picture         =   "frmListadoDocPago.frx":6F4DF
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   900
         Width           =   1110
      End
      Begin VB.CommandButton cmdanular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anular"
         Enabled         =   0   'False
         Height          =   645
         Left            =   2340
         Picture         =   "frmListadoDocPago.frx":75D31
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   645
         Left            =   3465
         Picture         =   "frmListadoDocPago.frx":7C583
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   225
         Width           =   1110
      End
      Begin VB.CommandButton cmdAbono 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Crear Abono"
         Enabled         =   0   'False
         Height          =   645
         Left            =   2340
         Picture         =   "frmListadoDocPago.frx":82DD5
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   900
         Width           =   1110
      End
      Begin VB.CommandButton cmbCobrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos Cobro"
         Height          =   645
         Left            =   75
         Picture         =   "frmListadoDocPago.frx":89627
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   225
         Width           =   1110
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6300
      Left            =   45
      TabIndex        =   27
      Top             =   3645
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   11113
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
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
      Height          =   2940
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   15615
      Begin VB.CheckBox chkFPrevista 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   108
         Top             =   2115
         Width           =   285
      End
      Begin VB.TextBox txtConcepto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5400
         TabIndex        =   92
         Top             =   1395
         Width           =   3795
      End
      Begin VB.CheckBox chkIberia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iberia"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8820
         TabIndex        =   91
         Top             =   945
         Width           =   780
      End
      Begin VB.CheckBox chkFCobro 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   83
         Top             =   1755
         Width           =   285
      End
      Begin VB.CheckBox chkVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   53
         Top             =   1395
         Width           =   285
      End
      Begin VB.CheckBox chkAirbus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8820
         TabIndex        =   48
         Top             =   630
         Width           =   780
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enviada"
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
         Height          =   825
         Left            =   12240
         TabIndex        =   41
         Top             =   1350
         Width           =   1620
         Begin VB.OptionButton opEnviada 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   240
            Index           =   2
            Left            =   810
            TabIndex        =   44
            Top             =   270
            Width           =   555
         End
         Begin VB.OptionButton opEnviada 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   43
            Top             =   270
            Width           =   555
         End
         Begin VB.OptionButton opEnviada 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   540
            Value           =   -1  'True
            Width           =   825
         End
      End
      Begin VB.CheckBox chkSinPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo las que no tienen pedido asociado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9630
         TabIndex        =   39
         Top             =   2565
         Width           =   3705
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de documento"
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
         Height          =   2175
         Left            =   10125
         TabIndex        =   10
         Top             =   225
         Width           =   1890
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fact. y Albaran"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   90
            TabIndex        =   101
            Top             =   1620
            Width           =   1590
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Proforma"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   90
            TabIndex        =   90
            Top             =   1845
            Width           =   1590
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fact. y Abonos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   38
            Top             =   1395
            Width           =   1590
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fact. Abonadas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   15
            Top             =   1170
            Width           =   1680
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fact. Anuladas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   14
            Top             =   945
            Width           =   1590
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abono"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   13
            Top             =   705
            Width           =   1005
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Albaran"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   12
            Top             =   480
            Width           =   1050
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Factura"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4950
         TabIndex        =   5
         Top             =   1005
         Width           =   735
      End
      Begin VB.TextBox txtanno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1005
         Width           =   570
      End
      Begin VB.CheckBox chktodospedidos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8010
         TabIndex        =   32
         Top             =   2520
         Width           =   780
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8820
         TabIndex        =   23
         Top             =   315
         Width           =   780
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pendiente Cobro"
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
         Height          =   960
         Left            =   12240
         TabIndex        =   20
         Top             =   225
         Width           =   1620
         Begin VB.OptionButton opPendiente 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   24
            Top             =   630
            Width           =   825
         End
         Begin VB.OptionButton opPendiente 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   22
            Top             =   315
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton opPendiente 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   240
            Index           =   1
            Left            =   810
            TabIndex        =   21
            Top             =   315
            Width           =   555
         End
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1125
         TabIndex        =   3
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   2745
         TabIndex        =   4
         Top             =   990
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   345
         Left            =   6660
         TabIndex        =   34
         Top             =   990
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   609
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196664
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   855
         TabIndex        =   36
         Top             =   285
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   345
         Left            =   855
         TabIndex        =   40
         Top             =   2475
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesdev 
         Height          =   330
         Left            =   1125
         TabIndex        =   49
         Top             =   1350
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhastav 
         Height          =   330
         Left            =   2745
         TabIndex        =   50
         Top             =   1350
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbclientesFact 
         Height          =   345
         Left            =   855
         TabIndex        =   81
         Top             =   630
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fCobroDesde 
         Height          =   330
         Left            =   1125
         TabIndex        =   84
         Top             =   1710
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fCobroHasta 
         Height          =   330
         Left            =   2745
         TabIndex        =   85
         Top             =   1710
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   5400
         TabIndex        =   88
         Top             =   1755
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   1050
         Left            =   14265
         Picture         =   "frmListadoDocPago.frx":8FE79
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker fprevistadesde 
         Height          =   330
         Left            =   1125
         TabIndex        =   109
         Top             =   2070
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fprevistahasta 
         Height          =   330
         Left            =   2745
         TabIndex        =   110
         Top             =   2070
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   11
         Left            =   2535
         TabIndex        =   112
         Top             =   2115
         Width           =   105
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Prevista"
         Height          =   195
         Index           =   10
         Left            =   375
         TabIndex        =   111
         Top             =   2145
         Width           =   705
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   195
         Index           =   9
         Left            =   4305
         TabIndex        =   93
         Top             =   1455
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   16
         Left            =   4320
         TabIndex        =   89
         Top             =   1815
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Cobro"
         Height          =   195
         Index           =   8
         Left            =   375
         TabIndex        =   87
         Top             =   1785
         Width           =   555
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   7
         Left            =   2535
         TabIndex        =   86
         Top             =   1755
         Width           =   105
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clien. Fact."
         Height          =   195
         Index           =   6
         Left            =   45
         TabIndex        =   82
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   5
         Left            =   2535
         TabIndex        =   52
         Top             =   1395
         Width           =   105
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Vencim."
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   51
         Top             =   1425
         Width           =   705
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   1
         Left            =   5760
         TabIndex        =   35
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   3
         Left            =   4320
         TabIndex        =   33
         Top             =   1065
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   31
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Factura"
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   8
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   2
         Left            =   2535
         TabIndex        =   7
         Top             =   1035
         Width           =   105
      End
   End
   Begin XtremeSuiteControls.PushButton cmdMarcar 
      Height          =   300
      Left            =   45
      TabIndex        =   64
      Top             =   3330
      Width           =   1500
      _Version        =   851970
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Marcar Todas"
      Appearance      =   5
      Picture         =   "frmListadoDocPago.frx":90743
   End
   Begin XtremeSuiteControls.PushButton cmdDesmarcar 
      Height          =   300
      Left            =   1575
      TabIndex        =   65
      Top             =   3330
      Width           =   1815
      _Version        =   851970
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Desmarcar Todas"
      Appearance      =   5
      Picture         =   "frmListadoDocPago.frx":96FA5
   End
   Begin VB.Label lblIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   13590
      TabIndex        =   78
      Top             =   10350
      Width           =   1980
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IVA"
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
      Height          =   300
      Index           =   2
      Left            =   13050
      TabIndex        =   77
      Top             =   10350
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Base"
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
      Height          =   300
      Index           =   0
      Left            =   13050
      TabIndex        =   76
      Top             =   10035
      Width           =   510
   End
   Begin VB.Label lblBase 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   13590
      TabIndex        =   75
      Top             =   10035
      Width           =   1980
   End
   Begin VB.Label lblrestan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
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
      Height          =   300
      Index           =   1
      Left            =   13050
      TabIndex        =   74
      Top             =   10665
      Width           =   510
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   13590
      TabIndex        =   73
      Top             =   10665
      Width           =   1980
   End
   Begin VB.Shape Shape1 
      Height          =   1005
      Left            =   13005
      Top             =   9990
      Width           =   2625
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   16
      Top             =   3330
      Width           =   15630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de documentos de Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   16080
   End
End
Attribute VB_Name = "frmListadoDocPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum COLS
    ID_DOC = 9
    COL_FECHA_PREVISTA_COBRO = 20
    COL_FECHA_COBRO = 21
    COL_CCC = 22
    ID_PEDIDO = 23
    ID_CLIENTE_FACTURA = 24
    COL_CIF = 25
End Enum

Private Sub chkFCobro_Click()
    If chkFCobro.Value = Checked Then
        fCobroDesde.Enabled = True
        fCobroHasta.Enabled = True
    Else
        fCobroDesde.Enabled = False
        fCobroHasta.Enabled = False
    End If
End Sub

Private Sub chkFPrevista_Click()
    fPrevistaDesde.Enabled = chkFPrevista.Value
    fPrevistaHasta.Enabled = chkFPrevista.Value
End Sub

Private Sub chkVencimiento_Click()
    If chkVencimiento.Value = Checked Then
        fdesdev.Enabled = True
        fhastav.Enabled = True
    Else
        fdesdev.Enabled = False
        fhastav.Enabled = False
    End If
End Sub

Private Sub cmbclientesFact_change()
    If cmbclientesFact.getPK_SALIDA <> 0 Then
        pedidos (cmbclientesFact.getPK_SALIDA)
    End If
End Sub

Private Sub cmdAIM_Click()
    If lista.ListItems.Count > 0 Then
        frmAirbus_ListadoMuestras.ID_FACTURA = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmAirbus_ListadoMuestras.Show
    End If
End Sub

Private Sub cmdAlbaran_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(12)) >= 2 Then
        Label4.Caption = "Factura Nº" & lista.ListItems(lista.selectedItem.Index)
        Frame9.visible = True
    Else
        MsgBox "Solo se pueden marcar facturas.", vbExclamation, App.Title
    End If
End Sub

Private Sub cmdAlbaranCerrar_Click()
    Frame9.visible = False
End Sub

Private Sub cmdAlbaranCrear_Click()
            Dim oDoc As New clsDocs_pago
            Dim oAlbaran As New clsDocs_pago
            Dim omuestras As New clsDocs_pago_muestras
            Dim oConceptos As New clsDocs_pago_conceptos
            Dim rs As New ADODB.Recordset
            Dim i As Long
            Dim num_doc As Long
            Dim idDoc As Long
   On Error GoTo cmdAlbaranCrear_Click_Error
    Me.MousePointer = 11
            idDoc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
'            For i = CLng(Text1(3)) To CLng(Text1(2))
                num_doc = 0
                oDoc.CargarDocumento idDoc
                With oAlbaran
                    .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
                    .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                    .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                    .setEMPLEADO_ID = oDoc.getEMPLEADO_ID
                    .setCLIENTE_ID = oDoc.getCLIENTE_ID
                    .setCLIENTE_ID_FACTURA = oDoc.getCLIENTE_ID_FACTURA
                    .setTOTAL = moneda_bd(oDoc.getTOTAL)
                    .setDESCUENTO = Replace(oDoc.getDESCUENTO, ",", ".")
'                    .setIVA = 0
                    .setANULADO = 0
                    .setFP_ID = oDoc.getFP_ID
                    .setPEDIDO_ID = oDoc.getPEDIDO_ID
                    .setFACTURA_CONCEPTOS = oDoc.getFACTURA_CONCEPTOS
                    .setPAGADO = oDoc.getID_DOC
                    ' Insertamos el documento de pago
                    num_doc = .InsertarDocPago
                    If num_doc = 0 Then
                         MsgBox "Error al insertar el albaran.", vbExclamation, App.Title
                    End If
                End With
                ' Insertamos el detalle de la factura de conceptos
                Set rs = oConceptos.ConceptosDocumento(idDoc)
                If rs.RecordCount > 0 Then
                    Do
                        With oConceptos
                            .setDOC_ID = num_doc
                            .setDESCRIPCION = rs("DESCRIPCION")
                            .setFECHA = Format(rs("FECHA"), "yyyy-mm-dd")
                            .setPRECIO = Replace(Format(rs("precio"), "0.00"), ",", ".")
                            .setCANTIDAD = rs("CANTIDAD")
                            .setAPARTADO = rs("APARTADO")
                            .setSUBTOTAL = Replace(Format(rs("subtotal"), "0.00"), ",", ".")
                            .setTOTAL = Replace(Format(rs("total"), "0.00"), ",", ".")
                            .setDTO = Replace(Format(rs("dto"), "0.00"), ",", ".")
                            .setFAMILIA_ID = rs("familia_id")
                            If .Insertar = False Then
                                Exit Sub
                            End If
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
                ' Insertamos el detalle de la factura de muestras
                Set rs = omuestras.MuestrasDocumento(idDoc)
                If rs.RecordCount > 0 Then
                    Do
                        With omuestras
                            .setDOC_ID = num_doc
                            .setMUESTRA_ID = rs(6)
'                            .setORDEN = RS(8)
                            .setORDEN = .CalcularOrden(num_doc)

                            .setCODIGO = rs(9)
                            .setFECHA = Format(rs(2), "yyyy-mm-dd")
                            .setTIPO_ANALISIS = rs(3)
                            .setREFERENCIA_CLIENTE = rs(4)
    '                        .setPRECIO = rs(5)
                            .setPRECIO = Replace(Format(rs(5), "0.00"), ",", ".")
                            If .Insertar_doc_pago_muestra(chkDeterminaciones.Value) = -1 Then
                                MsgBox "Error al insertar en doc_pago_muestra", vbCritical, App.Title
                                Exit Sub
                            End If
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
'            Next
            Set oMuestra = Nothing
                Me.MousePointer = 0

            MsgBox "Albaran creado correctamente.", vbInformation, App.Title
            Frame9.visible = False
'        End If

   On Error GoTo 0
   Exit Sub

cmdAlbaranCrear_Click_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAlbaranCrear_Click of Formulario frmListadoDocPago"

End Sub

Private Sub cmdAsignarPedido_Click()
    Dim oDP As New clsDocs_pago
    If cmdPedidosAsginar.getTEXTO = "" Then
        oDP.informar_pedido CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_DOC)), 0
    Else
        oDP.informar_pedido CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_DOC)), cmdPedidosAsginar.getPK_SALIDA
    End If
    Set oDP = Nothing
    frmDatosEspeciales.visible = False
    actualizar_lista CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_DOC)), lista.selectedItem.Index
    gdoc = 0
    
End Sub

Private Sub cmdCliente_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDoc As New clsDocs_pago
    oDoc.CargarDocumento lista.ListItems(lista.selectedItem.Index).SubItems(9)
    frmClientes.PK = oDoc.getCLIENTE_ID
    frmClientes.Show 1
End Sub

Private Sub cmdDesglose_Click()
    If lista.ListItems.Count > 0 Then
        frmFacturacion_Desglose.PK = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmFacturacion_Desglose.Show 1
    End If
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    ' 10 : Factura solo por conceptos
    ' 12 : Tipo de documento
'    If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(10)) <> 1 Or CInt(lista.ListItems(lista.selectedItem.Index).SubItems(12)) > 2 Then
    If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(10)) <> 1 Then
        MsgBox "Solo se pueden duplicar las facturas por conceptos.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("Va a duplicar la factura por conceptos. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oDoc As New clsDocs_pago
        Dim documento As Long
        If oDoc.CargarDocumento(CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9))) = True Then
            oDoc.setOBSERVACIONES = ""
            oDoc.setCOMENTARIO = ""
            oDoc.setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
            oDoc.setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
            oDoc.setTOTAL = moneda_bd(oDoc.getTOTAL)
            oDoc.setPAGADO = 0
            documento = oDoc.InsertarDocPago
        End If
        If documento = 0 Then
            MsgBox "Error al insertar el documento duplicado.", vbExclamation, App.Title
            Exit Sub
        Else
        Dim odd As New clsDocs_pago_conceptos
        Dim rs As ADODB.Recordset
        Set rs = odd.ConceptosDocumento(CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)))
        If rs.RecordCount > 0 Then
            Do
                odd.setDOC_ID = documento
                odd.setALBARAN_ID = 0
                odd.setABONADO = 0
                odd.setDESCRIPCION = rs("descripcion")
                odd.setFECHA = Format(rs("fecha"), "yyyy-mm-dd")
                odd.setCANTIDAD = rs("cantidad")
                odd.setPRECIO = moneda_bd(rs("precio"))
                odd.setSUBTOTAL = moneda_bd(rs("subtotal"))
                odd.setDTO = moneda_bd(rs("dto"))
                odd.setTOTAL = moneda_bd(rs("total"))
                odd.setFAMILIA_ID = rs("familia_id")
                odd.setAPARTADO = rs("apartado")
                odd.Insertar
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        MsgBox "Factura duplicada correctamente.", vbInformation, App.Title
        gdoc = documento
        frmFacturaConceptos.Show 1
        gdoc = 0
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDuplicar_Click of Formulario frmListadoDocPago"
End Sub

Private Sub cmdEditar_Click()
   On Error GoTo cmdEditar_Click_Error

    If lista.ListItems.Count > 0 Then
     If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(10)) <> 1 Then
        frmDocumento_Edicion.PK_DOCUMENTO = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmDocumento_Edicion.Show 1
        actualizar_lista CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)), lista.selectedItem.Index
     Else
        MsgBox "Sólo se editan facturas con muestras. Para facturas de conceptos, pulse conceptos.", vbInformation, App.Title
     End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEditar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEditar_Click of Formulario frmListadoDocPago"
End Sub
'Private Sub cmdFactorizadas_Click()
'    Dim oDoc As New clsDocs_pago
'
'    If chkTodos.value = vbChecked Then
'        Call oDoc.ImprimirListadoFactorizada_TodosClientes(Format(fdesde, "yyyy-mm-dd"), Format(fhasta, "yyyy-mm-dd"), TIPO_DOCUMENTO, pendiente_cobro, Estado_documento, chkSinPedido.value)
'    Else
'        Call oDoc.ImprimirListadoFactorizada_ClienteUnico(cmbclientes.getPK_SALIDA, Format(fdesde, "yyyy-mm-dd"), Format(fhasta, "yyyy-mm-dd"), TIPO_DOCUMENTO, pendiente_cobro, Estado_documento, chkSinPedido.value)
'    End If
'
'    Set oDoc = Nothing
'End Sub

Private Sub cmdIberia_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim cont As Integer
    Dim oDoc_pago As New clsDocs_pago
    If contar_marcados = 0 Then
      If oDoc_pago.validar_previos_documento(lista.ListItems(lista.selectedItem.Index).SubItems(9)) Then
        oDoc_pago.generar_factura lista.ListItems(lista.selectedItem.Index).SubItems(9), False, "", "rptFacturaIberia"
      End If
    Else
      Dim i As Integer
      For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
          If oDoc_pago.validar_previos_documento(lista.ListItems(i).SubItems(9)) Then
              oDoc_pago.generar_factura lista.ListItems(i).SubItems(9), True, "", "rptFacturaIberia"
          End If
        End If
      Next
    End If

End Sub

Private Sub cmdImprimir2_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim cont As Integer
    Dim oDoc_pago As New clsDocs_pago
    If contar_marcados = 0 Then
      If oDoc_pago.validar_previos_documento(lista.ListItems(lista.selectedItem.Index).SubItems(9)) Then
        oDoc_pago.generar_factura lista.ListItems(lista.selectedItem.Index).SubItems(9), False, "", "rptFactura"
      End If
    Else
      Dim i As Integer
      For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
          If oDoc_pago.validar_previos_documento(lista.ListItems(i).SubItems(9)) Then
              oDoc_pago.generar_factura lista.ListItems(i).SubItems(9), True, "", "rptFactura"
          End If
        End If
      Next
    End If
'    Set oDoc_pago = Nothing
End Sub

Private Sub cmdInformarPedido_Click()
    frmDatosEspeciales.visible = Not frmDatosEspeciales.visible
'    If lista.ListItems.Count = 0 Then
'       frmDatosEspeciales.top = Me.Height / 2 - frmDatosEspeciales.Height
'    Else
'        frmDatosEspeciales.top = lista.ListItems(lista.selectedItem.Index).top
'    End If
End Sub

Private Sub cmdPedidosCliente_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDoc As New clsDocs_pago
    oDoc.CargarDocumento lista.ListItems(lista.selectedItem.Index).SubItems(9)
    frmClientes_Pedidos.PK = oDoc.getCLIENTE_ID
    frmClientes_Pedidos.Show 1
End Sub

'Private Sub cmdRecibo_Click()
'    If lista.ListItems.Count > 0 Then
'        If contar_marcados = 0 Then
'            MsgBox "Marque el/los documentos para los que desea generar el recibo.", vbExclamation, App.Title
'            Exit Sub
'        End If
'        Dim i As Integer
'        ' Verificar si todos son del mismo cliente
'        Dim distintos As Boolean
'        Dim cliente As String
'        Dim ID_DOCUMENTO As Long
'        Dim NUMERO_DOCUMENTOS As String
'        Dim IMPORTE As Currency
'        distintos = False
'        For i = 1 To lista.ListItems.Count
'            If lista.ListItems(i).Checked = True Then
'                If lista.ListItems(i).SubItems(1) <> cliente Or distintos Then
'                    If cliente = "" Then
'                        cliente = lista.ListItems(i).SubItems(1)
'                        ID_DOCUMENTO = lista.ListItems(i).SubItems(9)
'                        NUMERO_DOCUMENTOS = NUMERO_DOCUMENTOS & lista.ListItems(i).Text & ","
'                        IMPORTE = IMPORTE + CCur(lista.ListItems(i).SubItems(8))
'                    Else
'                        distintos = True
'                    End If
'                Else
'                    NUMERO_DOCUMENTOS = NUMERO_DOCUMENTOS & lista.ListItems(i).Text & ","
'                    IMPORTE = IMPORTE + CCur(lista.ListItems(i).SubItems(8))
'                End If
'            End If
'        Next
'        If distintos Then
'            MsgBox "Sólo puede marcar documentos del mismo cliente.", vbExclamation, App.Title
'            Exit Sub
'        Else
'            NUMERO_DOCUMENTOS = Left(NUMERO_DOCUMENTOS, Len(NUMERO_DOCUMENTOS) - 1)
'        End If
'        Dim consulta As String
'        Dim tNum2Text As New cNum2Text
'        Dim oDOCUMENTO As New clsDocs_pago
'        oDOCUMENTO.CargarDocumento ID_DOCUMENTO
'        Dim oCliente As New clsCliente
'        Dim oMunicipio As New clsMunicipios
'        Dim oprovincia As New clsProvincias
'        oCliente.CargaCliente (oDOCUMENTO.getCLIENTE_ID)
'        oMunicipio.CargarMunicipio oCliente.getMUNICIPIO_ID
'        oprovincia.CargarProvincia oCliente.getPROVINCIA_ID
'        Dim PAGO As String
'        PAGO = "CCC : " & oCliente.getCUENTA & " BANCO : " & oCliente.getBANCO
'        consulta = "SELECT '" & NUMERO_DOCUMENTOS & "' AS NUMERO,'" & _
'                               Format(Date, "dd-mm-yyyy") & "' AS FECHA,'" & _
'                               Format(IMPORTE, "currency") & "' AS IMPORTE,'" & _
'                               cliente & "' AS NOMBRE_CLIENTE,'" & _
'                               Format(Date, "dd-mm-yyyy") & "' AS VENCIMIENTO,'" & _
'                               UCase(tNum2Text.Numero2Letra(IMPORTE, , 2, "euro", "céntimo", Masculino, Masculino)) & "' AS IMPORTE_LETRAS,' " & _
'                               PAGO & "' AS PAGO,'" & _
'                               oCliente.getDIRECCION & "' AS DIRECCION,'" & _
'                               oCliente.getCOD_POSTAL & "' AS CP,'" & _
'                               oCliente.getCIF & "' AS CIF,'" & _
'                               oMunicipio.getNOMBRE & "' AS MUNICIPIO,'" & _
'                               oprovincia.getNOMBRE & "' AS PROVINCIA"
'        log (consulta)
'        frmReport.iniciar
'        frmReport.informe = "Facturacion\rptrecibo"
'        frmReport.consulta = consulta
'        frmReport.imprimir = False
'        frmReport.PDF = ""
'        frmReport.generar
'        frmReport.Visible = True
'    End If
'
'End Sub

'Private Sub chkfactorizar_Click()
'    Dim oDoc_pago As New clsDocs_pago
'    If chkfactorizar.value = Checked Then
'        ffactorizada.Enabled = True
'        cmdCambiar.Enabled = True
'        oDoc_pago.FACTORIZAR lista.ListItems(lista.selectedItem.Index).SubItems(9), Format(ffactorizada.value, "dd-mm-yyyy")
'        lista.ListItems(lista.selectedItem.Index).SubItems(13) = Format(ffactorizada.value, "dd-mm-yyyy")
'    Else
'        ffactorizada.Enabled = False
'        cmdCambiar.Enabled = False
'        oDoc_pago.FACTORIZAR lista.ListItems(lista.selectedItem.Index).SubItems(9), ""
'        lista.ListItems(lista.selectedItem.Index).SubItems(13) = ""
'    End If
'End Sub

Private Sub chktodospedidos_Click()
    cmbPedidos.limpiar
    If chktodospedidos.Value = Checked Then
        cmbPedidos.desactivar
    Else
        cmbPedidos.activar
    End If
End Sub

Private Sub cmbClientes_change()
'    If cmbclientes.getPK_SALIDA <> 0 Then
'        pedidos (cmbclientes.getPK_SALIDA)
'    End If
End Sub
Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbClientes.limpiar
        cmbClientes.desactivar
        chkAirbus.Enabled = True
        chkIberia.Enabled = True
        pedidos (0)
    Else
        cmbClientes.activar
        chkAirbus.Enabled = False
        chkIberia.Enabled = False
    End If
End Sub

Private Sub cmbCobrar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
'    Dim cont As Integer
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Checked = True Then
'            cont = cont + 1
'        End If
'    Next
'    If cont = 0 Then
'        MsgBox "Debe marcar las facturas a las que informara los datos del cobro.", vbExclamation, App.Title
'        Exit Sub
'    End If
    frmFacturacion_Envio.PK = lista.ListItems(lista.selectedItem.Index).SubItems(9)
    frmFacturacion_Envio.Show 1
    cmdBuscar_Click
End Sub

Private Sub cmdAbono_Click()
    If lista.ListItems.Count > 0 Then
        frmFacturaAbonar.PK = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmFacturaAbonar.Show 1
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdAnular_Click()
    Dim oDoc As New clsDocs_pago
      If contabilizado(lista.ListItems(lista.selectedItem.Index).SubItems(9)) Then
         Exit Sub
      End If
      If MsgBox("Va a anular el documento, ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        frmMotivo.Show 1
        If Trim(MOTIVO) = "" Then
            MsgBox "Para anular el documento es necesario introducir un motivo.", vbInformation, App.Title
            Exit Sub
        Else
            MOTIVO = "FACTURA ANULADA. " & MOTIVO
        End If
        ' Anulamos el docmuento de pago ANULADO = 1 (docs_pago)
        If oDoc.Anular(CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)), MOTIVO) Then
            ' Modificamos el documento de pago de las muestras
            Dim odocm As New clsDocs_pago_muestras
            Dim oMuestra As New clsMuestra
            Dim rs As New ADODB.Recordset
            Set rs = odocm.MuestrasDocumento(CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)))
            If rs.RecordCount <> 0 Then
                Do
                    oMuestra.Informar_Documento_Pago rs("muestra_id"), 0
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' En el caso de probetas, recuperar las muestras y las probetas
            Dim c As String
            Set rs = datos_bd("select distinct muestra_id from ce_resultados where doc_id = " & CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)))
            If rs.RecordCount <> 0 Then
                Do
                    oMuestra.Informar_Documento_Pago rs("muestra_id"), 0
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Limpiar doc_id probetas
            execute_bd "update ce_resultados set doc_id = 0 where doc_id = " & CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9))
            
            Set oMuestra = Nothing
            Set odocm = Nothing
            cmdBuscar_Click
        End If
      End If
    Set oDoc = Nothing
End Sub

Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub buscar()
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
    
    Dim T_BASE As Currency
    Dim T_IVA As Currency
    Dim T_TOTAL As Currency
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim oDoc As New clsDocs_pago
    Me.MousePointer = 11
    If txtNumero <> "" And txtanno <> "" Then
        Set rs = oDoc.Documento_por_numero(txtNumero, txtanno, TIPO_DOCUMENTO, pendiente_cobro, Estado_documento)
'        txtNumero = ""
    Else
        If chkTodos.Value = Unchecked And cmbClientes.getPK_SALIDA = 0 Then
            MsgBox "Seleccione un cliente.", vbInformation, App.Title
            Me.MousePointer = 0
            Exit Sub
        Else
            Dim cliente As Long
            Dim clienteFact As Long
            If cmbClientes.getTEXTO <> "" Then
                cliente = cmbClientes.getPK_SALIDA
            End If
            If cmbclientesFact.getTEXTO <> "" Then
                clienteFact = cmbclientesFact.getPK_SALIDA
            End If
            Set rs = oDoc.Documentos(cliente, clienteFact, Format(fdesde, "yyyy-mm-dd"), Format(fhasta, "yyyy-mm-dd"), TIPO_DOCUMENTO, pendiente_cobro, Estado_documento, Estado_Enviada, chkAirbus.Value, chkVencimiento.Value, fdesdev.Value, fhastav.Value, chkFCobro, fCobroDesde, fCobroHasta, chkFPrevista.Value, fPrevistaDesde, fPrevistaHasta, IIf(cmbFP.Text = "", 0, cmbFP.BoundText), chkIberia.Value, txtConcepto.Text)
        End If
    End If
    
    desactivar_controles
    If rs.RecordCount <> 0 Then
        cmdListado2.Enabled = True
        cmdFactorizadas.Enabled = True
        formatea_titulo
        Dim NUMERO As String
        Do
            If cmbPedidos.getTEXTO = "" Or (cmbPedidos.getTEXTO <> "" And cmbPedidos.getPK_SALIDA = rs(10)) Then
                Select Case rs(6)
                Case 1
                    NUMERO = "A-" & Format(rs(1), "0000")
                Case 2
                    NUMERO = "F-" & Format(rs(1), "0000")
                Case 3
                    NUMERO = "B-" & Format(rs(1), "0000")
                Case Else
                    NUMERO = Format(rs(1), "0000")
                End Select
                If Left(rs(11), 1) = "-" Then
                    NUMERO = NUMERO & rs(11)
                End If
                If chkSinPedido.Value = Unchecked Or _
                   (chkSinPedido.Value = Checked And rs(10) = 0) Then
                    With lista.ListItems.Add(, , NUMERO)
                        .SubItems(1) = rs.Fields(2)
                        .SubItems(2) = rs.Fields(3)
                        .SubItems(9) = rs.Fields(0)
                        IMPORTE = Format(rs(8), "currency")
                        If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
                            BASE = Format(IMPORTE, "0.00")
                        Else
                            BASE = Format(IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100), "0.00")
                        End If
                        IVA = Format((BASE * rs.Fields("iva")) / 100, "0.00")
                        .SubItems(3) = Format(IMPORTE, "currency")
                        .SubItems(4) = Format(rs.Fields("descuento"), "Standard")
                        .SubItems(5) = Format(BASE, "currency")
                        .SubItems(6) = rs.Fields("iva")
                        .SubItems(7) = Format(IVA, "currency")
                        .SubItems(8) = Format(BASE + IVA, "currency")
                        .SubItems(10) = rs(9)
                        .SubItems(11) = rs(7)
                        .SubItems(12) = rs(6)
                        .SubItems(13) = rs(11) ' Factorizada
                        .SubItems(14) = rs(12) ' PEdido
                        .SubItems(15) = rs(13) ' Forma de pago
                        .SubItems(16) = rs(16) ' Asiento
                        If rs(13) <> 0 Then
                            .SubItems(17) = Format(CDate(rs(3)) + rs(17), "dd/mm/yyyy") ' F.Vencimiento
                        Else
                            .SubItems(17) = ""
                        End If
                        .SubItems(18) = rs(18)
                        .SubItems(19) = rs(19)
                        If IsNull(rs(23)) Then
                            .SubItems(COLS.COL_FECHA_PREVISTA_COBRO) = ""
                        Else
                            .SubItems(COLS.COL_FECHA_PREVISTA_COBRO) = rs(23)
                        End If
                        If IsNull(rs(20)) Then
                            .SubItems(COLS.COL_FECHA_COBRO) = ""
                        Else
                            .SubItems(COLS.COL_FECHA_COBRO) = rs(20)
                        End If
                        If Not IsNull(rs(21)) Then
                            .SubItems(COLS.COL_CCC) = rs(21)  ' CCC
                        End If
                        T_BASE = T_BASE + CCur(.SubItems(5))
                        T_IVA = T_IVA + CCur(.SubItems(7))
                        T_TOTAL = T_TOTAL + CCur(.SubItems(8))
                        ' 14 CLIENTE_ID, 15 CLIENTE_ID_FACTURA
                        If rs(14) <> rs(15) Then
                            colorear lista.ListItems.Count, vbRed
                        End If
                        .SubItems(COLS.ID_PEDIDO) = rs(10)
                        .SubItems(COLS.ID_CLIENTE_FACTURA) = rs(15)
                        .SubItems(COLS.COL_CIF) = rs(22) ' CIF
                    End With
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
        lblBase = moneda(CStr(T_BASE))
        lblIVA = moneda(CStr(T_IVA))
        lbltotal = moneda(CStr(T_TOTAL))
        lista_Click
        
    Else
        lblMsg = "No existen registros con estos criterios."
    End If
    Me.MousePointer = 0
    Set oDoc = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar los documentos del cliente.", vbCritical, Err.Description
End Sub

'Private Sub cmdCambiar_Click()
'    Dim oDoc_pago As New clsDocs_pago
'    oDoc_pago.FACTORIZAR lista.ListItems(lista.selectedItem.Index).SubItems(9), Format(ffactorizada.value, "dd-mm-yyyy")
'    lista.ListItems(lista.selectedItem.Index).SubItems(13) = Format(ffactorizada.value, "dd-mm-yyyy")
'End Sub

Private Sub cmdCartaPago_Click()
    Dim s As String
    Dim i As Integer
    Dim j As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            s = s & lista.ListItems(i).SubItems(9) & ","
        End If
    Next
    If s <> "" Then
        s = Left(s, Len(s) - 1)
        Dim oDoc As New clsDocs_pago
        If MsgBox("¿Va a generar un total de " & oDoc.Numero_Cartas_Pago(s) & " carta/s de pago. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim rs As ADODB.Recordset
            Dim rs_fact As ADODB.Recordset
            Dim oCliente As New clsCliente
            Dim total As Currency
            Dim oMunicipio As New clsMunicipios
            Dim oProvincia As New clsProvincias
            Set rs = oDoc.Clientes_Distintos(s)
            If rs.RecordCount > 0 Then
                i = 1
                Do
                    Dim appword As Word.Application
                    Dim docword As Word.Document
                    ' Crear copia para su uso
                    Set appword = CreateObject("word.application")
                    Set docword = appword.Documents.Open(copiar_plantilla("CARTA_PAGO", CLng(i), 1))
                    ' Datos del cliente
                    oCliente.CargaCliente (rs(0))
                    oMunicipio.CargarMunicipio (oCliente.getMUNICIPIO_ID)
                    oProvincia.CargarProvincia oCliente.getPROVINCIA_ID
                    With docword.Sections(1).Headers(1).Range.Tables(2)
                        .Rows(1).Cells(2).Range.Text = oCliente.getNOMBRE
                        .Rows(2).Cells(2).Range.Text = oCliente.getDIRECCION
                        .Rows(3).Cells(2).Range.Text = oCliente.getCOD_POSTAL & " " & oMunicipio.getNOMBRE
                        .Rows(4).Cells(2).Range.Text = oProvincia.getNOMBRE
                    End With
                    docword.Tables(1).Rows(1).Cells(1).Range.InsertAfter fecha_larga(Date)
                    ' Facturas
                    Set rs_fact = oDoc.Facturas_Pendientes_Cliente(s, rs(0))
                    If rs_fact.RecordCount > 0 Then
                        j = 2
                        total = 0
                        Do
                            With docword.Tables(3)
                                .Rows(j).Cells(1).Range.Text = rs_fact(1)
                                .Rows(j).Cells(2).Range.Text = rs_fact(2)
                                oDoc.CargarDocumento (rs_fact(0))
                                'cIVA
                                Dim totalfactura As Currency
                                totalfactura = Replace(oDoc.getTOTAL, ".", ",") + ((Replace(oDoc.getTOTAL, ".", ",") * oDoc.getIVA) / 100)
                                .Rows(j).Cells(3).Range.Text = moneda(CStr(totalfactura))
                                total = total + totalfactura
                            End With
                            rs_fact.MoveNext
                            If rs_fact.EOF = False Then
                                docword.Tables(3).Rows.Add
                                j = j + 1
                            End If
                        Loop Until rs_fact.EOF
                    End If
                    ' Total
                    docword.Tables(2).Rows(1).Cells(1).Range.InsertAfter Format(total, "Currency")
                    docword.Save
                    appword.visible = True
                    Set docword = Nothing
                    Set appword = Nothing
                    rs.MoveNext
                    i = i + 1
                Loop Until rs.EOF
            End If
        End If
    Else
        MsgBox "Marque algún documento para generar la carta de pago.", vbInformation, App.Title
    End If
'    If MsgBox("Va a cobrar los documentos marcados, ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'        For i = 1 To lista.ListItems.Count
'            If lista.ListItems(i).Checked = True Then
'                oDoc.Cobrar CLng(lista.ListItems(i).SubItems(9))
'            End If
'        Next
'        cmdBuscar_Click
'    End If
'    Set oDoc = Nothing
End Sub

Private Sub cmdConceptos_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
'    If contabilizado(lista.ListItems(lista.SelectedItem.Index).SubItems(9)) Then
'       Exit Sub
'    End If
    gdoc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
    If lista.ListItems(lista.selectedItem.Index).SubItems(10) <> 1 Then
'        Nuevo
'        If UCase(USUARIO.getUSUARIO) = "JULIO" Then
'            frmDocumento_Conceptos.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(9)
'            frmDocumento_Conceptos.Show 1
'        Else
            frmConceptosFactura.Show 1
'        End If
        actualizar_lista CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)), lista.selectedItem.Index
        gdoc = 0
    End If
End Sub

Private Sub cmddatos_Click()
    If lista.ListItems.Count > 0 Then
'        gdoc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
'        frmFacturacion_Cobro.Show 1
        frmFacturacion_Envio.PK = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmFacturacion_Envio.Show 1
        actualizar_lista CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)), lista.selectedItem.Index
    End If
End Sub

Private Sub cmdDescobrar_Click()
    Dim oDoc As New clsDocs_pago
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a DESCOBRAR el documento SELECCIONADO, ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'        For i = 1 To lista.ListItems.Count
'            If lista.ListItems(i).Checked = True Then
                oDoc.DesCobrar CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9))
'            End If
'        Next
        cmdBuscar_Click
    End If
    Set oDoc = Nothing
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdListado2_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    On Error GoTo fallo
        
    If MsgBox("¿Desea exportar a excel?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        generar_excel_listado
    Else
    
        ' Nuevo listado en Crystal Report
        Dim cliente As Long
        Dim clienteFact As Long
        If cmbClientes.getTEXTO <> "" Then
            cliente = cmbClientes.getPK_SALIDA
        End If
        If cmbclientesFact.getTEXTO <> "" Then
            clienteFact = cmbclientesFact.getPK_SALIDA
        End If
        Dim oDoc As New clsDocs_pago
        oDoc.ImprimirListado cliente, clienteFact, Format(fdesde, "yyyy-mm-dd"), Format(fhasta, "yyyy-mm-dd"), TIPO_DOCUMENTO, pendiente_cobro, Estado_documento, Estado_Enviada, chkAirbus.Value, chkVencimiento.Value, fdesdev.Value, fhastav.Value, chkIberia.Value
       
        Set oDoc = Nothing
    End If
    
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
End Sub

Private Sub cmdmail_Click()
   On Error GoTo cmdmail_Click_Error

    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque el/los documentos que desea enviar por correo.", vbExclamation, App.Title
            Exit Sub
        End If
        Dim alguno_generado As Boolean
        alguno_generado = False
        Dim destino_documento As String
        Dim destino_documento_todas As String
        Dim oDoc_pago As New clsDocs_pago
        Dim i As Integer
        Dim ref As String
        Dim oD As New clsDocumentacion
        Me.MousePointer = 11
        Dim NUMERO As Integer
        Dim facturas As String
        Dim pedido As String
        Dim tipoDocumento As String
        Dim oCP As New clsClientes_pedidos
        facturas = ""
        NUMERO = 1
        For i = 1 To lista.ListItems.Count
         If lista.ListItems(i).Checked = True Then
          If oDoc_pago.validar_previos_documento(lista.ListItems(i).SubItems(9)) Then
            alguno_generado = True
            If chkFirma.Value = Checked Then
                'M1275-I
                Dim ocli As New clsCliente
                oDoc_pago.CargarDocumento lista.ListItems(i).SubItems(9)
                Select Case oDoc_pago.getTIPO
                    Case 1
                        tipoDocumento = "Albaran/es"
                    Case 2
                        tipoDocumento = "Factura/s"
                    Case 3
                        tipoDocumento = "Abono/s"
                End Select
                
                ' Detalle
                pedido = ""
                If oDoc_pago.getPEDIDO_ID <> 0 Then
                    oCP.Carga oDoc_pago.getPEDIDO_ID
'                    PEDIDO = oCP.getDESCRIPCION & " (" & oCP.getCODIGO & ")"
                    pedido = oCP.getCODIGO
                End If
                
                facturas = facturas & "<tr>"
                facturas = facturas & "<td>" & oDoc_pago.getNUMERO & "</td>"
                facturas = facturas & "<td>" & Format(oDoc_pago.getFECHA_FACTURA, "dd/mm/yyyy") & "</td>"
                facturas = facturas & "<td>" & pedido & "</td>"
                facturas = facturas & "</tr>"
                
'                ocli.CargaCliente oDoc_pago.getCLIENTE_ID
                ocli.CargaCliente oDoc_pago.getCLIENTE_ID_FACTURA
                If ocli.getFACTURA_ELECTRONICA = 0 Then
                    alguno_generado = False
                    Me.MousePointer = 0
                    MsgBox "El cliente " & lista.ListItems(i).SubItems(1) & ", NO desea recibir factura electrónica.", vbExclamation, App.Title
                    'JGM
                    Exit Sub
                Else
                'M1275-F
                'JGM-I
                ' Generar el documento firmado usando el servidor de impresion
                frmFirma.visible = True
                lblFirma.Caption = "Generando firma DOC : " & lista.ListItems(i).Text
                lblFirma2.Caption = NUMERO & " de " & contar_marcados
                NUMERO = NUMERO + 1
                DoEvents
                oD.EliminarDOC_PAGO lista.ListItems(i).SubItems(9)
                
                Dim oimp As New clsImpresion
                Dim ID As Integer
                With oimp
                    .setMUESTRA_ID = lista.ListItems(i).SubItems(9)
                    .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    .setTIPO = 65
                    .setPUESTO = USUARIO.getUSO
                    ID = .Insertar
                End With
                Dim r As Integer
                Dim intentos As Integer
                r = 1
                intentos = 0
                Do While r = 1 And intentos < 10
                    Espera (1)
                    r = verificar_impresion(lista.ListItems(i).SubItems(9))
                    intentos = intentos + 1
                Loop
                If r = 0 Then
                    Espera (1)
                'JGM-F
                    destino_documento = oD.CargarDOC_PAGO(lista.ListItems(i).SubItems(9), False)
                    If destino_documento <> "" Then
                        ref = ref & lista.ListItems(i).Text & ", "
                        destino_documento_todas = destino_documento_todas & destino_documento & ";"
                    Else
                        Me.MousePointer = 0
                        frmFirma.visible = False
                        MsgBox "El informe con la factura no se ha generado correctamente. Verifique el Servidor de Impresión", vbCritical, App.Title
                        Exit Sub
                    End If
                Else
                    Me.MousePointer = 0
                    frmFirma.visible = False
                    MsgBox "El informe con la factura no se ha generado correctamente. Verifique el Servidor de Impresión", vbCritical, App.Title
                    Exit Sub
                End If
                'M1275-I
                End If
                'M1275-F
            Else
                oDoc_pago.CargarDocumento lista.ListItems(i).SubItems(9)
                Select Case oDoc_pago.getTIPO
                    Case 1
                        tipoDocumento = "Albaran/es"
                    Case 2
                        tipoDocumento = "Factura/s"
                    Case 3
                        tipoDocumento = "Abono/s"
                End Select
                ' DETALLE
                pedido = ""
                If oDoc_pago.getPEDIDO_ID <> 0 Then
                    oCP.Carga oDoc_pago.getPEDIDO_ID
'                    PEDIDO = oCP.getDESCRIPCION & " (" & oCP.getCODIGO & ")"
                    pedido = oCP.getCODIGO
                End If
                facturas = facturas & "<tr>"
                facturas = facturas & "<td>" & oDoc_pago.getNUMERO & "</td>"
                facturas = facturas & "<td>" & Format(oDoc_pago.getFECHA_FACTURA, "dd/mm/yyyy") & "</td>"
                facturas = facturas & "<td>" & pedido & "</td>"
                facturas = facturas & "</tr>"
                
                destino_documento = App.Path & "\" & oDoc_pago.getNUMERO & ".pdf"
                On Error Resume Next
                If Dir(destino_documento) <> "" Then
                    Kill destino_documento
                End If
                On Error GoTo cmdmail_Click_Error
                ' Generamos el pdf
                oDoc_pago.generar_factura lista.ListItems(i).SubItems(9), False, destino_documento, "rptFactura"
                ' Obtenemos los datos del correo
                Dim oCliente As New clsCliente
'                oCliente.CargaCliente oDoc_pago.getCLIENTE_ID
                oCliente.CargaCliente oDoc_pago.getCLIENTE_ID_FACTURA
                If Dir(destino_documento) = "" Then
                    Me.MousePointer = 0
                    frmFirma.visible = False
                    MsgBox "El informe con la factura no se ha generado correctamente.", vbInformation, App.Title
                    Exit Sub
                End If
                'FIRMADIGITAL
                If chkFirma.Value = Checked Then
                    Dim firma As String
                    firma = firmarPdf(destino_documento)
                    If firma <> "" Then
                        Me.MousePointer = 0
                        frmFirma.visible = False
                        MsgBox "ERROR AL FIRMAR DIGITALMENTE EL DOCUMENTO : " & firma
                        Exit Sub
                    End If
                End If
                ref = ref & oDoc_pago.getNUMERO & ", "
                destino_documento_todas = destino_documento_todas & destino_documento & ";"
            End If
           End If
          End If
        Next
        If alguno_generado Then
            ref = "Envío de " & tipoDocumento & " número/s: " & Left(ref, Len(ref) - 2)
            Dim body As String
            body = ""
            body = body & "<html>" & vbNewLine
            Dim hora As Date
            hora = Time
            If hora > "12:00:00" Then
                body = body & "Buenas tardes,<br />"
            Else
                body = body & "Buenos días,<br />"
            End If
            body = body & "Adjunto le envío la siguiente relación de facturas electrónicas: <br /><br />"
            body = body & "<table border='1'>"
            body = body & " <tr><th>Nº Factura</th><th>Fecha Factura</th><th>Nº de pedido</th></tr>"
            body = body & facturas
            body = body & "</table><br />"
            body = body & "Si desea recibir las facturas en papel, por favor comuníquenoslo respondiendo a este correo.<br />"
            body = body & "Quedamos a su disposición para cualquier consulta al respecto.<br /><br />"
            body = body & "Un saludo y gracias."
            body = body & "</html>"
            
            
            'M1357-I
            'genera_correo oCliente.getEMAIL2, ref, "", destino_documento_todas, Me.hdc
            If chkFirma.Value = Checked Then
                genera_correo ocli.getEMAIL_FACTURACION, ref, body, destino_documento_todas, Me.hdc, True
            Else
                genera_correo oCliente.getEMAIL_FACTURACION, ref, body, destino_documento_todas, Me.hdc, True
            End If
            'M1357-F
        End If
    End If
    frmFirma.visible = False
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdmail_Click_Error:

    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmail_Click of Formulario frmListadoDocPago"
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
'    If contabilizado(lista.ListItems(lista.SelectedItem.Index).SubItems(9)) Then
'        Exit Sub
'    End If
    gdoc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
    If lista.ListItems(lista.selectedItem.Index).SubItems(10) <> 1 Then
        frmModificarFactura.Show 1
    Else
        frmFacturaConceptos.Show 1
    End If
    actualizar_lista CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)), lista.selectedItem.Index
    gdoc = 0
End Sub

Private Sub cmdSal_Click()
    log ("Cerrando listado de documentos de pago")
    Unload Me
End Sub

Private Sub cmdSalir_Click()
End Sub
Private Sub cmdVerPedido_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDoc As New clsDocs_pago
    oDoc.CargarDocumento lista.ListItems(lista.selectedItem.Index).SubItems(9)
    If oDoc.getPEDIDO_ID = 0 Then
        MsgBox "El documento no tiene pedido asociado.", vbExclamation, App.Title
    Else
        frmClientes_Detalle_Pedido.PK = oDoc.getPEDIDO_ID
        frmClientes_Detalle_Pedido.Show 1
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
    If KeyCode = 116 Then 'F5 Informar pedido
        cmdInformarPedido_Click
    End If

End Sub

Private Sub Form_Load()
    log (Me.Name)
'    cargar_botones Me
    Me.Left = 10
    Me.top = 50
    txtanno = Year(Date)
    cambiar.Max = Year(Date)

    cabecera
    rellenar_clientes
    pedidos (0)
    fhasta = Now
    fdesde = Now
    fdesdev = Now
    fhastav = Now
    fCobroDesde = Now
    fCobroHasta = Now
    fPrevistaDesde = Now
    fPrevistaHasta = Now
'    ffactorizada = Now
    ' Viene del código de barras
    If gdoc <> 0 Then
        cargar_documento
    End If
End Sub
Private Sub rellenar_clientes()
    cargar_combo cmbFP, New clsFP
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbclientesFact, New clsCliente, 0, frmClientes, ""
End Sub
Private Sub pedidos(ID As Integer)
    cmbPedidos.limpiar
    Dim filtro As String
    If ID <> 0 Then
        filtro = " AND CLIENTE_ID = " & ID
    End If
    llenar_combo cmbPedidos, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub
Private Sub pedidosAsignar(ID As Integer)
    cmdPedidosAsginar.limpiar
    Dim filtro As String
    If ID <> 0 Then
        filtro = " AND CLIENTE_ID = " & ID
    End If
    llenar_combo cmdPedidosAsginar, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºDoc", 900, lvwColumnLeft
        .Add , , "Cliente", 2500, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Importe", 1100, lvwColumnRight
        .Add , , "Dto%", 500, lvwColumnCenter
        .Add , , "Base", 1100, lvwColumnRight
        .Add , , "IVA%", 500, lvwColumnRight
        .Add , , "Cuota I.V.A.", 1100, lvwColumnRight
        .Add , , "Total", 1200, lvwColumnRight
        .Add , , "ID_DOC", 1, lvwColumnCenter
        .Add , , "FACTURA_CONCEPTOS", 1, lvwColumnCenter
        .Add , , "PAGADO", 1, lvwColumnCenter
        .Add , , "TIPO_DOCUMENTO", 1, lvwColumnCenter
        .Add , , "FACTORIZADA", 1, lvwColumnCenter
        .Add , , "PEDIDO", 1500, lvwColumnLeft
        .Add , , "FP", 1, lvwColumnCenter
        .Add , , "Asiento", 700, lvwColumnCenter
        .Add , , "F.Vencim.", 1050, lvwColumnCenter
        .Add , , "Comentario", 1, lvwColumnLeft
        .Add , , "ClienteFactura", 1, lvwColumnLeft
        .Add , , "F.Prev.Cobro", 1050, lvwColumnCenter
        .Add , , "F.Cobro", 1050, lvwColumnCenter
        .Add , , "CCC", 1, lvwColumnCenter
        .Add , , "ID_PEDIDO", 1, lvwColumnCenter
        .Add , , "ID_CLIENTE_FACTURA", 1, lvwColumnCenter
        .Add , , "CIF", 1, lvwColumnCenter
    End With
End Sub

Public Function TIPO_DOCUMENTO() As String
    If opTipo(0).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA  ' Factura
        lblMsg = "Facturas "
    ElseIf opTipo(1).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN  ' Albaran
        lblMsg = "Albaranes "
    ElseIf opTipo(2).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ABONO   ' Abono
        lblMsg = "Abonos "
    ElseIf opTipo(3).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA  ' Facturas Anuladas
        lblMsg = "Facturas Anuladas "
    ElseIf opTipo(4).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA  ' Facturas Abonadas
        lblMsg = "Facturas Abonadas "
    ElseIf opTipo(5).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & "," & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ABONO
        lblMsg = " TODAS "
    ElseIf opTipo(6).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA  ' Proformas
        lblMsg = "Proformas "
    ElseIf opTipo(7).Value = True Then
        TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA & "," & C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
        lblMsg = "Facturas y Albaranes "
    End If
End Function
Public Function Estado_documento() As String
    If opTipo(0).Value = True Then
        Estado_documento = "0,2"
    ElseIf opTipo(1).Value = True Then
        Estado_documento = "0"
    ElseIf opTipo(2).Value = True Then
        Estado_documento = "0"
    ElseIf opTipo(3).Value = True Then
        Estado_documento = "1"
    ElseIf opTipo(4).Value = True Then
        Estado_documento = "2"
    ElseIf opTipo(5).Value = True Then
        Estado_documento = "0,2"
    ElseIf opTipo(6).Value = True Then
        Estado_documento = "0,2"
    ElseIf opTipo(7).Value = True Then ' Facturas y Albaranes
        Estado_documento = "0,2"
    End If
End Function
Private Function Estado_Enviada() As String
    If opEnviada(0).Value = True Then
        Estado_Enviada = "0,1"
    ElseIf opEnviada(1).Value = True Then
        Estado_Enviada = "1"
    ElseIf opEnviada(2).Value = True Then
        Estado_Enviada = "0"
    End If
End Function

Private Function pendiente_cobro() As Integer
    If opPendiente(0).Value = True Then
        pendiente_cobro = 0 ' Pendiente
        lblMsg = lblMsg & "pendientes de cobro "
    ElseIf opPendiente(1).Value = True Then
        pendiente_cobro = 1  ' Cobrada
        lblMsg = lblMsg & "cobrados "
    Else
        pendiente_cobro = 2  ' Todos
    End If
End Function

Private Sub formatea_titulo()
    lblMsg = lblMsg & "entre " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy")
End Sub
Private Sub desactivar_controles()
    cmbCobrar.Enabled = False
    cmdDescobrar.Enabled = False
    cmdDatos.Enabled = False
    cmdAnular.Enabled = False
    cmdEditar.Enabled = False
    cmdduplicar.Enabled = False
    cmdAbono.Enabled = False
    cmdImprimir2.Enabled = False
    cmdIberia.Enabled = False
    cmdModificar.Enabled = False
    cmdduplicar.Enabled = False
    cmdEditar.Enabled = False
    cmdCartaPago.Enabled = False
    cmdMail.Enabled = False
'    cmdRecibo.Enabled = False
'    cmdConceptos.Enabled = False
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    desactivar_controles
    cmdImprimir2.Enabled = True
    cmdIberia.Enabled = True
    cmdMail.Enabled = True
    cmdDatos.Enabled = True
    cmdduplicar.Enabled = False
    If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(10)) = 1 And CInt(lista.ListItems(lista.selectedItem.Index).SubItems(12)) <> 3 Then
        cmdduplicar.Enabled = True
    End If
    If TIPO_DOCUMENTO <> "3" And Estado_documento <> "1" Then
      cmbCobrar.Enabled = True
      If lista.ListItems(lista.selectedItem.Index).SubItems(11) = 0 Then
        cmdCartaPago.Enabled = True
        cmdModificar.Enabled = True
        cmdEditar.Enabled = True
'        cmbCobrar.Enabled = True
        cmdAnular.Enabled = True
        cmdEditar.Enabled = True
        cmdduplicar.Enabled = True
        cmdAbono.Enabled = True
      ElseIf lista.ListItems(lista.selectedItem.Index).SubItems(11) <> 0 Then
        If TIPO_DOCUMENTO = "2" Then
            cmdDescobrar.Enabled = True
            cmdAbono.Enabled = True
        End If
      End If
    End If
    ' Poder enviar los abonos por mail
    If TIPO_DOCUMENTO = "3" Then
        cmdModificar.Enabled = True
    End If
    ' Si es una factura abonada, habilitar el botón de Abonar para los parciales
    If TIPO_DOCUMENTO = "2" And Estado_documento = "2" Then
        cmdAbono.Enabled = True
    End If
    If lista.ListItems.Count > 0 Then
        If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(10)) <> 1 Then
            cmdEditar.Enabled = True
        Else
            cmdEditar.Enabled = False
        End If
        ' Informar pedido
        pedidosAsignar lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_CLIENTE_FACTURA)
        If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_PEDIDO) <> 0 Then
            cmdPedidosAsginar.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_PEDIDO)
        End If
    End If

End Sub

Private Sub lista_DblClick()
'    cmdImprimir2_Click
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub opTipo_Click(Index As Integer)
    If Index > 1 Then
        opPendiente(2).Value = True
    End If
End Sub
Private Sub actualizar_lista(documento As Long, fila As Integer)
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT d.ID_DOC,d.NUMERO,cl.nombre,d.FECHA_FACTURA,d.DESCUENTO,d.IVA,d.TIPO,d.PAGADO,d.total,d.factura_conceptos,d.PEDIDO_ID,d.FACTORIZADA,cp.codigo,d.FP_ID,d.CLIENTE_ID,d.CLIENTE_ID_FACTURA, d.COMENTARIO,clifact.nombre " & _
               "  FROM docs_pago d,clientes cl,clientes_pedidos cp, clientes clifact " & _
               "WHERE d.id_doc = " & documento & _
               "  AND d.cliente_id = cl.id_cliente " & _
               "  AND d.cliente_id_factura = clifact.id_cliente " & _
               "  AND d.pedido_id = cp.id_pedido "
    Set rs = datos_bd(consulta)
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
    If rs.RecordCount <> 0 Then
         
                Select Case rs(6)
                Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
                    NUMERO = "A-" & Format(rs(1), "0000")
                Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
                    NUMERO = "F-" & Format(rs(1), "0000")
                Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
                    NUMERO = "B-" & Format(rs(1), "0000")
                Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA
                    NUMERO = "P-" & Format(rs(1), "0000")
                Case Else
                    NUMERO = Format(rs(1), "0000")
                End Select
         
         lista.ListItems(fila).Text = NUMERO
         
         lista.ListItems(fila).SubItems(1) = rs.Fields(2)
         lista.ListItems(fila).SubItems(2) = rs.Fields(3)
         lista.ListItems(fila).SubItems(9) = rs.Fields(0)
         IMPORTE = rs.Fields(8)
         If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
               BASE = IMPORTE
         Else
               BASE = IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100)
         End If
         IVA = (BASE * rs.Fields("iva")) / 100
         lista.ListItems(fila).SubItems(3) = Format(IMPORTE, "currency")
         lista.ListItems(fila).SubItems(4) = Format(rs.Fields("descuento"), "Standard")
         lista.ListItems(fila).SubItems(5) = Format(BASE, "currency")
         lista.ListItems(fila).SubItems(6) = rs.Fields("iva")
         lista.ListItems(fila).SubItems(7) = Format(IVA, "currency")
         lista.ListItems(fila).SubItems(8) = Format(BASE + IVA, "currency")
         lista.ListItems(fila).SubItems(10) = rs(9)
    
' VENC-I
         lista.ListItems(fila).SubItems(11) = rs(7)
         lista.ListItems(fila).SubItems(12) = rs(6)
         lista.ListItems(fila).SubItems(13) = rs(11) ' Factorizada
         lista.ListItems(fila).SubItems(14) = rs(12) ' PEdido
         lista.ListItems(fila).SubItems(15) = rs(13) ' Forma de pago
         lista.ListItems(fila).SubItems(18) = rs(16) ' Forma de pago
         lista.ListItems(fila).SubItems(19) = rs(17) ' Cliente Factura
' VENC-F
        lista.ListItems(fila).SubItems(COLS.ID_PEDIDO) = rs(10)
        lista.ListItems(fila).SubItems(COLS.ID_CLIENTE_FACTURA) = rs(15)

       ' 14 CLIENTE_ID, 15 CLIENTE_ID_FACTURA
         If rs(14) <> rs(15) Then
            colorear fila, vbRed
         End If
        
    
    End If
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Function contar_marcados() As Integer
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cont = cont + 1
        End If
    Next
    contar_marcados = cont
End Function
Private Sub cargar_documento()
    With lista.ListItems.Add(, , "")
    End With
    actualizar_lista CLng(gdoc), 1
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
Private Function contabilizado(documento As Long) As Boolean
      Dim oDoc As New clsDocs_pago
   On Error GoTo contabilizado_Error

      contabilidad = oDoc.esta_contabilidado(documento)

   On Error GoTo 0
   Exit Function

contabilizado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure contabilizado of Formulario frmListadoDocPago"
End Function
Private Sub colorear(fila As Integer, color As Long)
    Dim i As Integer
    On Error Resume Next
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub


Private Sub generar_excel_listado()
    Dim i As Integer
    Dim oFP As New clsFP
    Dim oDPE As New clsDocs_pago_envios
    Dim rs As ADODB.Recordset
    
    Dim cobros As Boolean
    cobros = False
    If MsgBox("¿Desea exportar los comentarios de los cobros?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        cobros = True
    End If
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo generar_excel_listado_Error

    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Add
    Set XLS = XLW.Worksheets(1)
    XLW.Worksheets(3).Delete
    XLW.Worksheets(2).Delete
    XLW.Worksheets(1).Name = "Listado de facturas"
    'XLA.visible = True
    Me.MousePointer = 11
    With XLS.Range("A1:U1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With XLS.Range("A1:U1").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = &HC0C0FF
    End With
    With XLS.Range("A1:U1").Borders
        .LineStyle = vbSolid
    End With
    XLS.Range("C1:C1").ColumnWidth = 20
    XLS.Range("D1:D1").ColumnWidth = 20
'    XLS.Range("F1:F1").ColumnWidth = 20
    XLS.Range("O1:O1").ColumnWidth = 20
    XLS.Range("P1:P1").ColumnWidth = 20
    XLS.Range("R1:R1").ColumnWidth = 20
    XLS.Range("S1:S1").ColumnWidth = 20
'    XLS.Range("1:1").HorizontalAlignment = xlCenter
'    XLS.Range("1:1").VerticalAlignment = xlCenter
'    XLS.Range("1:1").RowHeight = 30
'    XLS.Range("1:1").WrapText = True
    'Cabecera
    XLS.Cells(1, 1) = "Tipo"
    XLS.Cells(1, 2) = "Documento"
    XLS.Cells(1, 3) = "Cliente"
    XLS.Cells(1, 4) = "Cliente Factura"
    XLS.Cells(1, 5) = "CIF"
    XLS.Cells(1, 6) = "Fecha"
    XLS.Cells(1, 7) = "Vencimiento"
    XLS.Cells(1, 8) = "Base"
    XLS.Cells(1, 9) = "Dto"
    XLS.Cells(1, 10) = "Base"
    XLS.Cells(1, 11) = "Iva"
    XLS.Cells(1, 12) = "Imp.Iva"
    XLS.Cells(1, 13) = "Total"
    XLS.Cells(1, 14) = "Pagada"
    XLS.Cells(1, 15) = "Pedido"
    XLS.Cells(1, 16) = "Comentario"
    XLS.Cells(1, 17) = "C.C."
    XLS.Cells(1, 18) = "F.Prevista Cobro"
    XLS.Cells(1, 19) = "F.Cobro"
    If cobros = True Then
        XLS.Cells(1, 20) = "F.Envío"
        XLS.Cells(1, 21) = "Envío"
        XLS.Cells(1, 22) = "Usuario"
    End If
    fila = 2
    Dim num_doc As String
    ' Datos
    For i = 1 To lista.ListItems.Count
        XLS.Range(XLS.Cells(fila, 8), XLS.Cells(fila, 13)).NumberFormat = "0.00"
        If lista.ListItems(i).SubItems(12) = 1 Then
            XLS.Cells(fila, 1) = "A" ' Tipo
        Else
            XLS.Cells(fila, 1) = "F" ' Tipo
        End If
        ' Numero Doc
        num_doc = lista.ListItems(i).Text
        num_doc = Replace(num_doc, "F-", "")
        num_doc = Replace(num_doc, "A-", "")
        num_doc = Format(num_doc, "#,###")
        XLS.Cells(fila, 2) = num_doc & "/" & Format(lista.ListItems(i).SubItems(2), "yyyy")  ' Documento
        If Trim(lista.ListItems(i).SubItems(1)) <> Trim(lista.ListItems(i).SubItems(19)) Then
            XLS.Cells(fila, 3).Font.color = vbRed
            XLS.Cells(fila, 4).Font.color = vbRed
        End If
        XLS.Cells(fila, 3) = lista.ListItems(i).SubItems(1) ' Cliente
        XLS.Cells(fila, 4) = lista.ListItems(i).SubItems(19) ' Cliente Factura
        XLS.Cells(fila, 5) = lista.ListItems(i).SubItems(COLS.COL_CIF)  ' Cif
        XLS.Cells(fila, 6) = Format(lista.ListItems(i).SubItems(2), "mm/dd/yyyy") ' Fecha
        If lista.ListItems(i).SubItems(12) = 2 And _
           CInt(lista.ListItems(i).SubItems(15)) <> 0 Then
             oFP.CARGAR lista.ListItems(i).SubItems(15)
             XLS.Cells(fila, 7) = Format(CDate(lista.ListItems(i).SubItems(2)) + oFP.getDIAS, "mm/dd/yyyy") ' Vencimiento
        Else
             XLS.Cells(fila, 7) = Format(lista.ListItems(i).SubItems(2), "mm/dd/yyyy")
        End If
        
        XLS.Cells(fila, 8) = CSng(lista.ListItems(i).SubItems(3)) ' Importe
        XLS.Cells(fila, 9) = CSng(lista.ListItems(i).SubItems(4)) ' Dto
        XLS.Cells(fila, 10) = CSng(lista.ListItems(i).SubItems(5)) ' Base con DTo
        XLS.Cells(fila, 11) = CSng(lista.ListItems(i).SubItems(6)) ' Iva
        XLS.Cells(fila, 12) = CSng(lista.ListItems(i).SubItems(7)) ' Imp.Iva
        XLS.Cells(fila, 13) = CSng(lista.ListItems(i).SubItems(8)) ' Total
        If lista.ListItems(i).SubItems(11) = 0 Then
            XLS.Cells(fila, 14) = "N" ' Pagada
        Else
            XLS.Cells(fila, 14) = "S"
        End If
        XLS.Cells(fila, 15) = lista.ListItems(i).SubItems(14) ' Pedido
        XLS.Cells(fila, 16) = lista.ListItems(i).SubItems(18)  ' Comentario
        XLS.Cells(fila, 17) = lista.ListItems(i).SubItems(COLS.COL_CCC)   ' CC
        ' Cobro
        XLS.Cells(fila, 18) = Format(lista.ListItems(i).SubItems(COLS.COL_FECHA_PREVISTA_COBRO), "mm/dd/yyyy")    ' Fecha Prevista
        XLS.Cells(fila, 19) = Format(lista.ListItems(i).SubItems(COLS.COL_FECHA_COBRO), "mm/dd/yyyy")   ' Fecha Cobro
        If cobros = True Then
            Set rs = oDPE.Listado(lista.ListItems(i).SubItems(9))
            If rs.RecordCount > 0 Then
                Do
                    XLS.Cells(fila, 20) = Format(rs(1), "mm/dd/yyyy") ' Fecha Envio
                    XLS.Cells(fila, 21) = rs(2) ' Detalle
                    If Not IsNull(rs(3)) Then
                        XLS.Cells(fila, 22) = rs(3) ' Usuario
                    End If
                    rs.MoveNext
                    If Not rs.EOF Then
                        fila = fila + 1
                    End If
                Loop Until rs.EOF
            End If
        End If
        fila = fila + 1
    Next
    Me.MousePointer = 0
    MsgBox "Listado generado correctamente.", vbOKOnly + vbInformation, App.Title
    XLA.visible = True
   On Error GoTo 0
   Exit Sub

generar_excel_listado_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_excel_listado of Formulario frmListadoDocPago"
End Sub


Private Sub txtconcepto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtConcepto <> "" Then
            cmdBuscar_Click
        End If
    End If
End Sub

Private Sub txtnumero_GotFocus()
    txtNumero.SelStart = 0
    txtNumero.SelLength = Len(txtNumero)
End Sub

Private Sub txtnumero_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If txtNumero <> "" Then
            cmdBuscar_Click
        End If
    End If
End Sub

