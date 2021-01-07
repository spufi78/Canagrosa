VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmClientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envío de Informes"
   ClientHeight    =   11955
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   13635
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11955
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Metrología"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1080
      Left            =   45
      TabIndex        =   92
      Top             =   6705
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   22
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   18
         Top             =   270
         Width           =   7920
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   21
         Left            =   1155
         TabIndex        =   19
         Top             =   630
         Width           =   7920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   28
         Left            =   150
         TabIndex        =   94
         Top             =   315
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail"
         Height          =   195
         Index           =   27
         Left            =   150
         TabIndex        =   93
         Top             =   675
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   10980
      Width           =   1365
   End
   Begin VB.CommandButton cmdOfertas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ofertas"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   11025
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   11025
      Width           =   1050
   End
   Begin VB.Frame frmIndicadores 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Indicadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   9270
      TabIndex        =   78
      Top             =   3870
      Width           =   4320
      Begin MSComctlLib.ListView listaIndicadores 
         Height          =   6015
         Left            =   135
         TabIndex        =   79
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
   End
   Begin VB.Frame frmanalisis 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsables del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   9270
      TabIndex        =   73
      Top             =   405
      Width           =   4320
      Begin MSComctlLib.ListView listaResponsables 
         Height          =   2280
         Left            =   135
         TabIndex        =   74
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   135
         TabIndex        =   75
         Top             =   2565
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirNorma 
         Height          =   435
         Left            =   630
         TabIndex        =   76
         Top             =   2925
         Width           =   1545
         _Version        =   851970
         _ExtentX        =   2725
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmClientes.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarNorma 
         Height          =   435
         Left            =   2205
         TabIndex        =   77
         Top             =   2925
         Width           =   1590
         _Version        =   851970
         _ExtentX        =   2805
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmClientes.frx":712C
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dirección Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   60
      TabIndex        =   64
      Top             =   1890
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1065
         MaxLength       =   150
         TabIndex        =   4
         Top             =   300
         Width           =   7980
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1065
         TabIndex        =   5
         Top             =   705
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo cmbPais 
         Height          =   315
         Left            =   5265
         TabIndex        =   6
         Top             =   705
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   1065
         TabIndex        =   7
         Top             =   1125
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMunicipios 
         Height          =   315
         Left            =   5265
         TabIndex        =   8
         Top             =   1125
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   345
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   68
         Top             =   795
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pais"
         Height          =   195
         Index           =   3
         Left            =   4365
         TabIndex        =   67
         Top             =   765
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   66
         Top             =   1185
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   4365
         TabIndex        =   65
         Top             =   1185
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Envío de Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1080
      Left            =   45
      TabIndex        =   60
      Top             =   5625
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1155
         TabIndex        =   17
         Top             =   630
         Width           =   7920
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   16
         Top             =   270
         Width           =   7920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   62
         Top             =   675
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "AT. Informes"
         Height          =   195
         Index           =   21
         Left            =   150
         TabIndex        =   61
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Acceso Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Index           =   1
      Left            =   6030
      TabIndex        =   56
      Top             =   9585
      Width           =   3210
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enviar por Correo"
         Height          =   285
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   990
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar clave"
         Height          =   285
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   990
         Width           =   1410
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   270
         Width           =   1785
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   17
         Left            =   1080
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   29
         Top             =   630
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   19
         Left            =   225
         TabIndex        =   58
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clave"
         Height          =   195
         Index           =   18
         Left            =   225
         TabIndex        =   57
         Top             =   690
         Width           =   405
      End
   End
   Begin VB.CommandButton cmddirecciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Direcciones"
      Height          =   870
      Left            =   2235
      Picture         =   "frmClientes.frx":D98E
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   11025
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos financieros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   0
      Left            =   45
      TabIndex        =   51
      Top             =   7785
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1260
         TabIndex        =   24
         Top             =   945
         Width           =   5820
      End
      Begin VB.CheckBox chkFacturaElectronica 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura Electrónica"
         Height          =   240
         Left            =   7335
         TabIndex        =   25
         Top             =   990
         Width           =   1680
      End
      Begin pryCombo.miCombo cmbtarifa 
         Height          =   345
         Left            =   5175
         TabIndex        =   23
         Top             =   585
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   609
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   10
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   20
         Top             =   225
         Width           =   3105
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   11
         Left            =   5175
         MaxLength       =   30
         TabIndex        =   21
         Top             =   225
         Width           =   3885
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1260
         TabIndex        =   22
         Top             =   585
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmdClienteFactura 
         Height          =   345
         Left            =   1260
         TabIndex        =   26
         Top             =   1350
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   609
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente Factura"
         Height          =   195
         Index           =   24
         Left            =   135
         TabIndex        =   84
         Top             =   1410
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail (Fact.)"
         Height          =   195
         Index           =   23
         Left            =   135
         TabIndex        =   82
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         Height          =   195
         Index           =   17
         Left            =   4590
         TabIndex        =   55
         Top             =   645
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   54
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   53
         Top             =   315
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CCC"
         Height          =   195
         Index           =   9
         Left            =   4590
         TabIndex        =   52
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   10980
      Width           =   1050
   End
   Begin VB.CommandButton cmdPedidos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pedidos"
      Height          =   870
      Left            =   45
      Picture         =   "frmClientes.frx":E258
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   11025
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11430
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10980
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   13
      Left            =   45
      TabIndex        =   45
      Top             =   9585
      Width           =   5955
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1050
         Index           =   13
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   225
         Width           =   5745
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   60
      TabIndex        =   39
      Top             =   3510
      Width           =   9195
      Begin VB.OptionButton opIdiomaFactura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Inglés"
         Height          =   195
         Index           =   1
         Left            =   2385
         TabIndex        =   88
         Top             =   1755
         Width           =   870
      End
      Begin VB.OptionButton opIdiomaFactura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Español"
         Height          =   195
         Index           =   0
         Left            =   1305
         TabIndex        =   87
         Top             =   1755
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   20
         Left            =   1140
         TabIndex        =   9
         Top             =   210
         Width           =   7940
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   19
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   10
         Top             =   600
         Width           =   3105
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   18
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   15
         Top             =   1335
         Width           =   3675
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   14
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   12
         Top             =   945
         Width           =   3105
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   15
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   11
         Top             =   615
         Width           =   3675
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   13
         Top             =   975
         Width           =   3675
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   9
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1305
         Width           =   3105
      End
      Begin pryCombo.miCombo cmbPlant 
         Height          =   345
         Left            =   5400
         TabIndex        =   90
         Top             =   1710
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   609
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus Plant"
         Height          =   195
         Index           =   26
         Left            =   4410
         TabIndex        =   91
         Top             =   1770
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Idioma Factura"
         Height          =   195
         Index           =   25
         Left            =   90
         TabIndex        =   89
         Top             =   1755
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail"
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   63
         Top             =   285
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C. Contable"
         Height          =   195
         Index           =   20
         Left            =   4395
         TabIndex        =   59
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label Centro 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   990
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sección"
         Height          =   195
         Index           =   14
         Left            =   4380
         TabIndex        =   47
         Top             =   675
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagina Web"
         Height          =   195
         Index           =   10
         Left            =   4380
         TabIndex        =   44
         Top             =   1035
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   41
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cargo"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   40
         Top             =   1350
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Cliente "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   60
      TabIndex        =   36
      Top             =   390
      Width           =   9195
      Begin VB.CheckBox chkIntra 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intrancomunitario"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   86
         Top             =   1125
         Width           =   1545
      End
      Begin VB.CheckBox chkIberia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iberia"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7515
         TabIndex        =   85
         Top             =   900
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chkAgroalimentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agroalimentario"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   72
         Top             =   900
         Width           =   1410
      End
      Begin VB.CheckBox chkExtranjero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Extrancomunitario"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7515
         TabIndex        =   71
         Top             =   1125
         Width           =   1770
      End
      Begin VB.CheckBox chkAirbus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus Military"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7515
         TabIndex        =   70
         Top             =   690
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CheckBox chkeads 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aeronautico"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   50
         Top             =   675
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CheckBox chkfd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura por determinaciones"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2970
         TabIndex        =   49
         Top             =   675
         Width           =   2925
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   3420
         MaxLength       =   30
         TabIndex        =   3
         Top             =   990
         Width           =   1665
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1020
         Width           =   1710
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   1
         Top             =   645
         Width           =   1710
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   0
         Top             =   270
         Width           =   7980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "FAX"
         Height          =   195
         Index           =   15
         Left            =   2970
         TabIndex        =   43
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   42
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.I.F."
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   38
         Top             =   690
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   37
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   45
      TabIndex        =   46
      Top             =   -15
      Width           =   13590
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODO As String
Public PK As Long
Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_CLIENTE
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Cliente " & txtDatos(1)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdAdjuntos_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_CLIENTE
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub

Private Sub cmbPais_LostFocus()
    cargar_provincias
End Sub
Private Sub cmbProvincia_LostFocus()
    cargar_municipios
End Sub

Private Sub cmdAnadirNorma_Click()
    If cmbResponsable.getTEXTO <> "" Then
        Dim i As Integer
        Dim encontrado As Boolean
        encontrado = False
        For i = 1 To listaResponsables.ListItems.Count
            If CLng(listaResponsables.ListItems(i).Text) = cmbResponsable.getPK_SALIDA Then
                encontrado = True
            End If
        Next
        If Not encontrado Then
            With listaResponsables.ListItems.Add(, , cmbResponsable.getPK_SALIDA)
                .SubItems(1) = cmbResponsable.getTEXTO
            End With
        End If
    End If
End Sub

Private Sub cmddirecciones_Click()
    frmClientes_Direcciones.PK = PK
    frmClientes_Direcciones.Show 1
    consulta_Cliente
End Sub

Private Sub cmdEliminarNorma_Click()
    If listaResponsables.ListItems.Count > 0 Then
        listaResponsables.ListItems.Remove listaResponsables.selectedItem.Index
    End If
End Sub

Private Sub cmdOfertas_Click()
    frmOferta_Listado_Modal.pk_CLIENTE = PK
    frmOferta_Listado_Modal.Show 1
End Sub

'Private Sub cmdModificar_Click()
'    desbloquear_controles
'End Sub
Private Sub cmdok_Click()
    If PK <> 0 Then
        modificar_cliente
    Else
        insertar_cliente
    End If
End Sub

Private Sub cmdPedidos_Click()
    frmClientes_Pedidos.PK = PK
    frmClientes_Pedidos.Show 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    txtDatos(17) = generar_clave_web
End Sub

Private Sub Command2_Click()
    Dim DOC As String
    Dim ASUNTO As String
    Dim texto As String
    ASUNTO = "Canagrosa.com - Nueva Contraseña para la intranet"
    texto = "Estimado cliente," & vbNewLine
    texto = texto & "Se ha procesado al cambio de contraseña para la Intranet de cliente." & vbNewLine & vbNewLine
    texto = texto & "Pulse para acceder a la intranet: http://www.canagrosa.com/geslab" & vbNewLine & vbNewLine
    texto = texto & "Usuario: " & txtDatos(16) & vbNewLine
    texto = texto & "Contraseña: " & txtDatos(17) & vbNewLine & vbNewLine
    texto = texto & "Reciba un cordial saludo," & vbNewLine
    texto = texto & "mailto:canagrosa@canagrosa.com" & vbNewLine
    genera_correo txtDatos(7), ASUNTO, texto, DOC, Me.Hwnd
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_paises
    cargar_combo cmbFP, New clsFP
    llenar_combo cmbtarifa, New clsTarifas, 0, Me, ""
    llenar_combo cmbResponsable, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmdClienteFactura, New clsCliente, 0, frmClientes, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbPlant, DECODIFICADORA.AIRBUS_PLANT
    Set oDeco = Nothing
    cabecera
    If PK <> 0 Then
        cmddirecciones.Enabled = True
        cmdAdjuntos.Enabled = True
        consulta_Cliente
    Else
        cmddirecciones.Enabled = False
        cmdAdjuntos.Enabled = False
        Dim oCliente As New clsCliente
        oCliente.CrearId_Cliente
        txtDatos(16) = oCliente.getID_CLIENTE
        Dim oParametro As New clsParametros
        oParametro.Carga parametros.CUENTA_CONTABLE_CLIENTE, ""
        txtDatos(18) = oParametro.getVALOR
'        cmbCC.MostrarElemento familia.cliente
    End If
    permisos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PK = 0
    Set frmClientes = Nothing
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
       If Index <> 13 Then
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtDatos(13).SetFocus
       Else
        If Index <> 13 Then
        SendKeys "+{Tab}", True
        End If
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdSalir_Click
    End Select
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
    If Index = 17 Then
        If Not IsNumeric(txtDatos(17)) Then
            MsgBox "La contraseña del usuario web debe ser numérica.", vbCritical, App.Title
            txtDatos(17) = ""
        End If
    End If
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 18
        txtDatos(i) = ""
    Next
    cmbPais.Text = ""
    cmbProvincia.Text = ""
    cmbMunicipios.Text = ""
    cmbtarifa.limpiar
    cmbPlant.limpiar
    chkFD.Value = Unchecked
    'E0200-I
    'chkSubcontrata.value = Unchecked
    'E0200-F
    txtDatos(1).SetFocus
End Sub

'Public Sub bloquear_campos()
'    Dim i As Integer
'    For i = 1 To 13
'        txtdatos(i).Locked = True
'    Next
'    chkfd.Enabled = False
'    cmbMunicipios.Locked = True
'    cmbProvincia.Locked = True
'    cmbPais.Locked = True
'    cmbtarifa.desactivar
'End Sub

Public Sub insertar_cliente()
   On Error GoTo insertar_cliente_Error

    If valida_datos = False Then
        Exit Sub
    End If
    If MsgBox("Va a dar de alta el Cliente. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        Set oCliente = mover_datos
        Dim ID As Long
        ID = oCliente.insertar_cliente
        If ID > 0 Then
            Dim ohc As New clsHistorial_cambios
            With ohc
                .setTIPO = HC_TIPOS.HC_CLIENTE
                .setIDENTIFICADOR = ID
                .setIDENTIFICADOR_TEXTO = txtDatos(1)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setMOTIVO = HC_CREACION
                .Insertar
            End With
            Set ohc = Nothing
            almacenarResponsables ID
            borrar_campos
        End If
        Set oCliente = Nothing
    End If

   On Error GoTo 0
   Exit Sub

insertar_cliente_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_cliente of Formulario frmClientes"
End Sub

Public Sub modificar_cliente()
   On Error GoTo modificar_cliente_Error

    If USUARIO.getPER_MOD_CLIENTE = False Then
        MsgBox "Su usuario no tiene permisos para modificar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If


    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim cliente As Integer
    If MsgBox("Va a modificar los datos del Cliente. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
         frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del cliente."
         frmMotivo.Show 1
         If Trim(MOTIVO) = "" Then
            MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
            Exit Sub
        End If
        Set oCliente = mover_datos
        oCliente.setID_CLIENTE = PK
        If oCliente.modificar_cliente = True Then
            Dim ohc As New clsHistorial_cambios
            With ohc
                .setTIPO = HC_TIPOS.HC_CLIENTE
                .setIDENTIFICADOR = PK
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setIDENTIFICADOR_TEXTO = txtDatos(1)
                .setMOTIVO = Trim(MOTIVO)
                .Insertar
            End With
            Set ohc = Nothing
            almacenarResponsables PK
            Unload Me
        End If
        Set oCliente = Nothing
    End If

   On Error GoTo 0
   Exit Sub

modificar_cliente_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure modificar_cliente of Formulario frmClientes"

End Sub


Public Function valida_datos() As Boolean
    valida_datos = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre del cliente no puede estar en blanco.", vbCritical, "Error"
        txtDatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(2) = "" Then
        MsgBox "La dirección no puede estar en blanco.", vbCritical, "Error"
        txtDatos(2).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(3) = "" Then
        MsgBox "El codigo postal no puede estar en blanco.", vbCritical, "Error"
        txtDatos(3).SetFocus
        valida_datos = False
        Exit Function
    Else
        If Not IsNumeric(txtDatos(3)) Then
            MsgBox "El codigo postal debe ser numérico.", vbCritical, "Error"
            txtDatos(3).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If txtDatos(6) = "" Then
        MsgBox "El CIF no puede estar en blanco.", vbCritical, "Error"
        txtDatos(6).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbPais.Text = "" Then
        MsgBox "El pais no puede estar en blanco.", vbCritical, "Error"
        cmbPais.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbProvincia.Text = "" Then
        MsgBox "La provincia no puede estar en blanco.", vbCritical, "Error"
        cmbProvincia.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipios.Text = "" Then
        MsgBox "El municipio no puede estar en blanco.", vbCritical, "Error"
        cmbMunicipios.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(18) = "" Then
        MsgBox "La cuenta contable no puede estar en blanco.", vbCritical, "Error"
        txtDatos(18).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbtarifa.getTEXTO = "" Then
        MsgBox "Indique la tarifa del cliente.", vbCritical, "Error"
        cmbtarifa.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta_Cliente()
    On Error GoTo fallo
    Dim oCliente As New clsCliente
    lbltitulo.Caption = "Modificacion de Cliente"
    lbltitulo.BackColor = &H80C0FF
    oCliente.CargaCliente (PK)
    With oCliente
        txtDatos(1) = .getNOMBRE
        txtDatos(2) = .getDIRECCION
        txtDatos(3) = .getCOD_POSTAL
        txtDatos(6) = .getCIF
        txtDatos(4) = .getTELEFONO
        txtDatos(5) = .getFAX
        txtDatos(8) = .getRESPONSABLE
        txtDatos(19) = .getRESPONSABLE_OTROS
        txtDatos(22) = .getRESPONSABLE_METROLOGIA
        txtDatos(9) = .getCARGO
        'txtdatos(1) = .getTIPO = ""    ' Ojo
        txtDatos(7) = .getEMAIL
        txtDatos(20) = .getEMAIL2
        txtDatos(0) = .getEMAIL_FACTURACION
        txtDatos(21) = .getEMAIL_METROLOGIA
        txtDatos(10) = .getBANCO
        txtDatos(11) = .getCUENTA
        txtDatos(13) = .getOBSERVACIONES
        txtDatos(12) = .getWEB
        'txtdatos(1) = .getANULADO = 0   ' Ojo
        txtDatos(14) = .getCENTRO
        txtDatos(15) = .getSECCION
        txtDatos(16) = PK
        txtDatos(17) = .getCLAVEWEB
        cmbPlant.MostrarElemento .getPLANT_ID
        If .getFACTURA_DETERMINACIONES = 0 Then
            chkFD.Value = Unchecked
        Else
            chkFD.Value = Checked
        End If
        If .getEADS = 0 Then
            chkEADS.Value = Unchecked
        Else
            chkEADS.Value = Checked
        End If
        chkAgroalimentario.Value = .getAGROALIMENTARIO
        chkAirbus = .getAIRBUS
        chkIberia = .getIBERIA
        chkExtranjero = .getEXTRANJERO
        chkIntra = .getINTRA
        opIdiomaFactura(.getIDIOMA_FACTURA).Value = True
        ' Pais
        Dim opais As New clsPais
        opais.CargarPais (.getPAIS_ID)
        cmbPais.BoundText = opais.getNOMBRE
        cmbPais.Text = opais.getNOMBRE
        Set opais = Nothing
        ' Provincia
        Dim oProvincia As New clsProvincias
        oProvincia.CargarProvincia (.getPROVINCIA_ID)
        cmbProvincia.BoundText = oProvincia.getNOMBRE
        cmbProvincia.Text = oProvincia.getNOMBRE
        Set oProvincia = Nothing
        ' Municipio
        Dim oMunicipio As New clsMunicipios
        oMunicipio.CargarMunicipio (.getMUNICIPIO_ID)
        cmbMunicipios.BoundText = oMunicipio.getNOMBRE
        cmbMunicipios.Text = oMunicipio.getNOMBRE
        Set oMunicipio = Nothing
        ' FP
        cmbFP.BoundText = .getFP_ID
        ' Tarifa
        cmbtarifa.MostrarElemento .getTARIFA_ID
        txtDatos(18) = .getCC
        cargarResponsables PK
        cargar_indicadores
        'E0110-I
        'E0200-I
'        If .getES_SUBCONTRATA = 0 Then
'            chkSubcontrata.value = Unchecked
'        Else
'            chkSubcontrata.value = Checked
'        End If
        'E0200-I
        'E0110-F
        'M1275-I
        chkFacturaElectronica.Value = .getFACTURA_ELECTRONICA
        'M1275-F
    End With
    Set oCliente = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del cliente.", vbCritical, Err.Description
End Sub

Public Sub desbloquear_controles()
    Dim i As Integer
    For i = 1 To 18
        txtDatos(i).Locked = False
    Next
    chkFD.Enabled = True
    cmbMunicipios.Locked = False
    cmbProvincia.Locked = False
    cmbPais.Locked = False
    cmbPlant.activar
End Sub

Public Function mover_datos() As clsCliente
    On Error GoTo fallo
    Dim oCliente As New clsCliente
    With oCliente
        .setNOMBRE = txtDatos(1)
        .setDIRECCION = txtDatos(2)
        If txtDatos(3) <> "" Then
            .setCOD_POSTAL = CLng(txtDatos(3))
        Else
            .setCOD_POSTAL = 0
        End If
        .setCIF = UCase(txtDatos(6))
        .setTELEFONO = txtDatos(4)
        .setFAX = txtDatos(5)
        .setRESPONSABLE = txtDatos(8)
        .setRESPONSABLE_OTROS = txtDatos(19)
        .setRESPONSABLE_METROLOGIA = txtDatos(22)
        .setCARGO = txtDatos(9)
        .setTIPO = "" ' Ojo
        .setEMAIL = txtDatos(7)
        .setEMAIL2 = txtDatos(20)
        .setEMAIL_FACTURACION = txtDatos(0)
        .setEMAIL_METROLOGIA = txtDatos(21)
        .setBANCO = txtDatos(10)
        .setCUENTA = txtDatos(11)
        .setOBSERVACIONES = txtDatos(13)
        .setWEB = txtDatos(12)
        .setANULADO = 0 ' Ojo
        .setCENTRO = txtDatos(14)
        .setSECCION = txtDatos(15)
        If cmbPlant.getTEXTO = "" Then
            .setPLANT_ID = 0
        Else
            .setPLANT_ID = cmbPlant.getPK_SALIDA
        End If
        If txtDatos(17) <> "" Then
            .setCLAVEWEB = txtDatos(17) ' ClaveWeb
        Else
            .setCLAVEWEB = generar_clave_web
        End If
        If chkFD.Value = Checked Then
            .setFACTURA_DETERMINACIONES = 1
        Else
            .setFACTURA_DETERMINACIONES = 0
        End If
        If chkEADS.Value = Checked Then
            .setEADS = 1
        Else
            .setEADS = 0
        End If
        .setAGROALIMENTARIO = chkAgroalimentario.Value
        .setAIRBUS = chkAirbus.Value
        .setIBERIA = chkIberia.Value
        .setEXTRANJERO = chkExtranjero.Value
        .setINTRA = chkIntra.Value
        .setIDIOMA_FACTURA = 0
        If opIdiomaFactura(1).Value = True Then
            .setIDIOMA_FACTURA = 1
        End If
        ' Pais
        If cmbPais.Text <> "" Then
            If IsNumeric(cmbPais.BoundText) Then
                .setPAIS_ID = cmbPais.BoundText
            Else
                Dim opais As New clsPais
                Dim pais As Long
                pais = opais.buscar(cmbPais.Text)
                If pais = 0 Then
                    opais.setNOMBRE = cmbPais.Text
                    .setPAIS_ID = opais.Insertar
                Else
                    .setPAIS_ID = pais
                End If
            End If
        End If
        ' Provincia
        If cmbProvincia.Text <> "" Then
            If IsNumeric(cmbProvincia.BoundText) Then
                .setPROVINCIA_ID = cmbProvincia.BoundText
            Else
                Dim oprov As New clsProvincias
                Dim PROVINCIA As Long
                PROVINCIA = oprov.buscar(cmbProvincia.Text)
                If PROVINCIA = 0 Then
                    oprov.setPAIS_ID = .getPAIS_ID
                    oprov.setNOMBRE = cmbProvincia.Text
                    .setPROVINCIA_ID = oprov.Insertar
                Else
                    .setPROVINCIA_ID = PROVINCIA
                End If
            End If
        End If
        ' Municipio
        If cmbMunicipios.Text <> "" Then
            If IsNumeric(cmbMunicipios.BoundText) Then
                .setMUNICIPIO_ID = cmbMunicipios.BoundText
            Else
                Dim omun As New clsMunicipios
                Dim municipio As Long
                municipio = omun.buscar(cmbMunicipios.Text)
                If municipio = 0 Then
                    omun.setPROVINCIA_ID = .getPROVINCIA_ID
                    omun.setNOMBRE = cmbMunicipios.Text
                    .setMUNICIPIO_ID = omun.Insertar
                Else
                    .setMUNICIPIO_ID = municipio
                End If
            End If
        End If
        ' Fp
        If cmbFP.BoundText = "" Then
            .setFP_ID = 0
        Else
            .setFP_ID = cmbFP.BoundText
        End If
        ' Tarifa
        .setTARIFA_ID = cmbtarifa.getPK_SALIDA
        .setCC = txtDatos(18)
        'E0111-I
        'E0200-I
'        If chkSubcontrata.value = Checked Then
'            .setES_SUBCONTRATA = 1
'        Else
'            .setES_SUBCONTRATA = 0
'        End If
        'E0200-F
        'E0111-F
        'M1275-I
        If chkFacturaElectronica.Value = Checked Then
            .setFACTURA_ELECTRONICA = 1
        Else
            .setFACTURA_ELECTRONICA = 0
        End If
        'M1275-F
    End With
    Set mover_datos = oCliente
    Set oCliente = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del cliente.", vbCritical, Err.Description
End Function
Public Sub cargar_paises()
    Dim opais As New clsPais
    Set cmbPais.RowSource = opais.Listado  'recorset devuelto por la funcion
    cmbPais.ListField = "nombre" 'campo que veo
    cmbPais.DataField = "nombre" 'campo asociado
    cmbPais.BoundColumn = "id_pais" 'lo que realmente envia
    Set opais = Nothing
End Sub
Public Sub cargar_provincias()
'    cmbProvincia.Text = ""
    If cmbPais.Text <> "" Then
     If IsNumeric(cmbPais.BoundText) Then
        Dim oProvincia As New clsProvincias
        Set cmbProvincia.RowSource = oProvincia.Listado(CInt(cmbPais.BoundText))  'recorset devuelto por la funcion
        cmbProvincia.ListField = "nombre" 'campo que veo
        cmbProvincia.DataField = "nombre" 'campo asociado
        cmbProvincia.BoundColumn = "id_provincia" 'lo que realmente envia
        Set oProvincia = Nothing
     End If
    End If
End Sub
Private Sub cargar_municipios()
'    cmbMunicipios.Text = ""
    If cmbProvincia.Text <> "" Then
     If IsNumeric(cmbProvincia.BoundText) Then
        Dim omuni As New clsMunicipios
        Set cmbMunicipios.RowSource = omuni.Listado(CInt(cmbProvincia.BoundText))
        cmbMunicipios.ListField = "nombre" 'campo que veo
        cmbMunicipios.DataField = "nombre" 'campo asociado
        cmbMunicipios.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
    End If
End Sub

Private Sub permisos()
    chkEADS.visible = True
    chkAirbus.visible = True
    chkIberia.visible = True
    cmdPedidos.visible = True
    cmdOfertas.visible = True
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        txtDatos(17).PasswordChar = ""
    End If
    frmIndicadores.visible = USUARIO.getPER_INDICADORES_CLIENTE
End Sub

Private Sub cabecera()
    With listaResponsables.ColumnHeaders
        .Add , , "ID_RESPONSABLE", 1, lvwColumnLeft
        .Add , , "Nombre", 3800, lvwColumnLeft
    End With
    With listaIndicadores.ColumnHeaders
        .Add , , "Año", 800, lvwColumnLeft
        .Add , , "NºMuestras", 1500, lvwColumnCenter
        .Add , , "Importe", 1500, lvwColumnRight
    End With

End Sub

Private Sub cargar_indicadores()
    Dim c As String
    c = " SELECT '1',M.ANNO, COUNT(*) " & _
        "   FROM CLIENTES C LEFT JOIN MUESTRAS M ON C.ID_CLIENTE = M.CLIENTE_ID " & _
        "  where c.id_cliente = " & PK & _
        "  GROUP BY M.ANNO " & _
        "  Union " & _
        " SELECT '2',YEAR(DP.FECHA_FACTURA), SUM(DP.TOTAL) " & _
        "   FROM CLIENTES C LEFT JOIN DOCS_PAGO DP ON C.ID_CLIENTE = DP.CLIENTE_ID " & _
        "  where c.id_cliente = " & PK & " And DP.ANULADO = 0 And DP.tipo = 2 " & _
        "  GROUP BY YEAR(DP.FECHA_FACTURA) " & _
        "  ORDER BY 2 DESC,1 "
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    Dim i As Integer
    If rs.RecordCount > 0 Then
        Do
            Dim encontrado As Boolean
            encontrado = False
            For i = 1 To listaIndicadores.ListItems.Count
                If Not IsNull(rs(1)) Then
                    If CInt(rs(1)) = CInt(listaIndicadores.ListItems(i)) Then
                        encontrado = True
                        Exit For
                    End If
                End If
            Next
            If encontrado Then
                If Not IsNull(rs.Fields(2)) Then
                    listaIndicadores.ListItems(i).SubItems(CInt(rs(0))) = rs.Fields(2)
                Else
                    listaIndicadores.ListItems(i).SubItems(CInt(rs(0))) = "0"
                End If
            Else
                If Not IsNull(rs.Fields(1)) Then
                    With listaIndicadores.ListItems.Add(, , rs.Fields(1))
                        .SubItems(CInt(rs(0))) = rs.Fields(2)
                    End With
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    For i = 1 To listaIndicadores.ListItems.Count
        If listaIndicadores.ListItems(i).SubItems(2) <> "" Then
            listaIndicadores.ListItems(i).SubItems(2) = moneda(listaIndicadores.ListItems(i).SubItems(2))
        End If
    Next
    Set rs = Nothing
End Sub

Private Sub almacenarResponsables(cliente As Long)
            ' Responsables
            Dim oCR As New clsClientes_responsables
            oCR.Eliminar cliente
            Dim i As Integer
            For i = 1 To listaResponsables.ListItems.Count
                With oCR
                    .setCLIENTE_ID = cliente
                    .setRESPONSABLE_ID = listaResponsables.ListItems(i).Text
                    .Insertar
                End With
            Next
            Set oCR = Nothing

End Sub
Private Sub cargarResponsables(cliente As Long)
            ' Responsables
        Dim oCR As New clsClientes_responsables
        listaResponsables.ListItems.Clear
        Dim rs As ADODB.Recordset
        Set rs = oCR.Listado(cliente)
        If rs.RecordCount > 0 Then
            Do
                With listaResponsables.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                End With
            
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oCR = Nothing
End Sub

