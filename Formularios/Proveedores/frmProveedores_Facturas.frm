VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProveedores_Facturas 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   12870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17700
   Icon            =   "frmProveedores_Facturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12870
   ScaleWidth      =   17700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmRevision 
      BackColor       =   &H00C0C0C0&
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
      Height          =   4425
      Left            =   3375
      TabIndex        =   103
      Top             =   3285
      Visible         =   0   'False
      Width           =   10005
      Begin VB.TextBox txtRevisionEnvio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   119
         Top             =   2475
         Width           =   3720
      End
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENVIAR A REVISAR"
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
         Left            =   1620
         TabIndex        =   118
         Top             =   1035
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.TextBox txtRevisionFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   117
         Top             =   2880
         Width           =   3720
      End
      Begin VB.TextBox txtRevisionMotivo 
         Appearance      =   0  'Flat
         Height          =   1035
         Left            =   1620
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   115
         Top             =   1395
         Width           =   7005
      End
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RECHAZADA"
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
         Index           =   2
         Left            =   5760
         TabIndex        =   113
         Top             =   1035
         Width           =   1860
      End
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "APROBADA"
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
         Left            =   4005
         TabIndex        =   112
         Top             =   1035
         Width           =   1725
      End
      Begin VB.CommandButton cmdRevisionCancelar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   8010
         Picture         =   "frmProveedores_Facturas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   3555
         Width           =   1860
      End
      Begin VB.CommandButton cmdRevisionModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   735
         Left            =   6075
         Picture         =   "frmProveedores_Facturas.frx":6B5C
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   3555
         Width           =   1860
      End
      Begin pryCombo.miCombo cmbRevisada_por 
         Height          =   345
         Left            =   1620
         TabIndex        =   106
         Top             =   540
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Envío"
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
         Index           =   7
         Left            =   90
         TabIndex        =   120
         Top             =   2565
         Width           =   1110
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Revisión"
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
         Index           =   6
         Left            =   90
         TabIndex        =   116
         Top             =   2970
         Width           =   1335
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Motivo Rechazo"
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
         Index           =   5
         Left            =   90
         TabIndex        =   114
         Top             =   1755
         Width           =   1395
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
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
         Index           =   4
         Left            =   90
         TabIndex        =   111
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label lblRevision 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REVISION DE FACTURA"
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
         Left            =   135
         TabIndex        =   109
         Top             =   90
         Width           =   9735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario Revisor"
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
         Left            =   90
         TabIndex        =   108
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   90
         TabIndex        =   107
         Top             =   540
         Width           =   9735
      End
   End
   Begin VB.CheckBox chkNoEnviadas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar NO ENVIADAS"
      Height          =   240
      Left            =   12825
      TabIndex        =   99
      Top             =   990
      Width           =   2175
   End
   Begin VB.CheckBox chkIncidencias 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar INCIDENCIAS"
      Height          =   240
      Left            =   10755
      TabIndex        =   90
      Top             =   990
      Width           =   1950
   End
   Begin VB.CheckBox chkPagoPrevisto 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar Pago Previsto"
      Height          =   240
      Left            =   8685
      TabIndex        =   89
      Top             =   990
      Width           =   1995
   End
   Begin VB.CheckBox chkVencidas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar las vencidas"
      Height          =   240
      Left            =   6750
      TabIndex        =   88
      Top             =   990
      Width           =   1860
   End
   Begin VB.CheckBox chkPendientesPago 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar sólo pdtes. pago"
      Height          =   240
      Left            =   4545
      TabIndex        =   87
      Top             =   990
      Width           =   2130
   End
   Begin VB.Frame frmGenera 
      BackColor       =   &H00C0C0C0&
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
      Height          =   2490
      Left            =   3375
      TabIndex        =   73
      Top             =   3870
      Visible         =   0   'False
      Width           =   10005
      Begin VB.CommandButton cmdGenerarFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   735
         Left            =   6075
         Picture         =   "frmProveedores_Facturas.frx":D3AE
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1620
         Width           =   1860
      End
      Begin VB.CommandButton cmdCancelarFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   8010
         Picture         =   "frmProveedores_Facturas.frx":13C00
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1620
         Width           =   1860
      End
      Begin pryCombo.miCombo cmbProveedorCambio 
         Height          =   345
         Left            =   1035
         TabIndex        =   76
         Top             =   990
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   609
      End
      Begin VB.Label lblCambioProveedor 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   90
         TabIndex        =   79
         Top             =   540
         Width           =   9735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
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
         Index           =   1
         Left            =   90
         TabIndex        =   78
         Top             =   1035
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "INDIQUE EL NUEVO PROVEEDOR"
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
         Index           =   0
         Left            =   180
         TabIndex        =   77
         Top             =   90
         Width           =   9420
      End
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   16020
      TabIndex        =   46
      Top             =   10710
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Adodc1"
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   16020
      TabIndex        =   45
      Top             =   10395
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   5055
      Left            =   15525
      TabIndex        =   34
      Top             =   540
      Width           =   2085
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   735
         Index           =   0
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2115
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   690
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4275
         Width           =   915
      End
      Begin VB.CommandButton cmdGestor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestor"
         Height          =   690
         Left            =   90
         Picture         =   "frmProveedores_Facturas.frx":1A452
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3555
         Width           =   915
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escanear"
         Height          =   690
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4275
         Width           =   915
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   690
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3555
         Width           =   915
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Factura en Acrobat Reader"
         Height          =   960
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2565
         Width           =   1905
      End
      Begin AcroPDFLibCtl.AcroPDF pdf1 
         Height          =   2265
         Left            =   90
         TabIndex        =   35
         Top             =   180
         Width           =   1860
         _cx             =   5080
         _cy             =   5080
      End
   End
   Begin VB.TextBox txtiva 
      Height          =   375
      Left            =   16425
      TabIndex        =   29
      Top             =   11070
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txtmov 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   16020
      TabIndex        =   26
      Top             =   10035
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   16515
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   11880
      Width           =   1095
   End
   Begin VB.Frame fmov 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   45
      TabIndex        =   18
      Top             =   6930
      Width           =   15435
      Begin VB.CheckBox chkPrevista 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pago Previsto"
         Height          =   195
         Left            =   3375
         TabIndex        =   8
         Top             =   2070
         Width           =   1320
      End
      Begin VB.CheckBox chkIncidencia 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INCIDENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3960
         TabIndex        =   13
         Top             =   4905
         Width           =   2040
      End
      Begin VB.CheckBox chkEnviada 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENVIADA"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   4905
         Width           =   2040
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   4650
         Left            =   6300
         TabIndex        =   59
         Top             =   135
         Width           =   8880
         Begin VB.TextBox txtmov 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   5
            Left            =   3900
            TabIndex        =   61
            Top             =   900
            Width           =   975
         End
         Begin pryCombo.miCombo cmbCC 
            Height          =   345
            Left            =   1275
            TabIndex        =   62
            Top             =   180
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbSubcuenta 
            Height          =   345
            Left            =   1275
            TabIndex        =   63
            Top             =   540
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   609
         End
         Begin MSComctlLib.ListView listaFamilias 
            Height          =   3240
            Left            =   135
            TabIndex        =   64
            Top             =   1305
            Width           =   8250
            _ExtentX        =   14552
            _ExtentY        =   5715
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
         Begin XtremeSuiteControls.PushButton cmdAnadirFamilia 
            Height          =   345
            Left            =   8415
            TabIndex        =   65
            ToolTipText     =   "Añadir registro"
            Top             =   1935
            Width           =   390
            _Version        =   851970
            _ExtentX        =   688
            _ExtentY        =   609
            _StockProps     =   79
            Appearance      =   5
            Picture         =   "frmProveedores_Facturas.frx":1AD1C
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarFamilia 
            Height          =   345
            Left            =   8415
            TabIndex        =   66
            ToolTipText     =   "Eliminar registro seleccionado"
            Top             =   3510
            Width           =   390
            _Version        =   851970
            _ExtentX        =   688
            _ExtentY        =   609
            _StockProps     =   79
            Appearance      =   5
            Picture         =   "frmProveedores_Facturas.frx":2157E
         End
         Begin XtremeSuiteControls.PushButton cmdModificarFamilia 
            Height          =   345
            Left            =   8415
            TabIndex        =   71
            ToolTipText     =   "Modificar Registro Seleccionado"
            Top             =   2700
            Width           =   390
            _Version        =   851970
            _ExtentX        =   688
            _ExtentY        =   609
            _StockProps     =   79
            Appearance      =   5
            Picture         =   "frmProveedores_Facturas.frx":27DE0
         End
         Begin VB.TextBox txtmov 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   1275
            TabIndex        =   60
            Top             =   900
            Width           =   1605
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Subcuenta"
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   70
            Top             =   585
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Familia"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   69
            Top             =   225
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "I.V.A. (%):"
            Height          =   195
            Index           =   7
            Left            =   3075
            TabIndex        =   68
            Top             =   960
            Width           =   690
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Base:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   67
            Top             =   945
            Width           =   405
         End
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   9
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   4860
         Width           =   1200
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   8
         Left            =   8895
         TabIndex        =   14
         Top             =   4860
         Width           =   1200
      End
      Begin VB.TextBox txtmov 
         Appearance      =   0  'Flat
         Height          =   1035
         Index           =   7
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3780
         Width           =   5970
      End
      Begin VB.CheckBox chkPago 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagada"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   2385
         Width           =   960
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   6
         Left            =   1230
         MaxLength       =   25
         TabIndex        =   4
         Top             =   945
         Width           =   4845
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   4
         Left            =   13050
         TabIndex        =   16
         Top             =   4860
         Width           =   1515
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   3
         Left            =   11205
         TabIndex        =   15
         Top             =   4860
         Width           =   1245
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   3
         Top             =   585
         Width           =   4845
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   315
         Left            =   1230
         TabIndex        =   1
         Top             =   225
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
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
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   1305
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fVencimiento 
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   2025
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fPrevista 
         Height          =   315
         Left            =   4830
         TabIndex        =   9
         Top             =   2025
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin VB.Frame frmPagada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   90
         TabIndex        =   91
         Top             =   2385
         Width           =   6090
         Begin VB.TextBox txtmov 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Index           =   10
            Left            =   3060
            TabIndex        =   96
            Top             =   315
            Width           =   1155
         End
         Begin MSComCtl2.DTPicker fPago 
            Height          =   315
            Left            =   1125
            TabIndex        =   92
            Top             =   315
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
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
            Format          =   51445761
            CurrentDate     =   38002
         End
         Begin pryCombo.miCombo cmbSubcuentaPago 
            Height          =   345
            Left            =   1125
            TabIndex        =   93
            Top             =   675
            Width           =   4860
            _ExtentX        =   8573
            _ExtentY        =   609
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   300
            Left            =   4230
            TabIndex        =   98
            ToolTipText     =   "Modificar Registro Seleccionado"
            Top             =   315
            Width           =   1785
            _Version        =   851970
            _ExtentX        =   3149
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "Consultar Remesa"
            Appearance      =   5
            Picture         =   "frmProveedores_Facturas.frx":2E642
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Remesa"
            Height          =   195
            Index           =   17
            Left            =   2430
            TabIndex        =   97
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            Height          =   195
            Index           =   16
            Left            =   135
            TabIndex        =   95
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sub.Pago"
            Height          =   195
            Index           =   15
            Left            =   135
            TabIndex        =   94
            Top             =   720
            Width           =   705
         End
      End
      Begin MSDataListLib.DataCombo cmbVencimiento 
         Height          =   315
         Left            =   1230
         TabIndex        =   6
         Top             =   1665
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fechaFactura 
         Height          =   315
         Left            =   4740
         TabIndex        =   2
         Top             =   225
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
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
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Factura"
         Height          =   195
         Index           =   19
         Left            =   3600
         TabIndex        =   102
         Top             =   255
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vencimiento"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   100
         Top             =   1725
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Base"
         Height          =   195
         Index           =   14
         Left            =   6345
         TabIndex        =   54
         Top             =   4920
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   12510
         TabIndex        =   23
         Top             =   4905
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Retención"
         Height          =   195
         Index           =   13
         Left            =   8055
         TabIndex        =   49
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   44
         Top             =   3555
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   43
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Vencimiento"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   37
         Top             =   2070
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A. Importe"
         Height          =   195
         Index           =   4
         Left            =   10170
         TabIndex        =   22
         Top             =   4905
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Contable"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   255
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N. Factura"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   19
         Top             =   1365
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5220
      Left            =   45
      TabIndex        =   0
      Top             =   1335
      Width           =   15450
      _ExtentX        =   27252
      _ExtentY        =   9208
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   15615
      Top             =   10260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin pryCombo.miCombo cmbProveedor 
      Height          =   345
      Left            =   990
      TabIndex        =   48
      Top             =   585
      Width           =   14490
      _ExtentX        =   25559
      _ExtentY        =   609
   End
   Begin XtremeSuiteControls.PushButton cmdRelacionar 
      Height          =   795
      Left            =   15525
      TabIndex        =   52
      Top             =   5670
      Width           =   2100
      _Version        =   851970
      _ExtentX        =   3704
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "RELACIONES"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":34EA4
   End
   Begin XtremeSuiteControls.PushButton cmdAnadir 
      Height          =   480
      Index           =   0
      Left            =   13455
      TabIndex        =   55
      Top             =   12330
      Width           =   2010
      _Version        =   851970
      _ExtentX        =   3545
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Añadir Factura"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":3B706
   End
   Begin XtremeSuiteControls.PushButton cmdAnadir 
      Height          =   480
      Index           =   1
      Left            =   11430
      TabIndex        =   56
      Top             =   12330
      Width           =   2010
      _Version        =   851970
      _ExtentX        =   3545
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Modificar"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":41F68
   End
   Begin XtremeSuiteControls.PushButton cmdLimpiar 
      Height          =   480
      Left            =   9405
      TabIndex        =   57
      Top             =   12330
      Width           =   2010
      _Version        =   851970
      _ExtentX        =   3545
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Limpiar Campos"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":487CA
   End
   Begin XtremeSuiteControls.PushButton cmdborrar 
      Height          =   480
      Left            =   45
      TabIndex        =   58
      Top             =   12330
      Width           =   2010
      _Version        =   851970
      _ExtentX        =   3545
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Eliminar Factura"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":4F02C
   End
   Begin XtremeSuiteControls.PushButton cmdCambiarProveedor 
      Height          =   795
      Left            =   15525
      TabIndex        =   72
      Top             =   6525
      Width           =   2100
      _Version        =   851970
      _ExtentX        =   3704
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Cambiar Factura de Proveedor"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":5588E
   End
   Begin XtremeSuiteControls.PushButton cmdEnviarMail 
      Height          =   795
      Left            =   15525
      TabIndex        =   80
      Top             =   7380
      Width           =   2100
      _Version        =   851970
      _ExtentX        =   3704
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Enviar Marcadas por Correo"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":5C0F0
   End
   Begin MSComCtl2.DTPicker fdesde 
      Height          =   330
      Left            =   990
      TabIndex        =   81
      Top             =   945
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
      Format          =   51445761
      CurrentDate     =   38002
   End
   Begin MSComCtl2.DTPicker fhasta 
      Height          =   330
      Left            =   2970
      TabIndex        =   82
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
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
      Format          =   51445761
      CurrentDate     =   38002
   End
   Begin XtremeSuiteControls.PushButton cmdMarcar 
      Height          =   300
      Index           =   0
      Left            =   45
      TabIndex        =   85
      Top             =   6570
      Width           =   1560
      _Version        =   851970
      _ExtentX        =   2752
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Marcar Todas"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":62952
   End
   Begin XtremeSuiteControls.PushButton cmdMarcar 
      Height          =   300
      Index           =   1
      Left            =   1620
      TabIndex        =   86
      Top             =   6570
      Width           =   1695
      _Version        =   851970
      _ExtentX        =   2990
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Desmarcar Todas"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":691B4
   End
   Begin XtremeSuiteControls.PushButton cmdRevisionFactura 
      Height          =   795
      Left            =   15525
      TabIndex        =   110
      Top             =   8235
      Width           =   2100
      _Version        =   851970
      _ExtentX        =   3704
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "REVISIÓN DE FACTURA"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas.frx":6FA16
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Facturas.frx":76278
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Facturas.frx":7670F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Facturas.frx":76BA5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "hasta"
      Height          =   195
      Index           =   2
      Left            =   2430
      TabIndex        =   84
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "F.Factura"
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
      Index           =   0
      Left            =   45
      TabIndex        =   83
      Top             =   1035
      Width           =   825
   End
   Begin VB.Label lblRetencion 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10530
      TabIndex        =   51
      Top             =   6570
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retención"
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
      Height          =   285
      Index           =   3
      Left            =   9270
      TabIndex        =   50
      Top             =   6570
      Width           =   1275
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedor"
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
      Index           =   12
      Left            =   45
      TabIndex        =   47
      Top             =   630
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Base"
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
      Height          =   285
      Index           =   1
      Left            =   3690
      TabIndex        =   33
      Top             =   6570
      Width           =   870
   End
   Begin VB.Label lblBase 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4545
      TabIndex        =   32
      Top             =   6570
      Width           =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Height          =   285
      Index           =   0
      Left            =   6570
      TabIndex        =   31
      Top             =   6570
      Width           =   1005
   End
   Begin VB.Label lblIVA 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7560
      TabIndex        =   30
      Top             =   6570
      Width           =   1725
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique las facturas y documentos del proveedor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   28
      Top             =   270
      Width           =   3435
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas del proveedor : "
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
      Left            =   90
      TabIndex        =   27
      Top             =   0
      Width           =   2625
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13320
      TabIndex        =   25
      Top             =   6570
      Width           =   1950
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Facturado"
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
      Height          =   285
      Index           =   2
      Left            =   12510
      TabIndex        =   24
      Top             =   6570
      Width           =   825
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   17640
   End
   Begin VB.Label lblContabilizada 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2430
      TabIndex        =   101
      Top             =   12375
      Width           =   6810
   End
End
Attribute VB_Name = "frmProveedores_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public PK_FACTURA_ID As Long
'M1257-I
Public TOBJETO As Long
Public COBJETO As Long
'M1257-F
Private Enum COLS
    C_ID = 0
    C_fecha = 1
    C_concepto = 2
    C_NUMERO = 3
    C_FAMILIA = 4
    C_SUBCUENTA = 5
    C_BASE = 6
    C_IVA_PORCENTAJE = 7
    C_IVA = 8
    C_RETENCION = 9
    C_total = 10
    C_FP = 11
    C_vencimiento = 12
    C_FP_ID = 13
    C_TOBJETO = 14
    C_cOBJETO = 15
    C_PAGADA = 16
    C_ENVIADA = 17
End Enum
Private Enum COLS_F
    CF_ID = 0
    CF_FAMILIA = 1
    CF_SUBCUENTA = 2
    CF_BASE = 3
    CF_IVA_PORCENTAJE = 4
    CF_IVA = 5
    CF_Importe = 6
    CF_FAMILIA_ID = 7
    CF_SUBCUENTA_ID = 8
End Enum
       
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº", 1200, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Concepto", 3000, lvwColumnCenter
        .Add , , "Numero", 1800, lvwColumnCenter
        .Add , , "Familia", 1, lvwColumnLeft
        .Add , , "Subcuenta", 1, lvwColumnLeft
        .Add , , "Base", 1200, lvwColumnRight
        .Add , , "Iva %", 1, lvwColumnCenter
        .Add , , "Iva", 1100, lvwColumnRight
        .Add , , "Retención", 1100, lvwColumnRight
        .Add , , "Total", 1200, lvwColumnRight
        .Add , , "F.P.", 1300, lvwColumnCenter
        .Add , , "F.Vencimiento", 1050, lvwColumnCenter
        .Add , , "F.Pago", 1, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "Pagada", 350, lvwColumnCenter
        .Add , , "Enviada", 350, lvwColumnCenter
        .Add , , "REV.", 350, lvwColumnCenter
    End With
    With listaFamilias.ColumnHeaders
        .Add , , "Id", 1, lvwColumnLeft
        .Add , , "Familia", 2000, lvwColumnLeft
        .Add , , "Subcuenta", 2150, lvwColumnLeft
        .Add , , "Base", 1000, lvwColumnRight
        .Add , , "Iva %", 700, lvwColumnCenter
        .Add , , "Iva", 1000, lvwColumnRight
        .Add , , "Importe", 1000, lvwColumnRight
        .Add , , "FamiliaID", 1, lvwColumnLeft
        .Add , , "SubcuentaID", 1, lvwColumnLeft
    End With
End Sub

Private Sub chkIncidencias_Click()
    cargar_lista
End Sub

Private Sub chkNoEnviadas_Click()
    cargar_lista
End Sub

Private Sub chkPago_Click()
    If chkPago.Value = Checked Then
        frmPagada.Enabled = True
        fPago.Value = Date
        chkPrevista.Value = Unchecked
'        fPago.Enabled = True
'        cmbSubcuentaPago.activar
    Else
        frmPagada.Enabled = False
'        fPago.Enabled = False
'        cmbSubcuentaPago.desactivar
    End If
End Sub

Private Sub chkPagoPrevisto_Click()
    cargar_lista
End Sub

Private Sub chkPendientesPago_Click()
    cargar_lista
End Sub

Private Sub chkPrevista_Click()
    If chkPrevista.Value = Checked Then
'        fPrevista.value = Date
        fPrevista.Enabled = True
'        chkPago.value = Unchecked
    Else
        fPrevista.Enabled = False
    End If
End Sub

Private Sub chkVencidas_Click()
    cargar_lista
End Sub

Private Sub cmbFP_Change()
    Dim oDeco As New clsDecodificadora
    cmbVencimiento.Enabled = False
    cmbVencimiento.BoundText = "0"
    
    If cmbFP.Text <> "" Then
        oDeco.Carga_valor DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP, cmbFP.BoundText
        Dim s() As String
        Dim s2() As String
        s = Split(oDeco.getPARAMETROS, ";")
        ' Vencimiento
        s2 = Split(s(0), "=")
        If s2(1) = "1" Then
            cmbVencimiento.Enabled = True
        End If
    End If
    Set oDeco = Nothing
End Sub

Private Sub cmbProveedor_change()
    If cmbProveedor.getTEXTO <> "" Then
        PK = cmbProveedor.getPK_SALIDA
        cmdLimpiar_Click
        
        cargar_proveedor
    End If
End Sub

Private Sub cmbVencimiento_Change()
    calcularVencimiento
End Sub
Private Sub cmdAnadir_Click(Index As Integer)
   On Error GoTo cmdAnadir_Click_Error
    If Index = 1 And lista.ListItems.Count = 0 Then Exit Sub
    If validar = False Then
        Exit Sub
    End If
    Dim oProveedor_Factura As New clsProveedores_Facturas
    With oProveedor_Factura
        .setPROVEEDOR_ID = PK
        .setFECHA = fecha_bd(fecha)
        .setFECHA_FACTURA = fecha_bd(fechaFactura)
        .setNUMERO = txtmov(0)
        .setCONCEPTO = txtmov(6)
        .setFORMAPAGO = cmbFP.BoundText
        If cmbVencimiento.Text = "" Then
            .setVENCIMIENTO_ID = 0
        Else
            .setVENCIMIENTO_ID = cmbVencimiento.BoundText
        End If
        .setFAMILIA_ID = cmbCC.getPK_SALIDA
        .setSUBCUENTA = cmbSubcuenta.getPK_SALIDA
        .setBI = moneda_bd(txtmov(9))
        .setIVA_PORCENTAJE = 0
        .setRETENCION = moneda_bd(txtmov(8))
        .setIVA = moneda_bd(txtmov(3))
        .setTOTAL = moneda_bd(txtmov(4))
        
        .setF_VENCIMIENTO = fecha_bd(fVencimiento)
        If chkPago.Value = Checked Then
            .setF_PAGO = fecha_bd(fPago)
            If cmbSubcuentaPago.getTEXTO = "" Then
                .setSUBCUENTA_PAGO = 0
            Else
                .setSUBCUENTA_PAGO = cmbSubcuentaPago.getPK_SALIDA
            End If
        Else
            .setF_PAGO = "0000-00-00"
            .setSUBCUENTA_PAGO = 0
        End If
        If chkPrevista.Value = Checked Then
            .setF_PREVISTA = "'" & fecha_bd(fPrevista) & "'"
        Else
            .setF_PREVISTA = "null"
        End If
        .setOBSERVACIONES = txtmov(7)
        If Index = 0 Then
            .setENVIADA = 0
        Else
            .setENVIADA = chkEnviada.Value
        End If
        .setINCIDENCIA = chkIncidencia.Value
        'M1257-I
'        .setTOBJETO = TOBJETO
'        .setCOBJETO = COBJETO
        'M1257-F
        .setREMESA_ID = 0
        Dim ID As Long
        If Index = 0 Then
            ID = .Insertar
        Else
            ID = .Modificar(lista.ListItems(lista.selectedItem.Index).Text)
        End If
        If ID = 0 Then
            MsgBox "Error al actualizar la factura de proveedor.", vbCritical, App.Title
            Exit Sub
        End If
        ' CUENTAS CONTABLES
        Dim oPFF As New clsProveedores_facturas_fam
        If Index <> 0 Then
            oPFF.Eliminar ID
        End If
        Dim i As Integer
        For i = 1 To listaFamilias.ListItems.Count
            With oPFF
                .setFACTURA_ID = ID
                .setFAMILIA_ID = listaFamilias.ListItems(i).SubItems(COLS_F.CF_FAMILIA_ID)
                .setSUBCUENTA_ID = listaFamilias.ListItems(i).SubItems(COLS_F.CF_SUBCUENTA_ID)
                
                .setBI = moneda_bd(listaFamilias.ListItems(i).SubItems(COLS_F.CF_BASE))
                .setIVA_PORCENTAJE = moneda_bd(listaFamilias.ListItems(i).SubItems(COLS_F.CF_IVA_PORCENTAJE))
                .setIVA = moneda_bd(listaFamilias.ListItems(i).SubItems(COLS_F.CF_IVA))
                .setIMPORTE = moneda_bd(listaFamilias.ListItems(i).SubItems(COLS_F.CF_Importe))
                
                If .Insertar = 0 Then
                    MsgBox "Error al insertar las SUBCUENTAS de las facturas.", vbCritical, App.Title
                End If
            End With
        Next
        Set oPFF = Nothing
        ' CREAR LA RELACION SI VIENE DESDE UN TOBJETO
        If TOBJETO <> 0 And COBJETO <> 0 Then
            Dim oPFR As New clsProveedores_facturas_rel
            With oPFR
                .setFACTURA_ID = ID
                .setTOBJETO = TOBJETO
                .setCOBJETO = COBJETO
                .Insertar
            End With
        End If
        ' ABRIR EL ESCANER PARA ADJUNTAR LA FACTURA
        If ID <> 0 And Index = 0 Then
            frmEscaner.Show 1
            If documento_escaner <> "" Then
                Dim nombreNuevo As String
                nombreNuevo = txtmov(6)
                If Trim(nombreNuevo) <> "" Then
                    datos(4).Text = documento_escaner
                    datos(0).Text = nombreNuevo & ".pdf"
                    adjuntar ID
                End If
            End If
        End If
        If ID <> 0 Then
            cmdLimpiar_Click
            cargar_lista
            For i = 1 To lista.ListItems.Count
                If CLng(lista.ListItems(i)) = ID Then
                    lista.ListItems(i).EnsureVisible
                    lista.ListItems(i).Selected = True
                    Exit For
                End If
            Next
        End If
    
    End With
    Set oProveedor_Factura = Nothing

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmProveedores_Facturas"

End Sub

Private Sub cmdAnadirFamilia_Click()
   On Error GoTo cmdAnadirFamilia_Click_Error
    Dim i As Integer
    Dim encontrado As Boolean
    Dim IMPORTE As Double
    
    If Trim(cmbCC.getTEXTO) = "" Then
        MsgBox "Debe indicar la Familia.", vbExclamation, App.Title
        Exit Sub
    End If
    If Trim(cmbSubcuenta.getTEXTO) = "" Then
        MsgBox "Debe indicar la Subcuenta.", vbExclamation, App.Title
        Exit Sub
    End If
    If Trim(txtmov(2)) = "" Then
        MsgBox "Debe indicar la base.", vbExclamation, App.Title
        Exit Sub
    End If
    If Trim(txtmov(5)) = "" Then
        MsgBox "Debe indicar el porcentaje de IVA.", vbExclamation, App.Title
        Exit Sub
    End If
    
    'Búsqueda de repeticiones
    encontrado = False
    i = 1
    IMPORTE = moneda(txtmov(2) + ((txtmov(2) * CInt(txtmov(5)) / 100)))
    
    Do While i <= listaFamilias.ListItems.Count And encontrado = False
        If listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_FAMILIA_ID) = cmbCC.getPK_SALIDA And listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_SUBCUENTA_ID) = cmbSubcuenta.getPK_SALIDA Then
            encontrado = True
'JGM            If MsgBox("La subcuenta" & cmbSubcuenta.getTEXTO & " ya existe. ¿Desea reemplazar el actual precio base: " & listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_BASE) & ", por  " & moneda(txtmov(2)) & "?", vbYesNo) = vbYes Then
'JGM                listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_BASE) = moneda(txtmov(2))
'JGM                listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_IVA_PORCENTAJE) = txtmov(5)
'JGM                listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_IVA) = moneda((txtmov(2) * CInt(txtmov(5)) / 100))
'JGM                listaFamilias.ListItems.Item(i).SubItems(COLS_F.CF_Importe) = IMPORTE
'JGM            End If
            MsgBox "Información : La familia " & cmbCC.getTEXTO & " y la subcuenta " & cmbSubcuenta.getTEXTO & " ya existen en la lista.", vbCritical, App.Title
'            Exit Sub
        End If
        i = i + 1
    Loop
    
'    If Not encontrado Then
        With listaFamilias.ListItems.Add(, , "0") ' ID
            .SubItems(COLS_F.CF_FAMILIA) = cmbCC.getTEXTO
            .SubItems(COLS_F.CF_FAMILIA_ID) = cmbCC.getPK_SALIDA
            .SubItems(COLS_F.CF_SUBCUENTA) = cmbSubcuenta.getTEXTO
            .SubItems(COLS_F.CF_SUBCUENTA_ID) = cmbSubcuenta.getPK_SALIDA
            .SubItems(COLS_F.CF_BASE) = moneda(txtmov(2))
            .SubItems(COLS_F.CF_IVA_PORCENTAJE) = txtmov(5)
            .SubItems(COLS_F.CF_IVA) = moneda((txtmov(2) * CInt(txtmov(5)) / 100))
            .SubItems(COLS_F.CF_Importe) = moneda(CStr(IMPORTE))
        End With
'    End If
    calcularTotal
    
    cmbCC.limpiar
    cmbSubcuenta.limpiar
    txtmov(2) = ""
  '  txtmov(5) = ""
    txtmov(2).BackColor = &H80000005
    txtmov(5).BackColor = &H80000005
    
    cmbCC.SetFocus

   On Error GoTo 0
   Exit Sub

cmdAnadirFamilia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirFamilia_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub cmdCambiarProveedor_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oFP As New clsProveedores_Facturas
    oFP.Carga lista.ListItems(lista.selectedItem.Index)
    If oFP.getF_CONTABILIZADA <> 0 Then
        MsgBox "No se le puede cambiar el proveedor. La factura ya esta contabilizada.", vbCritical, App.Title
        Exit Sub
    End If
    Set oFP = Nothing
    llenar_combo cmbProveedorCambio, New clsProveedor, 0, frmProveedores_Detalle, ""
    lblCambioProveedor = "Asiento : " & lista.ListItems(lista.selectedItem.Index) & " -> " & lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_concepto) & ", de fecha : " & lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_fecha)
    frmGenera.visible = True
End Sub

Private Sub cmdCancelarFactura_Click()
    frmGenera.visible = False
End Sub

Private Sub cmdEliminarFamilia_Click()
    If listaFamilias.ListItems.Count > 0 Then
        listaFamilias.ListItems.Remove listaFamilias.selectedItem.Index
        calcularTotal
    End If
End Sub

Private Sub cmdEnviarMail_Click()
    Dim i As Integer
    Dim existen As Boolean
   On Error GoTo cmdEnviarMail_Click_Error

    existen = False
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            existen = True
        End If
    Next
    
    If existen Then
        Dim ADJUNTOS As String
        Dim salida As String
        Dim oD As New clsDocumentacion
        Dim oPF As New clsProveedores_Facturas
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                salida = oD.CargarProveedorFacturas(lista.ListItems(i).Text, False)
                If Dir(salida) = "" Then
                    MsgBox "Error al cargar las facturas.", vbCritical, App.Title
                    Exit Sub
                End If
                ADJUNTOS = ADJUNTOS & salida & ";"
                ' Marcar factura como enviada
                oPF.marcarEnviada CLng(lista.ListItems(i).Text)
            End If
        Next
        Set oD = Nothing
        genera_correo "", "", "", ADJUNTOS, Me.hdc, True
    End If
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdEnviarMail_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEnviarMail_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle
        datos(4).Text = cd.FileName
    End If
End Sub

Private Sub cmdGenerarFactura_Click()
   On Error GoTo cmdGenerarFactura_Click_Error

    If cmbProveedorCambio.getTEXTO = "" Then
        MsgBox "Indique el nuevo proveedor.", vbCritical, App.Title
        Exit Sub
    End If
    ' Cambiar el proveedor
    Dim oFP As New clsProveedores_Facturas
    oFP.modificarProveedor lista.ListItems(lista.selectedItem.Index), cmbProveedorCambio.getPK_SALIDA
    Set oFP = Nothing
    cargar_lista
    frmGenera.visible = False

   On Error GoTo 0
   Exit Sub

cmdGenerarFactura_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGenerarFactura_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub cmdLimpiar_Click()
    txtmov(0) = ""
    txtmov(6) = ""
    cmbCC.limpiar
    cmbSubcuenta.limpiar
    txtmov(2) = ""
    txtmov(3) = ""
    txtmov(4) = ""
    chkPago.Value = Unchecked
    chkPrevista.Value = Unchecked
 '   fPago.Enabled = False
    fPrevista.Enabled = False
  '  cmbSubcuenta.desactivar
    txtmov(7) = ""
    txtmov(8) = ""
    txtmov(9) = ""
    pdf1.LoadFile vbNullString
    pdf1.visible = False
    listaFamilias.ListItems.Clear
    On Error Resume Next
    txtmov(0).SetFocus
End Sub

Private Sub cmdMarcar_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If Index = 0 Then
            lista.ListItems(i).Checked = True
        Else
            lista.ListItems(i).Checked = False
        End If
    Next
End Sub

Private Sub cmdModificarFamilia_Click()
   On Error GoTo cmdModificarFamilia_Click_Error

    If listaFamilias.ListItems.Count = 0 Then Exit Sub
    If Trim(cmbCC.getTEXTO) = "" Then
        MsgBox "Debe indicar la Familia.", vbExclamation, App.Title
        Exit Sub
    End If
    If Trim(cmbSubcuenta.getTEXTO) = "" Then
        MsgBox "Debe indicar la Subcuenta.", vbExclamation, App.Title
        Exit Sub
    End If
    If Trim(txtmov(2)) = "" Then
        MsgBox "Debe indicar la base.", vbExclamation, App.Title
        Exit Sub
    End If
    If Trim(txtmov(5)) = "" Then
        MsgBox "Debe indicar el porcentaje de IVA.", vbExclamation, App.Title
        Exit Sub
    End If
    
    txtmov(2) = Replace(txtmov(2), ".", "")
    txtmov(2) = Replace(txtmov(2), "€", "")
    txtmov(2) = Trim(txtmov(2))
    
    With listaFamilias.ListItems(listaFamilias.selectedItem.Index)
        .SubItems(COLS_F.CF_FAMILIA) = cmbCC.getTEXTO
        .SubItems(COLS_F.CF_FAMILIA_ID) = cmbCC.getPK_SALIDA
        .SubItems(COLS_F.CF_SUBCUENTA) = cmbSubcuenta.getTEXTO
        .SubItems(COLS_F.CF_SUBCUENTA_ID) = cmbSubcuenta.getPK_SALIDA
        .SubItems(COLS_F.CF_BASE) = moneda(txtmov(2))
        .SubItems(COLS_F.CF_IVA_PORCENTAJE) = txtmov(5)
        .SubItems(COLS_F.CF_IVA) = moneda((txtmov(2) * CInt(txtmov(5)) / 100))
        .SubItems(COLS_F.CF_Importe) = moneda(txtmov(2) + ((txtmov(2) * CInt(txtmov(5)) / 100)))
    End With
    calcularTotal
    
    cmbCC.limpiar
    cmbSubcuenta.limpiar
    txtmov(2) = ""
    txtmov(2).BackColor = &H80000005
    txtmov(5).BackColor = &H80000005
    
'    If listaFamilias.ListItems.Count > listaFamilias.selectedItem.Index Then
'         Set listaFamilias.selectedItem = listaFamilias.ListItems(listaFamilias.selectedItem.Index + 1)
'         listaFamilias_Click
'    Else
        cmbCC.SetFocus
'    End If
    

   On Error GoTo 0
   Exit Sub

cmdModificarFamilia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarFamilia_Click of Formulario frmProveedores_Facturas"

End Sub

Private Sub cmdRelacionar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    frmProveedores_Facturas_Rel.PK_PROVEEDOR_ID = PK
    frmProveedores_Facturas_Rel.PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
    frmProveedores_Facturas_Rel.Show 1
End Sub

Private Sub cmdRevisionCancelar_Click()
    frmRevision.visible = False
End Sub

Private Sub cmdRevisionFactura_Click()
   On Error GoTo cmdRevisionFactura_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    lblRevision = "Asiento : " & lista.ListItems(lista.selectedItem.Index) & " -> " & lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_concepto) & ", de fecha : " & lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_fecha)
    Dim oPF As New clsProveedores_Facturas
    With oPF
        .Carga lista.ListItems(lista.selectedItem.Index)
        If .getREVISADA_POR <> 0 Then
            cmbRevisada_por.MostrarElemento .getREVISADA_POR
            opSituacion(.getSITUACION).Value = True
            txtRevisionMotivo = .getMOTIVO_RECHAZO
            txtRevisionEnvio = .getF_REVISION_ENVIO
            txtRevisionFecha = .getF_REVISION
        Else
            cmbRevisada_por.limpiar
            txtRevisionMotivo = ""
            txtRevisionEnvio = ""
            txtRevisionFecha = ""
        End If
    End With
    Set oPF = Nothing
    frmRevision.visible = True

   On Error GoTo 0
   Exit Sub

cmdRevisionFactura_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRevisionFactura_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub cmdRevisionModificar_Click()
   On Error GoTo cmdRevisionModificar_Click_Error

    If cmbRevisada_por.getTEXTO = "" Then
        MsgBox "Indique el usuario revisor.", vbCritical, App.Title
        Exit Sub
    End If
    Dim oFP As New clsProveedores_Facturas
    With oFP
        .setREVISADA_POR = cmbRevisada_por.getPK_SALIDA
        If opSituacion(0).Value = True Then
            .setSITUACION = 0
        ElseIf opSituacion(1).Value = True Then
            .setSITUACION = 1
        ElseIf opSituacion(2).Value = True Then
            .setSITUACION = 2
        End If
        .setMOTIVO_RECHAZO = txtRevisionMotivo
        .modificarRevision lista.ListItems(lista.selectedItem.Index)
    End With
    Set oFP = Nothing
    cargar_lista
    frmRevision.visible = False
   On Error GoTo 0
   Exit Sub

cmdRevisionModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRevisionModificar_Click of Formulario frmProveedores_Facturas"

End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fecha_Change()
    calcularVencimiento
    fechaFactura = fecha
End Sub

Private Sub fecha_LostFocus()
    If Year(fecha) <> Year(Date) Then
        MsgBox "Atención, el año es distinto al actual.", vbInformation, App.Title
    End If
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    fecha = Date
    cargar_combos
    If PK_FACTURA_ID <> 0 Then
        Dim oPF As New clsProveedores_Facturas
        oPF.Carga PK_FACTURA_ID
        fdesde = oPF.getFECHA
    Else
        fdesde = "01/01/" & Year(Date)
    End If
    fhasta = "31/12/" & Year(Date)
    Dim op As New clsParametros
    op.Carga parametros.IVA, ""
    txtmov(5) = op.getVALOR
    txtmov(8) = moneda(0)
'M1257-I
'JGM-I
'    cmbSubcuentaPago.desactivar
    If PK <> 0 Then
        cmbProveedor.MostrarElemento PK
        cargar_proveedor
        cargar_lista
    End If
    If TOBJETO <> 0 Then
        cmbProveedor.desactivar
    End If
    ' Viene buscando una factura en concreto, la busca y se situa
    If PK_FACTURA_ID <> 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If CLng(lista.ListItems(i).Text) = PK_FACTURA_ID Then
                Set lista.selectedItem = lista.ListItems(i)
                lista.ListItems(i).EnsureVisible
                lista_Click
                Exit For
            End If
        Next
    End If
'        cargar_proveedor
'     Else
'        cargar_subcontratacion
'     End If
'JGM-F
    If USUARIO.getPER_TESORERIA_MENU = False Then
        cmdborrar.visible = False
        cmdLimpiar.visible = False
        cmdAnadir(1).visible = False
        cmdAnadir(0).visible = False
        cmdAnadirFamilia.visible = False
        cmdModificarFamilia.visible = False
        cmdEliminarFamilia.visible = False
        cmdCambiarProveedor.visible = False
        cmdAnular.visible = False
        
    End If
End Sub
Private Sub cargar_combos()
    llenar_combo cmbProveedor, New clsProveedor, 0, frmProveedores_Detalle, ""
'    cargar_combo cmbFP, New clsFP
    llenar_combo cmbCC, New clsFamilias, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbSubcuenta, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS
    oDeco.cargar_mi_combo cmbSubcuentaPago, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS
    
    oDeco.cargar_combo cmbFP, DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP
    oDeco.cargar_combo cmbVencimiento, DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS
    
    Set oDeco = Nothing
    
    llenar_combo cmbRevisada_por, New clsUsuarios, 0, frmUsuarios, ""
End Sub
Private Sub cmdAdjuntar_Click(Index As Integer)
   On Error GoTo cmdAdjuntar_Click_Error
    If lista.ListItems.Count = 0 Then Exit Sub
    If datos(4).Text = "" Then
        cmdEXplorar_Click (0)
    End If
    adjuntar lista.ListItems(lista.selectedItem.Index).Text
   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub cmdAnular_Click()
   On Error GoTo cmdAnular_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oD As New clsDocumentacion
    oD.EliminarProveedorFactura lista.ListItems(lista.selectedItem.Index).Text
    Set oD = Nothing
    MsgBox "Documento eliminado correctamente.", vbInformation, App.Title
    cargar_lista

   On Error GoTo 0
   Exit Sub

cmdAnular_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnular_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub cmdEscaner_Click(Index As Integer)
   On Error GoTo cmdEscaner_Click_Error

    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_concepto)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"

            cmdAdjuntar_Click (Index)
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmProveedores_Facturas"

End Sub

Private Sub cmdGestor_Click()
   On Error GoTo cmdGestor_Click_Error
    If lista.ListItems.Count = 0 Then Exit Sub
    documento_escaner_eliminar = False
    frmGestorDocumentos.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_concepto)
'        nombreNuevo = ""
'        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
'        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"
            cmdAdjuntar_Click (Index)
            If documento_escaner_eliminar = True Then
                On Error Resume Next
                Kill documento_escaner
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub
cmdGestor_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGestor_Click of Formulario frmProveedores_Facturas"
End Sub
Private Sub cmdBorrar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar la factura seleccionada?.", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oProveedor_Factura As New clsProveedores_Facturas
            oProveedor_Factura.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            Set oProveedor_Factura = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdMostrar_Click()
   On Error GoTo CMDMOSTRAR_Click_Error

    If lista.ListItems.Count = 0 Then
        MsgBox "Seleccione algún archivo de la lista.", vbExclamation, App.Title
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index).Checked = True Then
        Dim oD As New clsDocumentacion
        oD.CargarProveedorFacturas lista.ListItems(lista.selectedItem.Index).Text, True
        Set oD = Nothing
    End If

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrar_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub lista_Click()
   On Error GoTo lista_Click_Error

    If lista.ListItems.Count > 0 Then
        ' Cargar datos de la factura para poder modificar
        Dim oPF As New clsProveedores_Facturas
        With oPF
            If .Carga(lista.ListItems(lista.selectedItem.Index)) Then
                fecha = .getFECHA
                fechaFactura = .getFECHA_FACTURA
                txtmov(0) = .getNUMERO
                txtmov(6) = .getCONCEPTO
                cmbFP.BoundText = .getFORMAPAGO
                cmbVencimiento.BoundText = .getVENCIMIENTO_ID
'                cmbCC.MostrarElemento .getFAMILIA_ID
'                cmbSubcuenta.MostrarElemento .getSUBCUENTA
'JGM-I
                txtmov(9) = moneda(.getBI)
                txtmov(8) = moneda(.getRETENCION)
                txtmov(3) = moneda(.getIVA)
                txtmov(4) = moneda(.getTOTAL)
'JGM-F
'                txtmov(5) = .getIVA_PORCENTAJE
                If Not IsNull(.getF_VENCIMIENTO) And .getF_VENCIMIENTO <> "" Then
                    fVencimiento = .getF_VENCIMIENTO
                End If
                chkPago.Value = Unchecked
'                fPago.Enabled = False
                chkPrevista.Value = Unchecked
                fPrevista.Enabled = False
'                cmbSubcuentaPago.desactivar
                cmbSubcuentaPago.MostrarElemento 0
                If Not IsNull(.getF_PAGO) Then
                    If .getF_PAGO <> "0000-00-00" And .getF_PAGO <> "" Then
'                        fPago.Enabled = True
                        chkPago.Value = Checked
'                        cmbSubcuentaPago.activar
                        fPago = .getF_PAGO
                        cmbSubcuentaPago.MostrarElemento .getSUBCUENTA_PAGO
                    End If
                End If
                If Not IsNull(.getF_PREVISTA) Then
                    If .getF_PREVISTA <> "0000-00-00" And .getF_PREVISTA <> "" Then
                        fPrevista.Enabled = True
                        chkPrevista.Value = Checked
                        fPrevista = .getF_PREVISTA
                    End If
                End If
                txtmov(7) = .getOBSERVACIONES
                
                chkEnviada.Value = .getENVIADA
                chkIncidencia.Value = .getINCIDENCIA
                
                If .getREMESA_ID <> 0 Then
                    Dim oRemesa As New clsRemesas
                    oRemesa.Carga .getREMESA_ID
                    txtmov(10) = oRemesa.getNUMERO & "/" & oRemesa.getANNO
                Else
                    txtmov(10) = ""
                End If
                cargarListaFamilias lista.ListItems(lista.selectedItem.Index)
                ' Si contabilizado, proteger botones
                If .getF_CONTABILIZADA <> 0 Then
                    lblContabilizada.Caption = "FACTURA CONTABILIZADA : " & .getF_CONTABILIZADA
                    cmdAnadirFamilia.Enabled = False
                    cmdEliminarFamilia.Enabled = False
                    cmdModificarFamilia.Enabled = False
                    cmdborrar.Enabled = False
                Else
                    lblContabilizada.Caption = ""
                    cmdAnadirFamilia.Enabled = True
                    cmdEliminarFamilia.Enabled = True
                    cmdModificarFamilia.Enabled = True
                    cmdborrar.Enabled = True
                End If
            End If
        End With
        Set oPF = Nothing
'        If lista.ListItems(lista.selectedItem.Index).Checked = True Then
            mostrar_pdf lista.ListItems(lista.selectedItem.Index).Text
'        Else
'            pdf1.LoadFile vbNullString
'            pdf1.Visible = False
'        End If
    End If

   On Error GoTo 0
   Exit Sub

lista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmProveedores_Facturas"
End Sub
Private Sub cargarListaFamilias(ID_FACTURA As Long)
    Dim oPFF As New clsProveedores_facturas_fam
    Dim rs As ADODB.Recordset
    listaFamilias.ListItems.Clear
    Set rs = oPFF.Listado(ID_FACTURA)
    If rs.RecordCount > 0 Then
        Do
            With listaFamilias.ListItems.Add(, , rs(0)) ' ID
                If Not IsNull(rs(1)) Then
                    .SubItems(COLS_F.CF_FAMILIA) = rs(1)
                Else
                    .SubItems(COLS_F.CF_FAMILIA) = ""
                End If
                .SubItems(COLS_F.CF_FAMILIA_ID) = rs(2)
                .SubItems(COLS_F.CF_SUBCUENTA) = rs(3)
                .SubItems(COLS_F.CF_SUBCUENTA_ID) = rs(4)
                .SubItems(COLS_F.CF_BASE) = moneda(rs(5))
                .SubItems(COLS_F.CF_IVA_PORCENTAJE) = rs(6)
                .SubItems(COLS_F.CF_IVA) = moneda(rs(7))
                .SubItems(COLS_F.CF_Importe) = moneda(rs(8))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
'JGM    calcularTotal
    Set rs = Nothing
    Set oPFF = Nothing
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then Exit Sub
'    If TOBJETO = 0 Then
'        Dim oFactura As New clsProveedores_Facturas
        
'        oFactura.Carga lista.ListItems(lista.selectedItem.Index).SubItems(6)
'        TOBJETO = oFactura.getTOBJETO
'        COBJETO = oFactura.getCOBJETO
        TOBJETO = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_TOBJETO)
        COBJETO = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_cOBJETO)

  'M1274-I
        Select Case TOBJETO
        Case TOBJETO_SC_DETERMINACIONES
            frmSC_Paquete_Detalle.PK = COBJETO
            frmSC_Paquete_Detalle.Show 1
        Case TOBJETO_SC_EFICACIA
            frmSC_Paquete_Detalle_CE.PK = COBJETO
            frmSC_Paquete_Detalle_CE.Show 1
        Case TOBJETO_SC_GENERICA, TOBJETO_SC_PEACH
            frmSC_Paquete_Detalle_Generico.PK = COBJETO
            frmSC_Paquete_Detalle_Generico.Show 1
        End Select
 'M1274-F
'    End If
End Sub
Private Sub cargar_proveedor()
    Dim oProveedor As New clsProveedor
    With oProveedor
        .Carga (PK)
        cmbFP.BoundText = .getFP_ID
        cmbVencimiento.BoundText = .getVENCIMIENTO_ID
        lbltitulo(0) = "Documentos del Proveedor : " & .getNOMBRE
        Me.Caption = lbltitulo(0)
        cargar_lista
    End With
    Set oProveedor = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtmov(0)) = "" Then
        MsgBox "Debe indicar el número de factura.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtmov(6)) = "" Then
        MsgBox "Debe indicar un concepto.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(cmbFP.Text) = "" Then
        MsgBox "Debe indicar la Forma de Pago.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If listaFamilias.ListItems.Count = 0 Then
        MsgBox "Debe incluir alguna FAMILIA-SUBCUENTA", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtmov(3)) = "" Then
        MsgBox "Debe indicar el IVA.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtmov(8)) = "" Then
        MsgBox "Debe indicar la RETENCION.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtmov(4)) = "" Then
        MsgBox "Debe indicar el total.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    ' Fecha de Pago no puede ser mayor a la del día
    If chkPago.Value = Checked Then
        If fPago > Date Then
            MsgBox "La fecha de Pago no puede ser mayor que la del día.", vbExclamation, App.Title
            validar = False
            Exit Function
        End If
    End If
End Function
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub listaFamilias_Click()
'        .Add , , "Id", 1, lvwColumnLeft
'        .Add , , "Familia", 2000, lvwColumnLeft
'        .Add , , "Subcuenta", 2150, lvwColumnLeft
'        .Add , , "Base", 1000, lvwColumnRight 10
'        .Add , , "Iva %", 700, lvwColumnCenter
'        .Add , , "Iva", 1000, lvwColumnRight 11
'        .Add , , "Importe", 1000, lvwColumnRight 13
'        .Add , , "FamiliaID", 1, lvwColumnLeft
'        .Add , , "SubcuentaID", 1, lvwColumnLeft &HBCF3EF

    cmbCC.MostrarElemento listaFamilias.selectedItem.SubItems(7)
    cmbSubcuenta.MostrarElemento listaFamilias.selectedItem.SubItems(8)
    txtmov(2).Text = listaFamilias.selectedItem.SubItems(3)
    txtmov(2).BackColor = &HBCF3EF
    txtmov(5).Text = listaFamilias.selectedItem.SubItems(4)
    txtmov(5).BackColor = &HBCF3EF
    txtmov(2).SetFocus
End Sub

Private Sub PushButton1_Click()
   On Error GoTo PushButton1_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim oRF As New clsRemesas_documentos
        If oRF.Carga(CLng(lista.ListItems(lista.selectedItem.Index).Text)) Then
            Dim oRemesa As New frmRemesas_Detalle
            Dim oFRM As New frmRemesas_Detalle
            With oFRM
                .PK = oRF.getREMESA_ID
                .MODO = "C"
                .Show 1
            End With
        End If
    End If

   On Error GoTo 0
   Exit Sub

PushButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton1_Click of Formulario frmProveedores_Facturas"
End Sub

Private Sub txtmov_GotFocus(Index As Integer)
    If Index = 2 Or Index = 5 Or Index = 8 Then
        txtmov(Index).SelStart = 0
        txtmov(Index).SelLength = Len(txtmov(Index))
    End If
End Sub

'M1257-F
Private Sub txtmov_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Or Index = 5 Or Index = 8 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
        ' Si pulsa intro, emular el modificar
        If Index = 2 Then
            If KeyAscii = 13 Then
                cmdModificarFamilia_Click
                If listaFamilias.ListItems.Count > listaFamilias.selectedItem.Index Then
                     Set listaFamilias.selectedItem = listaFamilias.ListItems(listaFamilias.selectedItem.Index + 1)
                     listaFamilias_Click
                End If
            End If
        End If
    End If
End Sub

Private Sub txtmov_LostFocus(Index As Integer)
   On Error GoTo txtmov_LostFocus_Error

'JGM    calcularTotal
'    If (Index = 2 Or Index = 5 Or Index = 8) And (txtmov(2) <> "" And txtmov(5) <> "" And txtmov(8) <> "") Then
'        txtmov(2) = Format(txtmov(2), "currency")
'        txtmov(8) = Format(txtmov(8), "currency")
'
'        txtmov(3) = Format((txtmov(2) * CInt(txtmov(5)) / 100), "currency")
'
'        txtmov(4) = Format(CCur(txtmov(2)) + CCur(txtmov(3)) - CCur(txtmov(8)), "currency")
'    End If

    If Index = 8 And (txtmov(9) <> "" And txtmov(3) <> "") Then
        txtmov(8) = Format(txtmov(8), "currency")
        txtmov(4) = Format(CCur(txtmov(9)) + CCur(txtmov(3)) - CCur(txtmov(8)), "currency")
    End If

   On Error GoTo 0
   Exit Sub

txtmov_LostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtmov_LostFocus of Formulario frmProveedores_Facturas"
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPF As New clsProveedores_Facturas
    Set rs = oPF.Listado(PK, TOBJETO, COBJETO, fdesde, fhasta, chkPendientesPago.Value, chkVencidas.Value, chkPagoPrevisto.Value, chkIncidencias.Value, chkNoEnviadas.Value)
    Dim BASE As Currency
    Dim IVA As Currency
    Dim retencion As Currency
    Dim total As Currency
    BASE = 0
    IVA = 0
    retencion = 0
    total = 0
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000000")) ' ID
            .SubItems(COLS.C_fecha) = Format(rs(1), "dd/mm/yyyy")  ' Fecha
            If Not IsNull(rs(2)) Then
                .SubItems(COLS.C_concepto) = rs(2)  ' Concepto
            End If
            If Not IsNull(rs(3)) Then
                .SubItems(COLS.C_NUMERO) = rs(3)  ' Numero
            End If
            If Not IsNull(rs(4)) Then
                .SubItems(COLS.C_FAMILIA) = rs(4)  ' Familia
            End If
            If Not IsNull(rs(5)) Then
                .SubItems(COLS.C_SUBCUENTA) = rs(5)  ' Subcuenta
            End If
            .SubItems(COLS.C_BASE) = Format(rs(6), "currency")  ' BI
            .SubItems(COLS.C_IVA_PORCENTAJE) = rs(7)  ' IVA PORCENTAJE
            .SubItems(COLS.C_IVA) = Format(rs(8), "currency")  ' IVA
            .SubItems(COLS.C_total) = Format(rs(9), "currency")  ' TOTAL
            BASE = BASE + rs(6)
            IVA = IVA + rs(8)
            retencion = retencion + rs(16)
            total = total + rs(9)
            If Not IsNull(rs(10)) Then
                .SubItems(COLS.C_FP) = rs(10)  ' FP
            End If
            If Not IsNull(rs(11)) Then
                .SubItems(COLS.C_vencimiento) = rs(11)  ' F.Vencimiento
            End If
            If Not IsNull(rs(12)) Then
                .SubItems(COLS.C_FP_ID) = rs(12)  ' F.Pago
            End If
            If Not IsNull(rs(13)) Then
                .SubItems(COLS.C_TOBJETO) = rs(13)  ' Tobjeto
            End If
            If Not IsNull(rs(14)) Then
                .SubItems(COLS.C_cOBJETO) = rs(14)  ' Cobjeto
            End If
            If Not IsNull(rs(15)) Then
                .Checked = True
            Else
                .Checked = False
            End If
            If Not IsNull(rs(16)) Then
                .SubItems(COLS.C_RETENCION) = Format(rs(16), "currency") ' RETENCION
            End If
            If IsNull(rs(12)) Or rs(12) = "0000-00-00" Then
                .SubItems(COLS.C_PAGADA) = ""
            Else
                .SubItems(COLS.C_PAGADA) = "X"
            End If
            If rs(18) = 0 Then
                .SubItems(COLS.C_ENVIADA) = ""
            Else
                .SubItems(COLS.C_ENVIADA) = "X"
            End If
            ' REVISION 21: REVISADA_POR, 22: SITUACION
            If rs(19) = 0 Then ' Si no hay revisor, no bola
                .ListSubItems.Add , , "", vbNothing
            Else
                If rs(20) = 0 Then 'ENVIADA
                    .ListSubItems.Add , , "", 2
                ElseIf rs(20) = 1 Then 'aprobada
                    .ListSubItems.Add , , "", 1
                ElseIf rs(20) = 2 Then 'Rechazada
                    .ListSubItems.Add , , "", 3
                End If
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
'        lista_Click
    End If
    lblBase = Format(BASE, "currency")
    lblIVA = Format(IVA, "currency")
    lblRetencion = Format(retencion, "currency")
    lbltotal = Format(total, "currency")
End Sub
Private Sub mostrar_pdf(ID As Long)
    Dim oD As New clsDocumentacion
    Dim destino As String
    destino = oD.CargarProveedorFacturas(ID, False)
    If Dir(destino) <> "" Then
        pdf1.visible = True
        pdf1.LoadFile destino
        pdf1.setShowToolbar False
    Else
        pdf1.visible = False
        pdf1.LoadFile vbNullString
    End If
End Sub
Private Sub calcularVencimiento()
    fVencimiento = fecha
    If cmbVencimiento.Text <> "" Then
        Dim oDeco As New clsDecodificadora
        oDeco.Carga_valor DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS, cmbVencimiento.BoundText
        If IsNumeric(oDeco.getPARAMETROS) Then
            fVencimiento.Value = fecha + CInt(oDeco.getPARAMETROS)
        End If
    End If
    ' Calcular fecha prevista de pago, días cercano a 10 o 25
    Dim dia As Integer
    Dim MES As Integer
    Dim ANNO As Integer
    dia = Day(fVencimiento)
    MES = Month(fVencimiento)
    ANNO = Year(fVencimiento)
    If dia <= 10 Then
        fPrevista.Value = "10/" & MES & "/" & ANNO
    ElseIf dia > 10 And dia <= 25 Then
        fPrevista.Value = "25/" & MES & "/" & ANNO
    Else
        If MES < 12 Then
            fPrevista.Value = "10/" & MES + 1 & "/" & ANNO
        Else
            fPrevista.Value = "10/01/" & ANNO + 1
        End If
    End If
'    If cmbFP.BoundText <> "" Then
'        Dim oFP As New clsFP
'        oFP.CARGAR cmbFP.BoundText
'        If oFP.getDIAS <> 0 Then
'            fVencimiento.value = fecha + oFP.getDIAS
'        End If
'    End If
End Sub
Private Sub adjuntar(ID As Long)
    If datos(4).Text <> "" Then
        Dim oD As New clsDocumentacion
        oD.SubirProveedorFacturas ID, datos(4), datos(0)
        Set oD = Nothing
        cargar_lista
        datos(0) = ""
        datos(4) = ""
        MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub calcularTotal()
    On Error Resume Next
    Dim i As Integer
    Dim BASE As Currency
    Dim IVA As Currency
    Dim total As Currency
    For i = 1 To listaFamilias.ListItems.Count
        BASE = BASE + listaFamilias.ListItems(i).SubItems(COLS_F.CF_BASE)
        IVA = IVA + listaFamilias.ListItems(i).SubItems(COLS_F.CF_IVA)
    Next
    Dim retencion As Currency
    If IsNumeric(txtmov(8)) Then
        retencion = txtmov(8)
    Else
        retencion = 0
    End If
    total = moneda(CCur(BASE) + CCur(IVA) - CCur(retencion))
    txtmov(9) = moneda(CStr(BASE))
    txtmov(3) = moneda(CStr(IVA))
    txtmov(4) = moneda(CStr(total))
    
End Sub
