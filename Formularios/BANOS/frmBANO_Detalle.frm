VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmBANO_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Baños"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBANO_Detalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   15780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Periodicidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   45
      TabIndex        =   71
      Top             =   7020
      Width           =   10005
      Begin MSComctlLib.ListView listaFrecuencias 
         Height          =   1605
         Left            =   90
         TabIndex        =   72
         Top             =   270
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2831
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
      Begin pryCombo.miCombo cmbperiodicidad 
         Height          =   375
         Left            =   1125
         TabIndex        =   73
         Top             =   1935
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTarifaF 
         Height          =   375
         Left            =   1125
         TabIndex        =   75
         Top             =   2340
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   661
      End
      Begin XtremeSuiteControls.PushButton cmdPeriodicidad 
         Height          =   525
         Index           =   0
         Left            =   8595
         TabIndex        =   77
         Top             =   270
         Width           =   1320
         _Version        =   851970
         _ExtentX        =   2328
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmBANO_Detalle.frx":000C
      End
      Begin XtremeSuiteControls.PushButton cmdPeriodicidad 
         Height          =   525
         Index           =   1
         Left            =   8595
         TabIndex        =   78
         Top             =   810
         Width           =   1320
         _Version        =   851970
         _ExtentX        =   2328
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmBANO_Detalle.frx":686E
      End
      Begin XtremeSuiteControls.PushButton cmdPeriodicidad 
         Height          =   525
         Index           =   2
         Left            =   8595
         TabIndex        =   79
         Top             =   1350
         Width           =   1320
         _Version        =   851970
         _ExtentX        =   2328
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmBANO_Detalle.frx":D0D0
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   76
         Top             =   2385
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   74
         Top             =   1965
         Width           =   870
      End
   End
   Begin VB.Frame frmAIM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Datos ADS (Sólo clientes Airbus)"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   10080
      TabIndex        =   53
      Top             =   2070
      Width           =   5640
      Begin pryCombo.miCombo cmbPrograma 
         Height          =   330
         Left            =   900
         TabIndex        =   54
         Top             =   765
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbEnsayo 
         Height          =   330
         Left            =   900
         TabIndex        =   55
         Top             =   360
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSection 
         Height          =   330
         Left            =   900
         TabIndex        =   56
         Top             =   1170
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFluid 
         Height          =   330
         Left            =   900
         TabIndex        =   57
         Top             =   1575
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFacility 
         Height          =   330
         Left            =   900
         TabIndex        =   58
         Top             =   1980
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facility"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   63
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fluid"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   62
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Section"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   61
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   60
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Programa"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   135
         TabIndex        =   59
         Top             =   810
         Width           =   870
      End
   End
   Begin VB.CheckBox chkAnulado 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   12645
      TabIndex        =   52
      Top             =   180
      Width           =   1860
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones para la recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   45
      TabIndex        =   49
      Top             =   5040
      Width           =   10005
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   6
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   270
         Width           =   9840
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Airbus (Aplicación web MTQM)"
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
      Left            =   10080
      TabIndex        =   45
      Top             =   615
      Width           =   5640
      Begin pryCombo.miCombo cmbAirbusArea 
         Height          =   375
         Left            =   855
         TabIndex        =   20
         Top             =   600
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbCentro 
         Height          =   375
         Left            =   855
         TabIndex        =   19
         Top             =   225
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbAirbusLinea 
         Height          =   375
         Left            =   855
         TabIndex        =   21
         Top             =   990
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Línea P."
         Height          =   195
         Index           =   19
         Left            =   90
         TabIndex        =   48
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   47
         Top             =   225
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Area"
         Height          =   195
         Index           =   20
         Left            =   90
         TabIndex        =   46
         Top             =   630
         Width           =   330
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8910
      Width           =   1365
   End
   Begin VB.CheckBox chkCE 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Control de Eficacia"
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
      Left            =   180
      TabIndex        =   12
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Frame frmCE 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   870
      Left            =   45
      TabIndex        =   43
      Top             =   6120
      Width           =   10005
      Begin VB.CommandButton cmdficha 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ficha"
         Height          =   555
         Left            =   8730
         Picture         =   "frmBANO_Detalle.frx":13932
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1185
      End
      Begin pryCombo.miCombo cmbFicha 
         Height          =   375
         Left            =   1125
         TabIndex        =   13
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ficha CE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   135
         TabIndex        =   44
         Top             =   405
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Facturación Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   10080
      TabIndex        =   41
      Top             =   4545
      Width           =   5640
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   5
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2025
         Width           =   1200
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   4
         Left            =   765
         TabIndex        =   65
         Top             =   2025
         Width           =   1155
      End
      Begin VB.CheckBox chkFD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura por Determinaciones"
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
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   2940
      End
      Begin VB.CheckBox chkrevisarfactura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisar factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3420
         TabIndex        =   16
         Top             =   270
         Width           =   1770
      End
      Begin VB.TextBox txttarifa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3690
         TabIndex        =   18
         Top             =   990
         Width           =   1590
      End
      Begin MSComctlLib.ListView tarifas 
         Height          =   750
         Left            =   135
         TabIndex        =   17
         Top             =   585
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   1323
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
      Begin pryCombo.miCombo cmbtarifa 
         Height          =   375
         Left            =   135
         TabIndex        =   69
         Top             =   1620
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod. Tarifa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   135
         TabIndex        =   70
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   2385
         TabIndex        =   68
         Top             =   2070
         Width           =   1530
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   66
         Top             =   2070
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio Tarifa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   3690
         TabIndex        =   42
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8910
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14535
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8910
      Width           =   1050
   End
   Begin VB.CommandButton cmdDeterminaciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Determinaciones"
      Height          =   870
      Left            =   11580
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7920
      Width           =   1950
   End
   Begin VB.CommandButton cmdDatosEspecificos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Específicos"
      Height          =   870
      Left            =   13590
      Picture         =   "frmBANO_Detalle.frx":13BA3
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7920
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   4365
      Left            =   45
      TabIndex        =   27
      Top             =   615
      Width           =   10005
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1740
         TabIndex        =   8
         Top             =   3150
         Width           =   7800
      End
      Begin pryCombo.miCombo cmbenvases 
         Height          =   375
         Left            =   1740
         TabIndex        =   7
         Top             =   2790
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbprocedencia 
         Height          =   375
         Left            =   1740
         TabIndex        =   9
         Top             =   3555
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbsolucion 
         Height          =   375
         Left            =   1740
         TabIndex        =   6
         Top             =   2430
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbProcesos 
         Height          =   375
         Left            =   1740
         TabIndex        =   5
         Top             =   2070
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbInstalacion 
         Height          =   375
         Left            =   1740
         TabIndex        =   4
         Top             =   1710
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbLineas 
         Height          =   375
         Left            =   1740
         TabIndex        =   3
         Top             =   1350
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTipos 
         Height          =   375
         Left            =   1740
         TabIndex        =   2
         Top             =   990
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   8055
         TabIndex        =   11
         Top             =   3960
         Width           =   1515
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   4110
         TabIndex        =   10
         Top             =   3945
         Width           =   2235
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1740
         MaxLength       =   100
         TabIndex        =   0
         Top             =   210
         Width           =   7815
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   375
         Left            =   1740
         TabIndex        =   1
         Top             =   630
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   6795
         TabIndex        =   64
         Top             =   4005
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Instalación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   90
         TabIndex        =   51
         Top             =   1710
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conservación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   35
         Top             =   3210
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   90
         TabIndex        =   38
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   6795
         TabIndex        =   37
         Top             =   3975
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo entre toma de muestra y fin análisis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   90
         TabIndex        =   36
         Top             =   3975
         Width           =   3870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   90
         TabIndex        =   34
         Top             =   2865
         Width           =   690
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solución"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   33
         Top             =   2445
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso Base"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   32
         Top             =   2100
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   31
         Top             =   615
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   30
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Muestra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   29
         Top             =   1005
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de las aguas y baños pertenecientes a los tipos de muestra especiales"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   40
      Top             =   330
      Width           =   5505
   End
   Begin VB.Image imagen 
      Height          =   720
      Left            =   14850
      Picture         =   "frmBANO_Detalle.frx":1446D
      Top             =   -75
      Width           =   720
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle Aguas y Baños"
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
      TabIndex        =   39
      Top             =   30
      Width           =   2385
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   16065
   End
End
Attribute VB_Name = "frmBANO_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private tarifa_modificada As Boolean

Private Enum COLS
    C_ORDEN = 0
    C_PERIODICIDAD = 1
    C_CODIGO = 2
    C_TIPO_FRECUENCIA_ID = 3
    C_CODIGO_ID = 4
End Enum
Private Sub cargar_aim()
    On Error GoTo cargar_aim_Error

    frmAIM.Enabled = False
    Frame3.Enabled = False
    cmbEnsayo.limpiar
    cmbPrograma.limpiar
    cmbSection.limpiar
    cmbFluid.limpiar
    cmbFacility.limpiar
    
    If cmbclientes.getTEXTO <> "" Then
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbclientes.getPK_SALIDA
        Dim ID_PLANTA As String
        ID_PLANTA = CStr(oCliente.getPLANT_ID)
        If oCliente.getAIRBUS = 1 Then
            If ID_PLANTA = "0" Then
                MsgBox "El cliente ADS no tiene informada la planta. Es necesario informarla en la ficha de cliente.", vbCritical, App.Title
                Exit Sub
            Else
                frmAIM.Enabled = True
                Frame3.Enabled = True
                Dim oDeco As New clsDecodificadora
                oDeco.cargar_mi_combo_parametro cmbEnsayo, DECODIFICADORA.AIRBUS_TIPOS_ENSAYOS, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbPrograma, DECODIFICADORA.AIRBUS_PROGRAMAS, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbSection, DECODIFICADORA.AIRBUS_SECTION, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbFluid, DECODIFICADORA.AIRBUS_FLUID, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbFacility, DECODIFICADORA.AIRBUS_FACILITY, ID_PLANTA
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cargar_aim_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_aim of Formulario frmSE_Recepcion"
End Sub


Private Sub cmbClientes_change()
   On Error GoTo cmbClientes_change_Error

    cargar_tarifas
    cargar_aim

   On Error GoTo 0
   Exit Sub

cmbClientes_change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbClientes_change of Formulario frmBANO_Detalle"
    
End Sub

Private Sub cmbFicha_change()
    Dim oBANO As New clsBanos
    If cmbFicha.getTEXTO = "" Then
        oBANO.Modificar_Ficha CInt(PK), 0
    Else
        oBANO.Modificar_Ficha CInt(PK), CLng(cmbFicha.getPK_SALIDA)
    End If
End Sub

Private Sub cmbLineas_change()
    On Error Resume Next
    If cmbInstalacion.getTEXTO = "" Then
        cmbInstalacion.MostrarElemento cmbLineas.getPK_SALIDA
    End If
End Sub

Private Sub cmdFicha_Click()
    frmCE_Ficha_Bano.PK = PK
    frmCE_Ficha_Bano.Show 1
End Sub

Private Sub chkCE_Click()
    If chkCE.Value = Checked Then
        frmCE.Enabled = True
    Else
        frmCE.Enabled = False
        cmbFicha.limpiar
    End If
End Sub


Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_BANO
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Baño " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub cmdPeriodicidad_Click(Index As Integer)
    Select Case Index
        Case 0, 1 ' Añadir, Modificar
            If cmbperiodicidad.getTEXTO = "" Then
                MsgBox "Debe indicar la Periodicidad.", vbExclamation, App.Title
            Else
                ' Comprobar duplicado
                Dim fila As Integer
                If listaFrecuencias.ListItems.Count > 0 Then
                    For fila = 1 To listaFrecuencias.ListItems.Count
                        If CLng(listaFrecuencias.ListItems(fila).SubItems(COLS.C_TIPO_FRECUENCIA_ID)) = cmbperiodicidad.getPK_SALIDA Then
                            MsgBox "La periodicidad ya existe.", vbExclamation, App.Title
                            Exit Sub
                        End If
                    Next
                End If
                If Index = 0 Then
                    listaFrecuencias.ListItems.Add , , ""
                    fila = listaFrecuencias.ListItems.Count
                Else
                    fila = listaFrecuencias.selectedItem.Index
                End If
                With listaFrecuencias.ListItems(fila)
                    .SubItems(COLS.C_PERIODICIDAD) = cmbperiodicidad.getTEXTO
                    .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = cmbperiodicidad.getPK_SALIDA
                    If cmbTarifaF.getTEXTO = "" Then
                        .SubItems(COLS.C_CODIGO) = ""
                        .SubItems(COLS.C_CODIGO_ID) = "0"
                    Else
                        .SubItems(COLS.C_CODIGO) = cmbTarifaF.getTEXTO
                        .SubItems(COLS.C_CODIGO_ID) = cmbTarifaF.getPK_SALIDA
                    End If
                End With
            End If
            cmbperiodicidad.limpiar
            cmbTarifaF.limpiar
        Case 2 ' Eliminar
            If listaFrecuencias.ListItems.Count = 0 Then Exit Sub
            listaFrecuencias.ListItems.Remove listaFrecuencias.selectedItem.Index
    End Select
End Sub

Private Sub listaFrecuencias_Click()
    If listaFrecuencias.ListItems.Count = 0 Then Exit Sub
    cmbperiodicidad.MostrarElemento listaFrecuencias.ListItems(listaFrecuencias.selectedItem.Index).SubItems(COLS.C_TIPO_FRECUENCIA_ID)
    If listaFrecuencias.ListItems(listaFrecuencias.selectedItem.Index).SubItems(COLS.C_CODIGO_ID) <> "" Then
        cmbTarifaF.MostrarElemento listaFrecuencias.ListItems(listaFrecuencias.selectedItem.Index).SubItems(COLS.C_CODIGO_ID)
    End If
End Sub

Private Sub listaFrecuencias_DblClick()
    If listaFrecuencias.ListItems.Count = 0 Then Exit Sub
    If listaFrecuencias.ListItems(listaFrecuencias.selectedItem.Index).SubItems(COLS.C_CODIGO_ID) <> "" Then
        frmTarifas_Codigos_Precios.PK = listaFrecuencias.ListItems(listaFrecuencias.selectedItem.Index).SubItems(COLS.C_CODIGO_ID)
        frmTarifas_Codigos_Precios.Show 1
    End If
End Sub
Private Sub tarifas_Click()
    If tarifas.ListItems.Count > 0 Then
         txttarifa = Trim(tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2))
         txttarifa.SetFocus
    End If
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub txttarifa_GotFocus()
    txttarifa.SelStart = 0
    txttarifa.SelLength = Len(txttarifa.Text)
End Sub

Private Sub txttarifa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
       KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        anadir_precio
'        KeyAscii = 0
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDatosEspecificos_Click()
    With frmTDE_Analisis
        .PK_ANALISIS = 0
        .PK_BANO = PK
        .Show 1
    End With
End Sub

Private Sub cmdDeterminaciones_Click()
    With frmDeterminaciones_analisis
        .PK_ANALISIS = 0
        .PK_BANO = PK
        .Show 1
    End With
End Sub

Private Sub cmdok_Click()
    If validar = True Then
        On Error GoTo fallo
        Dim BANO As Long
        Dim oBANO As New clsBanos
        Dim i As Integer
        Dim ORDEN As Integer
        With oBANO
            .setNOMBRE = txtDatos(0)
            .setCLIENTE_ID = cmbclientes.getPK_SALIDA
            .setTIPO_MUESTRA_ID = cmbTipos.getPK_SALIDA
            .setID_LINEA = cmbLineas.getPK_SALIDA
            .setINSTALACION_ID = cmbInstalacion.getPK_SALIDA
            .setID_PROCESO_BASE = cmbProcesos.getPK_SALIDA
            .setID_SOLUCION = cmbsolucion.getPK_SALIDA
            ' Temporalmente, la primera periodicidad
            .setTIPO_FRECUENCIA_ID = listaFrecuencias.ListItems(1).SubItems(COLS.C_TIPO_FRECUENCIA_ID)
            .setTIPO_FRECUENCIA_ID = cmbperiodicidad.getPK_SALIDA
            .setFORMATO_ID = cmbenvases.getPK_SALIDA
            .setCONSERVACION = txtDatos(1)
            .setTOMA_FIN = txtDatos(2)
            .setVOLUMEN = txtDatos(3)
            If cmbprocedencia.getTEXTO = "" Then
                .setSOLUCION_PROCEDENCIA_ID = 0
            Else
                .setSOLUCION_PROCEDENCIA_ID = cmbprocedencia.getPK_SALIDA
            End If
            If cmbtarifa.getTEXTO = "" Then
                .setTARIFA_CODIGO_ID = 0
            Else
                .setTARIFA_CODIGO_ID = cmbtarifa.getPK_SALIDA
            End If
            If txtDatos(4) <> "" Then
                .setPRECIO = moneda_bd(txtDatos(4))
            Else
                .setPRECIO = moneda_bd("0")
            End If
            .setREVISAR_FACTURA = chkrevisarfactura.Value
            .setFACTURA_DETERMINACIONES = chkFD.Value
            If chkCE.Value = Checked Then
                .setFICHA_ID = cmbFicha.getPK_SALIDA
            Else
                .setFICHA_ID = 0
            End If
            .setCENTRO_ID = cmbCentro.getPK_SALIDA
            .setAIRBUS_AREA_ID = cmbAirbusArea.getPK_SALIDA
            .setAIRBUS_LINEA_ID = cmbAirbusLinea.getPK_SALIDA
            .setOBSERVACIONES = txtDatos(6)
            
            If chkAnulado.Value = Checked Then
                .setANULADO = 1
            Else
                .setANULADO = 0
            End If
        End With
        Dim ohc As New clsHistorial_cambios
        If PK <> 0 Then
            If MsgBox("Va a modificar el baño. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                frmMotivo.Caption = "Indique detalladamente el motivo de modificación del baño."
                frmMotivo.Show 1
                If Trim(MOTIVO) = "" Then
                    MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                    Exit Sub
                End If
                BANO = oBANO.Modificar(PK)
                If BANO <> 0 Then
                    With ohc
                        .setTIPO = HC_TIPOS.HC_BANO
                        .setIDENTIFICADOR = PK
                        .setIDENTIFICADOR_TEXTO = txtDatos(0)
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setMOTIVO = Trim(MOTIVO)
                        .Insertar
                    End With
                End If
            End If
        Else
            If MsgBox("Va a introducir un nuevo baño. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                BANO = oBANO.Insertar
                If BANO <> 0 Then
                    With ohc
                        .setTIPO = HC_TIPOS.HC_BANO
                        .setIDENTIFICADOR = BANO
                        .setIDENTIFICADOR_TEXTO = txtDatos(0)
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setMOTIVO = HC_CREACION
                        .Insertar
                    End With
                End If
            End If
        End If
        ' Tarifas
        Set ohc = Nothing
        Me.MousePointer = 11
      ' Enviar correo si se modifica la tarifa
'      If tarifa_modificada = True Then
'            Dim oParametro As New clsParametros
'            oParametro.Carga PARAM_USUARIO_VIGILADO, ""
'            If USUARIO.getID_EMPLEADO = oParametro.getVALOR Then
'                Dim asunto As String
'                Dim DETALLE As String
'                asunto = "El usuario " & USUARIO.getUSUARIO & " ha modificado la tarifa de un BAÑO."
'
'                DETALLE = "" & vbNewLine
'                DETALLE = DETALLE & " Fecha : " & Format(Date, "dd-mm-yyyy") & vbNewLine
'                DETALLE = DETALLE & " Hora  : " & Time & vbNewLine & vbNewLine
'                DETALLE = DETALLE & " BAÑO : " & txtDatos(0) & vbNewLine & vbNewLine
'                DETALLE = DETALLE & " Cliente : " & cmbClientes.getTEXTO & vbNewLine & vbNewLine
'
'                DETALLE = DETALLE & " Cambios en la tarifa " & vbNewLine
'                DETALLE = DETALLE & " -------------------- " & vbNewLine
'
'                Dim CO As String
'                Dim rs2 As ADODB.RecordSet
'                CO = "SELECT A.ID_TARIFA, A.NOMBRE, B.PRECIO " & _
'                     "  FROM TARIFAS A LEFT JOIN TARIFAS_PRECIOS B ON A.ID_TARIFA = B.TARIFA_ID  AND B.BANO_ID = " & BANO & _
'                     " where A.EN_VIGOR = 1 "
'                Set rs2 = datos_bd(CO)
'                Dim PRECIO As String
'                Dim precio_ant As String
'                If rs2.RecordCount > 0 Then
'                    Do
'                            For i = 1 To tarifas.ListItems.Count
'                              If tarifas.ListItems(i).Text = rs2(0) Then
'                                If IsNull(rs2(2)) Then
'                                    precio_ant = moneda("0")
'                                Else
'                                    precio_ant = moneda(rs2(2))
'                                End If
'                                If Trim(tarifas.ListItems(i).SubItems(2)) = "" Then
'                                    PRECIO = moneda("0")
'                                Else
'                                    PRECIO = moneda(tarifas.ListItems(i).SubItems(2))
'                                End If
'                                If PRECIO <> precio_ant Then
'                                    DETALLE = DETALLE & tarifas.ListItems(i).SubItems(1) & " : " & precio_ant & " -> " & PRECIO & vbNewLine
'                                End If
'                              End If
'                            Next
'                        rs2.MoveNext
'                    Loop Until rs2.EOF
'                End If
'
'                oParametro.Carga PARAM_USUARIO_VIGILADO_CORREO, ""
'                ret = Enviar_Mail_CDO(oParametro.getVALOR, asunto, DETALLE, vbNullString)
'            End If
'      End If
        ' Periodicidades
        guardar_frecuencias BANO
        
        If USUARIO.getPER_FACTURACION = True Then
            Dim oTP As New clsTarifas_precios
            If PK <> 0 Then
              oTP.Eliminar_por_bano (PK)
            End If
            If tarifas.ListItems.Count > 0 Then
              For i = 1 To tarifas.ListItems.Count
                If Trim(tarifas.ListItems(i).SubItems(2)) <> "" Then
                    With oTP
                        .setBANO_ID = BANO
                        .setTARIFA_ID = tarifas.ListItems(i).Text
                        .setPRECIO = moneda_bd(tarifas.ListItems(i).SubItems(2))
                        .Insertar
                    End With
                End If
              Next
            End If
        End If
        ' Almacenar datos AIM
        Dim oAO As New clsAirbus_objetos
        With oAO
            .setTOBJETO = TOBJETO.TOBJETO_BANO
            .setCOBJETO = BANO
            .setENSAYO_ID = IIf(cmbEnsayo.getTEXTO = "", 0, cmbEnsayo.getPK_SALIDA)
            .setPROGRAMA_ID = IIf(cmbPrograma.getTEXTO = "", 0, cmbPrograma.getPK_SALIDA)
            .setSECTION_ID = IIf(cmbSection.getTEXTO = "", 0, cmbSection.getPK_SALIDA)
            .setFLUID_ID = IIf(cmbFluid.getTEXTO = "", 0, cmbFluid.getPK_SALIDA)
            .setFACILITY_ID = IIf(cmbFacility.getTEXTO = "", 0, cmbFacility.getPK_SALIDA)
            .Insertar
        End With
        
        Me.MousePointer = 0
        If PK = 0 Then
            MsgBox "El baño se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
            PK = BANO
            cargar_bano
        Else
            MsgBox "El baño se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
        End If
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al insertar el baño. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub guardar_frecuencias(BANO_ID As Long)
    Dim i As Integer
    Dim oBF As New clsBanos_frecuencias
   On Error GoTo guardarPeriodicidades_Error
    oBF.Eliminar BANO_ID
    For i = 1 To listaFrecuencias.ListItems.Count
        With oBF
            .setBANO_ID = BANO_ID
            .setTIPO_FRECUENCIA_ID = listaFrecuencias.ListItems(i).SubItems(COLS.C_TIPO_FRECUENCIA_ID)
            .setCODIGO_ID = listaFrecuencias.ListItems(i).SubItems(COLS.C_CODIGO_ID)
            .setORDEN = i
            .Insertar
        End With
    Next

   On Error GoTo 0
   Exit Sub

guardarPeriodicidades_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure guardarPeriodicidades of Formulario frmBANO_Detalle"
    
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    cabecera
    perfil
    If PK <> 0 Then
        cargar_bano
    Else
        lbltitulo.Caption = "Alta de Aguas y Baños"
        cmdDeterminaciones.Enabled = False
        cmdDatosEspecificos.Enabled = False
    End If
    tarifa_modificada = False
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 4 Then
        txtDatos(Index) = moneda(txtDatos(Index))
    End If
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al baño", vbInformation, App.Title
        validar = False
        Exit Function
    End If
'    If cmbClientes.BoundText = "" Then
    If cmbclientes.getTEXTO = "" Then
        MsgBox "Debe seleccionar un cliente.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbTipos.getTEXTO = "" Then
        MsgBox "Debe seleccionar un Tipo de Muestra.", vbInformation, App.Title
        validar = False
        cmbTipos.SetFocus
        Exit Function
    End If
    If cmbLineas.getTEXTO = "" Then
        MsgBox "Debe seleccionar una línea.", vbInformation, App.Title
        validar = False
        cmbLineas.SetFocus
        Exit Function
    End If
    If cmbInstalacion.getTEXTO = "" Then
        MsgBox "Debe seleccionar una Instalación.", vbInformation, App.Title
        validar = False
        cmbInstalacion.SetFocus
        Exit Function
    End If
    If cmbProcesos.getTEXTO = "" Then
        MsgBox "Debe seleccionar un Proceso Base.", vbInformation, App.Title
        validar = False
        cmbProcesos.SetFocus
        Exit Function
    End If
    If cmbsolucion.getTEXTO = "" Then
        MsgBox "Debe seleccionar una Solucion.", vbInformation, App.Title
        validar = False
        cmbsolucion.SetFocus
        Exit Function
    End If
    If listaFrecuencias.ListItems.Count = 0 Then
        MsgBox "Debe tener al menos una periodicidad.", vbInformation, App.Title
        validar = False
        cmbperiodicidad.SetFocus
        Exit Function
    End If
    If cmbenvases.getTEXTO = "" Then
        MsgBox "Debe seleccionar un envase.", vbInformation, App.Title
        validar = False
        cmbenvases.SetFocus
        Exit Function
    End If
    If chkCE.Value = Checked Then
        If cmbFicha.getTEXTO = "" Then
            MsgBox "Indique la ficha del CE.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
    End If
End Function
Public Sub cargar_bano()
   On Error GoTo cargar_bano_Error

    lbltitulo.Caption = "Modificación de Aguas y Baños"
    cmdDeterminaciones.Enabled = True
    cmdDatosEspecificos.Enabled = True
    Dim BANO As New clsBanos
    With BANO
        If .cargar_bano(PK) = True Then
            txtDatos(0) = .getNOMBRE
            txtDatos(1) = .getCONSERVACION
            txtDatos(2) = .getTOMA_FIN
            txtDatos(3) = .getVOLUMEN
            txtDatos(4) = moneda(.getPRECIO)
            txtDatos(6) = .getOBSERVACIONES
            ' Precio por determinaciones
            Dim oDA As New clsDeterminaciones_analisis
            txtDatos(5) = moneda(oDA.precio_bano(PK, 0))
            cmbclientes.MostrarElemento .getCLIENTE_ID
            cmbTipos.MostrarElemento .getTIPO_MUESTRA_ID
            cmbInstalacion.MostrarElemento .getINSTALACION_ID
            cmbLineas.MostrarElemento .getID_LINEA
            cmbProcesos.MostrarElemento .getID_PROCESO_BASE
            cmbsolucion.MostrarElemento .getID_SOLUCION
            cmbprocedencia.MostrarElemento .getSOLUCION_PROCEDENCIA_ID
'            cmbperiodicidad.MostrarElemento .getTIPO_FRECUENCIA_ID
            cmbenvases.MostrarElemento .getFORMATO_ID
            cmbtarifa.MostrarElemento .getTARIFA_CODIGO_ID
            chkrevisarfactura.Value = .getREVISAR_FACTURA
            chkFD.Value = .getFACTURA_DETERMINACIONES
            If .getFICHA_ID = 0 Then
                chkCE.Value = Unchecked
            Else
                chkCE.Value = Checked
                cmbFicha.MostrarElemento .getFICHA_ID
            End If
            If .getANULADO <> 0 Then
                chkAnulado.Value = Checked
            Else
                chkAnulado.Value = Unchecked
            End If
            cmbCentro.MostrarElemento .getCENTRO_ID
            cmbAirbusArea.MostrarElemento .getAIRBUS_AREA_ID
            cmbAirbusLinea.MostrarElemento .getAIRBUS_LINEA_ID
            ' Frecuencias (Periodicidad)
            cargar_frecuencias PK
            ' Multitarifa
            Dim oMT As New clsTarifas_precios
            Set rs = oMT.Listado_por_bano_Cliente(PK, .getCLIENTE_ID)
            If rs.RecordCount <> 0 Then
                Dim i As Integer
                Do
                        For i = 1 To tarifas.ListItems.Count
                            If CInt(tarifas.ListItems(i).Text) = CInt(rs("TARIFA_ID")) Then
                                tarifas.ListItems(i).SubItems(2) = moneda(CStr(rs("PRECIO")))
                            End If
                        Next
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' AIM
            Dim oAO As New clsAirbus_objetos
            With oAO
                If .Carga(TOBJETO.TOBJETO_BANO, PK) = True Then
                    If .getENSAYO_ID <> 0 Then
                        cmbEnsayo.MostrarElemento .getENSAYO_ID
                    End If
                    If .getPROGRAMA_ID <> 0 Then
                        cmbPrograma.MostrarElemento .getPROGRAMA_ID
                    End If
                    If .getSECTION_ID <> 0 Then
                        cmbSection.MostrarElemento .getSECTION_ID
                    End If
                    If .getFLUID_ID <> 0 Then
                        cmbFluid.MostrarElemento .getFLUID_ID
                    End If
                    If .getFACILITY_ID <> 0 Then
                        cmbFacility.MostrarElemento .getFACILITY_ID
                    End If
                End If
            End With
        End If
    End With

   On Error GoTo 0
   Exit Sub

cargar_bano_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_bano of Formulario frmBANO_Detalle"
End Sub

Private Sub cargar_combos()
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTipos, New clsTipos_muestra, 0, frmTM_Detalle, " TIPO_ESPECIAL_ID <> 0 "
    llenar_combo cmbLineas, New clsLineas, 0, frmLineas, ""
    llenar_combo cmbProcesos, New clsProceso_base, 0, frmProcesosBase, ""
    llenar_combo cmbsolucion, New clsSoluciones, 0, frmSoluciones, ""
    llenar_combo cmbperiodicidad, New clsTipos_Frecuencia, 0, frmTipos_Frecuencia, ""
    llenar_combo cmbenvases, New clsformatos, 0, frmformatos, ""
    llenar_combo cmbprocedencia, New clsSoluciones, 0, frmSoluciones, ""
    llenar_combo cmbtarifa, New clsTarifas_codigos, 0, Me, ""
    llenar_combo cmbTarifaF, New clsTarifas_codigos, 0, Me, ""
    llenar_combo cmbFicha, New clsCe_ficha, 0, frmCE_Ficha, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbCentro, DECODIFICADORA.BANOS_CENTROS
    oDeco.cargar_mi_combo cmbAirbusArea, DECODIFICADORA.BANOS_AREAS
    oDeco.cargar_mi_combo cmbAirbusLinea, DECODIFICADORA.BANOS_LINEAS
    
    oDeco.cargar_mi_combo cmbInstalacion, DECODIFICADORA.BANOS_INSTALACIONES
End Sub
Private Sub anadir_precio()
    If tarifas.ListItems.Count > 0 Then
        If txttarifa.Text = "" Then
            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
            txttarifa.SetFocus
        Else
            tarifa_modificada = True
            tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2) = moneda(txttarifa)
            txttarifa = ""
            If tarifas.ListItems.Count > tarifas.selectedItem.Index Then
                Set tarifas.selectedItem = tarifas.ListItems(tarifas.selectedItem.Index + 1)
                tarifas.SetFocus
                tarifas_Click
            End If
        End If
    End If
End Sub
Private Sub cargar_tarifas()
    Dim oTarifa As New clsTarifas
    Dim rs As ADODB.Recordset
    tarifas.ListItems.Clear
    If cmbclientes.getTEXTO <> "" Then
        Set rs = oTarifa.Listado_por_nombre_Cliente(cmbclientes.getPK_SALIDA)
        If rs.RecordCount <> 0 Then
            Do
                With tarifas.ListItems.Add(, , rs(3))
                    .SubItems(1) = rs(0)
                    .SubItems(2) = " "
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
    End If
End Sub
Private Sub cargar_frecuencias(BANO_ID As Long)
    Dim oBF As New clsBanos_frecuencias
    Dim rs As ADODB.Recordset
    listaFrecuencias.ListItems.Clear
    Set rs = oBF.Listado(PK)
    If rs.RecordCount <> 0 Then
        Do
            With listaFrecuencias.ListItems.Add(, , rs("ORDEN"))
                .SubItems(COLS.C_PERIODICIDAD) = texto(rs("TIPO_FRECUENCIA"))
                .SubItems(COLS.C_CODIGO) = texto(rs("CODIGO"))
                .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = entero(rs("TIPO_FRECUENCIA_ID"))
                .SubItems(COLS.C_CODIGO_ID) = entero(rs("CODIGO_ID"))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cabecera()
    With tarifas.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Tarifa", 1600, lvwColumnLeft
        .Add , , "Precio", 1600, lvwColumnRight
    End With
    With listaFrecuencias.ColumnHeaders
        .Add , , "ORDEN", 1, lvwColumnLeft
        .Add , , "Periodicidad", 3000, lvwColumnCenter
        .Add , , "Cod.Tarifa", 5000, lvwColumnLeft
        .Add , , "TIPO_FRECUENCIA_ID", 1, lvwColumnCenter
        .Add , , "CODIGO_ID", 1, lvwColumnCenter
    End With
End Sub
Private Sub perfil()
    If USUARIO.getPER_FACTURACION = False Then
        Frame2.visible = False
    End If
End Sub
