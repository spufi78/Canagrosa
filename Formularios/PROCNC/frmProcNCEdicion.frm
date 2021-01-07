VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.2#0"; "Codejock.ReportControl.v13.2.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProcNCEdicion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Procedimiento de No Conformidad"
   ClientHeight    =   10035
   ClientLeft      =   1245
   ClientTop       =   930
   ClientWidth     =   14445
   Icon            =   "frmProcNCEdicion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmProcNCEdicion"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   14445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmRevisada 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Revisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4050
      TabIndex        =   69
      Top             =   9045
      Visible         =   0   'False
      Width           =   8025
      Begin VB.CommandButton cmdRevisar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Revisar"
         Height          =   630
         Left            =   6840
         Picture         =   "frmProcNCEdicion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   180
         Width           =   1050
      End
      Begin VB.TextBox txtRevisionFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4095
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   360
         Width           =   2445
      End
      Begin VB.TextBox txtRevisionUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Index           =   10
         Left            =   3420
         TabIndex        =   73
         Top             =   405
         Width           =   450
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   71
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.Frame fraAccionesCorrectivas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Acciones Correctoras / Preventivas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   30
      TabIndex        =   18
      Top             =   6840
      Width           =   14355
      Begin VB.CommandButton cmdEliminarAccCorrectiva 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   630
         Left            =   13245
         Picture         =   "frmProcNCEdicion.frx":6B5C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1500
         Width           =   1050
      End
      Begin VB.CommandButton cmdAnadirAccCorrectiva 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   630
         Left            =   13245
         Picture         =   "frmProcNCEdicion.frx":D3AE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdModificarAccCorrectiva 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   630
         Left            =   13245
         Picture         =   "frmProcNCEdicion.frx":13C00
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1050
      End
      Begin MSComctlLib.ListView lstAccionesCorrectivas 
         Height          =   1875
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   13110
         _ExtentX        =   23125
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
   End
   Begin VB.Frame fraCalidad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Investigación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   12825
      TabIndex        =   16
      Top             =   1650
      Width           =   1635
      Begin XtremeSuiteControls.PushButton cmdOrigen 
         Height          =   765
         Left            =   90
         TabIndex        =   57
         Top             =   300
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Origen de la Incidencia"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProcNCEdicion.frx":1A452
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton cmdPersonal 
         Height          =   765
         Left            =   90
         TabIndex        =   58
         Top             =   1110
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Personal de la Investigación"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProcNCEdicion.frx":20CB4
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton cmdInvestigacionCalidad 
         Height          =   765
         Left            =   90
         TabIndex        =   60
         Top             =   1905
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Investigacion Escena"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProcNCEdicion.frx":27516
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton cmdClasificacion 
         Height          =   765
         Left            =   90
         TabIndex        =   61
         Top             =   2700
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Identif. y clasificación del Problema"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProcNCEdicion.frx":2DD78
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton cmdCausas 
         Height          =   765
         Left            =   90
         TabIndex        =   62
         Top             =   3495
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Causas"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProcNCEdicion.frx":345DA
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton cmdEval 
         Height          =   765
         Left            =   90
         TabIndex        =   59
         Top             =   4290
         Width           =   1455
         _Version        =   851970
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Evaluación Final"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProcNCEdicion.frx":3AE3C
         TextImageRelation=   4
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archivos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   6930
      TabIndex        =   47
      Top             =   4800
      Width           =   5865
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntos"
         Height          =   885
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   585
         Width           =   810
      End
      Begin MSComctlLib.ListView lstDocumentacion 
         Height          =   1590
         Left            =   135
         TabIndex        =   49
         Top             =   225
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame fraInfoBasica 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Acciones Inmediatas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   6930
      TabIndex        =   10
      Top             =   1650
      Width           =   5895
      Begin XtremeReportControl.ReportControl rptAccionesInmediatas 
         Height          =   2025
         Left            =   120
         TabIndex        =   52
         Top             =   270
         Width           =   5670
         _Version        =   851970
         _ExtentX        =   10001
         _ExtentY        =   3572
         _StockProps     =   64
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
         ShowHeader      =   0   'False
         InitialSelectionEnable=   0   'False
         HeaderRowsEnableSelection=   0   'False
      End
      Begin VB.CommandButton cmdAnadirAccInmediata 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5220
         Picture         =   "frmProcNCEdicion.frx":4169E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Añadir Acción Inmediata"
         Top             =   2520
         Width           =   315
      End
      Begin VB.TextBox txtAccionInmediata 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2340
         Width           =   5040
      End
      Begin VB.CommandButton cmdEliminarAccInmediata 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5535
         Picture         =   "frmProcNCEdicion.frx":418C3
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Eliminar acción Inmediata"
         Top             =   2520
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos básicos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3750
      Left            =   0
      TabIndex        =   41
      Top             =   3000
      Width           =   6900
      Begin VB.TextBox txtTitulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   810
         TabIndex        =   3
         Top             =   1425
         Width           =   5955
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   1905
         Left            =   810
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1740
         Width           =   5955
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   360
         Left            =   810
         TabIndex        =   2
         Top             =   1035
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbAuditoria 
         Height          =   360
         Left            =   810
         TabIndex        =   1
         Top             =   630
         Visible         =   0   'False
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbOrigen 
         Height          =   360
         Left            =   810
         TabIndex        =   0
         Top             =   225
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   420
         Left            =   810
         TabIndex        =   65
         Top             =   630
         Visible         =   0   'False
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   741
      End
      Begin pryCombo.miCombo cmbProveedores 
         Height          =   420
         Left            =   810
         TabIndex        =   66
         Top             =   630
         Visible         =   0   'False
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   741
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   56
         Top             =   285
         Width           =   465
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   44
         Top             =   1455
         Width           =   390
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   43
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   42
         Top             =   1080
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha y Responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   0
      TabIndex        =   33
      Top             =   1650
      Width           =   6900
      Begin VB.TextBox txtNIncidencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtResponsableApertura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   585
         Width           =   5610
      End
      Begin VB.TextBox txtDepartamentoResponsable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   900
         Width           =   5610
      End
      Begin MSComCtl2.DTPicker txtFechaPuestaEnMarcha 
         Height          =   315
         Left            =   3465
         TabIndex        =   46
         Top             =   225
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   40156
      End
      Begin MSComCtl2.DTPicker fecha_cierre 
         Height          =   315
         Left            =   5580
         TabIndex        =   67
         Top             =   225
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   40156
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.Cierre"
         Height          =   195
         Index           =   2
         Left            =   4860
         TabIndex        =   68
         Top             =   270
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Particular"
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
         Index           =   9
         Left            =   4635
         TabIndex        =   45
         Top             =   585
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Incidencia"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   40
         Top             =   315
         Width           =   960
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.Apertura"
         Height          =   195
         Index           =   7
         Left            =   2655
         TabIndex        =   39
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resp. Apertura"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   38
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lbCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   945
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1005
      Left            =   2250
      TabIndex        =   19
      Top             =   600
      Width           =   9825
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   25
         Top             =   210
         Visible         =   0   'False
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   1
         Left            =   585
         TabIndex        =   26
         Top             =   540
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         ForeColor       =   255
         Appearance      =   2
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   3
         Left            =   2535
         TabIndex        =   27
         Top             =   540
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   4
         Left            =   3165
         TabIndex        =   28
         Top             =   540
         Visible         =   0   'False
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   5
         Left            =   4725
         TabIndex        =   29
         Top             =   540
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   6
         Left            =   6780
         TabIndex        =   30
         Top             =   540
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   7
         Left            =   8175
         TabIndex        =   31
         Top             =   540
         Visible         =   0   'False
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.RadioButton opEstado 
         Height          =   285
         Index           =   8
         Left            =   8940
         TabIndex        =   32
         Top             =   540
         Width           =   195
         _Version        =   851970
         _ExtentX        =   344
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "RadioButton1"
         Appearance      =   2
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   8190
         TabIndex        =   55
         Top             =   300
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   54
         Top             =   750
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1650
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Line Line1 
         X1              =   750
         X2              =   9060
         Y1              =   675
         Y2              =   690
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abierta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Index           =   1
         Left            =   285
         TabIndex        =   24
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tramitación"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2220
         TabIndex        =   23
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pte. Plan Acc. Correctoras"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3870
         TabIndex        =   22
         Top             =   270
         Width           =   1890
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pte. Cierre"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   6465
         TabIndex        =   21
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cerrado"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   8715
         TabIndex        =   20
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdInformeParcial 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Parcial"
      Height          =   915
      Left            =   1890
      Picture         =   "frmProcNCEdicion.frx":41A57
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9060
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Completo"
      Height          =   915
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9060
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   13365
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9090
      Width           =   1020
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   12315
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9090
      Width           =   1020
   End
   Begin XtremeSuiteControls.PushButton cmdRechazarEstado 
      Height          =   915
      Left            =   30
      TabIndex        =   63
      Top             =   690
      Width           =   2205
      _Version        =   851970
      _ExtentX        =   3889
      _ExtentY        =   1614
      _StockProps     =   79
      Caption         =   "Atrás"
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
      Picture         =   "frmProcNCEdicion.frx":42321
   End
   Begin XtremeSuiteControls.PushButton cmdAvanzarEstado 
      Height          =   945
      Left            =   12150
      TabIndex        =   64
      Top             =   660
      Width           =   2235
      _Version        =   851970
      _ExtentX        =   3942
      _ExtentY        =   1667
      _StockProps     =   79
      Caption         =   "Avanzar"
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
      Picture         =   "frmProcNCEdicion.frx":42BFB
      TextImageRelation=   4
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rellene los datos básicos y las acciones inmediatas para generar una nueva no conformidad"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   51
      Top             =   360
      Width           =   6540
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13815
      Picture         =   "frmProcNCEdicion.frx":434D5
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Procedimiento de No Conformidad"
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
      TabIndex        =   50
      Top             =   45
      Width           =   4320
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "frmProcNCEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarenuNivelAcceso As C_PROCNC_NIVELES_ACCESO
Private mvarenuEstado As C_PROCNC_ESTADOS
' Niveles Acceso
' 0.- No puede hacer nada, solo ver lo que hay pero todo desactivado -> Usuarios sin permisos de no conformidad
' 1.- Puede ver toda la informacion que existe, pero no puede modificar nada -> Usuarios con permiso de No Conformidad, pero no son Responsables de Dpto
' 2.- Puede ver todas la informacion que existe y acceder a las Correctivas que tenga asignadas
' 3.- Responsables de equipo, encargados de Investigacion y Plan inicial de acciones correctivas. No pueden crear, pero si están designados, pueden acceder a sus parcelas
' 4.- Responsables de equipo, Que además son Jefes de Equipo
' 5.- Responsables de Departamento, pueden ver en solo lectura todo, y Crear PNC
' 6.- Acceso Total

Private mvarobjProcNC As New clsProcNc
Const TITULO_FRAME_CALIDAD = "A Completar por los Resposables de Calidad"


Private Sub cargar_datos()

Dim x As Long
Dim oDeco As clsDecodificadora

    
    ' Si es nuevo, solo carga el responsable de la apertura, fecha, etc
    If PK = 0 Then
        txtFechaPuestaEnMarcha = Format(Now, "dd/mm/yyyy")
        txtResponsableApertura.Text = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        txtResponsableApertura.Tag = USUARIO.getID_EMPLEADO
        txtNIncidencia.Text = Format(mvarobjProcNC.CrearID, "000000")
        
        txtDepartamentoResponsable.Text = ""
        txtDepartamentoResponsable.Tag = "0"
        
        For x = 1 To enumDPTO.TOTAL_DEPARTAMENTOS
            If USUARIO.getRESPONSABLE_DEPARTAMENTOS(x) Then
                Set oDeco = New clsDecodificadora
                oDeco.Carga_valor DECODIFICADORA.PROCNC_DEPARTAMENTOS, x
                txtDepartamentoResponsable.Text = oDeco.getDESCRIPCION
                txtDepartamentoResponsable.Tag = CStr(x)
                Set oDeco = Nothing
                Exit For
            End If
        Next x
        Exit Sub
    End If
    
    
    ' Presenta los datos del PNC
    With mvarobjProcNC
'        txtIdParticular.Text = Format(mvarobjProcNC.getID_PARTICULAR, "000000")
        cmbOrigen.BoundText = .getORIGEN_ID
        If .getORIGEN_ID = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE Or _
           .getORIGEN_ID = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_INCIDENCIA_MENOR_CLIENTE Then
            cmbClientes.MostrarElemento .getAUDITORIA_ID
        Else
            If .getORIGEN_ID = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_PROVEEDOR Then
                cmbProveedores.MostrarElemento .getAUDITORIA_ID
            Else
                cmbAuditoria.BoundText = .getAUDITORIA_ID
            End If
        End If
        cmbTipo.BoundText = .getTIPO_ID
        txtNIncidencia.Text = Format(.getID_PROCNC, "0000000")
        txtFechaPuestaEnMarcha = Format(.getFECHA_ALTA, "dd/mm/yyyy")
        
        If Not IsNull(.getFECHA_CIERRE) Then
            fecha_cierre = Format(.getFECHA_CIERRE, "dd/mm/yyyy")
        End If
        If .getREVISADA_USUARIO_ID <> 0 Then
            Dim oUsuario As New clsUsuarios
            oUsuario.CARGAR .getREVISADA_USUARIO_ID
            txtRevisionUsuario = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
            Set oUsuario = Nothing
            txtRevisionFecha = .getREVISADA_FECHA
        End If
        
        txtResponsableApertura.Text = .getRESPONSABLE_NOMBRE_APELLIDOS
        txtResponsableApertura.Tag = .getRESPONSABLE_ID
        txtDepartamentoResponsable.Text = .getRESPONSABLE_DEPARTAMENTO
        txtDepartamentoResponsable.Tag = .getRESPONSABLE_ID_DEPARTAMENTO
        txtTitulo.Text = .getRESUMEN
        txtdescripcion.Text = .getDESCRIPCION_INCIDENCIA
        
        
    End With
    
    
    PresentarDatos_AccionesInmediatas
    PresentarDatos_DocumentosAdjuntos
    PresentarDatos_AccionesCorrectivas

    

End Sub

Private Sub cmborigen_Change()
    Dim oDeco As New clsDecodificadora
    cmbProveedores.visible = False
    cmbClientes.visible = False
    cmbAuditoria.visible = False
    If cmbOrigen.Text <> "" Then
        Select Case cmbOrigen.BoundText
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE, ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_INCIDENCIA_MENOR_CLIENTE
                llenar_combo cmbClientes, New clsCliente, 0, Me, " ANULADO = 0 "
                cmbClientes.visible = True
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_PROVEEDOR
                llenar_combo cmbProveedores, New clsProveedor, 0, Me, " ANULADO = 0 "
                cmbProveedores.visible = True
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_AUDITORIA_INTERNA
                cmbAuditoria.visible = True
                oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_INTERNA
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_AUDITORIA_EXTERNA
                cmbAuditoria.visible = True
                oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_EXTERNA
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_DETECCION_INTERNA
                cmbAuditoria.visible = True
                oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_DETECCION
        End Select
        
    End If
End Sub

Private Sub cmdAnadirAccCorrectiva_Click()

On Error GoTo cmdAnadirAccionCorrectivas_Click_Error

    If PK = 0 Then
        MsgBox "Debe estar creada la NC para poder añadir Acciones.", vbExclamation, App.Title
        Exit Sub
    End If
    
    Dim objfrm As New frmProcNCEdicion_AccionCorrectiva
    Dim objPnc As New clsProcNc
    
    objfrm.PK_PNC = PK
    objfrm.PK = 0
    objfrm.NivelAcceso = mvarenuNivelAcceso
    objfrm.estado_pnc = mvarenuEstado
    
    objfrm.Show vbModal
    
    PresentarDatos_AccionesCorrectivas
    
    Unload objfrm
    Set objfrm = Nothing

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.cmdAnadirAccionCorrectivas_Click"
    Exit Sub
cmdAnadirAccionCorrectivas_Click_Error:
    Set objfrm = Nothing
    Set objPnc = Nothing
    
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.cmdAnadirAccionCorrectivas_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadirAccionCorrectivas_Click of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""


End Sub

Private Sub cmdAnadirAccInmediata_Click()
    If txtAccionInmediata <> "" Then
        Dim Record As ReportRecord
        Dim Item As ReportRecordItem
        Set Record = rptAccionesInmediatas.Records.Add()
        Record.AddItem 0
        Record.AddItem CStr(txtAccionInmediata)
        rptAccionesInmediatas.PaintManager.FixedRowHeight = False
        rptAccionesInmediatas.Populate
        txtAccionInmediata = ""
        txtAccionInmediata.SetFocus
    End If
End Sub

Private Sub cmdAvanzarEstado_Click()
On Error GoTo cmdCambiarEstado_Click_Error

    If mvarenuEstado = C_PROCNC_ESTADOS.ABIERTA Then
'        Call CambiarEstado_a_VoBo
        Call CambiarEstado_a_Tramitacion
'    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
'        Call CambiarEstado_a_Tramitacion
    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.EN_TRAMITACION Then
'        Call CambiarEstado_a_PteConfirmacionCalidad
        Call CambiarEstado_a_PtePlanAccionesCorrectivas
'    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD Then
'        Call CambiarEstado_a_PtePlanAccionesCorrectivas
    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
        Call CambiarEstado_a_PteCierre
    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.pte_cierre Then
'        Call CambiarEstado_a_CerradaParcial
        Call CambiarEstado_a_CerradaTotal
'    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL Then
'        Call CambiarEstado_a_CerradaTotal
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdCambiarEstado_Click"
    Exit Sub
cmdCambiarEstado_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdCambiarEstado_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdCambiarEstado_Click of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdCausas_Click()
    frmProcNCCausasProblemas.PK = PK
    frmProcNCCausasProblemas.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL & mvarenuEstado < C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    frmProcNCCausasProblemas.Show vbModal
End Sub

Private Sub cmdClasificacion_Click()
    frmProcNCClasificacionProblema.PK = PK
    frmProcNCClasificacionProblema.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL & mvarenuEstado < C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    frmProcNCClasificacionProblema.Show vbModal
End Sub

Private Sub cmdEliminarAccCorrectiva_Click()
 
    If Not cmdEliminarAccCorrectiva.Enabled Then Exit Sub
 
    If lstAccionesCorrectivas.ListItems.Count = 0 Then Exit Sub
    
    lngid = lstAccionesCorrectivas.selectedItem
    
    mvarobjProcNC.eliminar_accion_correctiva lngid
        
    PresentarDatos_AccionesCorrectivas
    
End Sub

Private Sub cmdEval_Click()
frmProcNCEvaluacion.PK = PK
frmProcNCEvaluacion.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL) And (mvarenuEstado < C_PROCNC_ESTADOS.CERRADA)
frmProcNCEvaluacion.Show vbModal
End Sub

Private Sub cmdInforme_Click()
    
On Error GoTo cmdImprimir_Click_Error

    With frmReport
        .iniciar
        .informe = "/NC/rptProcNCCompleto"
        .criterio = "{procnc.ID_PROCNC} = " & CStr(mvarobjProcNC.getID_PROCNC) & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        '.Visible = True
        .Show vbModal
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdImprimir_Click"
    Exit Sub
cmdImprimir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdImprimir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub cmdInformeParcial_Click()
On Error GoTo cmdInformeParcial_Click_Error
    Dim c As String
    c = "{procnc.ID_PROCNC} = " & CStr(mvarobjProcNC.getID_PROCNC)
    c = c & " and {decodificadora.CODIGO}=110"
    c = c & " and {decodificadora_tipos.CODIGO}=119"
'    c = c & " and {decodificadora_tipos_acc.CODIGO}=" & DECODIFICADORA.PROCNC_ACCIONES_TIPOS

    With frmReport
        .iniciar
        .informe = "/NC/rptProcNC"
        .criterio = c
        .imprimir = False
        .generar
        '.Visible = True
        .Show vbModal
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdInformeParcial_Click"
    Exit Sub
cmdInformeParcial_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdInformeParcial_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdInformeParcial_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub cmdInvestigacionCalidad_Click()
    frmProcNCInvestigacionEscena.PK = PK
    
    If mvarenuEstado = C_PROCNC_ESTADOS.EN_TRAMITACION Then
        frmProcNCInvestigacionEscena.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.JEFE_EQUIPO_INVESTIGACION Or mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
    Else
        frmProcNCInvestigacionEscena.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL & mvarenuEstado < C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    End If
    
    frmProcNCInvestigacionEscena.Show vbModal

End Sub

Private Sub cmdOrigen_Click()
    frmProcNCOrigenes.PK = PK
    frmProcNCOrigenes.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL & mvarenuEstado < C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    frmProcNCOrigenes.Show vbModal
End Sub

Private Sub cmdPersonal_Click()
    frmProcNCPersonalInvestigacion.PK = PK
    frmProcNCPersonalInvestigacion.Editable = (mvarenuNivelAcceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL & mvarenuEstado < C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    frmProcNCPersonalInvestigacion.Show vbModal
End Sub

Private Sub cmdRechazarEstado_Click()
On Error GoTo cmdRechazarEstado_Click_Error

    If mvarenuEstado = C_PROCNC_ESTADOS.pte_cierre Then
        Call cambiarEstado_RechazarConfirmacionCalidad
    End If
    
'    If mvarenuEstado = C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
'        Call cambiarEstado_RechazarVoBo
'    ElseIf mvarenuEstado = C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD Then
'        Call cambiarEstado_RechazarConfirmacionCalidad
'    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdRechazarEstado_Click"
    Exit Sub
cmdRechazarEstado_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdRechazarEstado_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdRechazarEstado_Click of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub cmdEliminarAccInmediata_Click()
    If rptAccionesInmediatas.Rows.Count = 0 Then Exit Sub
    rptAccionesInmediatas.RemoveRowEx rptAccionesInmediatas.SelectedRows(0)
    rptAccionesInmediatas.Populate
End Sub

Private Sub cmdModificarAccCorrectiva_Click()

    Dim objfrm As New frmProcNCEdicion_AccionCorrectiva

   On Error GoTo cmdModificarAccCorrectiva_Click_Error

    If Not cmdModificarAccCorrectiva.Enabled Then Exit Sub
    If lstAccionesCorrectivas.ListItems.Count = 0 Then Exit Sub
    
    lngid = lstAccionesCorrectivas.selectedItem
    
    If lngid <= 0 Then Exit Sub
        
    objfrm.PK = lngid
    objfrm.PK_PNC = PK
    objfrm.NivelAcceso = mvarenuNivelAcceso
    objfrm.estado_pnc = mvarenuEstado
    
    objfrm.Show vbModal
    
    Unload objfrm
    Set objfrm = Nothing
    
    PresentarDatos_AccionesCorrectivas

   On Error GoTo 0
   Exit Sub

cmdModificarAccCorrectiva_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarAccCorrectiva_Click of Formulario frmProcNCEdicion"
End Sub

Private Sub cmdok_Click()
    If Not guardar_datos Then Exit Sub
    MsgBox "Datos almacenados correctamente.", vbApplicationModal + vbInformation, App.Title
    inicio_carga
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdRevisar_Click()
    If MsgBox("¿Esta seguro/a de marcar como revisada la incidencia?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oPROCNC As New clsProcNc
        oPROCNC.marcarRevisada PK
        MsgBox "La incidencia se ha marcado como revisada correctamente.", vbInformation, App.Title
        Unload Me
        Set oPROCNC = Nothing
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.MousePointer = 0
    Select Case KeyCode
        Case 27
            cmdcancel_Click
        Case 116 ' F5 Datos especiales
            cmdok.Enabled = Not cmdok.Enabled
    End Select

End Sub

Private Sub Form_Load()
    log Me.Name
    cabecera
    cargar_botones Me
    ' En la siguiente funcion se debe poner todo lo que se deba cargar cada vez que se guarde
    inicio_carga
End Sub

Private Sub inicio_carga()
    If PK <> 0 Then mvarobjProcNC.Carga PK
    establecer_nivel_acceso
    If PK = 0 Then
        mvarenuEstado = C_PROCNC_ESTADOS.ABIERTA
    Else
        mvarenuEstado = mvarobjProcNC.getESTADO_ID
    End If
    establecer_opciones_estado
    cargar_datos
End Sub
Private Sub establecer_nivel_acceso()
    Dim res As Integer
    res = mvarobjProcNC.establecer_nivel_acceso()
    If res = -1 Then Exit Sub
    mvarenuNivelAcceso = res
    Exit Sub
End Sub

Private Sub establecer_opciones_estado()
'    If PK = 0 Then
'        mvarenuEstado = C_PROCNC_ESTADOS.ABIERTA
'    Else
'        mvarenuEstado = mvarobjProcNC.getESTADO_ID
'    End If
    
    opEstado(mvarenuEstado).Value = True
    
    Select Case mvarenuEstado
        Case C_PROCNC_ESTADOS.ABIERTA
            establecer_opciones_estado_abierta
'        Case C_PROCNC_ESTADOS.PTE_VISTO_BUENO
'            establecer_opciones_estado_pte_VoBo
        Case C_PROCNC_ESTADOS.EN_TRAMITACION
            establecer_opciones_estado_tramitacion
'        Case C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD
'            establecer_opciones_estado_pte_confirmacion_calidad
        Case C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS
            establecer_opciones_estado_pte_plan_correctivas
        Case C_PROCNC_ESTADOS.pte_cierre
            establecer_opciones_estado_pte_cierre
'        Case C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL
'            establecer_opciones_estado_cierre_parcial
        Case C_PROCNC_ESTADOS.CERRADA
            establecer_opciones_estado_cierre_total
    End Select
End Sub

Private Sub establecer_opciones_estado_abierta()
'    fraCalidad.Visible = False
    cmdInforme.visible = False
    cmdInformeParcial.visible = False
        
    cmdAvanzarEstado.visible = True
    cmdAvanzarEstado.Caption = "Enviar a Tramitar"
    cmdRechazarEstado.visible = False
    ' OPCIONES DE EDICION
    'If fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas -1
    
    cmdok.Enabled = (mvarenuNivelAcceso >= RESPONSABLE_DEPARTAMENTO)
    cmdOrigen.Enabled = False
    cmdPersonal.Enabled = False
    cmdInvestigacionCalidad.Enabled = False
    cmdClasificacion.Enabled = False
    cmdCausas.Enabled = False
    cmdEval.Enabled = False
    
    cmdAdjuntar.Enabled = True
    cmdAnadirAccInmediata.Enabled = True
    cmdEliminarAccInmediata.Enabled = True
    
    cmdAnadirAccCorrectiva.Enabled = False
    cmdEliminarAccCorrectiva.Enabled = False
    cmdModificarAccCorrectiva.Enabled = False
    
    cmdInforme.Enabled = False
    cmdInformeParcial.Enabled = False
    
    cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso >= RESPONSABLE_DEPARTAMENTO)
    cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.RESPONSABLE_DEPARTAMENTO)
    cmbAuditoria.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.RESPONSABLE_DEPARTAMENTO)
    
    cmbOrigen.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.RESPONSABLE_DEPARTAMENTO)
End Sub

'Private Sub establecer_opciones_estado_pte_VoBo()
'    fraCalidad.Visible = True
'    fraCalidad.Caption = TITULO_FRAME_CALIDAD
'    cmdOrigen.Visible = True
'    cmdPersonal.Visible = True
'
'    Set cmdOrigen.Picture = frmMenu.botones.ListImages(6).Picture
'    Set cmdPersonal.Picture = frmMenu.botones.ListImages(6).Picture
'
'
'    cmdInvestigacionCalidad.Visible = False
'    cmdClasificacion.Visible = False
'    cmdCausas.Visible = False
'    cmdEval.Visible = False
'
'    cmdInforme.Visible = False
'    cmdInformeParcial.Visible = False
'
'
'    cmdAvanzarEstado.Visible = True
'    cmdAvanzarEstado.Caption = "Tramitar PNC"
'    cmdRechazarEstado.Visible = True
'    cmdRechazarEstado.Caption = "Rechazar VºBº"
'
'    ' OPCIONES DE EDICION
'
'    'If fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas -1
'
'    cmdok.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdOrigen.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdPersonal.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdClasificacion.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdCausas.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdEval.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdAnadirAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdEliminarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdInforme.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdInformeParcial.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdRechazarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    txtTitulo.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
'    txtDescripcion.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
'    cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
'
'End Sub

Private Sub establecer_opciones_estado_tramitacion()
    
' Depende del nivel de acceso, se ve una cosa y otra
    
    ' cuando es gerencia, calidad o responsable de departamento
    If mvarenuNivelAcceso = ACCESO_TOTAL Or mvarenuNivelAcceso = RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Then
        fraCalidad.visible = True
'        fraCalidad.Caption = TITULO_FRAME_CALIDAD
        
        cmdOrigen.visible = True
        cmdPersonal.visible = True
        
        cmdInvestigacionCalidad.visible = True
        
'        Set cmdOrigen.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdPersonal.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdInvestigacionCalidad.Picture = frmMenu.botones.ListImages(6).Picture
    
        
        cmdClasificacion.visible = False
        cmdCausas.visible = False
        cmdEval.visible = False
        
            
    ElseIf mvarenuNivelAcceso = JEFE_EQUIPO_INVESTIGACION Then
        fraCalidad.visible = False
        
    Else
        fraCalidad.visible = False
    End If
    
    cmdInforme.visible = True
    cmdInformeParcial.visible = True
    
    cmdAvanzarEstado.visible = True
    cmdAvanzarEstado.Caption = "Plan de Acc.Correctoras"
    cmdRechazarEstado.visible = False
    
    ' OPCIONES DE EDICION
    
    'If fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas -1
    
    cmdok.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdOrigen.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdPersonal.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdClasificacion.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdCausas.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdEval.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    
    cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    
    cmdAnadirAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdEliminarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    
    cmdInforme.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdInformeParcial.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    
    cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
    cmbAuditoria.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
    cmbOrigen.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
        
    
End Sub

'Private Sub establecer_opciones_estado_pte_confirmacion_calidad()
'
'        fraCalidad.Visible = True
'        fraCalidad.Caption = TITULO_FRAME_CALIDAD
'
'        cmdOrigen.Visible = True
'        cmdPersonal.Visible = True
'        cmdInvestigacionCalidad.Visible = True
'
'        cmdClasificacion.Visible = True
'        cmdCausas.Visible = True
'
'        Set cmdOrigen.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdPersonal.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdInvestigacionCalidad.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdClasificacion.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdCausas.Picture = frmMenu.botones.ListImages(8).Picture
'
'        If mvarenuNivelAcceso = ACCESO_TOTAL Then
'            Set cmdClasificacion.Picture = frmMenu.botones.ListImages(6).Picture
'            Set cmdCausas.Picture = frmMenu.botones.ListImages(6).Picture
'        End If
'
'        cmdEval.Visible = False
'
'
'        cmdInforme.Visible = True
'        cmdInformeParcial.Visible = True
'
'        cmdAvanzarEstado.Visible = True
'        cmdAvanzarEstado.Caption = "Solicitar Plan Acc. Correctivas"
'        cmdRechazarEstado.Visible = True
'        cmdRechazarEstado.Caption = "Rechazar Confirmación"
'
'    ' OPCIONES DE EDICION
'
'    'If fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas -1
'
'    cmdok.Enabled = (mvarenuNivelAcceso >= ACCESO_TOTAL)
'    cmdOrigen.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'    cmdPersonal.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'    cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'    cmdClasificacion.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'    cmdCausas.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'    cmdEval.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdAnadirAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdEliminarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'    cmdInforme.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'    cmdInformeParcial.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'
'    cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    cmdRechazarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'    txtTitulo.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
'    txtDescripcion.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
'    cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
'
'
'
'End Sub

Private Sub establecer_opciones_estado_pte_plan_correctivas()
        
'        If mvarenuNivelAcceso = JEFE_EQUIPO_INVESTIGACION Then
'            fraCalidad.Caption = "Informacion Proc. No Conformidad"
'        Else
'            fraCalidad.Caption = TITULO_FRAME_CALIDAD
'        End If
                
        fraCalidad.visible = True
        
        cmdOrigen.visible = True
        cmdPersonal.visible = True
        cmdInvestigacionCalidad.visible = True
        cmdClasificacion.visible = True
        cmdCausas.visible = True
        cmdEval.visible = False
        
'        Set cmdOrigen.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdPersonal.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdInvestigacionCalidad.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdClasificacion.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdCausas.Picture = frmMenu.botones.ListImages(8).Picture
        
        
    
        cmdInforme.visible = True
        cmdInformeParcial.visible = True
    
        cmdAvanzarEstado.visible = True
        cmdAvanzarEstado.Caption = "Enviar Plan A Calidad"
        cmdRechazarEstado.visible = False
        
    ' OPCIONES DE EDICION
    
    'If Not fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas 1
    
    cmdok.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdOrigen.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdPersonal.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdClasificacion.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdCausas.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdEval.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    
    cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    
    cmdAnadirAccCorrectiva.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdEliminarAccCorrectiva.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    
    cmdInforme.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdInformeParcial.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    
    cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION)
    cmdRechazarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    txtTitulo.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
    txtdescripcion.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
    cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
    cmbAuditoria.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
    cmbOrigen.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
        
End Sub

Private Sub establecer_opciones_estado_pte_cierre()
        
        fraCalidad.visible = True
'        fraCalidad.Caption = TITULO_FRAME_CALIDAD
        
        cmdOrigen.visible = True
        cmdPersonal.visible = True
        cmdInvestigacionCalidad.visible = True
        cmdClasificacion.visible = True
        cmdCausas.visible = True
        cmdEval.visible = True
        
'        Set cmdOrigen.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdPersonal.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdInvestigacionCalidad.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdClasificacion.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdCausas.Picture = frmMenu.botones.ListImages(8).Picture
            
    
        cmdInforme.visible = True
        cmdInformeParcial.visible = True
    
        cmdAvanzarEstado.visible = True
        cmdAvanzarEstado.Caption = "Cerrar P.N.C."
        cmdRechazarEstado.visible = True
        cmdRechazarEstado.Caption = "Plan de Acc.Correctoras"
        cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
        cmbAuditoria.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
        cmbOrigen.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
        
        
    ' OPCIONES DE EDICION
    
    'If Not fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas 1
    
    cmdok.Enabled = (mvarenuNivelAcceso >= ACCESO_TOTAL)
    cmdOrigen.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
    cmdPersonal.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
    cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
    cmdClasificacion.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
    cmdCausas.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
    cmdEval.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    
    cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    
    cmdAnadirAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdEliminarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
    
    cmdInforme.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    cmdInformeParcial.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
    
    cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    cmdRechazarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
    txtTitulo.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
    txtdescripcion.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
        
End Sub

'Private Sub establecer_opciones_estado_cierre_parcial()
'
'        fraCalidad.Visible = True
'        fraCalidad.Caption = TITULO_FRAME_CALIDAD
'
'        cmdOrigen.Visible = True
'        cmdPersonal.Visible = True
'        cmdInvestigacionCalidad.Visible = True
'        cmdClasificacion.Visible = True
'        cmdCausas.Visible = True
'        cmdEval.Visible = True
'
'        Set cmdOrigen.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdPersonal.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdInvestigacionCalidad.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdClasificacion.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdCausas.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdEval.Picture = frmMenu.botones.ListImages(6).Picture
'
'
'        cmdInforme.Visible = True
'        cmdInformeParcial.Visible = True
'
'        cmdAvanzarEstado.Visible = True
'        cmdAvanzarEstado.Caption = "Cerrar P.N.C."
'        cmdRechazarEstado.Visible = False
'        cmbTipo.Enabled = (mvarenuNivelAcceso >= C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL)
'        ' OPCIONES DE EDICION
'
'        'If Not fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas 1
'
'        cmdok.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdOrigen.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
'        cmdPersonal.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
'        cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
'        cmdClasificacion.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
'        cmdCausas.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
'        cmdEval.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'        cmdInvestigacion.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'
'        cmdAnadirAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdEliminarAccCorrectiva.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
'
'        cmdInforme.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'        cmdInformeParcial.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
'
'        cmdAvanzarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        cmdRechazarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
'        txtTitulo.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
'        txtDescripcion.Locked = (mvarenuNivelAcceso < ACCESO_TOTAL)
'
'End Sub

Private Sub establecer_opciones_estado_cierre_total()
        fraCalidad.visible = True
'        fraCalidad.Caption = "Informacion PNC"
        
        cmdOrigen.visible = True
        cmdPersonal.visible = True
        cmdInvestigacionCalidad.visible = True
        cmdClasificacion.visible = True
        cmdCausas.visible = True
        cmdEval.visible = True
        
        lblCap(2).visible = True
        fecha_cierre.visible = True
        Dim oPROCNC As New clsProcNc
        If oPROCNC.habilitarRevision(PK) Then
            frmRevisada.visible = True
            If mvarobjProcNC.getREVISADA_USUARIO_ID <> 0 Then
                cmdRevisar.visible = False
            End If
        End If
        
'        Set cmdOrigen.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdPersonal.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdInvestigacionCalidad.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdClasificacion.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdCausas.Picture = frmMenu.botones.ListImages(8).Picture
'        Set cmdEval.Picture = frmMenu.botones.ListImages(8).Picture
        
    
        cmdInforme.visible = True
        cmdInformeParcial.visible = True
    
        cmdAvanzarEstado.visible = False
        cmdRechazarEstado.visible = False
        
        ' Opciones de Id Particular
        'lblCap(10).Visible = True
'        txtIdParticular.Visible = False
        
        ' OPCIONES DE EDICION

        'If Not fraAccionesCorrectivas.Visible Then des_acoplar_acc_correctivas 1

        cmdok.Enabled = False ' (mvarenuNivelAcceso = ACCESO_TOTAL)
        cmdOrigen.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
        cmdPersonal.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
        cmdInvestigacionCalidad.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
        cmdClasificacion.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
        cmdCausas.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO Or mvarenuNivelAcceso = SOLO_CORRECTIVAS_ASIGNADAS)
        cmdEval.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
        
        cmdAdjuntar.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
        cmdAnadirAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
        cmdEliminarAccInmediata.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
        
        cmdAnadirAccCorrectiva.Enabled = False '(mvarenuNivelAcceso = ACCESO_TOTAL)
        cmdEliminarAccCorrectiva.Enabled = False '(mvarenuNivelAcceso = ACCESO_TOTAL)
        cmdModificarAccCorrectiva.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
        
        cmdInforme.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
        cmdInformeParcial.Enabled = (mvarenuNivelAcceso >= JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO)
        
        cmdAvanzarEstado.Enabled = False '(mvarenuNivelAcceso = ACCESO_TOTAL)
        cmdRechazarEstado.Enabled = (mvarenuNivelAcceso = ACCESO_TOTAL)
        txtTitulo.Locked = True ' (mvarenuNivelAcceso < ACCESO_TOTAL)
        txtdescripcion.Locked = True ' (mvarenuNivelAcceso < ACCESO_TOTAL)
        cmbTipo.Enabled = False
        cmbAuditoria.Enabled = False
        cmbOrigen.Enabled = False
End Sub

Private Sub cmdAdjuntar_Click()
    If PK = 0 Then
        If MsgBox("Debe Guardar el PNC para poder adjuntarle un archivo.", vbInformation + vbYesNo, "Adjuntar Archivo a PNC") = vbNo Then
            Exit Sub
        Else
            If Not guardar_datos Then Exit Sub
        End If
    End If
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PROCNC_INCIDENCIA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    
    Call PresentarDatos_DocumentosAdjuntos
End Sub

Private Function guardar_datos() As Boolean
    guardar_datos = False
    If Not comprobar_datos Then Exit Function
    
    recoger_datos
    If PK = 0 Then
        mvarobjProcNC.Insertar
        PK = mvarobjProcNC.getID_PROCNC
    Else
        mvarobjProcNC.Modificar
    End If
    ' Acciones Inmediatas
    Dim oPAI As New clsProcnc_accionesinmediatas
    oPAI.Eliminar PK
    Dim i As Integer
    For i = 0 To rptAccionesInmediatas.Records.Count - 1
        oPAI.setID_PROCNC = PK
        oPAI.setDESCRIPCION = rptAccionesInmediatas.Records(i).Item(1).Caption
        oPAI.Insertar
    Next
    guardar_datos = True
End Function

Private Function comprobar_datos() As Boolean
On Error GoTo comprobar_datos_Error
    
    comprobar_datos = False
    
    Dim strMsg  As String
    strMsg = ""
    
    
    If cmbOrigen.BoundText = "" Then
        strMsg = strMsg & vbCrLf & " - Debe indicar el Origen del PNC"
    End If
    
    If cmbOrigen.BoundText = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE Or _
       cmbOrigen.BoundText = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_INCIDENCIA_MENOR_CLIENTE Then
        If cmbClientes.getTEXTO = "" Then
            strMsg = strMsg & vbCrLf & " - Debe indicar el CLIENTE del PNC"
        End If
    Else
        If cmbOrigen.BoundText = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_PROVEEDOR Then
            If cmbProveedores.getTEXTO = "" Then
                strMsg = strMsg & vbCrLf & " - Debe indicar el PROVEEDOR del PNC"
            End If
        Else
            If cmbAuditoria.BoundText = "" Then
                strMsg = strMsg & vbCrLf & " - Debe indicar el ORIGEN (auditoria o deteccion interna) "
            End If
        End If
    End If
    
    If cmbTipo.BoundText = "" Then
        strMsg = strMsg & vbCrLf & " - Debe indicar el TIPO del PNC"
    End If
    
    If Trim(txtTitulo.Text) = "" Then
        strMsg = strMsg & vbCrLf & " - Debe indicar el Título del PNC"
    End If
    
    If Trim(txtdescripcion.Text) = "" Then
        strMsg = strMsg & vbCrLf & " - Debe indicar una Descripción para el PNC"
    End If
    
    
    If Trim(strMsg) <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strMsg, vbInformation, "Guardar Proc. No Conformidad"
        Exit Function
    End If
    
    comprobar_datos = True
    
On Error GoTo 0
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.comprobar_datos"
    Exit Function
comprobar_datos_Error:
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.comprobar_datos"
    'error_grave Err.Number & " (" & Err.Description & ") in procedure comprobar_datos of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    'G_TRAZABILIDAD_ERROR = ""
    comprobar_datos = False
End Function

Private Sub recoger_datos()
    With mvarobjProcNC
        .setRESUMEN = txtTitulo.Text
        .setTIPO_ID = getDataComboSel(cmbTipo, 0)
        .setDESCRIPCION_INCIDENCIA = txtdescripcion.Text
        .setESTADO_ID = mvarenuEstado
        If cmbOrigen.BoundText = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE Or _
           cmbOrigen.BoundText = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_INCIDENCIA_MENOR_CLIENTE Then
            .setAUDITORIA_ID = cmbClientes.getPK_SALIDA
        Else
            If cmbOrigen.BoundText = ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_PROVEEDOR Then
                .setAUDITORIA_ID = cmbProveedores.getPK_SALIDA
            Else
                .setAUDITORIA_ID = getDataComboSel(cmbAuditoria, 0)
            End If
        End If
        .setORIGEN_ID = getDataComboSel(cmbOrigen, 0)
        .setFECHA_ALTA = CDate(txtFechaPuestaEnMarcha)
        If PK = 0 Then
            .setRESPONSABLE_ID = CLng(txtResponsableApertura.Tag)
            .setRESPONSABLE_ID_DEPARTAMENTO = CLng(txtDepartamentoResponsable.Tag)
        End If
        'M1315-I
        .setN_ACCIONES = lstAccionesCorrectivas.ListItems.Count
        'M1315-F
        If opEstado(8).Value = True Then
            .setFECHA_CIERRE = fecha_cierre
        End If
    End With
End Sub
Private Sub lstAccionesCorrectivas_DblClick()
    cmdModificarAccCorrectiva_Click
End Sub

Private Sub PresentarDatos_AccionesInmediatas()
    Dim oPAI As New clsProcnc_accionesinmediatas
    Dim rs As ADODB.Recordset
    Dim Record As ReportRecord
    Dim Item As ReportRecordItem
    rptAccionesInmediatas.ClearContent
    
    Set rs = oPAI.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            Set Record = rptAccionesInmediatas.Records.Add()
            Record.AddItem rs("id_accion_inmediata")
            Record.AddItem CStr(rs("DESCRIPCION"))
            rs.MoveNext
        Loop Until rs.EOF
    End If
    rptAccionesInmediatas.PaintManager.FixedRowHeight = False
    rptAccionesInmediatas.Populate
End Sub

Private Sub PresentarDatos_AccionesCorrectivas()
    Dim rs As ADODB.Recordset
    Set rs = mvarobjProcNC.devolver_listado_acciones_correctivas
    lstAccionesCorrectivas.ListItems.Clear
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        Exit Sub
    End If
    
    rs.MoveFirst
    While Not rs.EOF
        With lstAccionesCorrectivas.ListItems.Add(, , rs("id_accion_correctiva"))
            .SubItems(1) = rs("responsable_id")
            .SubItems(2) = rs("tipo")
            .SubItems(3) = rs("titulo")
            .SubItems(4) = rs("RESPONSABLE")
            .SubItems(5) = rs("ESTADO")
            .SubItems(6) = rs("fecha_puesta_en_marcha")
            .SubItems(7) = rs("fecha_prevista")
        End With
        rs.MoveNext
    Wend
    
    Set rs = Nothing
    
End Sub


Private Sub cabecera()
    ' Acciones inmediatas
    rptAccionesInmediatas.AllowColumnRemove = False
    Dim Column As ReportColumn, RecordControl As ReportRecord, RecordHeader As ReportRecord, RecordPaintManager As ReportRecord
    Set Column = rptAccionesInmediatas.Columns.Add(0, "ID", 0, False)
    Column.Alignment = xtpAlignmentWordBreak
    Set Column = rptAccionesInmediatas.Columns.Add(1, "Acción Inmediata", 360, False)
    Column.Alignment = xtpAlignmentWordBreak
    
    
    ' Adjuntos
    With lstDocumentacion.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "Documento", lstDocumentacion.Width, lvwColumnLeft
    End With
    
    With lstAccionesCorrectivas.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "id_responsable", 0, lvwColumnLeft
        .Add , , "Tipo", 1300, lvwColumnCenter
        .Add , , "Accion", 4500, lvwColumnLeft
        .Add , , "Responsable", 2500, lvwColumnLeft
        .Add , , "Estado", 1500, lvwColumnLeft
        .Add , , "F.Puesta en Marcha", 1500, lvwColumnCenter
        .Add , , "F.Fin Prevista", 1500, lvwColumnCenter
    End With
    
    'carga los tipos de pnc
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.PROCNC_TIPOS_NO_CONFORMIDAD
'    oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_AUDITORIAS
    oDeco.cargar_combo cmbOrigen, DECODIFICADORA.PROCNC_ORIGEN
    Set oDeco = Nothing
End Sub

Private Sub PresentarDatos_DocumentosAdjuntos()
    
    lstDocumentacion.ListItems.Clear
    Dim oAdjunto As New clsAdjuntos
    Dim rs As ADODB.Recordset
    Set rs = oAdjunto.Listado(TOBJETO.TOBJETO_PROCNC_INCIDENCIA, PK, "", "")
    If rs.RecordCount > 0 Then
        Do
            With lstDocumentacion.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

End Sub

Private Sub lstDocumentacion_DblClick()
On Error GoTo fallo
    If lstDocumentacion.ListItems.Count = 0 Then Exit Sub
    Dim oAdjunto As New clsAdjuntos
    oAdjunto.CargarDocumento TOBJETO.TOBJETO_PROCNC_INCIDENCIA, PK, 0, lstDocumentacion.ListItems(lstDocumentacion.selectedItem.Index).Text, True
    Set oAdjunto = Nothing
    Exit Sub
fallo:
    MsgBox "No es posible mostrar el documento. Consulte con el Administrador del Sistema", vbInformation, "Mostrar Documento Adjunto"

End Sub

'Private Sub CambiarEstado_a_VoBo()
'
'On Error GoTo CambiarEstado_a_VoBo_Error
'
'    If Not guardar_datos Then Exit Sub
'
'    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.PTE_VISTO_BUENO)
'
'    Unload Me
'
'On Error GoTo 0
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_VoBo"
'    Exit Sub
'CambiarEstado_a_VoBo_Error:
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_VoBo"
'    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_VoBo of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
'    G_TRAZABILIDAD_ERROR = ""
'
'End Sub
Private Sub CambiarEstado_a_Tramitacion()
    
On Error GoTo CambiarEstado_a_Tramitacion_Error

'    If Not mvarobjProcNC.comprobar_paso_a_tramitacion Then Exit Sub
    If MsgBox("Va a enviar la incidencia a Tramitación. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.EN_TRAMITACION)
    
    Unload Me

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_Tramitacion"
    Exit Sub
CambiarEstado_a_Tramitacion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_Tramitacion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_Tramitacion of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub CambiarEstado_a_PteConfirmacionCalidad()
    
On Error GoTo CambiarEstado_a_PteConfirmacionCalidad_Error

    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD)
    
    Unload Me


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_PteConfirmacionCalidad"
    Exit Sub
CambiarEstado_a_PteConfirmacionCalidad_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_PteConfirmacionCalidad"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_PteConfirmacionCalidad of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub CambiarEstado_a_PtePlanAccionesCorrectivas()
On Error GoTo CambiarEstado_a_PtePlanAccionesCorrectivas_Error

    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS)
    
    Unload Me

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_PtePlanAccionesCorrectivas"
    Exit Sub
CambiarEstado_a_PtePlanAccionesCorrectivas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_PtePlanAccionesCorrectivas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_PtePlanAccionesCorrectivas of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub CambiarEstado_a_PteCierre()
On Error GoTo CambiarEstado_a_PteCierre_Error

    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.pte_cierre)
    
    Unload Me

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_PteCierre"
    Exit Sub
CambiarEstado_a_PteCierre_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_PteCierre"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_PteCierre of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub CambiarEstado_a_CerradaParcial()
On Error GoTo CambiarEstado_a_CerradaParcial_Error

    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    
    Unload Me

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_CerradaParcial"
    Exit Sub
CambiarEstado_a_CerradaParcial_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_CerradaParcial"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_CerradaParcial of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub CambiarEstado_a_CerradaTotal()
On Error GoTo CambiarEstado_a_CerradaTotal_Error

    If Not ComprobarCierreTotalPosible() Then Exit Sub
    
    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.CERRADA)
    
    Unload Me
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_CerradaTotal"
    Exit Sub
CambiarEstado_a_CerradaTotal_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.CambiarEstado_a_CerradaTotal"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_CerradaTotal of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub cambiarEstado_RechazarVoBo()

    Dim objfrm As New frmProcNC_CambioEstadoRazones
    Dim strMensajeRechazo As String
    
On Error GoTo cambiarEstado_RechazarVoBo_Error

    objfrm.titulo = "Rechazar Vº Bº"
    
    objfrm.Show vbModal
    
    If Not objfrm.resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    strMensajeRechazo = objfrm.MotivoRechazo
    
    'If MsgBox("Para Proceder a enviar la Notificación de Rechazo del VºBº Incidencia al Responsable su apertura, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Rechazar VºBº P.N.C.") = vbNo Then Exit Sub
    
    If Not guardar_datos Then Exit Sub
    
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.ABIERTA, strMensajeRechazo)
        
    Unload Me

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.cambiarEstado_RechazarVoBo"
    Exit Sub
cambiarEstado_RechazarVoBo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.cambiarEstado_RechazarVoBo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cambiarEstado_RechazarVoBo of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub


Private Sub cambiarEstado_RechazarConfirmacionCalidad()

    Dim objfrm As New frmProcNC_CambioEstadoRazones
    Dim strMensajeRechazo As String
    
On Error GoTo cambiarEstado_RechazarConfirmacionCalidad_Error

    objfrm.titulo = "Rechazar Confirmación Calidad"
    
    objfrm.Show vbModal
    
    If Not objfrm.resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    strMensajeRechazo = objfrm.MotivoRechazo
    
    If Not guardar_datos Then Exit Sub

'    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.EN_TRAMITACION, strMensajeRechazo)
    Call mvarobjProcNC.cambiarEstado(C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS, strMensajeRechazo)
    
    Unload Me

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.cambiarEstado_RechazarConfirmacionCalidad"
    Exit Sub
cambiarEstado_RechazarConfirmacionCalidad_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.cambiarEstado_RechazarConfirmacionCalidad"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cambiarEstado_RechazarConfirmacionCalidad of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub





Private Function ComprobarCierreTotalPosible() As Boolean

    Dim objAcc As clsProcNcAccionCorrectora
    Dim blnRes As Boolean
    Dim strMensaje As String, strCad As String
    
On Error GoTo ComprobarCierreTotalPosible_Error

    strMensaje = ""
    
'    If Not mvarobjProcNC.comprobar_acciones_correctoras_cerradas(strCad) Then
'        strMensaje = vbCrLf & " · Algunas Acciones Correctivas no se encuentren cerradas: " & strCad
'        ComprobarCierreTotalPosible = False
'        Exit Function
'    End If
        
'    If Not mvarobjProcNC.comprobar_evaluacion_realizada() Then
'        strMensaje = strMensaje & vbCrLf & " · No es posible cerrar la no conformidad mientras no haya sido completada la EVALUACIÓN de la misma."
'        ComprobarCierreTotalPosible = False
'        Exit Function
'    End If
        
    If CStr(mvarobjProcNC.getCAUSA_DIRECTA) = "" Then
        strMensaje = strMensaje & vbCrLf & " · No es posible cerrar ya que la CAUSA DIRECTA no esta informada."
        ComprobarCierreTotalPosible = False
'        Exit Function
    End If
    
    If CStr(mvarobjProcNC.getCAUSA_RAIZ) = "" Then
        strMensaje = strMensaje & vbCrLf & " · No es posible cerrar ya que la CAUSA RAIZ no esta informada."
        ComprobarCierreTotalPosible = False
'        Exit Function
    End If
    
    If mvarobjProcNC.getES_SOLUCION_ACEPTABLE < 0 Then
        strMensaje = strMensaje & vbCrLf & " · No es posible cerrar ya que la EVALUACION FINAL no esta informada."
        ComprobarCierreTotalPosible = False
'        Exit Function
    End If
    
    ' SI SE MARCA COMO CERRADO, VALIDAR QUE TENGA ARCHIVOS ADJUNTOS (914)
    If lstDocumentacion.ListItems.Count = 0 Then
        strMensaje = strMensaje & vbCrLf & " · Para cerrar el PNC es necesario adjuntar evidencias."
    End If
    
    
    If Trim(strMensaje) <> "" Then
        MsgBox "No es posible cerrar la No Confomidad por la/s siguiente/s razón/es: " & strMensaje, vbInformation, "Cerrar Procedimiento No Conformidad"
        ComprobarCierreTotalPosible = False
        Exit Function
    End If


    ComprobarCierreTotalPosible = True

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.ComprobarCierreTotalPosible"
    Exit Function
ComprobarCierreTotalPosible_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNCEdicion.ComprobarCierreTotalPosible"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ComprobarCierreTotalPosible of Formulario frmProcNCEdicion" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Private Sub opEstado_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To 8
        lblestado(i).BackColor = &HC0C0C0
        lblestado(i).ForeColor = vbBlack
        lblestado(i).BorderStyle = 0
        lblestado(i).FontBold = False
        lblestado(i).FontSize = 8
    Next
    lblestado(Index).BackColor = &H80FFFF
    lblestado(Index).ForeColor = &HC0&
    lblestado(Index).BorderStyle = 1
    lblestado(Index).FontSize = 10
    lblestado(Index).FontBold = True
    mvarenuEstado = Index
    establecer_opciones_estado
End Sub
