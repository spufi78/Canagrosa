VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmDeterminaciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro determinaciones"
   ClientHeight    =   10830
   ClientLeft      =   2265
   ClientTop       =   930
   ClientWidth     =   14700
   Icon            =   "frmDeterminaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmDeterminaciones.frx":1272
   ScaleHeight     =   10830
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSituaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   630
      TabIndex        =   56
      Top             =   3060
      Visible         =   0   'False
      Width           =   13410
      Begin VB.CommandButton cmdSituacionCerrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   12465
         Picture         =   "frmDeterminaciones.frx":15B4
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3195
         Width           =   870
      End
      Begin VB.CommandButton cmdSituacionModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   615
         Left            =   11565
         Picture         =   "frmDeterminaciones.frx":7E06
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3195
         Width           =   870
      End
      Begin VB.Frame frmSituacion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   225
         TabIndex        =   57
         Top             =   3330
         Width           =   8520
         Begin XtremeSuiteControls.RadioButton opSituacion 
            Height          =   240
            Index           =   0
            Left            =   450
            TabIndex        =   58
            Top             =   135
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "En rango"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opSituacion 
            Height          =   240
            Index           =   2
            Left            =   6255
            TabIndex        =   59
            Top             =   135
            Width           =   1635
            _Version        =   851970
            _ExtentX        =   2884
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Fuera de Rango"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opSituacion 
            Height          =   240
            Index           =   1
            Left            =   3015
            TabIndex        =   61
            Top             =   135
            Width           =   2175
            _Version        =   851970
            _ExtentX        =   3836
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Rango con Incertidumbre"
            BackColor       =   16777215
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   5760
            Picture         =   "frmDeterminaciones.frx":E658
            Stretch         =   -1  'True
            Top             =   0
            Width           =   435
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   2475
            Picture         =   "frmDeterminaciones.frx":EBD6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   435
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   -45
            Picture         =   "frmDeterminaciones.frx":F0F7
            Stretch         =   -1  'True
            Top             =   0
            Width           =   435
         End
      End
      Begin MSComctlLib.ListView listaRangos 
         Height          =   2715
         Left            =   90
         TabIndex        =   64
         Top             =   405
         Width           =   13260
         _ExtentX        =   23389
         _ExtentY        =   4789
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
      Begin VB.Label lblDeterminacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Determinaciones"
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
         Height          =   240
         Left            =   45
         TabIndex        =   62
         Top             =   135
         Width           =   13320
      End
   End
   Begin VB.CommandButton cmdCopiarResultados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copiar Equipos y Reactivos"
      Height          =   840
      Left            =   7470
      Picture         =   "frmDeterminaciones.frx":F632
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   9945
      Width           =   1005
   End
   Begin VB.CommandButton cmdImagen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagenes"
      Height          =   840
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   9945
      Width           =   1005
   End
   Begin VB.CommandButton cmdTA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Análisis"
      Height          =   840
      Left            =   1125
      Picture         =   "frmDeterminaciones.frx":15E84
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Ver el Tipo de Determinacion"
      Top             =   9945
      Width           =   1005
   End
   Begin VB.CommandButton cmdBano 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Baño"
      Height          =   840
      Left            =   1125
      Picture         =   "frmDeterminaciones.frx":1674E
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Ver el Tipo de Determinacion"
      Top             =   9945
      Width           =   1005
   End
   Begin MSComctlLib.ListView auxdatos 
      Height          =   4635
      Left            =   675
      TabIndex        =   10
      Top             =   2655
      Visible         =   0   'False
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso Modificar Cerrada"
      Height          =   195
      Left            =   7875
      TabIndex        =   47
      Top             =   10440
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CheckBox chkCalcular 
      Caption         =   "Check1"
      Height          =   225
      Left            =   7875
      TabIndex        =   45
      Top             =   10620
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdduplicados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicados"
      Height          =   840
      Left            =   4365
      Picture         =   "frmDeterminaciones.frx":17018
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Mostrar el Histórido de Duplicados"
      Top             =   9945
      Width           =   1005
   End
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      Height          =   840
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   9945
      Width           =   1095
   End
   Begin VB.TextBox txtcampomatraz 
      Height          =   330
      Left            =   7605
      TabIndex        =   38
      Text            =   "0"
      Top             =   10215
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtformulasolidos 
      Height          =   330
      Left            =   7515
      TabIndex        =   37
      Text            =   "0"
      Top             =   9810
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdFormula 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fórmula"
      Height          =   840
      Left            =   3330
      Picture         =   "frmDeterminaciones.frx":17322
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Ver la Formula"
      Top             =   9945
      Width           =   1005
   End
   Begin VB.CommandButton cmdTD 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Determinación"
      Height          =   840
      Left            =   2160
      Picture         =   "frmDeterminaciones.frx":17BEC
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Ver el Tipo de Determinacion"
      Top             =   9945
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   7065
      TabIndex        =   31
      Top             =   9000
      Width           =   7575
      Begin VB.TextBox txtgrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6210
         TabIndex        =   33
         Top             =   405
         Width           =   1170
      End
      Begin VB.TextBox txtmetodo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   900
         TabIndex        =   32
         Top             =   405
         Width           =   4320
      End
      Begin VB.Label lblgrado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grado"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5535
         TabIndex        =   46
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Metodo"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   34
         Top             =   450
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipos"
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
      Height          =   2490
      Left            =   7065
      TabIndex        =   26
      Top             =   4230
      Width           =   7575
      Begin VB.CommandButton cmdEliminarEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   900
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "Elimina el campo seleccionado"
         Top             =   270
         Width           =   915
      End
      Begin VB.CommandButton cmdAnadirEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   900
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   27
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   1485
         Width           =   915
      End
      Begin MSComctlLib.ListView listaEquipos 
         Height          =   1785
         Left            =   135
         TabIndex        =   29
         Top             =   270
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   3149
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
      Begin pryCombo.miCombo cmbEquipos 
         Height          =   330
         Left            =   135
         TabIndex        =   30
         Top             =   2070
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   582
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reactivos"
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
      Height          =   2220
      Left            =   7065
      TabIndex        =   21
      Top             =   6750
      Width           =   7575
      Begin VB.CommandButton cmdEliminarReactivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   795
         Left            =   6570
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "Elimina el campo seleccionado"
         Top             =   270
         Width           =   915
      End
      Begin VB.CommandButton cmdAnadirReactivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   810
         Left            =   6555
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   1215
         Width           =   915
      End
      Begin MSComctlLib.ListView listaReactivos 
         Height          =   1155
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   2037
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
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   330
         Left            =   750
         TabIndex        =   25
         Top             =   1500
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbReactivosInternos 
         Height          =   330
         Left            =   750
         TabIndex        =   42
         Top             =   1830
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externos"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   44
         Top             =   1530
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interno"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   43
         Top             =   1875
         Width           =   495
      End
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   10125
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "P.N.T."
      Height          =   840
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Ver el PNT asociado"
      Top             =   9945
      Width           =   1005
   End
   Begin VB.TextBox txtanalisis 
      Height          =   375
      Left            =   9450
      TabIndex        =   19
      Top             =   10260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbano 
      Height          =   375
      Left            =   9270
      TabIndex        =   18
      Top             =   9900
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9540
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdCurvas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Histórico"
      Height          =   840
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ver el Histórico de Resultados"
      Top             =   9945
      Width           =   1005
   End
   Begin VB.CommandButton cmdCambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anterior"
      Height          =   840
      Index           =   1
      Left            =   10665
      Picture         =   "frmDeterminaciones.frx":184B6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ir a la muestra anterior"
      Top             =   9945
      Width           =   960
   End
   Begin VB.CommandButton cmdCambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Siguiente"
      Height          =   840
      Index           =   0
      Left            =   11670
      Picture         =   "frmDeterminaciones.frx":18D80
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Ir a la muestra siguiente"
      Top             =   9945
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   840
      Left            =   12675
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9945
      Width           =   960
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   840
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9945
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   45
      TabIndex        =   6
      Top             =   9000
      Width           =   6930
      Begin VB.Frame frameRevision 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   2355
         Begin XtremeSuiteControls.RadioButton spinDuplicados 
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   49
            Top             =   135
            Width           =   1050
            _Version        =   851970
            _ExtentX        =   1852
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Conforme"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton spinDuplicados 
            Height          =   240
            Index           =   2
            Left            =   1080
            TabIndex        =   50
            Top             =   135
            Width           =   1185
            _Version        =   851970
            _ExtentX        =   2090
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "No conforme"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.CommandButton cmdcalcular 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   660
         Left            =   5985
         Picture         =   "frmDeterminaciones.frx":1964A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   870
      End
      Begin VB.TextBox txtvalor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3600
         TabIndex        =   0
         Top             =   405
         Width           =   2295
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   405
         Width           =   3510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   9
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Campo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView deter 
      Height          =   3465
      Left            =   45
      TabIndex        =   1
      Top             =   690
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   6112
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
   Begin MSComctlLib.ListView datos 
      Height          =   4470
      Left            =   45
      TabIndex        =   3
      Top             =   4500
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   7885
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   225
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones.frx":19954
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones.frx":19E9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones.frx":1A3D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEdiciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ediciones"
      Height          =   840
      Left            =   8505
      Picture         =   "frmDeterminaciones.frx":1A95E
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   9945
      Width           =   1005
   End
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   11880
      TabIndex        =   39
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   11880
      TabIndex        =   11
      Top             =   360
      Width           =   2805
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Determinaciones"
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
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14715
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Campos Fórmula"
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
      Height          =   300
      Index           =   0
      Left            =   45
      TabIndex        =   4
      Top             =   4230
      Width           =   6945
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Determinaciones"
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
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   14580
   End
End
Attribute VB_Name = "frmDeterminaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Dim primera_vez As Boolean
Private Enum COLS
    nombre = 1
    ID_DETERMINACION = 4
    FORMULA_ID = 5
    TIPO_DETERMINACION_ID = 6
    DIF_DUPLICADOS_NUMERICA = 11
    DIF_DUPLICADOS = 12
    INCERTIDUMBRE = 13
    GRADO = 14
    REV_DUPLICADO = 15
    SITUACION = 16
    DIF_AVISO = 17
    dif_historico = 18
    C_VARIACION = 19
End Enum

Private Enum COLS_RANGOS
    CR_ORDEN = 0
    CR_RANGO = 1
    CR_INCERTIDUMBRE = 2
    CR_DIF_DUPLICADOS_NUMERICA = 3
    CR_DIF_DUPLICADOS = 4
    CR_DIF_AVISO = 5
    CR_DIF_HISTORICO = 6
    CR_C_VARIACION = 7
End Enum

Private Sub cmdCopiarResultados_Click()
    If deter.ListItems.Count = 0 Then Exit Sub
    frmDeterminaciones_CopiaResultados.PK = deter.ListItems(deter.selectedItem.Index).SubItems(4)
    frmDeterminaciones_CopiaResultados.Show 1
End Sub

Private Sub cmdEdiciones_Click()
    With frmMuestras_Ediciones
        .PK = gmuestra
        .Show 1
    End With
End Sub

'Private WithEvents TecladoNumerico As frmTecladoNumerico
'Private blnEsTablet As Boolean
'Private blnTecladoNumericoPrimeraVez As Boolean
'Private blnTecladoNumerico_NoMostrar As Boolean
Private Sub cmdImagen_Click()
    With frmCE_Imagenes
        .PK = gmuestra
        .Show 1
    End With
End Sub
'Private Sub ConfigurarTablet()
'    blnEsTablet = pc_es_tablet
'    If blnEsTablet Then
'        Set TecladoNumerico = New frmTecladoNumerico
'        TecladoNumerico.OcultarConformidad = True
'        TecladoNumerico.posX = Screen.Width - TecladoNumerico.Width
'        TecladoNumerico.posY = 0
'        blnTecladoNumericoPrimeraVez = True
'        blnTecladoNumerico_NoMostrar = False
'    End If
'End Sub

Private Sub cmdbano_Click()
    frmBANO_Detalle.PK = txtbano
    frmBANO_Detalle.Show 1
    actualizarTipoDeterminacion
End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim oEquipo As New clsEquipos
        oEquipo.Carga_Datos_Basicos cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
        almacenar_equipos
        cmbEquipos.limpiar
    End If
End Sub

Private Sub cmdAnadirReactivo_Click()
    ' Externo (E)
    If cmbReactivos.getTEXTO <> "" Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
        Dim oTR As New clsTipos_reactivo_ex
        oBote.CARGAR cmbReactivos.getPK_SALIDA
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "E"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Interno (I)
    If cmbReactivosInternos.getTEXTO <> "" Then
        Dim oRPR As New clsRpr_botes
        Dim oTRPR As New clsRPR_Tipos
        oRPR.Carga cmbReactivosInternos.getPK_SALIDA
        oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
        With listaReactivos.ListItems.Add(, , oRPR.getID_BOTE_PR)
            .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    cmbReactivos.limpiar
    cmbReactivosInternos.limpiar
    almacenar_reactivos
End Sub

'Dim ce As Boolean
Private Sub cmdCalcular_Click()
    On Error GoTo fallo
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    Dim requeridos As Boolean
    requeridos = True
    ' Validamos los campos requeridos para el calculo
    For i = datos.selectedItem.Index To 1 Step -1
         If datos.ListItems(i).bold = False Then
             If Trim(datos.ListItems(i).SubItems(1)) = "" Then
                 requeridos = False
             End If
         End If
    Next
    ' Comprobamos que esten todos los campos requeridos
    If requeridos = False Then
'        If Not blnEsTablet Then
            MsgBox "Faltan campos requeridos por informar.", vbExclamation, App.Title
'        End If
        Exit Sub
    End If
    ' Hacemos el calculo si estan todos los requeridos
    Dim predijo As String
    Dim cadena As String
    Dim campo As String
    Dim Formula As String
    Dim pos As Integer
    Dim ofor As New clsFormulas
    Dim encontrado As Boolean
    prefijo = ""
    ofor.CARGAR (deter.ListItems(deter.selectedItem.Index).SubItems(5))
    cadena = ofor.getEXPRESION
    If Not IsNull(cadena) Then
        For i = 1 To Len(cadena)
            If Mid(cadena, i, 1) <> "C" Then
              If Mid(cadena, i, 1) = "," Then
                Formula = Formula & "."
              Else
                Formula = Formula & Mid(cadena, i, 1)
              End If
            Else
                pos = InStr(i + 2, cadena, "_")
                campo = Mid(cadena, i + 2, (pos) - (i + 2))
                j = datos.selectedItem.Index
                encontrado = False
                Do
                 If CInt(datos.ListItems(j).SubItems(3)) = CInt(campo) Then
                     Formula = Formula & Replace(datos.ListItems(j).SubItems(1), ",", ".")
                     encontrado = True
                 End If
                 j = j - 1
                Loop Until j = 0 Or encontrado = True
                i = pos
            End If
        Next
    End If
    Dim ocampos As New clsFormulas_campos
    ocampos.CARGAR (datos.ListItems(datos.selectedItem.Index).SubItems(3))
    
    txtvalor = formatear(sc.Eval(Formula), ocampos.getENTEROS, ocampos.getDECIMALES)
    txtvalor_KeyPress (13)
    
'    datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(sc.Eval(Formula), ocampos.getENTEROS, ocampos.getDECIMALES)
'    If UCase(lblestado.Caption) = "DUPLICADA" Then
'        visualizar_duplicados
'    End If
'    grabar_auxdatos
'    ' Pasar al siguiente campo
'    If datos.ListItems.Count > datos.selectedItem.Index Then
'         Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
'         datos_Click
'    Else
'         If deter.ListItems.Count > deter.selectedItem.Index Then
'             Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index + 1)
'             Dim oDeter As New clsDeterminaciones
'             oDeter.CargarDeterminacion (deter.ListItems(deter.selectedItem.Index).SubItems(4))
'             If oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma") And oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico") Then
'                deter_Click
'                datos_Click
'             End If
'         Else
'             txtdato = ""
'             txtValor = ""
'             If Not blnEsTablet Then datos.SetFocus
'         End If
'    End If
    Set ocampos = Nothing
    Exit Sub
fallo:
    If Formula <> "" Then
        MsgBox "Error en la formula. Formula : " & Formula, vbCritical, Err.Description
    Else
        MsgBox "Error al calcular la formula." & Err.Description, vbCritical, "Error"
    End If
End Sub

Private Sub cmdCambio_Click(Index As Integer)
    Dim omue As New clsMuestra
    If auxdatos.ListItems.Count > 0 Then
        If MsgBox("Al cambiar de muestra, perdera los datos si no graba. ¿Esta seguro de cambiar de muestra?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    If omue.CargaMuestra(gmuestra) = True Then
        datos.ListItems.Clear
        auxdatos.ListItems.Clear
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        If Index = 0 Then
            consulta = "select id_muestra from muestras where tipo_muestra_id = " & omue.getTIPO_MUESTRA_ID & " and id_muestra > " & gmuestra & " order by id_muestra asc"
        Else
            consulta = "select id_muestra from muestras where tipo_muestra_id = " & omue.getTIPO_MUESTRA_ID & " and id_muestra < " & gmuestra & " order by id_muestra desc"
        End If
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            gmuestra = rs.Fields(0)
            inicializa_ventana
         Else
            If Index = 0 Then
                MsgBox "No existen muestras con código superior.", vbInformation, App.Title
            Else
                MsgBox "No existen muestras con código inferior.", vbInformation, App.Title
            End If
        End If
        deter_Click
    End If
    Set omue = Nothing
End Sub

Private Sub cmdcancel_Click()
    
    If auxdatos.ListItems.Count > 0 Then
        If MsgBox("Va a salir sin guardar los datos de la muestra. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCurvas_Click()
    gdeterminacion = deter.ListItems(deter.selectedItem.Index).SubItems(4)
    If deter.ListItems(deter.selectedItem.Index).SubItems(8) = "" Then
        frmHistoricoDeterminacion.LIMITE_INFERIOR = 0
    Else
        frmHistoricoDeterminacion.LIMITE_INFERIOR = deter.ListItems(deter.selectedItem.Index).SubItems(8)
    End If
    If deter.ListItems(deter.selectedItem.Index).SubItems(9) = "" Then
        frmHistoricoDeterminacion.LIMITE_SUPERIOR = 0
    Else
        frmHistoricoDeterminacion.LIMITE_SUPERIOR = deter.ListItems(deter.selectedItem.Index).SubItems(9)
    End If
    frmHistoricoDeterminacion.Show 1
End Sub

Private Sub cmdduplicados_Click()
    frmDuplicados_Detalle.PK_ID_TIPO_DETERMINACION = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.TIPO_DETERMINACION_ID)
    frmDuplicados_Detalle.PK_DIF = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS)
    frmDuplicados_Detalle.Show 1
End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
        almacenar_equipos
    End If
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
        almacenar_reactivos
    End If
End Sub

Private Sub cmdFormula_Click()
    If deter.ListItems.Count > 0 Then
        frmFORMULA_Detalle.PK = deter.ListItems(deter.selectedItem.Index).SubItems(5)
        frmFORMULA_Detalle.Show 1
    End If
End Sub

Private Sub cmdObservador_Click()
    If deter.ListItems.Count = 0 Then
        MsgBox "No hay determinaciones para Observar/Formar.", vbExclamation, App.Title
        Exit Sub
    End If
Dim objfrm As New frmObservadorEnsayo

    'MANTIS-807-I
    objfrm.FORMULARIO_ORIGEN = 3
    'MANTIS-807-F
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = gmuestra ' Id de la muestra
    objfrm.TIPO_DETERMINACION_ENSAYO_ID = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(6)) ' tipo de la Determinacion
    objfrm.DETERMINACION_ENSAYO_ID = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
    objfrm.MUESTRA_CERRADA = (Not cmdok.Enabled)
    If CInt(txtbano.Text) <> 0 Then
        objfrm.TIPO_OBSERVACION_ID = MC_TIPOS_OBSERVACION.MCTO_BANO
    Else
        objfrm.TIPO_OBSERVACION_ID = MC_TIPOS_OBSERVACION.MCTO_DETERMINACION
    End If

    objfrm.Show vbModal
    
    Set objfrm = Nothing
    
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    ' Validar reactivos caducados (1090)
    Dim cont As Integer
    Dim existen As Boolean
    existen = False
    For cont = 1 To listaReactivos.ListItems.Count
        If Trim(listaReactivos.ListItems(cont).SubItems(2)) <> "" Then
            If Format(listaReactivos.ListItems(cont).SubItems(2), "yyyy-mm-dd") < Format(Date, "yyyy-mm-dd") Then
                existen = True
            End If
        End If
    Next
    If existen Then
        If MsgBox("Existen reactivos CADUCADOS. ¿ESTA SEGURO DE ALMACENAR LOS DATOS DE LA MUESTRA?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    
    If MsgBox("Se van a insertar los datos de las determinaciones.¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        Dim oDeter As New clsDeterminaciones
        ' Almacenamos los nuevos metodos
        Dim i As Integer
        If deter.ListItems.Count > 0 Then
            For i = 1 To deter.ListItems.Count
                oDeter.modificar_metodo deter.ListItems(i).SubItems(4), deter.ListItems(i).SubItems(10), deter.ListItems(i).SubItems(COLS.GRADO)
            Next
        End If
        Dim odd As New clsDatos_determinaciones
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            auxdatos.Sorted = True
            auxdatos.SortKey = 3
        End If
        ' Almacenar Datos Determinaciones
        For i = 1 To auxdatos.ListItems.Count
            If auxdatos.ListItems(i).SubItems(3) <> "" Then ' Para la media y diferencia de duplicados
                If odd.CARGAR(CLng(auxdatos.ListItems(i)), auxdatos.ListItems(i).SubItems(3)) = True Then
                    'JGM
                    If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                        odd.setVALOR_1 = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    Else
                        odd.setVALOR_1 = " "
                    End If
                    ' Valor duplicado
                    If UCase(lblestado.Caption) = "DUPLICADA" Then
                        i = i + 1
                        If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                           odd.setVALOR_2 = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                        Else
                           odd.setVALOR_2 = " "
                        End If
                    End If
                    odd.Insertar_Valores
                End If
            End If
        Next
        ' Verificamos si se trata de un trigo
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra (gmuestra)
'        Dim esTrigo As Boolean
'        esTrigo = False
'        Dim oTipoAnalisis As New clsTipos_analisis
'        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TRIGO Or _
'           oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.HARINA_BLANDA Or _
'           oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.HARINA_DURO Then
'            esTrigo = True
'            oTipoAnalisis.CARGAR oMuestra.getTIPO_ANALISIS_ID
'        End If
        ' Almacena determinacion (Solucion)
        For i = 1 To auxdatos.ListItems.Count
         If UCase(lblestado.Caption) = "DUPLICADA" Then
            If auxdatos.ListItems(i).SubItems(4) = "M" Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                    oDeter.setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    ' Diferencia de duplicados
                    oDeter.setDIF_DUPLICADOS = Replace(auxdatos.ListItems(i + 2).SubItems(1), ",", ".")
                    oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
                    oDeter.setHORA = Format(Time, "hh:mm")
                    oDeter.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
'                    oDeter.setGRADO = ""
'                    If esTrigo Then
'                        oDeter.setGRADO = calcularClasificacionTrigo(oTipoAnalisis.getTIPO_TRIGO, CLng(auxdatos.ListItems(i)), auxdatos.ListItems(i).SubItems(1))
'                    End If
                    'M1371-I
                    Select Case auxdatos.ListItems(i + 3).SubItems(1)
                        Case "CONFORME"
                            oDeter.setREVISION_DUPLICADO = 1
                        Case "NO CONFORME"
                            oDeter.setREVISION_DUPLICADO = 2
                        Case Else
                            oDeter.setREVISION_DUPLICADO = 0
                    End Select
                    'M1371-F
                    oDeter.InsertarSolucion (CLng(auxdatos.ListItems(i)))
                End If
            End If
         Else
            If auxdatos.ListItems(i).bold = True Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                    oDeter.setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    oDeter.setDIF_DUPLICADOS = ""
                    oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
                    oDeter.setHORA = Format(Time, "hh:mm")
                    oDeter.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
 '                   oDeter.setGRADO = ""
 '                   If esTrigo Then
 '                       oDeter.setGRADO = calcularClasificacionTrigo(oTipoAnalisis.getTIPO_TRIGO, CLng(auxdatos.ListItems(i)), auxdatos.ListItems(i).SubItems(1))
 '                   End If
                    oDeter.InsertarSolucion (CLng(auxdatos.ListItems(i)))
                End If
            End If
         End If
        Next
        ' Si es Trigo, calcular los grados
        oDeter.informarGrados gmuestra
        ' Si la muestra ya estaba cerrada, actualizar la fecha de cierre
        If oMuestra.getCERRADA = 1 Then
            oMuestra.actualizar_fecha_cierre (gmuestra)
        Else
            oMuestra.comprobar_cierre (gmuestra)
        End If
'        If Trim(omuestra.getFECHA_COMIENZO) = "" Then
'            omuestra.actualizar_fecha_comienzo (gmuestra)
'        End If
        
        Set oMuestra = Nothing
        Me.MousePointer = 0
        MsgBox "Determinaciones salvadas correctamente.", vbInformation + vbOKOnly, App.Title
        Set odd = Nothing
        Set oDeter = Nothing
        auxdatos.ListItems.Clear
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox Err.Description, vbCritical, Err.Number
End Sub
Private Sub cmdPNT_Click()
    If deter.ListItems.Count > 0 Then
        Dim oTD As New clsTipos_determinacion
        oTD.CargarTipoDeterminacion deter.ListItems(deter.selectedItem.Index).SubItems(6)
        If oTD.getPNT_VINCULADO <> 0 Then
            Dim oPNT As New clsCa_documentos
            oPNT.mostrar oTD.getPNT_VINCULADO, True
            Set oPNT = Nothing
        Else
            MsgBox "El Tipo de Determinación no tiene PNT Vínculado.", vbExclamation, App.Title
        End If
    End If
End Sub
Private Sub actualizarTipoDeterminacion()
   On Error GoTo actualizarTipoDeterminacion_Error

    If deter.ListItems.Count = 0 Then Exit Sub
    Dim oTD As New clsTipos_determinacion
    If oTD.CargarTipoDeterminacion(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.TIPO_DETERMINACION_ID)) Then
        deter.ListItems(deter.selectedItem.Index).SubItems(1) = oTD.getNOMBRE
        deter.ListItems(deter.selectedItem.Index).SubItems(2) = oTD.getDESCRIPCION
        deter.ListItems(deter.selectedItem.Index).SubItems(5) = oTD.getFORMULA_ID
        deter.ListItems(deter.selectedItem.Index).SubItems(7) = oTD.getPARTICULAS
        Dim oDA As New clsDeterminaciones_analisis
        Dim validar As Boolean
        If txtbano = 0 Then
           validar = oDA.Carga_por_tipo_analisis(CLng(txtanalis), deter.ListItems(deter.selectedItem.Index).SubItems(COLS.TIPO_DETERMINACION_ID))
        Else
           validar = oDA.Carga_por_BANO(CLng(txtbano), deter.ListItems(deter.selectedItem.Index).SubItems(COLS.TIPO_DETERMINACION_ID))
        End If
        If validar Then
            deter.ListItems(deter.selectedItem.Index).SubItems(8) = oDA.getMINIMO
            deter.ListItems(deter.selectedItem.Index).SubItems(9) = oDA.getMAXIMO
        End If
        Set oDA = Nothing
    End If

   On Error GoTo 0
   Exit Sub

actualizarTipoDeterminacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizarTipoDeterminacion of Formulario frmDeterminaciones"
End Sub

Private Sub cmdSituacionCerrar_Click()
    frmSituaciones.visible = False
End Sub

Private Sub cmdSituacionModificar_Click()
    Dim oDeter As New clsDeterminaciones
    Dim SITUACION As Integer
   On Error GoTo cmdSituacionModificar_Click_Error
    
    If frmSituacion.visible = True Then
        SITUACION = 0
        If opSituacion(1).Value = True Then
            SITUACION = 1
        End If
        If opSituacion(2).Value = True Then
            SITUACION = 2
        End If
        oDeter.modificar_situacion deter.ListItems(deter.selectedItem.Index).SubItems(COLS.ID_DETERMINACION), SITUACION
        cargar_determinaciones
        'M3299 Recalcular situación de la muestra
        Dim oMuestra As New clsMuestra
        oMuestra.informar_situacion (gmuestra)
    End If
    ' Cambiar rango
    'MYYYY
    If listaRangos.ListItems.Count > 0 Then
        Dim oTDC As New clsTipos_determinacion_conf
        If oTDC.CARGAR(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.TIPO_DETERMINACION_ID), listaRangos.ListItems(listaRangos.selectedItem.Index).Text) Then
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS_NUMERICA) = oTDC.getDIF_DUPLICADOS_NUMERICA ' Dif. dupliciados numerica
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS) = oTDC.getDIF_DUPLICADOS ' % Dif. Duplicado
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE) = oTDC.getINCERTIDUMBRE
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_AVISO) = oTDC.getDIF_AVISO
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.dif_historico) = oTDC.getDIF_HISTORICO
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.C_VARIACION) = oTDC.getC_VARIACION
        Else
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS_NUMERICA) = ""
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS) = ""
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE) = ""
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_AVISO) = ""
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.dif_historico) = ""
            deter.ListItems(deter.selectedItem.Index).SubItems(COLS.C_VARIACION) = ""
        End If
        Dim oDC As New clsDeterminaciones_conf
        oDC.Insertar deter.ListItems(deter.selectedItem.Index).SubItems(COLS.ID_DETERMINACION), listaRangos.ListItems(listaRangos.selectedItem.Index).Text
    End If
    
    frmSituaciones.visible = False
   On Error GoTo 0
   Exit Sub

cmdSituacionModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSituacionModificar_Click of Formulario frmDeterminaciones"
End Sub

Private Sub cmdTA_Click()
    frmTA_Detalle.PK = txtanalisis
    frmTA_Detalle.Show 1
    actualizarTipoDeterminacion
End Sub

Private Sub cmdTD_Click()
    If deter.ListItems.Count > 0 Then
        frmTD_Detalle.PK = deter.ListItems(deter.selectedItem.Index).SubItems(6)
        frmTD_Detalle.Show 1
        actualizarTipoDeterminacion
    End If

End Sub

Private Sub datos_Click()
    On Error Resume Next
    If datos.ListItems.Count > 0 Then
        datos.selectedItem.EnsureVisible
        cmdcalcular.Enabled = False
        chkCalcular.Value = Unchecked
        If datos.ListItems(datos.selectedItem.Index).bold = True Then
         If Trim(lblestado.Caption) = "" And datos.ListItems.Count > 1 Then
            cmdcalcular.Enabled = True
            chkCalcular.Value = Checked
         Else
            If Trim(lblestado.Caption) = "DUPLICADA" And datos.ListItems.Count > 4 Then
                cmdcalcular.Enabled = True
                chkCalcular.Value = Checked
            End If
         End If
        End If
        txtdato = datos.ListItems(datos.selectedItem.Index)
        'M1371-I
        ' Si es duplicada y es el último campo, mostramos el frame de revision
        If Trim(lblestado.Caption) = "DUPLICADA" And datos.selectedItem.Index = datos.ListItems.Count Then
            frameRevision.visible = True
            spinDuplicados(1).Value = False
            spinDuplicados(2).Value = False
            txtvalor = ""
'            If datos.ListItems(datos.selectedItem.Index).SubItems(1) = "CONFORME" Then
'                spinDuplicados(1).value = True
'            ElseIf datos.ListItems(datos.selectedItem.Index).SubItems(1) = "NO CONFORME" Then
'                spinDuplicados(2).value = True
'            End If
        Else
            frameRevision.visible = False
            txtvalor = datos.ListItems(datos.selectedItem.Index).SubItems(1)
            txtvalor.SetFocus
            txtvalor.SelStart = 0
            txtvalor.SelLength = Len(txtvalor)
        End If
        'M1371-F
    End If
End Sub

Private Sub deter_Click()
    If deter.ListItems.Count < 1 Then
        Exit Sub
    End If
    Dim oDeter As New clsDeterminaciones
    deter.selectedItem.EnsureVisible
    ' Por particulas
    If deter.ListItems(deter.selectedItem.Index).SubItems(7) = 1 Then
'        If blnEsTablet Then
'            If TecladoNumerico.Visible Then TecladoNumerico.Visible = False
'        End If
        frmFluidos_Resultados.MUESTRA = gmuestra
        frmFluidos_Resultados.DETERMINACION_ID = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
        frmFluidos_Resultados.txtparametro(0) = deter.ListItems(deter.selectedItem.Index).SubItems(1)
        frmFluidos_Resultados.Show 1
        oDeter.CargarDeterminacion (deter.ListItems(deter.selectedItem.Index).SubItems(4))
        deter.ListItems(deter.selectedItem.Index).SubItems(3) = Replace(oDeter.getRESULTADO, ".", ",")
        siguiente_campo
    Else
        gdeterminacion = deter.ListItems(deter.selectedItem.Index).SubItems(4)
        Select Case CLng(deter.ListItems(deter.selectedItem.Index).SubItems(6))
         Case CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))
'            If blnEsTablet Then
'                If TecladoNumerico.Visible Then TecladoNumerico.Visible = False
'            End If
            frmAlveograma.Show 1
            oDeter.CargarDeterminacion (deter.ListItems(deter.selectedItem.Index).SubItems(4))
            deter.ListItems(deter.selectedItem.Index).SubItems(3) = Replace(oDeter.getRESULTADO, ".", ",")
            siguiente_campo
         Case CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))
'            If blnEsTablet Then
'                If TecladoNumerico.Visible Then TecladoNumerico.Visible = False
'            End If
            frmOrganoleptico.Show 1
            oDeter.CargarDeterminacion (deter.ListItems(deter.selectedItem.Index).SubItems(4))
            deter.ListItems(deter.selectedItem.Index).SubItems(3) = Replace(oDeter.getRESULTADO, ".", ",")
            siguiente_campo
         Case Else
            Call cargar_campos
'            MostrarTecladoNumerico deter.selectedItem.SubItems(1), txtdato.Text, txtvalor.Text
        End Select
    End If
End Sub

Private Sub cabecera()
    ' Determinaciones
    With deter.ColumnHeaders
        .Add , , "Pnt", 1200, lvwColumnLeft
        .Add , , "Nombre", 3000, lvwColumnLeft
        .Add , , "Descripcion", 3200, lvwColumnLeft
        .Add , , "Solución", 1000, lvwColumnRight
        .Add , , "ID_DETERMINACION", 1, lvwColumnCenter
        .Add , , "ID_FORMULA", 1, lvwColumnCenter
        .Add , , "ID_TIPO_DETERMINACION", 1, lvwColumnCenter
        .Add , , "Particulas", 1, lvwColumnCenter
        .Add , , "Mínimo", 900, lvwColumnCenter
        .Add , , "Máximo", 900, lvwColumnCenter
        .Add , , "Metodo", 1, lvwColumnCenter 'OCULTA
        .Add , , "Dif.Dupl.", 1400, lvwColumnCenter
        .Add , , "% Dif.Dupl.", 800, lvwColumnCenter
        .Add , , "Incertidumbre", 800, lvwColumnCenter
        .Add , , "Grado", 800, lvwColumnCenter
        .Add , , "REV.DUPLICADO", 1, lvwColumnLeft ' OCULTA
        .Add , , "SITUACION", 1, lvwColumnLeft ' OCULTA
        .Add , , "DIF.AVISO", 1, lvwColumnLeft ' OCULTA
        .Add , , "%Dif.Histórico", 1300, lvwColumnCenter ' OCULTA
        .Add , , "Coef.Variacion", 1300, lvwColumnCenter ' OCULTA
    End With
    ' Datos
    With datos.ColumnHeaders
        .Add , , "Campo", 3550, lvwColumnLeft
        .Add , , "Valor", 2300, lvwColumnRight
        .Add , , "Unidad", 800, lvwColumnLeft
        .Add , , "ID_CAMPO", 0, lvwColumnCenter
        .Add , , "ENTEROS", 0, lvwColumnCenter
        .Add , , "DECIMALES", 0, lvwColumnCenter
    End With
    ' Auxiliar de calculos de Datos
    With auxdatos.ColumnHeaders
        .Add , , "Muestra", 1000, lvwColumnLeft
        .Add , , "Valor", 1000, lvwColumnLeft
        .Add , , "Linea", 1000, lvwColumnLeft
        .Add , , "Campo", 1000, lvwColumnLeft
        .Add , , "Media", 200, lvwColumnLeft
    End With
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 900, lvwColumnLeft
        .Add , , "Nombre", 3700, lvwColumnLeft
        .Add , , "NºSerie", 1300, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "ID", 800, lvwColumnLeft
        .Add , , "Reactivo", 3900, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
    
    With listaRangos.ColumnHeaders
        .Add , , "ORDEN", 1, lvwColumnLeft
        .Add , , "Rango", 3400, lvwColumnCenter
        .Add , , "Incertidumbre", 1500, lvwColumnCenter
        .Add , , "Dif.Duplicados", 2600, lvwColumnCenter
        .Add , , "%Dif.Duplicados", 1300, lvwColumnCenter
        .Add , , "%Aviso Rango", 1300, lvwColumnCenter
        .Add , , "%Dif.Histórico", 1300, lvwColumnCenter
        .Add , , "Coef.Variacion", 1300, lvwColumnCenter
    End With
End Sub

Private Sub cargar_determinaciones()
    Dim rs As New ADODB.Recordset
   On Error GoTo cargar_determinaciones_Error

    deter.ListItems.Clear
    ' Determinaciones de la muestra
    Dim oDeter As New clsDeterminaciones
    Dim oDA As New clsDeterminaciones_analisis
    Set rs = oDeter.lista_determinaciones(gmuestra)
    While Not rs.EOF
        With deter.ListItems.Add(, , rs("pnt")) ' Pnt
            .SubItems(1) = rs("nombre") ' nombre
            .SubItems(2) = rs("descripcion") ' des
            If Not rs("resultado") <> "" And Not IsNull(rs("resultado")) Then ' resultado
               .SubItems(3) = " "
            Else
               .SubItems(3) = Replace(rs("resultado"), ".", ",")
            End If
            .SubItems(4) = rs("id_determinacion") ' id_deter
            .SubItems(5) = rs("formula_id") ' formula_id
            .SubItems(6) = rs("id_tipo_determinacion") ' id_tipo_deter
            .SubItems(7) = rs("particulas") ' Por particulas
            If CInt(txtbano) <> 0 Then
                oDA.Carga_por_BANO CLng(txtbano), rs("id_tipo_determinacion")
            Else
                oDA.Carga_por_tipo_analisis CLng(txtanalisis), rs("id_tipo_determinacion")
            End If
            .SubItems(8) = Trim(oDA.getMINIMO)
            .SubItems(9) = Trim(oDA.getMAXIMO)
            ' Metodo
            If rs("metodo_deter") <> "" Then
                .SubItems(10) = rs("metodo_deter") ' Metodo tabla Determinaciones (Se inserta en recepción)
            Else
                .SubItems(10) = rs("metodo") ' Metodo tabla Tipos_Determinacion
            End If
            ' % Dif. Duplicados
            .SubItems(COLS.DIF_DUPLICADOS_NUMERICA) = texto(rs("dif_duplicados_numerica")) ' DIF_DUPLICADOS_NUMERICA
            .SubItems(COLS.DIF_DUPLICADOS) = texto(rs("dif_duplicados")) ' DIF_DUPLICADOS
            .SubItems(COLS.DIF_AVISO) = texto(rs("dif_aviso")) ' DIF_AVISO
            .SubItems(COLS.INCERTIDUMBRE) = texto(rs("incertidumbre")) ' Incertidumbre
            .SubItems(COLS.GRADO) = texto(rs("grado"))  ' GRADO
            .SubItems(COLS.REV_DUPLICADO) = texto(rs("revision_duplicado"))
            .SubItems(COLS.SITUACION) = texto(rs("situacion"))
            .SubItems(COLS.dif_historico) = texto(rs("dif_historico"))
            .SubItems(COLS.C_VARIACION) = texto(rs("c_variacion"))
        End With
        '&H00C9E2CC&
        If Trim(deter.ListItems(deter.ListItems.Count).SubItems(3)) <> "" Then
            deter.ListItems(deter.ListItems.Count).ListSubItems(3).ReportIcon = rs("situacion") + 1
        End If
        rs.MoveNext
    Wend
    Set oDeter = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargar_determinaciones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_determinaciones of Formulario frmDeterminaciones"
End Sub

Private Sub cargar_campos()
    Dim ocampos As New clsFormulas_campos
    Dim ouni As New clsUnidades
    Dim odd As New clsDatos_determinaciones
    Dim oddep As New clsTipos_determinacion_dep
    Dim rs As ADODB.Recordset
    Dim rs_dd As ADODB.Recordset
    Dim rs_ddep As ADODB.Recordset
    Dim duplicado As Integer
    Dim nombre As String
    Dim i As Integer
    Dim j As Integer
    Dim encontrado As Boolean
    datos.ListItems.Clear
    cmdcalcular.Enabled = False
    chkCalcular.Value = Unchecked

    Set rs = ocampos.Lista_Formulas_Unidades(deter.ListItems(deter.selectedItem.Index).SubItems(5))
    If UCase(lblestado.Caption) = "DUPLICADA" Then
        duplicado = 2
    Else
        duplicado = 1
    End If
    ' Cargamos los datos_deter (resultados)
    Set rs_dd = odd.cargar_determinacion(CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4)))
    ' Cargamos las determinaciones_dependientes
    Set rs_ddep = oddep.Listado_Dependencias(deter.ListItems(deter.selectedItem.Index).SubItems(6))
    If rs.RecordCount <> 0 Then
     For j = 1 To duplicado
        rs.MoveFirst
        While Not rs.EOF
                If duplicado = 2 Then
                    nombre = rs(1) & " (" & j & ")"
                Else
                    nombre = rs(1)
                End If
                With datos.ListItems.Add(, , nombre)
                    .SubItems(1) = " "
                    If rs_dd.RecordCount <> 0 Then
                        rs_dd.MoveFirst
                        encontrado = False
                        Do
                            If rs_dd("campo_id") = rs(0) Then
                                encontrado = True
                            Else
                                rs_dd.MoveNext
                            End If
                        Loop Until rs_dd.EOF Or encontrado = True
                        If encontrado Then
                            If j = 1 Then
                                If Left(rs_dd("VALOR_1"), 1) <> "I" Then
                                    .SubItems(1) = Replace(rs_dd("VALOR_1"), ".", ",")
                                End If
                            Else
                                If Left(rs_dd("VALOR_2"), 1) <> "I" Then
                                    .SubItems(1) = Replace(rs_dd("VALOR_2"), ".", ",")
                                End If
                            End If
                        End If
                    End If
                    .SubItems(2) = rs(3)
                    .SubItems(3) = rs(0)
                    .SubItems(4) = rs(4) ' ENTEROS
                    .SubItems(5) = rs(5) ' DECIMALES
                End With
            If rs(2) <> 0 Then ' Es solucion
                datos.ListItems.Item(datos.ListItems.Count).bold = True
            End If
            ' Verificar si hay dos determinaciones iguales (Dependientes)
            If rs_ddep.RecordCount <> 0 Then
                rs_ddep.MoveFirst
                encontrado = False
                Do
                    If rs_ddep("campo_id") = rs(0) Then
                        encontrado = True
                    Else
                        rs_ddep.MoveNext
                    End If
                Loop Until rs_ddep.EOF Or encontrado = True
                If encontrado Then
                    For i = 1 To deter.ListItems.Count
                        If rs_ddep("TIPO_DETERMINACION_ID_DEP") = deter.ListItems(i).SubItems(6) Then
                            datos.ListItems(datos.ListItems.Count).SubItems(1) = deter.ListItems(i).SubItems(3)
                        End If
                    Next
                End If
            End If
            ' Fin de verificacion
            rs.MoveNext
        Wend
     Next
    End If
    ' Resultados duplicados
    If duplicado = 2 Then
       With datos.ListItems.Add(, , "Resultado (MEDIA)")
          .SubItems(1) = " "
       End With
       With datos.ListItems.Add(, , "Resultado (Diferencia)")
          .SubItems(1) = " "
       End With
       With datos.ListItems.Add(, , "% Dif. entre duplicados")
          .SubItems(1) = " "
          .SubItems(2) = "%"
       End With
       'M1371-I
       With datos.ListItems.Add(, , "Revisión de Duplicados")
          .SubItems(1) = " "
          .bold = True
       End With
       'M1371-F
       visualizar_duplicados
    End If
    ' Comprobar si ya tiene datos
    For i = 1 To auxdatos.ListItems.Count
        If deter.ListItems(deter.selectedItem.Index).SubItems(4) = auxdatos.ListItems(i) Then
            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(2))).SubItems(1) = auxdatos.ListItems(i).SubItems(1)
        End If
    Next
    ' Equipos
    listaEquipos.ListItems.Clear
    Dim OTDEQUIPOS As New clsDeterminaciones_equipos
    Set rs = OTDEQUIPOS.Listado(CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4)))
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                'M1385-I
                If rs("EN_INFORME") = 1 Then
                    listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
                Else
                    listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = False
                End If
                'M1385-F
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Reactivos
    listaReactivos.ListItems.Clear
    Dim OTDR As New clsDeterminaciones_reactivos
    Dim oReactivo As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    Dim oRPR As New clsRpr_botes
    Dim oTRPR As New clsRPR_Tipos
    
    Set rs = OTDR.Listado(CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4)))
    If rs.RecordCount > 0 Then
        Do
            If rs(1) = "E" Then
               oReactivo.CARGAR CLng(rs(0))
               oTb.CARGAR oReactivo.getTIPO_BOTE_EX_ID
               oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
               With listaReactivos.ListItems.Add(, , rs(0))
                  .SubItems(1) = oTR.getNOMBRE
                  .SubItems(2) = Format(oReactivo.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                  .SubItems(3) = "E"
               End With
            Else
                oRPR.Carga CLng(rs(0))
                oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
                With listaReactivos.ListItems.Add(, , rs(0))
                    .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
                    .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                    .SubItems(3) = "I"
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Metodo
    txtmetodo = deter.ListItems(deter.selectedItem.Index).SubItems(10)
    txtgrado = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.GRADO)
    Set odd = Nothing
    Set rs = Nothing
    Set ocampos = Nothing
    Set ouni = Nothing
    datos_Click
End Sub

Private Sub deter_DblClick()
   On Error GoTo deter_DblClick_Error

    If deter.ListItems.Count = 0 Then Exit Sub
    If USUARIO.getPER_DATOS_ESPECIALES Then
        If lblCerrada = "CERRADA" Then
            Dim SITUACION As String
            SITUACION = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.SITUACION)
            If SITUACION <> "" Then
                If IsNumeric(SITUACION) Then
                    opSituacion(SITUACION).Value = True
                End If
            End If
            frmSituacion.visible = True
        Else
            frmSituacion.visible = False
        End If
    End If
    cargarListaRangos deter.ListItems(deter.selectedItem.Index).SubItems(COLS.TIPO_DETERMINACION_ID)
    lblDeterminacion = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.nombre)
    frmSituaciones.visible = True

   On Error GoTo 0
   Exit Sub

deter_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure deter_DblClick of Formulario frmDeterminaciones"
End Sub
Private Sub cargarListaRangos(TIPO_DETERMINACION_ID As Long)
    Dim oCONF As New clsTipos_determinacion_conf
    Dim rs As ADODB.Recordset
   On Error GoTo cargarListaRangos_Error
    
    listaRangos.ListItems.Clear
    Set rs = oCONF.Listado(TIPO_DETERMINACION_ID)
    If rs.RecordCount > 0 Then
        Do
            With listaRangos.ListItems.Add(, , rs("ORDEN"))
                  .SubItems(COLS_RANGOS.CR_RANGO) = texto(rs("RANGO"))
                  .SubItems(COLS_RANGOS.CR_DIF_DUPLICADOS) = numerico(texto(rs("DIF_DUPLICADOS")), 2)
                  .SubItems(COLS_RANGOS.CR_DIF_DUPLICADOS_NUMERICA) = texto(rs("DIF_DUPLICADOS_NUMERICA"))
                  .SubItems(COLS_RANGOS.CR_DIF_AVISO) = numerico(texto(rs("DIF_AVISO")), 2)
                  .SubItems(COLS_RANGOS.CR_INCERTIDUMBRE) = numerico(texto(rs("INCERTIDUMBRE")), 3)
                  .SubItems(COLS_RANGOS.CR_DIF_HISTORICO) = numerico(texto(rs("DIF_HISTORICO")), 2)
                  .SubItems(COLS_RANGOS.CR_C_VARIACION) = numerico(texto(rs("C_VARIACION")), 2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargarListaRangos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarListaRangos of Formulario frmDeterminaciones"
End Sub
'Private Sub MostrarTecladoNumerico(cab, subcab, res)
'
'   If Not blnEsTablet Then Exit Sub
'
'    If blnTecladoNumerico_NoMostrar Then
'        blnTecladoNumerico_NoMostrar = False
'    Else
'        TecladoNumerico.cabecera = cab
'        TecladoNumerico.Subcabecera = subcab
'        TecladoNumerico.TextoInicial = res
'
'        blnTecladoNumericoPrimeraVez = False
'
'        If Not TecladoNumerico.Visible Then
'            TecladoNumerico.Show 1
'        End If
'    End If
'End Sub
Private Sub Form_Activate()
'    If Not blnEsTablet Then
        If deter.ListItems.Count > 0 Then
           If deter.ListItems.Count >= deter.selectedItem.Index Then
            If (CLng(deter.ListItems(deter.selectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))) And _
               (CLng(deter.ListItems(deter.selectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))) And _
                deter.ListItems(deter.selectedItem.Index).SubItems(7) <> "1" Then
                    deter_Click
                    
'                    If blnTecladoNumericoPrimeraVez And blnEsTablet Then
'                        blnTecladoNumericoPrimeraVez = False
'                        MostrarTecladoNumerico deter.selectedItem.SubItems(1), txtdato.Text, txtvalor.Text
'                    End If
            End If
           End If
        End If
'    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
'    If UCase(usuario.getUSUARIO) = "JULIO" Then
'        auxdatos.Visible = True
'    End If
    cargar_botones Me
'    primera_vez = True
    cabecera
    cargar_combos
    inicializa_ventana
'    inicia_balanza
    cargar_parametros
'    ConfigurarTablet
'    If Not blnEsTablet Then
'        If deter.ListItems.Count > 0 Then
'           If deter.ListItems.Count >= deter.SelectedItem.Index Then
'            If (CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))) And _
'               (CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))) And _
'                deter.ListItems(deter.SelectedItem.Index).SubItems(7) <> "1" Then
'                    deter_Click
'
'                    If blnTecladoNumericoPrimeraVez And blnEsTablet Then
'                        blnTecladoNumericoPrimeraVez = False
'                        MostrarTecladoNumerico deter.SelectedItem.SubItems(1), txtdato.Text, txtvalor.Text
'                    End If
'            End If
'           End If
'        End If
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If blnEsTablet Then
'        Unload TecladoNumerico
'    End If
End Sub

Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = CLng(listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text)
        frmEquipoEdicion.Show 1
    End If
End Sub

Private Sub listaEquipos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Me.MousePointer = vbHourglass
    almacenar_equipos
    Me.MousePointer = vbNormal
End Sub


Private Sub spinDuplicados_Click(Index As Integer)
    'M1371-I
'    visualizar_duplicados
    If Index = 1 Then
        datos.ListItems(datos.selectedItem.Index).SubItems(1) = "CONFORME"
    ElseIf Index = 2 Then
        datos.ListItems(datos.selectedItem.Index).SubItems(1) = "NO CONFORME"
    End If
    grabar_auxdatos
    ' Pasar al siguiente campo
    If datos.ListItems.Count > datos.selectedItem.Index Then
        Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
        datos_Click
    Else
        If deter.ListItems.Count > deter.selectedItem.Index Then
            Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index + 1)
            Dim oDeter As New clsDeterminaciones
            oDeter.CargarDeterminacion (deter.ListItems(deter.selectedItem.Index).SubItems(4))
            If oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma") And oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico") Then
                deter_Click
                datos_Click
            Else
                deter_Click
            End If
        Else
            txtdato = ""
            txtvalor = ""
'            blnTecladoNumerico_NoMostrar = True
'            If Not blnEsTablet Then datos.SetFocus
            datos.SetFocus
        End If
    End If
    'M1371-F
End Sub


Private Sub txtgrado_Change()
    If deter.ListItems.Count > 0 Then
        deter.ListItems(deter.selectedItem.Index).SubItems(COLS.GRADO) = txtgrado
    End If
End Sub

Private Sub txtgrado_GotFocus()
    txtgrado.BackColor = &H80C0FF
    txtgrado.SelStart = 0
    txtgrado.SelLength = Len(Trim(txtgrado))

End Sub

Private Sub txtgrado_LostFocus()
    txtgrado.BackColor = vbWhite

End Sub

Private Sub txtmetodo_Change()
    If deter.ListItems.Count > 0 Then
        deter.ListItems(deter.selectedItem.Index).SubItems(10) = txtmetodo
    End If
End Sub


'Private Sub TecladoNumerico_Change(ByVal res As String)
'    txtvalor.Text = res
'End Sub
'
'Private Sub TecladoNumerico_Salir()
'    blnTecladoNumerico_NoMostrar = True
''    txtvalor_KeyPress 13
'End Sub
'
'
'Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'    'desestimarEvento = True
'    txtvalor_KeyPress 13
'    If chkCalcular.value = Checked Then
'        cmdCalcular_Click
'    End If
'    If datos.ListItems.Count >= datos.selectedItem.Index Then
'        cabecera = deter.selectedItem.SubItems(1)
'        Subcabecera = txtdato.Text
'        RESULTADO = txtvalor.Text
'    ElseIf deter.ListItems.Count > deter.selectedItem.Index Then
'        cabecera = deter.selectedItem.SubItems(1)
'        Subcabecera = txtdato.Text
'        RESULTADO = txtvalor.Text
'    ElseIf deter.ListItems.Count = deter.selectedItem.Index Then
'        If Not blnTecladoNumerico_NoMostrar Then
'            cabecera = deter.selectedItem.SubItems(1)
'            Subcabecera = txtdato.Text
'            RESULTADO = txtvalor.Text
'        Else
'            Cerrar = True
'        End If
'    Else
'        Cerrar = True
'    End If
'End Sub
'Private Sub TecladoNumerico_AnteriorElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'    Dim sw As Boolean
'    sw = True
'    If datos.selectedItem.Index > 1 Then
'        Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index - 1)
'    Else
'        If deter.selectedItem.Index > 1 Then
'            Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index - 1)
'            deter_Click
''            Set datos.SelectedItem = deter.ListItems(1)
'        Else
'            sw = False
'        End If
'    End If
'    If sw = True Then
''        txtvalor_KeyPress 13
'        datos_Click
'        If datos.ListItems.Count >= datos.selectedItem.Index Then
'            cabecera = deter.selectedItem.SubItems(1)
'            Subcabecera = txtdato.Text
'            RESULTADO = txtvalor.Text
'        ElseIf deter.ListItems.Count > deter.selectedItem.Index Then
'            cabecera = deter.selectedItem.SubItems(1)
'            Subcabecera = txtdato.Text
'            RESULTADO = txtvalor.Text
'        ElseIf deter.ListItems.Count = deter.selectedItem.Index Then
'            If Not blnTecladoNumerico_NoMostrar Then
'                cabecera = deter.selectedItem.SubItems(1)
'                Subcabecera = txtdato.Text
'                RESULTADO = txtvalor.Text
'            Else
'                Cerrar = True
'            End If
'        Else
'            Cerrar = True
'        End If
'    Else
'        Cerrar = True
'    End If
'
'End Sub
Private Sub txtmetodo_GotFocus()
    txtmetodo.BackColor = &H80C0FF
    txtmetodo.SelStart = 0
    txtmetodo.SelLength = Len(Trim(txtmetodo))
End Sub

Private Sub txtmetodo_LostFocus()
    txtmetodo.BackColor = vbWhite
End Sub

Private Sub txtvalor_GotFocus()
    txtvalor.BackColor = &H80C0FF
    txtvalor.SelStart = 0
    txtvalor.SelLength = Len(Trim(txtvalor))
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    On Error GoTo fallo
    If KeyAscii = 13 And datos.ListItems.Count > 0 Then
        KeyAscii = 0
        
        If Trim(txtvalor) = "" Then
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = " "
        Else
            If Not IsNumeric(Replace(txtvalor, ".", ",")) Then
                datos.ListItems(datos.selectedItem.Index).SubItems(1) = txtvalor
            Else
                If Trim(datos.ListItems(datos.selectedItem.Index).SubItems(3)) <> "" Then
                    datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(Replace(txtvalor, ".", ","), datos.ListItems(datos.selectedItem.Index).SubItems(4), datos.ListItems(datos.selectedItem.Index).SubItems(5))
                Else
                    datos.ListItems(datos.selectedItem.Index).SubItems(1) = Replace(txtvalor, ".", ",")
                End If
            End If
        End If
        Set ocampos = Nothing
        ' Validar rangos maximos del baño
        Dim min As Single
        Dim Max As Single
        If Trim(txtvalor) <> "" And IsNumeric(Trim(txtvalor)) = True And datos.ListItems(datos.selectedItem.Index).bold = True Then
            Dim oDA As New clsDeterminaciones_analisis
            Dim validar As Boolean
            If txtbano = 0 Then
               validar = oDA.Carga_por_tipo_analisis(CLng(txtanalis), deter.ListItems(deter.selectedItem.Index).SubItems(6))
            Else
               validar = oDA.Carga_por_BANO(CLng(txtbano), deter.ListItems(deter.selectedItem.Index).SubItems(6))
            End If
            Dim SITUACION As Integer
            min = 0
            Max = 0
            SITUACION = C_SITUACION.S_EN_RANGO
            If validar Then
                If Trim(oDA.getMINIMO) <> "" Then
                 If IsNumeric(Trim(oDA.getMINIMO)) Then
                  min = CSng(Replace(oDA.getMINIMO, ".", ","))
                  If deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE) <> "" Then
                    If IsNumeric(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE)) Then
                        min = min - CSng(Replace(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE), ".", ","))
                    End If
                  End If
                  If CSng(datos.ListItems(datos.selectedItem.Index).SubItems(1)) < min Then
                    SITUACION = C_SITUACION.S_FUERA_RANGO
                    MsgBox "El valor introducido supera el mínimo exigido en los rangos. Considere revisar el histórico de resultados.", vbExclamation, App.Title
                  End If
                 End If
                End If
                If Trim(oDA.getMAXIMO) <> "" Then
                 If IsNumeric(Trim(oDA.getMAXIMO)) Then
                  Max = CSng(Replace(oDA.getMAXIMO, ".", ","))
                  If deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE) <> "" Then
                    If IsNumeric(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE)) Then
                        Max = Max + CSng(Replace(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.INCERTIDUMBRE), ".", ","))
                    End If
                  End If
                  If CSng(datos.ListItems(datos.selectedItem.Index).SubItems(1)) > Max Then
                    SITUACION = C_SITUACION.S_FUERA_RANGO
                    MsgBox "El valor introducido supera el máximo exigido en los rangos. Considere revisar el histórico de resultados.", vbExclamation, App.Title
                  End If
                 End If
                End If
                ' Verificar alerta de límites
                If SITUACION = C_SITUACION.S_EN_RANGO And (Trim(oDA.getMINIMO) <> "" Or Trim(oDA.getMAXIMO) <> "") Then
                    If Trim(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_AVISO)) <> "" Then
                        If Max > min Then
                            Dim dif As Single
                            dif = ((Max - min) * Trim(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_AVISO)) / 100)
                            If Trim(oDA.getMINIMO) <> "" Then
                                If IsNumeric(Trim(oDA.getMINIMO)) Then
                                    min = CSng(Replace(oDA.getMINIMO, ".", ",")) + dif
                                    If CSng(Replace(datos.ListItems(datos.selectedItem.Index).SubItems(1), ".", ",")) < min Then
                                       SITUACION = C_SITUACION.S_LIMITES
                                    End If
                                End If
                            End If
                            If Trim(oDA.getMAXIMO) <> "" Then
                                If IsNumeric(Trim(oDA.getMAXIMO)) Then
                                    Max = CSng(Replace(oDA.getMAXIMO, ".", ",")) - dif
                                    If CSng(Replace(datos.ListItems(datos.selectedItem.Index).SubItems(1), ".", ",")) > Max Then
                                       SITUACION = C_SITUACION.S_LIMITES
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                deter.ListItems(deter.selectedItem.Index).SubItems(COLS.SITUACION) = SITUACION
                deter.ListItems(deter.selectedItem.Index).ListSubItems(3).ReportIcon = SITUACION + 1
            End If
            ' Validar diferencia de resultados del baño (%DIF_HISTORICO)
            If txtbano <> 0 Then
                Dim oMuestra As New clsMuestra
                Dim dif_min As Single
                Dim dif_max As Single
                Dim resultadoActual As String
                Dim dif_historico As String
                
                resultadoActual = Replace(datos.ListItems(datos.selectedItem.Index).SubItems(1), ".", ",")
                dif_historico = Replace(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.dif_historico), ".", ",")
                
                If IsNumeric(dif_historico) And IsNumeric(resultadoActual) Then
                    Dim rs As ADODB.Recordset
                    Set rs = oMuestra.obtener_muestras_anteriores(CLng(gmuestra), CLng(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.ID_DETERMINACION)), True)
                    If rs.RecordCount <> 0 Then
                        Dim resultadoAnterior As String
                        resultadoAnterior = Replace(rs("resultado"), ".", ",")
                        
                        If IsNumeric(resultadoAnterior) Then
                            dif_min = CSng(resultadoAnterior) - (CSng(resultadoAnterior) * dif_historico / 100)
                            dif_max = CSng(resultadoAnterior) + (CSng(resultadoAnterior) * dif_historico / 100)
                            
                            If dif_min > CSng(resultadoActual) Or dif_max < CSng(resultadoActual) Then
                                If MsgBox("La diferencia respecto al baño anterior es mayor a la permitida del " & dif_historico & " %. ¿Mostrar histórico?", vbInformation + vbYesNo, App.Title) = vbYes Then
                                    gdeterminacion = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.ID_DETERMINACION)
                                    frmHistoricoDeterminacion.Show 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
'                    If Trim(oDA.getDIF_MINIMA) = "" Then
'                        min = CSng(datos.ListItems(datos.selectedItem.Index).SubItems(1))
'                    Else
'                        min = CSng(Replace(oDA.getDIF_MINIMA, ".", ","))
'                    End If
'                    If Trim(oDA.getDIF_MAXIMA) = "" Then
'                        Max = 99999
'                    Else
'                        If IsNumeric(oDA.getDIF_MAXIMA) Then
'                            Max = CSng(Replace(oDA.getDIF_MAXIMA, ".", ","))
'                        Else
'                            Max = 99999
'                        End If
'                    End If
'                    dif_min = CSng(datos.ListItems(datos.selectedItem.Index).SubItems(1)) - min
'                    dif_max = CSng(datos.ListItems(datos.selectedItem.Index).SubItems(1)) + Max
'                    If IsNumeric(rs(3)) Then
'                        If Trim(rs(3)) = "" Then
'                            RESULTADO = 0
'                        Else
'                            RESULTADO = CSng(Replace(rs(3), ".", ","))
'                        End If
'                        If RESULTADO <> 0 Then
'                            If dif_min > RESULTADO Or _
'                               dif_max < RESULTADO Then
'                               If MsgBox("La diferencia respecto al baño anterior es mayor a la permitida. ¿Mostrar histórico?", vbInformation + vbYesNo, App.Title) = vbYes Then
'                                    gdeterminacion = deter.ListItems(deter.selectedItem.Index).SubItems(COLS.ID_DETERMINACION)
'                                    frmHistoricoDeterminacion.Show 1
'                               End If
'                            End If
'                        End If
'                    End If
'                End If
'            End If
        End If
        ' Validación de resultado >= que el LC
'        If Trim(txtValor) <> "" And IsNumeric(Trim(txtValor)) = True And datos.ListItems(datos.selectedItem.Index).bold = True Then
'            If Trim(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.LC)) <> "" Then
'                If IsNumeric(Trim(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.LC))) Then
'                    If CSng(Trim(txtValor)) < CSng(Replace(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.LC), ".", ",")) Then
'                        MsgBox "El valor introducido es menor que el límite de cuantificación.", vbInformation, App.Title
'                    End If
'                End If
'            End If
'        End If
        ' JGM-I
        ' Nuevo cambio. Para la matraz en solidos disueltos, preguntar por el numero de matraz
        ' y si existe alguna muestra abierta con esa matraz, recuperar los datos
        buscar_matraz txtvalor
        ' JGM-F
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            visualizar_duplicados
        End If
        grabar_auxdatos
        ' Pasar al siguiente campo
        If datos.ListItems.Count > datos.selectedItem.Index Then
            Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
            datos_Click
        Else
            If deter.ListItems.Count > deter.selectedItem.Index Then
                Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index + 1)
                Dim oDeter As New clsDeterminaciones
                oDeter.CargarDeterminacion (deter.ListItems(deter.selectedItem.Index).SubItems(4))
                If oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma") And oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico") Then
                    deter_Click
                    datos_Click
                Else
                    deter_Click
                End If
            Else
                txtdato = ""
                txtvalor = ""
'                blnTecladoNumerico_NoMostrar = True
'                If Not blnEsTablet Then datos.SetFocus
                datos.SetFocus
            End If
        End If
    End If
    Exit Sub
fallo:
    error_grave "Error en frmDeterminaciones(txtvalor_KeyPress) : " & Err.Description
End Sub

Private Sub txtvalor_LostFocus()
    txtvalor.BackColor = vbWhite
End Sub

Public Sub grabar_auxdatos()
    Dim i As Integer
    For i = auxdatos.ListItems.Count To 1 Step -1
       If deter.ListItems(deter.selectedItem.Index).SubItems(4) = auxdatos.ListItems(i) Then
          auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , deter.ListItems(deter.selectedItem.Index).SubItems(4))
             .SubItems(1) = datos.ListItems(i).SubItems(1)
             .SubItems(2) = i
             .SubItems(3) = datos.ListItems(i).SubItems(3)
             If datos.ListItems(i).bold = True Then
                .bold = True
                ' Si es solucion, la subimoslas determinaciones
                If UCase(lblestado.Caption) <> "DUPLICADA" Then
                 If datos.ListItems(i).SubItems(1) <> "" Then
                    deter.ListItems(deter.selectedItem.Index).SubItems(3) = datos.ListItems(i).SubItems(1)
                 End If
               'M1371-I
                Else
                   'Marcador "REV." nos permitirá aislar el campo para su tratamiento (diferenciado de la MEDIA, etc.)
                    If datos.ListItems(i).Text = "Revisión de Duplicados" Then
                        .SubItems(4) = "REV."
                        .SubItems(1) = datos.ListItems(i).SubItems(1)
                    End If
                'M1371-F
                End If
             Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If datos.ListItems(i).Text = "Resultado (MEDIA)" Then
                        .SubItems(4) = "M"
                    End If
                    deter.ListItems(deter.selectedItem.Index).SubItems(3) = datos.ListItems(datos.ListItems.Count - 3).SubItems(1)
                    If datos.ListItems(datos.ListItems.Count).SubItems(1) = "CONFORME" Then
                        deter.ListItems(deter.selectedItem.Index).SubItems(COLS.REV_DUPLICADO) = 1
                    ElseIf datos.ListItems(datos.ListItems.Count).SubItems(1) = "NO CONFORME" Then
                        deter.ListItems(deter.selectedItem.Index).SubItems(COLS.REV_DUPLICADO) = 2
                    Else
                        deter.ListItems(deter.selectedItem.Index).SubItems(COLS.REV_DUPLICADO) = 0
                    End If
                End If
             End If
       End With
    Next
End Sub

Private Sub siguiente_campo()
    If deter.ListItems.Count > deter.selectedItem.Index Then
        Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index + 1)
        deter_Click
        datos_Click
    Else
        datos.ListItems.Clear
        txtdato = ""
        txtvalor = ""
'        If Not blnEsTablet Then
            datos.SetFocus
'        End If
    End If
End Sub

Public Sub inicializa_ventana()
    ' Título
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (gmuestra)
    txtbano = oMuestra.getBANO_ID
    txtanalisis = oMuestra.getTIPO_ANALISIS_ID
    lbltitulo = "Registro determinaciones muestra : " & Trim(str(oMuestra.getID_GENERAL)) & "/" & oMuestra.getANNO & " (" & oMuestra.CodigoParticular(gmuestra) & ")"
    If oMuestra.getBANO_ID = 0 Then
        cmdTA.visible = True
        cmdBano.visible = False
    Else
        cmdTA.visible = False
        cmdBano.visible = True
    End If
    
    ' Permiso para modificar la vida
    Dim op As New clsParametros
    Dim s() As String
    Dim i As Integer
    op.Carga parametros.PARAM_USUARIOS_MODIFICAN_EQUIPOS_MUESTRA_CERRADA, ""
    If op.getVALOR <> "" Then
        s = Split(op.getVALOR, ",")
        For i = LBound(s) To UBound(s)
            If USUARIO.getID_EMPLEADO = CInt(s(i)) Then
                chkModificar.Value = Checked
                Exit For
            End If
        Next
    End If
    Set op = Nothing
    
    
    Me.Caption = lbltitulo
    lblestado.Caption = ""
    ' Comprobar duplicada
    If oMuestra.getANALISIS_DUPLICADO = 1 Then
        lblestado.Caption = "DUPLICADA"
        lblestado.visible = True
    Else
        lblestado.visible = False
    End If
    ' Determinaciones
    cargar_determinaciones
    proteger_campos oMuestra.getCERRADA
End Sub

Private Sub visualizar_duplicados()
    On Error GoTo fallo
        ' Si la muestra es duplicada, visualizar resultados
        Dim numero_resultados As Integer
        Dim res1 As String
        Dim res2 As String
        Dim campo As Integer
        Dim ndecimales As Integer
        Dim nenteros As Integer
        numero_resultados = 0
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            For i = 1 To datos.ListItems.Count - 1
                If datos.ListItems(i).bold = True Then
                    If Trim(datos.ListItems(i).SubItems(1)) <> "" Then
                        numero_resultados = numero_resultados + 1
                        If Trim(res1) = "" Then
                            res1 = datos.ListItems(i).SubItems(1)
                            campo = datos.ListItems(i).SubItems(3)
                        Else
                            res2 = datos.ListItems(i).SubItems(1)
                        End If
                    End If
                End If
            Next
        End If
        
        If numero_resultados = 2 And IsNumeric(res1) And IsNumeric(res2) Then ' Calcular media y diferencia
            Dim media As Single
            Dim dif As Single
            Dim dif_media As Single
            Dim ocf As New clsFormulas_campos
            If campo <> 0 Then
                ocf.CARGAR (campo)
                ndecimales = ocf.getDECIMALES
                nenteros = ocf.getENTEROS
            Else
                nenteros = 5
                ndecimales = 2
            End If
            ' JGM Datos de la diferencia
            dif = Abs((CSng(res1) - CSng(res2)))
            'M1371-I
            'Todas las referencias basadas en datos.listitems.count varian en una unidad debido a la adición del nuevo campo.
            'Se modifican estas referencias añadiendo +1
            'datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = formatear(CStr(dif), nenteros, ndecimales)
            datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = formatear(CStr(dif), nenteros, ndecimales)
            'M1371-F
            ' Datos de la media
            media = (CSng(res1) + CSng(res2)) / 2
            'M1371-I
            'datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = formatear(CStr(media), nenteros, ndecimales)
            datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = formatear(CStr(media), nenteros, ndecimales)
            'M1371-F
            ' Se modifica la diferencia para que siempre se muestre en %
            If media = 0 Then
                dif_media = 0
            Else
                dif_media = (dif / media) * 100
            End If
            ' Se modifica a 1 decimal, petición de jesús
            'M1371-I
            'datos.ListItems(datos.ListItems.Count).SubItems(1) = formatear(CStr(dif_media), 2, 1)
            datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = formatear(CStr(dif_media), 2, 1)
            ' Datos de la repetibilidad de duplicados
            'M1371-F
            'M1371-I
            Select Case deter.ListItems(deter.selectedItem.Index).SubItems(COLS.REV_DUPLICADO)
            Case 0
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "NO REALIZADO"
            Case 1
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "CONFORME"
            Case 2
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "NO CONFORME"
            End Select
            'M1371-F
            ' Mensaje de Dif entre duplicados
            If deter.ListItems.Count > 0 Then
                If CSng(Format(Replace(dif_media, ".", ","), "#.0")) > CSng(Format(Replace(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS), ".", ","), "#.0")) And _
                   deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS) <> 0 Then
                    MsgBox "% de diferencia entre duplicados mayor que la permitida.", vbInformation, App.Title
                Else
                    If Trim(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS_NUMERICA)) <> "" Then
                        If IsNumeric(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS_NUMERICA)) Then
                          If CSng(dif) > CSng(deter.ListItems(deter.selectedItem.Index).SubItems(COLS.DIF_DUPLICADOS_NUMERICA)) Then
                            MsgBox "Diferencia entre duplicados mayor que la permitida.", vbInformation, App.Title
                          End If
                        End If
                    End If
                End If
            End If
        Else
            If res1 = "--" Or res2 = "--" Then
            'M1371-I
            '    datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = "--"
            '    datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = "--"
            '    datos.ListItems(datos.ListItems.Count).SubItems(1) = "--"
            
                datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = "--"
                datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = "--"
                datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = "--"
            'M1371-F
            Else
              If numero_resultados = 1 Then
                'M1371-I
                'datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = res1
                datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = res1
                'M1371-F
              Else
                If numero_resultados = 0 Then
                    'M1371-I
                    'datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = ""
                    'datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = ""
                    'datos.ListItems(datos.ListItems.Count).SubItems(1) = ""
                    datos.ListItems(datos.ListItems.Count - 3).SubItems(1) = ""
                    datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = ""
                    datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = ""
                    'M1371-F
                Else
                    If UCase(lblestado.Caption) = "DUPLICADA" Then
                        'M1371-I
                        'datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = deter.ListItems(deter.selectedItem.Index).SubItems(3)
                        datos.ListItems(datos.ListItems.Count - 2).SubItems(1) = deter.ListItems(deter.selectedItem.Index).SubItems(3)
                        'M1371-F
                    End If
                End If
              End If
            End If
        End If
        'M1371-I
        'deter.ListItems(deter.selectedItem.Index).SubItems(3) = datos.ListItems(datos.ListItems.Count - 2).SubItems(1)
        ' deter.ListItems(deter.selectedItem.Index).SubItems(3) = datos.ListItems(datos.ListItems.Count - 3).SubItems(1)
        'M1371-F
    Exit Sub
fallo:
    MsgBox "Error en visualizar_duplicados." & Err.Description, vbCritical, App.Title
End Sub
'Private Sub inicia_balanza()
'    On Error Resume Next
'    MSComm1.InputLen = 0 ' El valor 0 hace que se lea todo
'    MSComm1.RThreshold = 1 ' al recibir uno o mas caracteres
'    MSComm1.SThreshold = 1 ' al enviar uno o mas caracteres
'    MSComm1.CommPort = 1 'Paso 1: elijo el puerto COM 1
'    MSComm1.settings = "1200,O,7,1" ' Vel. 1200, paridad odd, 7 bits
'    MSComm1.PortOpen = True 'Abro el puerto
'End Sub
'Private Sub cerrar_balanza()
'    On Error Resume Next
'    MSComm1.PortOpen = False 'Puede haber error si
'End Sub
'Private Sub MSComm1_OnComm()
'    If MSComm1.CommEvent = comEvReceive Then
'     Sleep 500
'     txtvalor.Text = Trim(Replace(Mid(MSComm1.Input, 3, 8), ".", ","))
'    End If
'End Sub
Private Sub cargar_combos()
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, " AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
End Sub
Private Sub almacenar_equipos()
      ' Equipos
      Dim OTDE As New clsDeterminaciones_equipos
   On Error GoTo almacenar_equipos_Error
      Dim i As Integer
      OTDE.Eliminar CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
      For i = 1 To listaEquipos.ListItems.Count
        With OTDE
            .setDETERMINACION_ID = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setORDEN = i
            'M1385-I
            If listaEquipos.ListItems(i).Checked = True Then
                .setEN_INFORME = 1
            Else
                .setEN_INFORME = 0
            End If
            'M1385-F
            .Insertar_Determinacion
        End With
      Next
      ' Usos de los equipos
      Dim oEU As New clsEq_usos
      For i = 1 To listaEquipos.ListItems.Count
        With oEU
            .Eliminar gmuestra, CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
            .setMUESTRA_ID = gmuestra
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setDETERMINACION_ID = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
            .setUSOS = 1
            .Insertar
        End With
      Next

   On Error GoTo 0
   Exit Sub

almacenar_equipos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure almacenar_equipos of Formulario frmDeterminaciones"
End Sub
Private Sub almacenar_reactivos()
      ' Equipos
      Dim OTDE As New clsDeterminaciones_reactivos
      Dim i As Integer
   On Error GoTo almacenar_reactivos_Error

      OTDE.Eliminar CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
      For i = 1 To listaReactivos.ListItems.Count
        With OTDE
            .setDETERMINACION_ID = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(4))
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setTIPO = listaReactivos.ListItems(i).SubItems(3)
            .setORDEN = i
            .Insertar_Determinacion
        End With
      Next

   On Error GoTo 0
   Exit Sub

almacenar_reactivos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure almacenar_reactivos of Formulario frmDeterminaciones"

End Sub

Private Sub buscar_matraz(VALOR As String)
   ' Nuevo cambio. Para la matraz en solidos disueltos, preguntar por el numero de matraz
   ' y si existe alguna muestra abierta con esa matraz, recuperar los datos
   On Error GoTo buscar_matraz_Error
    If Trim(deter.ListItems(deter.selectedItem.Index).SubItems(5)) = "" Then
        Exit Sub
    End If
    If Trim(datos.ListItems(datos.selectedItem.Index).SubItems(3)) = "" Then
        Exit Sub
    End If
    If CInt(deter.ListItems(deter.selectedItem.Index).SubItems(5)) = CInt(txtformulasolidos) Then
        If CInt(datos.ListItems(datos.selectedItem.Index).SubItems(3)) = CInt(txtcampomatraz) Then
            If Trim(txtvalor) <> "" Then
              ' Busco los datos si no lo habia realizado anteriormente, mirando si el siguiente campo esta informado
              If Trim(datos.ListItems(datos.selectedItem.Index + 1).SubItems(1)) = "" Then
                Dim consulta As String
                consulta = " select campo_id,valor_1,valor_2 from datos_determinaciones " & _
                        "        where determinacion_id = ( " & _
                        "        select max(dd.determinacion_id) " & _
                        "         from muestras m, determinaciones d, datos_determinaciones dd " & _
                        "        Where m.id_muestra = d.MUESTRA_ID " & _
                        "          and d.ID_DETERMINACION = dd.DETERMINACION_ID " & _
                        "          and m.ANULADA=0 " & _
                        "          and m.CERRADA<>1 " & _
                        "          and d.FORMULA_ID = " & CInt(txtformulasolidos) & _
                        "          and dd.CAMPO_ID = " & CInt(txtcampomatraz) & _
                        "          and (VALOR_1= '" & VALOR & "' OR VALOR_2='" & VALOR & "')" & _
                        "          and m.ANNO = " & Year(Date) & _
                        "        )"
                Dim rs As ADODB.Recordset
                Dim VALOR1 As Boolean
                Dim DETERMINACIONES As Integer
                Set rs = datos_bd(consulta)
                DETERMINACIONES = 0
                If rs.RecordCount > 0 Then
'                    MsgBox "MATRAZ ENCONTRADA"
                    Do
                        VALOR1 = False
                        For i = datos.selectedItem.Index To datos.ListItems.Count
                            If datos.ListItems(i).SubItems(3) = rs(0) Then
                                If VALOR1 = False Then
                                    datos.ListItems(i).SubItems(1) = rs(1)
                                    VALOR1 = True
 '                               Else
 '                                   datos.ListItems(i).SubItems(1) = rs(2)
                                End If
                            End If
                        Next
                        rs.MoveNext
                        DETERMINACIONES = DETERMINACIONES + 1
                    Loop Until rs.EOF Or DETERMINACIONES = 5
'                Else
'                    MsgBox "MATRAZ NO ENCONTRADA"
                End If
               End If
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

buscar_matraz_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure buscar_matraz of Formulario frmDeterminaciones"
End Sub

Private Sub cargar_parametros()
    Dim oParametros As New clsParametros
   On Error GoTo cargar_parametros_Error

    oParametros.Carga parametros.MATRAZ_FORMULA, ""
    txtformulasolidos = oParametros.getVALOR
    oParametros.Carga parametros.MATRAZ_CAMPO, ""
    txtcampomatraz = oParametros.getVALOR


   On Error GoTo 0
   Exit Sub

cargar_parametros_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_parametros of Formulario frmDeterminaciones"
End Sub

Private Sub proteger_campos(CERRADA As Integer)
    If CERRADA = 1 And chkModificar.Value = Unchecked Then
        cmdok.Enabled = False
        txtmetodo.Locked = True
        txtgrado.Locked = True
        ' TEMPORTAL ELENA
        'cmdEliminarEquipo.Enabled = False
        'cmdAnadirEquipo.Enabled = False
        'cmdEliminarReactivo.Enabled = False
        'cmdAnadirReactivo.Enabled = False
        cmdEliminarEquipo.Enabled = True
        cmdAnadirEquipo.Enabled = True
        cmdEliminarReactivo.Enabled = True
        cmdAnadirReactivo.Enabled = True
        ' TEMP-F
        txtvalor.Enabled = False
    Else
        cmdok.Enabled = True
        txtmetodo.Locked = False
        txtgrado.Locked = False
        cmdEliminarEquipo.Enabled = True
        cmdAnadirEquipo.Enabled = True
        cmdEliminarReactivo.Enabled = True
        cmdAnadirReactivo.Enabled = True
        txtvalor.Enabled = True
    End If
    Select Case CERRADA
        Case 0
            lblCerrada = "ABIERTA"
        Case 1
            lblCerrada = "CERRADA"
        Case 2
            lblCerrada = "PTE. CIERRE"
        Case 3
            lblCerrada = "C.SIN INFORME"
    End Select
End Sub
