VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProcNC_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DETALLE DE INCIDENCIA"
   ClientHeight    =   9645
   ClientLeft      =   2265
   ClientTop       =   1740
   ClientWidth     =   11310
   Icon            =   "frmProcNC_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   11310
   Begin VB.CommandButton cmdInformeParcial 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Parcial"
      Height          =   870
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   8730
      Width           =   1020
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Completo"
      Height          =   870
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   8730
      Width           =   1020
   End
   Begin VB.CommandButton cmdRechazarEstado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rechazar VoBo"
      Height          =   870
      Left            =   2070
      Picture         =   "frmProcNC_Detalle.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   8730
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdCambiarEstado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitar PNC (Vº Bº)"
      Height          =   870
      Left            =   90
      Picture         =   "frmProcNC_Detalle.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   8730
      Width           =   1905
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción de la Incidencia"
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
      Height          =   2805
      Index           =   1
      Left            =   12960
      TabIndex        =   13
      Top             =   1320
      Width           =   10350
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Index           =   10
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1485
         Width           =   8910
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   825
         Index           =   0
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   630
         Width           =   8910
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   270
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cmborigen 
         Height          =   315
         Left            =   1215
         TabIndex        =   3
         Top             =   2385
         Width           =   8910
         _ExtentX        =   15716
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
      Begin VB.Label lblCapCausas 
         Caption         =   "Problemas de Aseguramiento de la Calidad"
         Height          =   585
         Index           =   3
         Left            =   0
         TabIndex        =   222
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Origen"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   36
         Top             =   2430
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Acc.inmediata"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   28
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   900
         Width           =   840
      End
   End
   Begin TabDlg.SSTab tabDatosNoConformidad 
      Height          =   6855
      Left            =   30
      TabIndex        =   41
      Top             =   1770
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      Tab             =   8
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "1.- Datos Apertura Incidencia"
      TabPicture(0)   =   "frmProcNC_Detalle.frx":3C8E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDptoImplicados"
      Tab(0).Control(1)=   "fraResponsableAperturaIncidencia"
      Tab(0).Control(2)=   "fraOrigen"
      Tab(0).Control(3)=   "ctlLinea"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "2.- Descripción y Documentación"
      TabPicture(1)   =   "frmProcNC_Detalle.frx":3CAA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraIncidencia"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "3.- Equipo Humano"
      TabPicture(2)   =   "frmProcNC_Detalle.frx":3CC6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraEquipoDesignado"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "4.- Identificación del Problema"
      TabPicture(3)   =   "frmProcNC_Detalle.frx":3CE2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lstDocumentacionIdentificacion"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdAsociarDocumentacionIdentificacion"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame18"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "5.- Análisis de la Escena"
      TabPicture(4)   =   "frmProcNC_Detalle.frx":3CFE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "6.- Recolección de Datos"
      TabPicture(5)   =   "frmProcNC_Detalle.frx":3D1A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label7"
      Tab(5).Control(1)=   "Label15"
      Tab(5).Control(2)=   "lstDocumentacionRecoleccionDatos"
      Tab(5).Control(3)=   "txtRecoleccionDatos"
      Tab(5).Control(4)=   "cmdAdjuntarRecoleccionDatos"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "7.- Causas"
      TabPicture(6)   =   "frmProcNC_Detalle.frx":3D36
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame3"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "8.- Resumen de Causas"
      TabPicture(7)   =   "frmProcNC_Detalle.frx":3D52
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label11"
      Tab(7).Control(1)=   "fraCausaRaiz"
      Tab(7).Control(2)=   "fraCausasContributibas"
      Tab(7).Control(3)=   "fraCausaDirecta"
      Tab(7).Control(4)=   "txtResumenCausas"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "9.- Plan de Acciones Correctivas"
      TabPicture(8)   =   "frmProcNC_Detalle.frx":3D6E
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "fraResumenAccionesCorrectivas"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Frame15"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "10.- Evaluación"
      TabPicture(9)   =   "frmProcNC_Detalle.frx":3D8A
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "fraResultado"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Frame16"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "Frame17"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "Frame19"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).ControlCount=   4
      Begin VB.Frame Frame19 
         Appearance      =   0  'Flat
         Caption         =   "Observaciones a la Evaluación"
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   -74880
         TabIndex        =   197
         Top             =   2910
         Width           =   10845
         Begin VB.TextBox txtObservaciones 
            Appearance      =   0  'Flat
            Height          =   3405
            Left            =   120
            MaxLength       =   65000
            MultiLine       =   -1  'True
            TabIndex        =   198
            Top             =   270
            Width           =   10605
         End
      End
      Begin VB.Frame Frame17 
         Appearance      =   0  'Flat
         Caption         =   "Solución Aceptable"
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   -67860
         TabIndex        =   194
         Top             =   1860
         Width           =   3825
         Begin VB.OptionButton optEval_res_no 
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   450
            TabIndex        =   196
            Top             =   570
            Width           =   885
         End
         Begin VB.OptionButton optEval_res_si 
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   450
            TabIndex        =   195
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
         Caption         =   "Evidencias"
         ForeColor       =   &H80000008&
         Height          =   2085
         Left            =   -74880
         TabIndex        =   176
         Top             =   750
         Width           =   6735
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   7
            Left            =   150
            TabIndex        =   189
            Top             =   1440
            Width           =   6345
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   7
               Left            =   5610
               TabIndex        =   191
               Top             =   30
               Width           =   255
            End
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   8
               Left            =   6030
               TabIndex        =   190
               Top             =   30
               Width           =   315
            End
            Begin VB.Label Label9 
               Caption         =   "¿Se han comunicado las modificaciones a todos los departamentos?"
               Height          =   495
               Index           =   7
               Left            =   30
               TabIndex        =   192
               Top             =   60
               Width           =   5025
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Height          =   465
            Index           =   6
            Left            =   150
            TabIndex        =   185
            Top             =   960
            Width           =   6345
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   5
               Left            =   5610
               TabIndex        =   187
               Top             =   30
               Width           =   315
            End
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   6
               Left            =   6030
               TabIndex        =   186
               Top             =   30
               Width           =   315
            End
            Begin VB.Label Label9 
               Caption         =   "¿Disponemos de las evidencias de todas las acciones tomadas para la corrección de la incidencia ?"
               Height          =   585
               Index           =   6
               Left            =   0
               TabIndex        =   188
               Top             =   30
               Width           =   5205
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   150
            TabIndex        =   181
            Top             =   630
            Width           =   6345
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   3
               Left            =   5610
               TabIndex        =   183
               Top             =   30
               Width           =   315
            End
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   4
               Left            =   6030
               TabIndex        =   182
               Top             =   30
               Width           =   315
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "¿Son Efectivas?"
               Height          =   195
               Index           =   5
               Left            =   0
               TabIndex        =   184
               Top             =   60
               Width           =   1170
            End
         End
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   177
            Top             =   330
            Width           =   6345
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   2
               Left            =   6030
               TabIndex        =   179
               Top             =   60
               Width           =   315
            End
            Begin VB.OptionButton optEvidencias 
               Height          =   255
               Index           =   1
               Left            =   5610
               TabIndex        =   178
               Top             =   60
               Width           =   315
            End
            Begin VB.Label Label9 
               Caption         =   "¿Las Acciones Correctivas han sido puestas en marcha en plazo?"
               Height          =   495
               Index           =   4
               Left            =   0
               TabIndex        =   180
               Top             =   60
               Width           =   4665
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Sí  -  No"
            Height          =   195
            Left            =   5790
            TabIndex        =   193
            Top             =   150
            Width           =   600
         End
      End
      Begin VB.Frame fraResultado 
         Appearance      =   0  'Flat
         Caption         =   "Resultado"
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   -67860
         TabIndex        =   173
         Top             =   750
         Width           =   3825
         Begin VB.OptionButton optEval_res_nc 
            Caption         =   "NO CONFORMIDAD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   450
            TabIndex        =   175
            Top             =   570
            Width           =   2925
         End
         Begin VB.OptionButton optEval_res_incidencia 
            Caption         =   "INCIDENCIA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   450
            TabIndex        =   174
            Top             =   210
            Width           =   2985
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Identificacion del Problema"
         Height          =   4305
         Left            =   -74940
         TabIndex        =   167
         Top             =   630
         Width           =   10995
         Begin VB.CommandButton cmdEliminarPregunta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar Pregunta/Respuesta"
            Height          =   480
            Left            =   5730
            Style           =   1  'Graphical
            TabIndex        =   199
            Top             =   3720
            Width           =   2415
         End
         Begin VB.CommandButton cmdAnadirPregunta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir Pregunta/Respuesta"
            Height          =   480
            Left            =   8220
            Style           =   1  'Graphical
            TabIndex        =   169
            Top             =   3720
            Width           =   2415
         End
         Begin MSFlexGridLib.MSFlexGrid lstCuestionesIdentProblema 
            Height          =   3435
            Left            =   90
            TabIndex        =   168
            Top             =   240
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   6059
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
      End
      Begin VB.CommandButton cmdAdjuntarRecoleccionDatos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asociar Documentación"
         Height          =   870
         Left            =   -65880
         Picture         =   "frmProcNC_Detalle.frx":3DA6
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   5880
         Width           =   1905
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Determinar las causas"
         ForeColor       =   &H80000008&
         Height          =   6105
         Left            =   -74910
         TabIndex        =   163
         Top             =   660
         Width           =   10935
         Begin VB.Frame fraProblemasHumanos 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   90
            TabIndex        =   226
            Top             =   5010
            Width           =   10755
            Begin VB.Frame Frame11 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   247
               Top             =   300
               Width           =   4515
               Begin VB.OptionButton OptProblemasHumanos_operador_sustituido_si 
                  Height          =   255
                  Left            =   3750
                  TabIndex        =   249
                  Top             =   30
                  Width           =   315
               End
               Begin VB.OptionButton OptProblemasHumanos_operador_sustituido_no 
                  Height          =   255
                  Left            =   4170
                  TabIndex        =   248
                  Top             =   30
                  Width           =   315
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "¿Ha sido sustituido el Operador?"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   250
                  Top             =   60
                  Width           =   2295
               End
            End
            Begin VB.Frame Frame12 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   0
               Left            =   1410
               TabIndex        =   243
               Top             =   540
               Width           =   4515
               Begin VB.OptionButton OptProblemasHumanos_instrucciones_incompletas_no 
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   245
                  Top             =   30
                  Width           =   315
               End
               Begin VB.OptionButton OptProblemasHumanos_instrucciones_incompletas_si 
                  Height          =   255
                  Left            =   3780
                  TabIndex        =   244
                  Top             =   30
                  Width           =   315
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "¿Están incompletas las Instrucciones de Tabajo?"
                  Height          =   195
                  Index           =   0
                  Left            =   30
                  TabIndex        =   246
                  Top             =   60
                  Width           =   3465
               End
            End
            Begin VB.Frame Frame11 
               BorderStyle     =   0  'None
               Height          =   165
               Index           =   1
               Left            =   1440
               TabIndex        =   239
               Top             =   870
               Width           =   4515
               Begin VB.OptionButton OptProblemasHumanos_herramientas_adecuadas_no 
                  Height          =   255
                  Left            =   4170
                  TabIndex        =   241
                  Top             =   -30
                  Width           =   315
               End
               Begin VB.OptionButton OptProblemasHumanos_herramientas_adecuadas_si 
                  Height          =   255
                  Left            =   3750
                  TabIndex        =   240
                  Top             =   -30
                  Width           =   315
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "¿Son adecuadas las Herramientas de Trabajo?"
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   242
                  Top             =   -30
                  Width           =   3330
               End
            End
            Begin VB.Frame Frame12 
               BorderStyle     =   0  'None
               Height          =   285
               Index           =   1
               Left            =   6180
               TabIndex        =   235
               Top             =   270
               Width           =   4515
               Begin VB.OptionButton OptProblemasHumanos_formacion_suficiente_si 
                  Height          =   255
                  Left            =   3780
                  TabIndex        =   237
                  Top             =   30
                  Width           =   315
               End
               Begin VB.OptionButton OptProblemasHumanos_formacion_suficiente_no 
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   236
                  Top             =   30
                  Width           =   315
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "¿Ha sido suficiente la formación?"
                  Height          =   195
                  Index           =   1
                  Left            =   30
                  TabIndex        =   238
                  Top             =   60
                  Width           =   2340
               End
            End
            Begin VB.Frame Frame12 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   2
               Left            =   6210
               TabIndex        =   231
               Top             =   540
               Width           =   4515
               Begin VB.OptionButton OptProblemasHumanos_objetivos_marcados_claramente_si 
                  Height          =   255
                  Left            =   3750
                  TabIndex        =   233
                  Top             =   0
                  Width           =   315
               End
               Begin VB.OptionButton OptProblemasHumanos_objetivos_marcados_claramente_no 
                  Height          =   255
                  Left            =   4170
                  TabIndex        =   232
                  Top             =   0
                  Width           =   315
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "¿Han quedado los objetivos claramente marcados?"
                  Height          =   195
                  Index           =   2
                  Left            =   0
                  TabIndex        =   234
                  Top             =   30
                  Width           =   3630
               End
            End
            Begin VB.Frame Frame11 
               BorderStyle     =   0  'None
               Height          =   255
               Index           =   3
               Left            =   6180
               TabIndex        =   227
               Top             =   780
               Width           =   4515
               Begin VB.OptionButton OptProblemasHumanos_proceso_inusual_complejo_no 
                  Height          =   255
                  Left            =   4200
                  TabIndex        =   229
                  Top             =   0
                  Width           =   315
               End
               Begin VB.OptionButton OptProblemasHumanos_proceso_inusual_complejo_si 
                  Height          =   255
                  Left            =   3780
                  TabIndex        =   228
                  Top             =   0
                  Width           =   315
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "¿Es un proceso inusual o complejo?"
                  Height          =   195
                  Index           =   3
                  Left            =   30
                  TabIndex        =   230
                  Top             =   30
                  Width           =   2550
               End
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Sí  -  No"
               Height          =   195
               Left            =   9960
               TabIndex        =   253
               Top             =   90
               Width           =   600
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Sí  -  No"
               Height          =   195
               Left            =   5190
               TabIndex        =   252
               Top             =   90
               Width           =   600
            End
            Begin VB.Label lblCapCausas 
               Caption         =   "Problemas Humanos"
               Height          =   465
               Index           =   6
               Left            =   0
               TabIndex        =   251
               Top             =   450
               Width           =   1215
            End
         End
         Begin VB.ListBox lstCausas 
            Appearance      =   0  'Flat
            Height          =   930
            Index           =   5
            Left            =   1380
            Style           =   1  'Checkbox
            TabIndex        =   225
            Top             =   4050
            Width           =   9495
         End
         Begin VB.ListBox lstCausas 
            Appearance      =   0  'Flat
            Height          =   930
            Index           =   4
            Left            =   1380
            Style           =   1  'Checkbox
            TabIndex        =   221
            Top             =   3090
            Width           =   9495
         End
         Begin VB.ListBox lstCausas 
            Appearance      =   0  'Flat
            Height          =   930
            Index           =   3
            Left            =   1380
            Style           =   1  'Checkbox
            TabIndex        =   220
            Top             =   2130
            Width           =   9495
         End
         Begin VB.ListBox lstCausas 
            Appearance      =   0  'Flat
            Height          =   930
            Index           =   2
            Left            =   1380
            Style           =   1  'Checkbox
            TabIndex        =   218
            Top             =   1170
            Width           =   9495
         End
         Begin VB.ListBox lstCausas 
            Appearance      =   0  'Flat
            Height          =   930
            Index           =   1
            Left            =   1380
            Style           =   1  'Checkbox
            TabIndex        =   215
            Top             =   210
            Width           =   9495
         End
         Begin VB.CommandButton cmdNuevaCausaProblema 
            Caption         =   "Nueva Causa Problema"
            Height          =   435
            Left            =   6180
            TabIndex        =   170
            Top             =   6030
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label lblCapCausas 
            Caption         =   "Problemas de Planificación"
            Height          =   495
            Index           =   5
            Left            =   90
            TabIndex        =   224
            Top             =   4350
            Width           =   1185
         End
         Begin VB.Label lblCapCausas 
            Caption         =   "Problemas de Aseguramiento de la Calidad"
            Height          =   585
            Index           =   4
            Left            =   90
            TabIndex        =   223
            Top             =   3270
            Width           =   1185
         End
         Begin VB.Label lblCapCausas 
            Caption         =   "Problemas de Producción"
            Height          =   585
            Index           =   2
            Left            =   120
            TabIndex        =   219
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblCapCausas 
            Caption         =   "Problemas de Equipamiento / Material"
            Height          =   585
            Index           =   1
            Left            =   90
            TabIndex        =   217
            Top             =   1350
            Width           =   1215
         End
         Begin VB.Label lblCapCausas 
            Caption         =   "Problemas de Requerimientos"
            Height          =   435
            Index           =   0
            Left            =   90
            TabIndex        =   216
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.TextBox txtResumenCausas 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   -74940
         MaxLength       =   65000
         MultiLine       =   -1  'True
         TabIndex        =   161
         Top             =   900
         Width           =   10995
      End
      Begin VB.Frame fraCausaDirecta 
         Caption         =   "Causa Directa"
         Height          =   1365
         Left            =   -74940
         TabIndex        =   158
         Top             =   2100
         Width           =   10995
         Begin VB.TextBox txtCausaDirecta 
            Appearance      =   0  'Flat
            Height          =   795
            Left            =   60
            MaxLength       =   65000
            TabIndex        =   159
            Top             =   510
            Width           =   10875
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "¿Que ha ocurrido?"
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   160
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame fraCausasContributibas 
         Caption         =   "Causas Contributivas"
         Height          =   1905
         Left            =   -74910
         TabIndex        =   136
         Top             =   3510
         Width           =   10965
         Begin TabDlg.SSTab tabCausasContributivas 
            Height          =   1575
            Left            =   60
            TabIndex        =   137
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2778
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Causa Contributiva 1ª"
            TabPicture(0)   =   "frmProcNC_Detalle.frx":4670
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblCampos(36)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblCampos(33)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "txtCC(1)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "txtCC_Desc(1)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Causa Contributiva 2ª"
            TabPicture(1)   =   "frmProcNC_Detalle.frx":468C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtCC_Desc(2)"
            Tab(1).Control(1)=   "txtCC(2)"
            Tab(1).Control(2)=   "lblCampos(34)"
            Tab(1).Control(3)=   "lblCampos(35)"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Causa Contributiva 3ª"
            TabPicture(2)   =   "frmProcNC_Detalle.frx":46A8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txtCC_Desc(3)"
            Tab(2).Control(1)=   "txtCC(3)"
            Tab(2).Control(2)=   "lblCampos(37)"
            Tab(2).Control(3)=   "lblCampos(38)"
            Tab(2).ControlCount=   4
            TabCaption(3)   =   "Causa Contributiva 4ª"
            TabPicture(3)   =   "frmProcNC_Detalle.frx":46C4
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txtCC_Desc(4)"
            Tab(3).Control(1)=   "txtCC(4)"
            Tab(3).Control(2)=   "lblCampos(39)"
            Tab(3).Control(3)=   "lblCampos(40)"
            Tab(3).ControlCount=   4
            TabCaption(4)   =   "Causa Contributiva 5ª"
            TabPicture(4)   =   "frmProcNC_Detalle.frx":46E0
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "txtCC_Desc(5)"
            Tab(4).Control(1)=   "txtCC(5)"
            Tab(4).Control(2)=   "lblCampos(41)"
            Tab(4).Control(3)=   "lblCampos(42)"
            Tab(4).ControlCount=   4
            Begin VB.TextBox txtCC_Desc 
               Appearance      =   0  'Flat
               Height          =   795
               Index           =   1
               Left            =   1830
               MultiLine       =   -1  'True
               TabIndex        =   147
               Top             =   720
               Width           =   8925
            End
            Begin VB.TextBox txtCC 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   2850
               TabIndex        =   146
               Top             =   390
               Width           =   7905
            End
            Begin VB.TextBox txtCC_Desc 
               Appearance      =   0  'Flat
               Height          =   795
               Index           =   2
               Left            =   -73170
               MultiLine       =   -1  'True
               TabIndex        =   145
               Top             =   720
               Width           =   8925
            End
            Begin VB.TextBox txtCC 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   -72150
               TabIndex        =   144
               Top             =   390
               Width           =   7905
            End
            Begin VB.TextBox txtCC_Desc 
               Appearance      =   0  'Flat
               Height          =   795
               Index           =   3
               Left            =   -73170
               MultiLine       =   -1  'True
               TabIndex        =   143
               Top             =   720
               Width           =   8925
            End
            Begin VB.TextBox txtCC 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   3
               Left            =   -72150
               TabIndex        =   142
               Top             =   390
               Width           =   7905
            End
            Begin VB.TextBox txtCC_Desc 
               Appearance      =   0  'Flat
               Height          =   795
               Index           =   4
               Left            =   -73170
               MultiLine       =   -1  'True
               TabIndex        =   141
               Top             =   720
               Width           =   8925
            End
            Begin VB.TextBox txtCC 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   4
               Left            =   -72150
               TabIndex        =   140
               Top             =   390
               Width           =   7905
            End
            Begin VB.TextBox txtCC_Desc 
               Appearance      =   0  'Flat
               Height          =   795
               Index           =   5
               Left            =   -73170
               MultiLine       =   -1  'True
               TabIndex        =   139
               Top             =   720
               Width           =   8925
            End
            Begin VB.TextBox txtCC 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   5
               Left            =   -72150
               TabIndex        =   138
               Top             =   390
               Width           =   7905
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "¿Por qué ha ocurrido?"
               Height          =   195
               Index           =   33
               Left            =   90
               TabIndex        =   157
               Top             =   750
               Width           =   1575
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion de la Causa Contributiva"
               Height          =   195
               Index           =   36
               Left            =   120
               TabIndex        =   156
               Top             =   450
               Width           =   2610
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "¿Por qué ha ocurrido?"
               Height          =   195
               Index           =   34
               Left            =   -74910
               TabIndex        =   155
               Top             =   750
               Width           =   1575
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion de la Causa Contributiva"
               Height          =   195
               Index           =   35
               Left            =   -74880
               TabIndex        =   154
               Top             =   450
               Width           =   2610
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "¿Por qué ha ocurrido?"
               Height          =   195
               Index           =   37
               Left            =   -74910
               TabIndex        =   153
               Top             =   750
               Width           =   1575
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion de la Causa Contributiva"
               Height          =   195
               Index           =   38
               Left            =   -74880
               TabIndex        =   152
               Top             =   450
               Width           =   2610
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "¿Por qué ha ocurrido?"
               Height          =   195
               Index           =   39
               Left            =   -74910
               TabIndex        =   151
               Top             =   750
               Width           =   1575
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion de la Causa Contributiva"
               Height          =   195
               Index           =   40
               Left            =   -74880
               TabIndex        =   150
               Top             =   450
               Width           =   2610
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "¿Por qué ha ocurrido?"
               Height          =   195
               Index           =   41
               Left            =   -74910
               TabIndex        =   149
               Top             =   750
               Width           =   1575
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion de la Causa Contributiva"
               Height          =   195
               Index           =   42
               Left            =   -74880
               TabIndex        =   148
               Top             =   450
               Width           =   2610
            End
         End
      End
      Begin VB.Frame fraCausaRaiz 
         Caption         =   "Causa Raiz"
         Height          =   1335
         Left            =   -74940
         TabIndex        =   133
         Top             =   5460
         Width           =   11025
         Begin VB.TextBox txtCausaRaiz 
            Appearance      =   0  'Flat
            Height          =   795
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   134
            Top             =   480
            Width           =   10905
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "¿Cual ha sido la causa raíz?"
            Height          =   195
            Index           =   43
            Left            =   90
            TabIndex        =   135
            Top             =   240
            Width           =   2010
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         Caption         =   "Resumen Acciones Correctivas"
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   90
         TabIndex        =   131
         Top             =   2220
         Width           =   10875
         Begin VB.CommandButton cmdEliminarAccionCorrectiva 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar Accion"
            Height          =   450
            Left            =   9180
            Style           =   1  'Graphical
            TabIndex        =   172
            Top             =   3990
            Width           =   1620
         End
         Begin VB.CommandButton cmdAnadirAccionCorrectivas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir Acción"
            Height          =   450
            Left            =   7500
            Style           =   1  'Graphical
            TabIndex        =   171
            Top             =   3990
            Width           =   1620
         End
         Begin MSFlexGridLib.MSFlexGrid lstAccionesCorrectivas 
            Height          =   3795
            Left            =   60
            TabIndex        =   132
            Top             =   180
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6694
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
      End
      Begin VB.Frame fraResumenAccionesCorrectivas 
         Appearance      =   0  'Flat
         Caption         =   "Resumen Acciones Inmediatas"
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   90
         TabIndex        =   129
         Top             =   660
         Width           =   10935
         Begin MSFlexGridLib.MSFlexGrid lstResumenAccionesInmediatas 
            Height          =   1365
            Left            =   60
            TabIndex        =   130
            Top             =   180
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   2408
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
      End
      Begin VB.TextBox txtRecoleccionDatos 
         Appearance      =   0  'Flat
         Height          =   3315
         Left            =   -74940
         MaxLength       =   65000
         MultiLine       =   -1  'True
         TabIndex        =   116
         Top             =   900
         Width           =   10995
      End
      Begin VB.Frame Frame2 
         Caption         =   "Análisis de la Escena"
         Height          =   6105
         Left            =   -74940
         TabIndex        =   92
         Top             =   690
         Width           =   10995
         Begin VB.ComboBox cmbAnalisisEscenaMinutos 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   2280
            Width           =   720
         End
         Begin VB.ComboBox cmbAnalisisEscenaHora 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmProcNC_Detalle.frx":46FC
            Left            =   4020
            List            =   "frmProcNC_Detalle.frx":46FE
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   2280
            Width           =   720
         End
         Begin VB.CommandButton cmdEliminarPersonalImplicado 
            Caption         =   "Eliminar"
            Height          =   315
            Left            =   9750
            TabIndex        =   207
            Top             =   570
            Width           =   1125
         End
         Begin VB.TextBox txtAnalisisEscenaFormacion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   115
            Top             =   5250
            Width           =   8475
         End
         Begin VB.TextBox txtAnalisisEscenaCambiosRecientes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   113
            Top             =   4830
            Width           =   8475
         End
         Begin VB.TextBox txtAnalisisEscenaEquiposImplicados 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   111
            Top             =   4410
            Width           =   8475
         End
         Begin VB.TextBox txtAnalisisEscenaSecuencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   109
            Top             =   3990
            Width           =   8475
         End
         Begin VB.ComboBox cmbPersonalImplicadoIncidencia 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3870
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   570
            Width           =   4545
         End
         Begin VB.CommandButton cmdAnadirPersonalImplicadoIncidencia 
            Caption         =   "Añadir"
            Height          =   315
            Left            =   8520
            TabIndex        =   104
            Top             =   570
            Width           =   1125
         End
         Begin VB.TextBox txtAnalisisEscenaComunicacion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   103
            Top             =   3570
            Width           =   8475
         End
         Begin VB.TextBox txtAnalisisEscenaCondicionesAmbientales 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   101
            Top             =   3150
            Width           =   8475
         End
         Begin VB.TextBox txtAnalisisEscenaCondicionesOperacion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            TabIndex        =   99
            Top             =   2730
            Width           =   8475
         End
         Begin VB.TextBox txtAnalisisEscenaLocalizacion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1410
            MaxLength       =   65000
            TabIndex        =   96
            Top             =   240
            Width           =   9465
         End
         Begin VB.TextBox txtAnalisisEscenaExperiencia 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2460
            MaxLength       =   65000
            MultiLine       =   -1  'True
            TabIndex        =   93
            Top             =   5670
            Width           =   8475
         End
         Begin MSComCtl2.DTPicker txtAnalisisEscenaFecha 
            Height          =   300
            Left            =   2460
            TabIndex        =   107
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
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
            Format          =   76087297
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSFlexGridLib.MSFlexGrid lstPersonalImplicado 
            Height          =   1365
            Left            =   90
            TabIndex        =   127
            Top             =   900
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2408
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "10.- Formación"
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   114
            Top             =   5280
            Width           =   1050
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "9.- Cambios Recientes"
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   112
            Top             =   4860
            Width           =   1590
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "8.- Equipos Implicados"
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   110
            Top             =   4440
            Width           =   1590
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "7.- Secuencia"
            Height          =   195
            Index           =   28
            Left            =   120
            TabIndex        =   108
            Top             =   4020
            Width           =   990
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "3.- Fecha y Hora"
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   106
            Top             =   2340
            Width           =   1185
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "6.- Comunicación"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   102
            Top             =   3600
            Width           =   1230
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "5.- Condiciones Ambientales"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   100
            Top             =   3180
            Width           =   1995
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "4.- Condiciones de Operación"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   98
            Top             =   2760
            Width           =   2100
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "2.- Nombre del Personal implicado en la incidencia"
            Height          =   195
            Index           =   21
            Left            =   150
            TabIndex        =   97
            Top             =   630
            Width           =   3555
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "1.- Localización"
            Height          =   195
            Index           =   26
            Left            =   150
            TabIndex        =   95
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            Caption         =   "11.- Experiencia"
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   94
            Top             =   5700
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdAsociarDocumentacionIdentificacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asociar Documentación"
         Height          =   870
         Left            =   -65850
         Picture         =   "frmProcNC_Detalle.frx":4700
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   5250
         Width           =   1905
      End
      Begin VB.Frame fraIncidencia 
         Caption         =   "Incidencia"
         Height          =   6165
         Left            =   -74910
         TabIndex        =   84
         Top             =   630
         Width           =   10965
         Begin VB.TextBox txtResumen 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2430
            TabIndex        =   212
            Top             =   180
            Width           =   8415
         End
         Begin VB.CommandButton cmdAnadirAccionInmediata 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir Acción"
            Height          =   720
            Left            =   9810
            Style           =   1  'Graphical
            TabIndex        =   123
            Top             =   2790
            Width           =   1050
         End
         Begin VB.CommandButton cmdEliminarSelAccInmediatas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar Accion"
            Height          =   720
            Left            =   9780
            Style           =   1  'Graphical
            TabIndex        =   122
            Top             =   3540
            Width           =   1080
         End
         Begin VB.CommandButton cmdAsociarDocumentacionIncidencia 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Asociar Documentación"
            Height          =   870
            Left            =   8970
            Picture         =   "frmProcNC_Detalle.frx":4FCA
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   4590
            Width           =   1905
         End
         Begin VB.TextBox txtDescripcionIncidencia 
            Appearance      =   0  'Flat
            Height          =   1575
            Left            =   120
            MaxLength       =   65000
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   780
            Width           =   10725
         End
         Begin MSFlexGridLib.MSFlexGrid lstAccionesInmediatas 
            Height          =   1575
            Left            =   30
            TabIndex        =   124
            Top             =   2760
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin MSFlexGridLib.MSFlexGrid lstDocumentacionIncidencia 
            Height          =   1575
            Left            =   60
            TabIndex        =   125
            Top             =   4560
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Breve Sinopsis de la Incidencia"
            Height          =   195
            Left            =   120
            TabIndex        =   211
            Top             =   240
            Width           =   2220
         End
         Begin VB.Label Label12 
            Caption         =   $"frmProcNC_Detalle.frx":5894
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   1650
            TabIndex        =   118
            Top             =   2370
            Width           =   6450
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Documentación de la Incidencia"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   4350
            Width           =   2295
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Acciones Inmediatas"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   2370
            Width           =   1470
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion de la Incidencia"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   540
            Width           =   2010
         End
      End
      Begin VB.Frame fraEquipoDesignado 
         Caption         =   "Equipo Designado"
         Height          =   6075
         Left            =   -74880
         TabIndex        =   79
         Top             =   660
         Width           =   10905
         Begin VB.CommandButton cmdRetirarDelEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quitar Seleccionado"
            Height          =   870
            Left            =   6900
            Picture         =   "frmProcNC_Detalle.frx":592E
            Style           =   1  'Graphical
            TabIndex        =   206
            Top             =   1350
            Width           =   1905
         End
         Begin VB.CommandButton cmdAnadirAlEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir Al Equipo"
            Height          =   870
            Left            =   8880
            Picture         =   "frmProcNC_Detalle.frx":61F8
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   1350
            Width           =   1905
         End
         Begin VB.OptionButton optDepartamento 
            Caption         =   "Departamento"
            Height          =   195
            Left            =   210
            TabIndex        =   83
            Top             =   540
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton optGrupo 
            Caption         =   "Grupo de Personas"
            Height          =   195
            Left            =   210
            TabIndex        =   82
            Top             =   1020
            Width           =   1785
         End
         Begin VB.ComboBox cmbDptoEquipo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   450
            Width           =   8475
         End
         Begin VB.ComboBox cmbGrupoPersonas 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2310
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   930
            Visible         =   0   'False
            Width           =   6525
         End
         Begin MSFlexGridLib.MSFlexGrid lstEquipoHumano 
            Height          =   3765
            Left            =   60
            TabIndex        =   128
            Top             =   2250
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6641
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin VB.Label Label14 
            Caption         =   "En caso de Departamento, el Responsable de Dpto será el Responsable del Equipo"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   2310
            TabIndex        =   121
            Top             =   210
            Width           =   6525
         End
      End
      Begin VB.Frame fraDptoImplicados 
         Appearance      =   0  'Flat
         Caption         =   "Departamentos Implicados"
         ForeColor       =   &H80000008&
         Height          =   5475
         Left            =   -66660
         TabIndex        =   68
         Top             =   1290
         Width           =   2715
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Informática"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   9
            Left            =   210
            TabIndex        =   78
            Top             =   3360
            Width           =   1305
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "I + D"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   8
            Left            =   210
            TabIndex        =   77
            Top             =   2970
            Width           =   1125
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Logística"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   7
            Left            =   210
            TabIndex        =   76
            Top             =   2580
            Width           =   1185
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Metrología"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   6
            Left            =   210
            TabIndex        =   75
            Top             =   2190
            Width           =   1275
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Laboratorio Aeronáutico"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   210
            TabIndex        =   74
            Top             =   1800
            Width           =   2400
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Laboratorio Agroalimentario"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   210
            TabIndex        =   73
            Top             =   1410
            Width           =   2400
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Administración y RRHH"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   210
            TabIndex        =   72
            Top             =   1020
            Width           =   2400
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Recepción"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   210
            TabIndex        =   71
            Top             =   3780
            Width           =   1305
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Calidad"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   210
            TabIndex        =   70
            Top             =   630
            Width           =   1845
         End
         Begin VB.CheckBox chkDpto 
            Appearance      =   0  'Flat
            Caption         =   "Gerencia"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   210
            TabIndex        =   69
            Top             =   240
            Width           =   1845
         End
      End
      Begin VB.Frame fraResponsableAperturaIncidencia 
         Appearance      =   0  'Flat
         Caption         =   "Responsable Apertura Incidencia"
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   -74940
         TabIndex        =   63
         Top             =   630
         Width           =   11025
         Begin VB.ComboBox cmbDptoResponsableApertura 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6960
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label lblCapDpto 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   5850
            TabIndex        =   66
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblResponsableAperturaIncidencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1200
            TabIndex        =   65
            Top             =   270
            Width           =   4455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   330
            Width           =   930
         End
      End
      Begin VB.Frame fraOrigen 
         Appearance      =   0  'Flat
         Caption         =   "Origen de la Incidencia"
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   -74940
         TabIndex        =   42
         Top             =   1260
         Width           =   8205
         Begin VB.Frame fraOrigenOtros 
            Appearance      =   0  'Flat
            Caption         =   "Otros"
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   90
            TabIndex        =   60
            Top             =   3630
            Width           =   7995
            Begin VB.TextBox txtOtros 
               Appearance      =   0  'Flat
               Height          =   1215
               Left            =   90
               MaxLength       =   65000
               MultiLine       =   -1  'True
               TabIndex        =   62
               Top             =   480
               Width           =   7785
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Caption         =   "Otros (Indicar)"
               Height          =   315
               Index           =   15
               Left            =   120
               TabIndex        =   61
               Top             =   180
               Width           =   1395
            End
         End
         Begin VB.Frame fraOrigenMetrología 
            Appearance      =   0  'Flat
            Caption         =   "Metrología"
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   90
            TabIndex        =   57
            Top             =   2820
            Width           =   3825
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Equipo"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   120
               TabIndex        =   59
               Top             =   180
               Width           =   1395
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Calibración"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   120
               TabIndex        =   58
               Top             =   450
               Width           =   1305
            End
         End
         Begin VB.Frame fraOrigenCliente 
            Appearance      =   0  'Flat
            Caption         =   "Cliente"
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   4140
            TabIndex        =   54
            Top             =   2820
            Width           =   3945
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Prop. Mejora"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   120
               TabIndex        =   56
               Top             =   450
               Width           =   1305
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Reclamación"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   120
               TabIndex        =   55
               Top             =   180
               Width           =   1395
            End
         End
         Begin VB.Frame fraOrigenFalloTecnico 
            Appearance      =   0  'Flat
            Caption         =   "Fallo Técnico"
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   4140
            TabIndex        =   48
            Top             =   240
            Width           =   3945
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Técnica Analítica"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   120
               TabIndex        =   53
               Top             =   180
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Informe"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   52
               Top             =   450
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Intercomparativo"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   120
               TabIndex        =   51
               Top             =   720
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Control Interno"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   120
               TabIndex        =   50
               Top             =   990
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Despliege Requisitos Cliente"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   49
               Top             =   1260
               Width           =   2325
            End
         End
         Begin VB.Frame fraOrigen_SistemaCalidad 
            Appearance      =   0  'Flat
            Caption         =   "Sistema de Calidad"
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   90
            TabIndex        =   43
            Top             =   240
            Width           =   3825
            Begin VB.Frame fraSegAccionCorrectiva 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Enabled         =   0   'False
               Height          =   915
               Left            =   120
               TabIndex        =   200
               Top             =   1260
               Width           =   3585
               Begin VB.TextBox txtPncOrigen 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   270
                  Locked          =   -1  'True
                  TabIndex        =   202
                  Top             =   570
                  Width           =   3285
               End
               Begin VB.CheckBox chkOrigenNoConformidad 
                  Appearance      =   0  'Flat
                  Caption         =   "Seguimiento Acción Correctiva"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Index           =   5
                  Left            =   0
                  TabIndex        =   201
                  Top             =   0
                  Width           =   2535
               End
               Begin VB.Label lblPncOrigen 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº  No Conformidad"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   270
                  TabIndex        =   203
                  Top             =   330
                  Width           =   1410
               End
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Formación"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   47
               Top             =   990
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Revisión por Dirección"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   46
               Top             =   720
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Auditoria Externa"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   45
               Top             =   450
               Width           =   2265
            End
            Begin VB.CheckBox chkOrigenNoConformidad 
               Appearance      =   0  'Flat
               Caption         =   "Auditoria Interna"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   180
               Width           =   2265
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid lstDocumentacionIdentificacion 
         Height          =   1575
         Left            =   -74940
         TabIndex        =   126
         Top             =   5250
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin MSFlexGridLib.MSFlexGrid lstDocumentacionRecoleccionDatos 
         Height          =   2115
         Left            =   -74970
         TabIndex        =   165
         Top             =   4680
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   3731
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Line ctlLinea 
         BorderColor     =   &H00000000&
         BorderWidth     =   4
         X1              =   -66690
         X2              =   -66690
         Y1              =   1350
         Y2              =   6780
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Recopilacion de Información"
         Height          =   195
         Left            =   -74970
         TabIndex        =   166
         Top             =   4470
         Width           =   2025
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Resumen de las causas"
         Height          =   195
         Left            =   -74910
         TabIndex        =   162
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Recolección de Datos"
         Height          =   195
         Left            =   -74910
         TabIndex        =   117
         Top             =   660
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Recopilacion de Información"
         Height          =   195
         Left            =   -74940
         TabIndex        =   91
         Top             =   5040
         Width           =   2025
      End
   End
   Begin VB.Frame frmevaluacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evaluación e impacto"
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
      Height          =   1905
      Left            =   12750
      TabIndex        =   21
      Top             =   6450
      Width           =   10350
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1005
         Index           =   1
         Left            =   5580
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   810
         Width           =   4680
      End
      Begin VB.CheckBox chkimpacto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Impacto"
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
         Left            =   5580
         TabIndex        =   7
         Top             =   225
         Width           =   1725
      End
      Begin VB.CheckBox cnknoprocede 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Procedente"
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
         Height          =   330
         Left            =   1710
         TabIndex        =   6
         Top             =   225
         Width           =   2040
      End
      Begin VB.CheckBox chkevaluada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Evaluada"
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
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1005
         Index           =   5
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   810
         Width           =   5160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   5445
         X2              =   5445
         Y1              =   90
         Y2              =   1890
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción del Impacto"
         Height          =   195
         Index           =   9
         Left            =   5580
         TabIndex        =   26
         Top             =   585
         Width           =   1710
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción de la Evaluación"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   25
         Top             =   585
         Width           =   2070
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Número y fechas de la Incidencia"
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
      Height          =   1005
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   750
      Width           =   11310
      Begin VB.TextBox txtNumeroMovimientos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   420
         Left            =   4500
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   204
         Top             =   480
         Width           =   705
      End
      Begin VB.TextBox txtFechaUltMov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   420
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   40
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtFechaCierre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   420
         Left            =   9900
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   38
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   420
         Left            =   6450
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   37
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtNGeneral 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   420
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   11
         Top             =   480
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker txtFechaAlta 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   76087297
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Modif."
         Height          =   195
         Index           =   11
         Left            =   4500
         TabIndex        =   205
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Ult. Mov."
         Height          =   195
         Index           =   14
         Left            =   2970
         TabIndex        =   39
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   5
         Left            =   6480
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº General"
         Height          =   195
         Index           =   7
         Left            =   1590
         TabIndex        =   19
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Alta"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Cierre"
         Height          =   195
         Index           =   10
         Left            =   9930
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   870
      Left            =   8985
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10110
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8730
      Width           =   1050
   End
   Begin VB.Frame frmanalisis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Análisis"
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
      Height          =   1950
      Left            =   12840
      TabIndex        =   27
      Top             =   4350
      Width           =   10350
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Index           =   3
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   945
         Width           =   9000
      End
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   1125
         TabIndex        =   29
         Top             =   225
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSDataListLib.DataCombo cmbDepartamento 
         Height          =   315
         Left            =   6390
         TabIndex        =   30
         Top             =   225
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSDataListLib.DataCombo cmbAfectado 
         Height          =   315
         Left            =   1125
         TabIndex        =   33
         Top             =   585
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "An. Causas"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   1215
         Width           =   1860
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Afectado"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   34
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Top             =   315
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Departamento"
         Height          =   195
         Index           =   1
         Left            =   5265
         TabIndex        =   31
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Describa los datos de la incidencia, rellenando los siguientes campos."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   23
      Top             =   360
      Width           =   4920
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10740
      Picture         =   "frmProcNC_Detalle.frx":6AC2
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Incidencia"
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
      TabIndex        =   22
      Top             =   90
      Width           =   2220
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   -45
      Width           =   11325
   End
   Begin VB.Menu mnuNuevaCausaProblema 
      Caption         =   "TipoCausaProblema"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnu_prob_requerimiento 
         Caption         =   "De Requerimiento"
      End
      Begin VB.Menu mnu_prob_equipo_material 
         Caption         =   "De Equipamiento/Material"
      End
      Begin VB.Menu mnu_prob_produccion 
         Caption         =   "De Producción"
      End
      Begin VB.Menu mnu_prob_aseguramiento_calidad 
         Caption         =   "De Aseguramiento de la Calidad"
      End
      Begin VB.Menu mnu_prob_planificacion 
         Caption         =   "De Planificación"
      End
      Begin VB.Menu mnu_prob_humanos 
         Caption         =   "Humanos"
      End
   End
End
Attribute VB_Name = "frmProcNC_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public PK As Long
Private mvarPK As Long

Private mvarobjPnc As New clsProcNc
Private mvarenumTipoEdicion As enumTipoEdicion
Private intCont As Integer
Private mvarblnResultado As Boolean
Private objUsuarios As New clsUsuarios
Private mvarlngIdJefeEquipo As Long

Private mvarNivelAcceso As Integer
Private mvarTipoEdicionNivelAcceso As Integer
Private fso As New FileSystemObject

' 0 Para solo ver
' 1 para el creador (solo cuando está abierta)
' 2 para componentes del equipo
' 3 para jefes de equipo
' 4 para responsables de calidad

Private Enum GridCols
    ' General
    ID = 0
    ' Equipo Humano
    EQUIPOS_RESPONSABLE = 1
    EQUIPOS_NOMBRE = 2
    ' Acciones Inmediatas
    ACC_INMEDIATAS_DESC = 1
    ' documentacion adjunta
    DOC_NOMBRE = 1
    DOC_OBSERVACIONES = 2
    'Cuestiones Identificacion Problemas
    IDPROBLEMA_REQUERIDA = 1
    IDPROBLEMA_PREGUNTA = 2
    IDPROBLEMA_RESPUESTA = 3
    IDPROBLEMA_TIPORESPUESTA = 4
    ' Personal Implicado
    PERSONAL_NOMBRE = 1
    'PERSONAL_DEPARTAMENTO = 2
    ' ACCIONES CORRECTIVAS
    ACC_CORRECTIVAS_TITULO = 1
    ACC_CORRECTIVAS_ESTADO = 2
    ACC_CORRECTIVAS_RESPONSABLE = 3
    'ACC_CORRECTIVAS_RESPONSABLE = 5
    'ACC_CORRECTIVAS_FECHAPUESTAMARCHA = 7
    'ACC_CORRECTIVAS_FECHAFIN = 8
    'ACC_CORRECTIVAS_DIASAVISOPREVIO = 9
End Enum

Private Function ComprobarDatos(Optional ByVal RechazandoEstado As Boolean = False) As Boolean

    Dim strCad As String, lf As String, blnErr As Boolean
On Error GoTo ComprobarDatos_Error

    lf = vbCrLf & " - "
    ComprobarDatos = True
    blnErr = False
    
    'strCad = ComprobarDatosTab(0) & ComprobarDatosTab(1)
    If cmbDptoResponsableApertura.ListIndex < 0 Then
        strCad = lf & "Debe señalar el departamento del que es Responsable"
        blnErr = True
    End If
    
    
    If mvarobjPnc.getESTADO_ID > C_PROCNC_ESTADOS.ABIERTA And mvarlngIdJefeEquipo <= 0 Then
        If Not RechazandoEstado Then
            strCad = lf & "Debe señalar el personal Jefe de Equipo encargado de la Investigacion de la Incidencia"
            blnErr = True
        End If
    End If
    
    
    If blnErr Then
        If blnErr Then MsgBox "Por favor, verifique los siguientes errores:" & strCad, vbInformation, "Guardar Incidencia"
        ComprobarDatos = Not blnErr
        Exit Function
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ComprobarDatos"
    Exit Function
ComprobarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ComprobarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function
Private Function ComprobarDatosTab(ByVal pestana As Integer) As String
Dim intCont As Integer
Dim strres As String, blnOk As Boolean

On Error GoTo ComprobarDatosTab_Error

strres = ""
Select Case pestana
    Case 0
        blnOk = False
        For intCont = 1 To 15
            If chkOrigenNoConformidad(intCont).value = vbChecked Then
                blnOk = True
            End If
        Next intCont
        
        If Not blnOk Then
            strres = strres & vbCrLf & " - Debe indicar al menos un Origen de No Conformidad."
        End If
        ' verifica que si se elige Otros tenga texto.
        If chkOrigenNoConformidad(15).value = vbChecked And Trim(txtOtros.Text) = "" Then
            strres = strres & vbCrLf & " - Si se indica 'Otros' como Origen de No Conformidad, debe especificarlo."
        End If
        
        ' Departamentos Implicados
        blnOk = False
        For intCont = 1 To 10
            If chkDpto(intCont).value = vbChecked Then
                blnOk = True
            End If
        Next intCont
        If Not blnOk Then
            strres = strres & vbCrLf & " - Debe indicar al menos un Departamento Implicado."
        End If
        
        
        ComprobarDatosTab = strres
    Case 1
        ComprobarDatosTab = ""
    Case Else
        
End Select

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ComprobarDatosTab"
    Exit Function
ComprobarDatosTab_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ComprobarDatosTab"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ComprobarDatosTab of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function

Private Sub Configurar_lista_Acciones_inmediatas()

On Error GoTo Configurar_lista_Acciones_inmediatas_Error

With lstAccionesInmediatas
    .Cols = 2
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.ACC_INMEDIATAS_DESC) = .Width * 0.99
    
    .TextMatrix(0, GridCols.ACC_INMEDIATAS_DESC) = "Acción Inmediata"
    .Rows = 1
End With

With lstResumenAccionesInmediatas
    .Cols = 2
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.ACC_INMEDIATAS_DESC) = .Width * 0.99
    
    .TextMatrix(0, GridCols.ACC_INMEDIATAS_DESC) = "Acción Inmediata"
    .Rows = 1
End With


With lstAccionesCorrectivas
    .Cols = 4
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.ACC_CORRECTIVAS_ESTADO) = .Width * 0.19
    .ColWidth(GridCols.ACC_CORRECTIVAS_TITULO) = .Width * 0.4
    .ColWidth(GridCols.ACC_CORRECTIVAS_RESPONSABLE) = .Width * 0.4
    
    
    .TextMatrix(0, GridCols.ACC_CORRECTIVAS_ESTADO) = "Estado"
    .TextMatrix(0, GridCols.ACC_CORRECTIVAS_TITULO) = "Acción Correctiva"
    .TextMatrix(0, GridCols.ACC_CORRECTIVAS_RESPONSABLE) = "Responsable Impl."
    
    .Rows = 1
End With



On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Acciones_inmediatas"
    Exit Sub
Configurar_lista_Acciones_inmediatas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Acciones_inmediatas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Configurar_lista_Acciones_inmediatas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub Configurar_lista_Cuestiones_Identificacion_Problema()

On Error GoTo Configurar_lista_Cuestiones_Identificacion_Problema_Error

    With lstCuestionesIdentProblema
        .Cols = 5
        .ColWidth(GridCols.ID) = 0
        .ColWidth(GridCols.IDPROBLEMA_REQUERIDA) = .Width * 0.1
        .ColWidth(GridCols.IDPROBLEMA_PREGUNTA) = .Width * 0.39
        .ColWidth(GridCols.IDPROBLEMA_RESPUESTA) = .Width * 0.49
        .ColWidth(GridCols.IDPROBLEMA_TIPORESPUESTA) = 0
        
        
        .TextMatrix(0, GridCols.IDPROBLEMA_REQUERIDA) = "¿Oblig?"
        .TextMatrix(0, GridCols.IDPROBLEMA_PREGUNTA) = "Pregunta"
        .TextMatrix(0, GridCols.IDPROBLEMA_RESPUESTA) = "Respuesta"
        .TextMatrix(0, GridCols.IDPROBLEMA_TIPORESPUESTA) = "Tipo Respuesta"
        .Rows = 1
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Cuestiones_Identificacion_Problema"
    Exit Sub
Configurar_lista_Cuestiones_Identificacion_Problema_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Cuestiones_Identificacion_Problema"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Configurar_lista_Cuestiones_Identificacion_Problema of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub Configurar_lista_Documentacion_Previa()

On Error GoTo Configurar_lista_Documentacion_Previa_Error

With lstDocumentacionIncidencia
    .Cols = 3
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.DOC_NOMBRE) = .Width * 0.4
    .ColWidth(GridCols.DOC_OBSERVACIONES) = .Width * 0.59
    
    .TextMatrix(0, GridCols.DOC_NOMBRE) = "Documento"
    .TextMatrix(0, GridCols.DOC_OBSERVACIONES) = "Observaciones"
    
    .Rows = 1
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Documentacion_Previa"
    Exit Sub
Configurar_lista_Documentacion_Previa_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Documentacion_Previa"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Configurar_lista_Documentacion_Previa of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub Configurar_lista_Personal_Implicado()

With lstPersonalImplicado
    .Cols = 3
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.PERSONAL_NOMBRE) = .Width * 0.99
    '.ColWidth(GridCols.PERSONAL_DEPARTAMENTO) = .Width * 0.39
    
    .TextMatrix(0, GridCols.PERSONAL_NOMBRE) = "Nombre y Apellidos"
    '.TextMatrix(0, GridCols.PERSONAL_DEPARTAMENTO) = "Departamento"
    
    .Rows = 1
End With

End Sub
Private Sub Configurar_lista_Documentacion_Identificacion_Problema()

On Error GoTo Configurar_lista_Documentacion_Identificacion_Problema_Error

With lstDocumentacionIdentificacion
    .Cols = 3
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.DOC_NOMBRE) = .Width * 0.4
    .ColWidth(GridCols.DOC_OBSERVACIONES) = .Width * 0.59
    
    .TextMatrix(0, GridCols.DOC_NOMBRE) = "Documento"
    .TextMatrix(0, GridCols.DOC_OBSERVACIONES) = "Observaciones"
    
    .Rows = 1
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Documentacion_Identificacion_Problema"
    Exit Sub
Configurar_lista_Documentacion_Identificacion_Problema_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Documentacion_Identificacion_Problema"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Configurar_lista_Documentacion_Identificacion_Problema of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub Configurar_lista_Documentacion_Recoleccion_Datos()

On Error GoTo Configurar_lista_Documentacion_Recoleccion_Datos_Error

With lstDocumentacionRecoleccionDatos
    .Cols = 3
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.DOC_NOMBRE) = .Width * 0.4
    .ColWidth(GridCols.DOC_OBSERVACIONES) = .Width * 0.59
    
    .TextMatrix(0, GridCols.DOC_NOMBRE) = "Documento"
    .TextMatrix(0, GridCols.DOC_OBSERVACIONES) = "Observaciones"
    
    .Rows = 1
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Documentacion_Recoleccion_Datos"
    Exit Sub
Configurar_lista_Documentacion_Recoleccion_Datos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Configurar_lista_Documentacion_Recoleccion_Datos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Configurar_lista_Documentacion_Recoleccion_Datos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub ConfigurarCabecerasListas()

On Error GoTo ConfigurarCabecerasListas_Error

    Call configurar_lista_equipos
    Call Configurar_lista_Acciones_inmediatas
    Call Configurar_lista_Documentacion_Previa
    Call Configurar_lista_Cuestiones_Identificacion_Problema
    Call Configurar_lista_Documentacion_Identificacion_Problema
    Call Configurar_lista_Personal_Implicado
    Call Configurar_lista_Documentacion_Recoleccion_Datos
    'Call Configurar_lista_Causas

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfigurarCabecerasListas"
    Exit Sub
ConfigurarCabecerasListas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfigurarCabecerasListas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ConfigurarCabecerasListas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub configurarProcNC_PteConfirmacionCalidad()
On Error GoTo configurarProcNC_Abierto_Error

    For intCont = 8 To 9
        tabDatosNoConformidad.TabVisible(intCont) = False
    Next intCont
    
    
    ' Botones
    cmdCambiarEstado.Caption = "Sol. Plan Acc. Correct."
    cmdRechazarEstado.Visible = True
    cmdRechazarEstado.Caption = "Devolver a Tramitación"
    
    ' Texto de estado
    txtestado.Text = "P.N.C. Pte. Confirmación Calidad"

   On Error GoTo 0
   Exit Sub

configurarProcNC_Abierto_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_Abierto of Formulario frmProcNC_Detalle"

End Sub

Private Sub ConfigurarSegunEstado()

On Error GoTo ConfigurarSegunEstado_Error

If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
    Call configurarProcNC_Abierto
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
    Call configurarProcNC_PteVoBo
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
    Call configurarProcNC_EnTramitacion
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD Then
    Call configurarProcNC_PteConfirmacionCalidad
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
    Call configurarProcNC_PtePlanAccionesCorrectivas
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.pte_cierre Then
    Call configurarProcNC_PteCierre
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL Then
    Call configurarProcNC_CierreParcialEval
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA Then
    Call configurarProcNC_CierreTotal
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfigurarSegunEstado"
    Exit Sub
ConfigurarSegunEstado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfigurarSegunEstado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ConfigurarSegunEstado of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Public Property Get idProcNC() As Long

    idProcNC = mvarlngidProcNC

End Property

Public Property Let idProcNC(ByVal lngidProcNC As Long)

    mvarlngidProcNC = lngidProcNC

End Property



Private Sub configurar_lista_equipos()

' Jonathan.
' Si la incidencia no está en estado Pte.
' VºBº, no configura nada.


On Error GoTo configurar_lista_equipos_Error

With lstEquipoHumano
    .Cols = 3
    .ColWidth(GridCols.ID) = 0
    .ColWidth(GridCols.EQUIPOS_RESPONSABLE) = .Width * 0.1
    .ColWidth(GridCols.EQUIPOS_NOMBRE) = .Width * 0.88
    
    .TextMatrix(0, GridCols.EQUIPOS_RESPONSABLE) = "Responsable Equipo"
    .TextMatrix(0, GridCols.EQUIPOS_NOMBRE) = "Personal/Departamento"
    .Rows = 1
End With


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurar_lista_equipos"
    Exit Sub
configurar_lista_equipos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurar_lista_equipos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurar_lista_equipos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Public Property Let PK(ByVal value As Long)
    Dim objPnc As New clsProcNc
    
On Error GoTo PK_Error

    mvarenumTipoEdicion = EDICION
    objPnc.Carga value
    Set mvarobjPnc = objPnc
    
    Form_Load
    
    mvarblnResultado = True
    cmdImprimir.Enabled = True

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PK"
    Exit Property
PK_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PK"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PK of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
        
End Property

Public Property Get PK() As Long

End Property

Private Sub PresentarDatos_PncAbierto_AccionesInmediatas()
Dim obj As clsProcNcAccionInmediata


On Error GoTo PresentarDatos_PncAbierto_AccionesInmediatas_Error

With lstAccionesInmediatas
    .Rows = 1
    For Each obj In mvarobjPnc.AccionesInmediatas.Iterator
        If obj.getID_AUX <> -1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = obj.getID_ACCION_INMEDIATA
            .TextMatrix(.Rows - 1, GridCols.ACC_INMEDIATAS_DESC) = obj.getDESCRIPCION
        End If
    Next obj
End With

With lstResumenAccionesInmediatas
    .Rows = 1
    For Each obj In mvarobjPnc.AccionesInmediatas.Iterator
        If obj.getID_AUX <> -1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = obj.getID_ACCION_INMEDIATA
            .TextMatrix(.Rows - 1, GridCols.ACC_INMEDIATAS_DESC) = obj.getDESCRIPCION
        End If
    Next obj
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto_AccionesInmediatas"
    Exit Sub
PresentarDatos_PncAbierto_AccionesInmediatas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto_AccionesInmediatas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncAbierto_AccionesInmediatas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncTramitacion()
Dim objCol As New clsGenericCollection, objEq As clsGenericClass

' Presenta el equipo humano designado

On Error GoTo PresentarDatos_PncTramitacion_Error

Call PresentarDatos_PncTramitacion_EquipoHumano

Call PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas

Call PresentarDatos_PncTramitacion_PersonalImplicado

Call PresentarDatos_PncTramitacion_DatosAnalisisEscena

Call PresentarDatos_PncTramitacion_Causas

Call PresentarDatos_PncTramitacion_RecoleccionDatos


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion"
    Exit Sub
PresentarDatos_PncTramitacion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncPlanAccionesCorrectivas()
On Error GoTo PresentarDatos_PncPlanAccionesCorrectivas_Error

    Call PresentarDatos_PncTramitacion_AccionesCorrectivas

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncPlanAccionesCorrectivas"
    Exit Sub
PresentarDatos_PncPlanAccionesCorrectivas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncPlanAccionesCorrectivas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncPlanAccionesCorrectivas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub PresentarDatos_PorDefecto()
Dim intCont As Integer
    
On Error GoTo PresentarDatos_PorDefecto_Error

    With mvarobjPnc
        txtFechaAlta.value = .getFECHA_ALTA
        txtFechaUltMov.Text = Format(.getFECHA_ULT_MOVIMIENTO, "dd/mm/yyyy")
        If CLng(.getFECHA_CIERRE) <> 0 Then
            txtFechaCierre.Text = Format(.getFECHA_CIERRE, "dd/mm/yyyy")
        Else
            txtFechaCierre.Text = "--/--/----"
        End If
        txtNumeroMovimientos.Text = mvarobjPnc.getTOTAL_MOVIMIENTOS
    End With
    
    'Rellena datos de Fecha y Hora
    For intCont = 0 To 23
        Call cmbAnalisisEscenaHora.AddItem(Format(intCont, "00"), intCont)
    Next intCont
    
    For intCont = 0 To 59
        Call cmbAnalisisEscenaMinutos.AddItem(Format(intCont, "00"), intCont)
    Next intCont
        
    Call PresentarDatos_PorDefecto_ListadoCausas
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PorDefecto"
    Exit Sub
PresentarDatos_PorDefecto_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PorDefecto"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PorDefecto of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Public Property Set ProcNC(ByRef value As Object)
    Set mvarobjPnc = value
End Property

Public Property Get ProcNC() As Object
    Set ProcNC = mvarobjPnc
End Property

Public Property Let TipoEdicion(ByVal valor As Integer)
    mvarenumTipoEdicion = valor
End Property

Public Property Get TipoEdicion() As Integer
    TipoEdicion = mvarenumTipoEdicion
End Property


Private Sub cmbtipo_Change()
On Error GoTo cmbtipo_Change_Error

    If cmbTipo.Text <> "" Then
        Dim oNC As New clsNc
        txtDatos(4) = oNC.Calcular_Numero(CLng(cmbTipo.BoundText))
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmbtipo_Change"
    Exit Sub
cmbtipo_Change_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmbtipo_Change"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmbtipo_Change of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdAdjuntarRecoleccionDatos_Click()
    Dim objfrm As New frmProcNC_Adjuntos
    
On Error GoTo cmdAdjuntarRecoleccionDatos_Click_Error

    Set objfrm.ArchivosAdjuntos = mvarobjPnc.AdjuntosRecoleccionDatos
    objfrm.SubRuta = "RECOL_DATOS"
    objfrm.TipoEdicion = EDICION
    objfrm.Show vbModal
    
    Set mvarobjPnc.AdjuntosRecoleccionDatos = objfrm.ArchivosAdjuntos
    
    Unload objfrm
    Set objfrm = Nothing
    
    Call PresentarDatos_PncAbierto_AdjuntosRecoleccionDatos

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAdjuntarRecoleccionDatos_Click"
    Exit Sub
cmdAdjuntarRecoleccionDatos_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAdjuntarRecoleccionDatos_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarRecoleccionDatos_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdAnadirAccionCorrectivas_Click()
Dim objfrm As New frmProcNC_AccCorrectivas
Dim objCol As clsGenericCollection
Dim objAC As clsProcNcAccionCorrectora
Dim strId As String



On Error GoTo cmdAnadirAccionCorrectivas_Click_Error

Set objfrm.Pnc = mvarobjPnc
objfrm.TipoEdicion = ALTA
objfrm.Show vbModal

If objfrm.Resultado = True Then
    
    mvarobjPnc.cargar_AccionesCorrectivas
    PresentarDatos_PncPlanAccionesCorrectivas
    
End If

Unload objfrm
Set objfrm = Nothing
Set objCol = Nothing
Set objAC = Nothing

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirAccionCorrectivas_Click"
    Exit Sub
cmdAnadirAccionCorrectivas_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirAccionCorrectivas_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadirAccionCorrectivas_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdAnadirAccionInmediata_Click()
Dim objfrm As New frmProcNC_AccInmediatas
Dim objCol As clsGenericCollection
Dim objAI As clsProcNcAccionInmediata
Dim strId As String


On Error GoTo cmdAnadirAccionInmediata_Click_Error

objfrm.TipoEdicion = ALTA
objfrm.Show vbModal

If objfrm.Resultado = True Then
    
    Set objCol = mvarobjPnc.AccionesInmediatas
    Set objAI = objfrm.AccionInmediata
    strId = objCol.Add(objAI)
    
    Set mvarobjPnc.AccionesInmediatas = objCol
    
    With lstAccionesInmediatas
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = strId
        .TextMatrix(.Rows - 1, 1) = objAI.getDESCRIPCION
    End With
End If

Unload objfrm
Set objfrm = Nothing
Set objCol = Nothing
Set objAI = Nothing

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirAccionInmediata_Click"
    Exit Sub
cmdAnadirAccionInmediata_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirAccionInmediata_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadirAccionInmediata_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdAnadirAlEquipo_Click()

On Error GoTo cmdAnadirAlEquipo_Click_Error

If optDepartamento.value Then
    If Trim(cmbDptoEquipo.List(cmbDptoEquipo.ListIndex)) <> "" Then
        Call AnadirDepartamentoEquipo
    End If
Else
    If Trim(cmbGrupoPersonas.List(cmbGrupoPersonas.ListIndex)) <> "" Then
        Call AnadirPersonaEquipo
    End If
End If





On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirAlEquipo_Click"
    Exit Sub
cmdAnadirAlEquipo_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirAlEquipo_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadirAlEquipo_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdAnadirPersonalImplicadoIncidencia_Click()
Dim strID_EMPLEADO As String, strNOMBRE_EMPLEADO  As String
Dim ya_existentes As String
Dim intCont As Long

On Error GoTo cmdAnadirPersonalImplicadoIncidencia_Click_Error

blnYaExisteJefeEquipo = False
ya_existentes = ""

If cmbPersonalImplicadoIncidencia.ListIndex < 0 Then
    Exit Sub
ElseIf CLng(cmbPersonalImplicadoIncidencia.ItemData(cmbPersonalImplicadoIncidencia.ListIndex)) = 0 Then
    Exit Sub
End If
    

' Recoge los que ya hay
For intCont = 1 To lstPersonalImplicado.Rows - 1
    ya_existentes = ya_existentes & ";" & lstPersonalImplicado.TextMatrix(intCont, GridCols.ID)
Next intCont


'lstEquipoHumano.Rows = 1
strID_EMPLEADO = cmbPersonalImplicadoIncidencia.ItemData(cmbPersonalImplicadoIncidencia.ListIndex)
strNOMBRE_EMPLEADO = cmbPersonalImplicadoIncidencia.List(cmbPersonalImplicadoIncidencia.ListIndex)

If InStr(1, ya_existentes, ";" & strID_EMPLEADO) <= 0 Then
    lstPersonalImplicado.Rows = lstPersonalImplicado.Rows + 1
    lstPersonalImplicado.TextMatrix(lstPersonalImplicado.Rows - 1, GridCols.ID) = strID_EMPLEADO
    lstPersonalImplicado.TextMatrix(lstPersonalImplicado.Rows - 1, GridCols.PERSONAL_NOMBRE) = strNOMBRE_EMPLEADO
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirPersonalImplicadoIncidencia_Click"
    Exit Sub
cmdAnadirPersonalImplicadoIncidencia_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirPersonalImplicadoIncidencia_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadirPersonalImplicadoIncidencia_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdAnadirPregunta_Click()
Dim objfrm As New frmProcNC_CuestionesIdentProblema
Dim objPR As New clsProcNcPreguntaRespuesta

On Error GoTo cmdAnadirPregunta_Click_Error

    objfrm.TipoEdicion = enumTipoEdicion.ALTA
    Set objfrm.PreguntaRespuesta = objPR
    objfrm.Show vbModal
    
    If objfrm.Resultado = True Then
        'mvarlngIndexPreguntasRespuestas = mvarlngIndexPreguntasRespuestas + 1
        ' Añade la clase
        Set objPR = objfrm.PreguntaRespuesta
        'objPR.setID_PREGUNTA_RESPUESTA = mvarlngIndexPreguntasRespuestas
        'Call mvarobjPnc.PreguntasRespuestas.Add(objPR, CStr(mvarlngIndexPreguntasRespuestas))
        Call mvarobjPnc.PreguntasRespuestas.Add(objPR)
        '
        With lstCuestionesIdentProblema
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = CStr(objPR.getID_PREGUNTA_RESPUESTA)
            Call GridBooleanCell(lstCuestionesIdentProblema, .Rows - 1, GridCols.IDPROBLEMA_REQUERIDA, objPR.getREQUERIDA)
            .TextMatrix(.Rows - 1, GridCols.IDPROBLEMA_PREGUNTA) = objPR.getPREGUNTA
            If (objPR.getTIPO_PREGUNTA_RESPUESTA = RESP_SINO) Then
                Call GridBooleanCell(lstCuestionesIdentProblema, .Rows - 1, GridCols.IDPROBLEMA_RESPUESTA, (objPR.getRESPUESTA = 1))
            Else
                .TextMatrix(.Rows - 1, GridCols.IDPROBLEMA_RESPUESTA) = objPR.getRESPUESTA
            End If
            .TextMatrix(.Rows - 1, GridCols.IDPROBLEMA_TIPORESPUESTA) = objPR.getTIPO_PREGUNTA_RESPUESTA
        End With
    End If
    
    Unload objfrm
    Set objfrm = Nothing

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirPregunta_Click"
    Exit Sub
cmdAnadirPregunta_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAnadirPregunta_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadirPregunta_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdAsociarDocumentacionIdentificacion_Click()
    Dim objfrm As New frmProcNC_Adjuntos
    
On Error GoTo cmdAsociarDocumentacionIdentificacion_Click_Error

    Set objfrm.ArchivosAdjuntos = mvarobjPnc.AdjuntosIndentificacionProblemas
    objfrm.SubRuta = "IDENT_PROBLEMA"
    objfrm.TipoEdicion = EDICION
    objfrm.Show vbModal
    
    Set mvarobjPnc.AdjuntosIndentificacionProblemas = objfrm.ArchivosAdjuntos
    
    Unload objfrm
    Set objfrm = Nothing
    
    
    Call PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAsociarDocumentacionIdentificacion_Click"
    Exit Sub
cmdAsociarDocumentacionIdentificacion_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAsociarDocumentacionIdentificacion_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAsociarDocumentacionIdentificacion_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    End Sub

Private Sub cmdAsociarDocumentacionIncidencia_Click()
    Dim objfrm As New frmProcNC_Adjuntos
    
On Error GoTo cmdAsociarDocumentacionIncidencia_Click_Error

    Set objfrm.ArchivosAdjuntos = mvarobjPnc.AdjuntosIncidencia
    objfrm.SubRuta = "DOC_INCIDENCIA"
    objfrm.TipoEdicion = EDICION
    objfrm.Show vbModal
    
    Set mvarobjPnc.AdjuntosIncidencia = objfrm.ArchivosAdjuntos
    
    Unload objfrm
    Set objfrm = Nothing
    
    
    Call PresentarDatos_PncAbierto_AdjuntosIncidencia

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAsociarDocumentacionIncidencia_Click"
    Exit Sub
cmdAsociarDocumentacionIncidencia_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdAsociarDocumentacionIncidencia_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAsociarDocumentacionIncidencia_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub cambiarEstado_RechazarVoBo()

    Dim objfrm As New frmProcNC_CambioEstadoRazones
    Dim strMensajeRechazo As String
    
On Error GoTo cambiarEstado_RechazarVoBo_Error

    objfrm.TITULO = "Rechazar Vº Bº"
    
    objfrm.Show vbModal
    
    If Not objfrm.Resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    strMensajeRechazo = objfrm.MotivoRechazo
    
    If MsgBox("Para Proceder a enviar la Notificación de Rechazo del VºBº Incidencia al Responsable su apertura, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Rechazar VºBº P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos(True) Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.ABIERTA, strMensajeRechazo)
    
    mvarblnResultado = True
    Me.Hide
    
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cambiarEstado_RechazarVoBo"
    Exit Sub
cambiarEstado_RechazarVoBo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cambiarEstado_RechazarVoBo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cambiarEstado_RechazarVoBo of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub cambiarEstado_RechazarConfirmacionCalidad()

    Dim objfrm As New frmProcNC_CambioEstadoRazones
    Dim strMensajeRechazo As String
    
On Error GoTo cambiarEstado_RechazarConfirmacionCalidad_Error

    objfrm.TITULO = "Rechazar Confirmación Calidad"
    
    objfrm.Show vbModal
    
    If Not objfrm.Resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    strMensajeRechazo = objfrm.MotivoRechazo
    
    If MsgBox("Para Proceder a enviar la Notificación de Rechazo de la Confirmación de la Incidencia al Responsable su apertura, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Rechazar VºBº P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.EN_TRAMITACION, strMensajeRechazo)
    
    mvarblnResultado = True
    Me.Hide
    
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cambiarEstado_RechazarConfirmacionCalidad"
    Exit Sub
cambiarEstado_RechazarConfirmacionCalidad_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cambiarEstado_RechazarConfirmacionCalidad"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cambiarEstado_RechazarConfirmacionCalidad of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub cambiarEstado_RechazarPteCierre()

    Dim objfrm As New frmProcNC_CambioEstadoRazones
    Dim strMensajeRechazo As String
    
On Error GoTo cambiarEstado_RechazarPteCierre_Error

    objfrm.TITULO = "Rechazar Pendiente de Cierre"
    
    objfrm.Show vbModal
    
    If Not objfrm.Resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    strMensajeRechazo = objfrm.MotivoRechazo
    
    If MsgBox("Para Proceder a enviar la Notificación de Rechazo del VºBº Incidencia al Responsable su apertura, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Rechazar VºBº P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS, strMensajeRechazo)
    
    mvarblnResultado = True
    Me.Hide
    
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cambiarEstado_RechazarPteCierre"
    Exit Sub
cambiarEstado_RechazarPteCierre_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cambiarEstado_RechazarPteCierre"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cambiarEstado_RechazarPteCierre of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub




Private Sub cmdCambiarEstado_Click()

On Error GoTo cmdCambiarEstado_Click_Error

    If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
        Call CambiarEstado_a_VoBo
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
        Call CambiarEstado_a_Tramitacion
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
        Call CambiarEstado_a_PteConfirmacionCalidad
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD Then
        Call CambiarEstado_a_PtePlanAccionesCorrectivas
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
        Call CambiarEstado_a_PteCierre
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.pte_cierre Then
        Call CambiarEstado_a_CerradaParcial
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL Then
        Call CambiarEstado_a_CerradaTotal
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdCambiarEstado_Click"
    Exit Sub
cmdCambiarEstado_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdCambiarEstado_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdCambiarEstado_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdEliminarAccionCorrectiva_Click()
Dim intRow As Integer, intCol As Integer, fila As Long
Dim strId As String, objCol As clsGenericCollection
Dim objAC As New clsProcNcAccionCorrectora


On Error GoTo cmdEliminarAccionCorrectiva_Click_Error

With lstAccionesCorrectivas
    If .RowSel <= 0 Then Exit Sub
    fila = .RowSel
    strId = .TextMatrix(.RowSel, 0)
    
    objAC.Eliminar CLng(strId)
    
    Call mvarobjPnc.cargar_AccionesCorrectivas
    Call PresentarDatos_PncPlanAccionesCorrectivas
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarAccionCorrectiva_Click"
    Exit Sub
cmdEliminarAccionCorrectiva_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarAccionCorrectiva_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdEliminarAccionCorrectiva_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdEliminarPersonalImplicado_Click()
Dim fila As Long

On Error GoTo cmdEliminarPersonalImplicado_Click_Error

fila = lstPersonalImplicado.RowSel

If fila <= 0 Then Exit Sub

Call EliminarFilaGrid(lstPersonalImplicado, fila)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarPersonalImplicado_Click"
    Exit Sub
cmdEliminarPersonalImplicado_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarPersonalImplicado_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdEliminarPersonalImplicado_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdEliminarPregunta_Click()
Dim strId As String, fila As Long

On Error GoTo cmdEliminarPregunta_Click_Error

With lstCuestionesIdentProblema
    If .RowSel <= 0 Then Exit Sub
    fila = .RowSel
    strId = .TextMatrix(fila, GridCols.ID)
    
    If GridBooleanCell_Estado(lstCuestionesIdentProblema, .RowSel, GridCols.IDPROBLEMA_REQUERIDA) Then
        Call MsgBox("No puede Eliminar una pregunta obligatoria", vbInformation, "Eliminar Pregunta")
        Exit Sub
    End If
        
    Call EliminarFilaGrid(lstCuestionesIdentProblema, .RowSel)
    ' Lo elimina de la colección
    Call mvarobjPnc.PreguntasRespuestas.Remove(strId)
        
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarPregunta_Click"
    Exit Sub
cmdEliminarPregunta_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarPregunta_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdEliminarPregunta_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdEliminarSelAccInmediatas_Click()
Dim intRow As Integer, intCol As Integer, fila As Long
Dim strId As String, objCol As clsGenericCollection


On Error GoTo cmdEliminarSelAccInmediatas_Click_Error

With lstAccionesInmediatas
    If .RowSel <= 0 Then Exit Sub
    fila = .RowSel
    strId = .TextMatrix(.RowSel, 0)
    
    Call mvarobjPnc.AccionesInmediatas.Remove(strId)
    'Call objCol.Remove(strId)
    
    Call EliminarFilaGrid(lstAccionesInmediatas, fila)
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarSelAccInmediatas_Click"
    Exit Sub
cmdEliminarSelAccInmediatas_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdEliminarSelAccInmediatas_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdEliminarSelAccInmediatas_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub cmdImprimir_Click()
    
On Error GoTo cmdImprimir_Click_Error

    With frmReport
        .iniciar
        .informe = "/NC/rptProcNCCompleto"
        .CRITERIO = "{procnc.ID_PROCNC} = " & CStr(mvarobjPnc.getID_PROCNC) & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
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

    With frmReport
        .iniciar
        .informe = "/NC/rptProcNC"
        .CRITERIO = "{procnc.ID_PROCNC} = " & CStr(mvarobjPnc.getID_PROCNC) & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
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


Private Sub cmdNuevaCausaProblema_Click()
On Error GoTo cmdNuevaCausaProblema_Click_Error

PopupMenu mnuNuevaCausaProblema

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdNuevaCausaProblema_Click"
    Exit Sub
cmdNuevaCausaProblema_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdNuevaCausaProblema_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdNuevaCausaProblema_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdRechazarEstado_Click()

On Error GoTo cmdRechazarEstado_Click_Error

If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
    Call cambiarEstado_RechazarVoBo
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.pte_cierre Then
    Call cambiarEstado_RechazarPteCierre
ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD Then
    Call cambiarEstado_RechazarConfirmacionCalidad
    
'ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
'ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
'ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
'ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
'ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
End If


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdRechazarEstado_Click"
    Exit Sub
cmdRechazarEstado_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdRechazarEstado_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdRechazarEstado_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdRetirarDelEquipo_Click()

Dim fila As Long

On Error GoTo cmdRetirarDelEquipo_Click_Error

fila = lstEquipoHumano.RowSel

If fila <= 0 Then Exit Sub

If GridBooleanCell_Estado(lstEquipoHumano, fila, GridCols.EQUIPOS_RESPONSABLE) Then
    MsgBox "No puede retirar al Jefe de Equipo", vbInformation, "Quitar Usuario del Equipo"
    Exit Sub
End If

Call EliminarFilaGrid(lstEquipoHumano, fila)


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdRetirarDelEquipo_Click"
    Exit Sub
cmdRetirarDelEquipo_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdRetirarDelEquipo_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdRetirarDelEquipo_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub chkevaluada_Click()
    On Error Resume Next
    If chkevaluada.value = Checked Then
        txtDatos(5).Enabled = True
        txtDatos(5).BackColor = vbWhite
        txtDatos(5).SetFocus
    Else
        txtDatos(5).Enabled = False
        txtDatos(5) = ""
        txtDatos(5).BackColor = &HE0E0E0
    End If
End Sub

Private Sub chkimpacto_Click()
'    txtDatos(1).Enabled = chkimpacto.value
    On Error Resume Next
    If chkimpacto.value = Checked Then
        txtDatos(1).Enabled = True
        txtDatos(1).BackColor = vbWhite
        txtDatos(1).SetFocus
    Else
        txtDatos(1).Enabled = False
        txtDatos(1) = ""
        txtDatos(1).BackColor = &HE0E0E0
    End If
End Sub

Private Sub cmbestados_Change()
On Error GoTo cmbestados_Change_Error

    If cmbestados.BoundText = C_NC_ESTADOS.CERRADA Then
        fecha_cierre.Enabled = True
        fecha_cierre = Date
    Else
        fecha_cierre.Enabled = False
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmbestados_Change"
    Exit Sub
cmbestados_Change_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmbestados_Change"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmbestados_Change of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub cmddocumentacion_Click()
On Error GoTo cmddocumentacion_Click_Error

    frmProcNC_Adjuntos.PK = PK
    frmProcNC_Adjuntos.Show 1

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmddocumentacion_Click"
    Exit Sub
cmddocumentacion_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmddocumentacion_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmddocumentacion_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdok_Click()

    
    Dim objPnc As New clsProcNc
    
On Error GoTo cmdok_Click_Error

    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    If mvarenumTipoEdicion = ALTA Then
        Call mvarobjPnc.Insertar
        
    Else
        Call mvarobjPnc.Modificar
    End If
    
    MsgBox "El Procedimiento de No Conformidad se ha guardado Correctamente", vbInformation, "Guardar P.N.C."
    
    mvarenumTipoEdicion = EDICION
    objPnc.Carga mvarobjPnc.getID_PROCNC
    Set mvarobjPnc = objPnc
    
    Form_Load
    
    mvarblnResultado = True
    cmdImprimir.Enabled = True


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdok_Click"
    Exit Sub
cmdok_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdok_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdSalir_Click()
    'mvarblnResultado = False
On Error GoTo cmdSalir_Click_Error

    Me.Hide

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdSalir_Click"
    Exit Sub
cmdSalir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cmdSalir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdSalir_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    cargar_combos
    
    If mvarenumTipoEdicion = enumTipoEdicion.ALTA Then
        cmdImprimir.Enabled = False
        mvarobjPnc.setESTADO_ID = C_PROCNC_ESTADOS.ABIERTA
        mvarobjPnc.setRESPONSABLE_ID = USUARIO.getID_EMPLEADO
        mvarobjPnc.setFECHA_ALTA = Now
        mvarobjPnc.setFECHA_ULT_MOVIMIENTO = Now
        mvarobjPnc.setFECHA_CIERRE = 0
        mvarobjPnc.setRESPONSABLE_ID = USUARIO.getID_EMPLEADO
        mvarobjPnc.setRESPONSABLE_NOMBRE_APELLIDOS = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        mvarlngIdJefeEquipo = 0
        Call RellenarPreguntasIdentificacionProblemaInicial
    End If
    
    ' Configura las Listas
    Call ConfigurarCabecerasListas
    
    Call ConfiguraCuestionesIdentificacionProblema
        
    Call ConfigurarSegunEstado
    
    Call PresentarDatos

    Call OpcionesEdicion

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Form_QueryUnload_Error

    If UnloadMode = vbFormControlMenu Then Cancel = True

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Form_QueryUnload"
    Exit Sub
Form_QueryUnload_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Form_QueryUnload"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_QueryUnload of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Error

    Set fso = Nothing

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Form_Unload"
    Exit Sub
Form_Unload_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.Form_Unload"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub lstAccionesCorrectivas_DblClick()
Dim intCont As Integer
Dim objfrm As New frmProcNC_AccCorrectivas
Dim fila As Long, objAI As clsProcNcAccionCorrectora
Dim objCol As clsGenericCollection

On Error GoTo lstAccionesCorrectivas_DblClick_Error

With lstAccionesCorrectivas
    
    If .RowSel <= 0 Then Exit Sub
    fila = .RowSel
    Set objCol = mvarobjPnc.AccionesCorrectoras
    Set objAI = objCol.Item(.TextMatrix(fila, GridCols.ID))
    
    Set objfrm.Pnc = mvarobjPnc
    Set objfrm.AccionCorrectiva = objAI
    If mvarNivelAcceso >= 3 Then
        objfrm.TipoEdicion = EDICION
    Else
        objfrm.TipoEdicion = visualizar
    End If
    
    objfrm.Show vbModal
    
    If Not objfrm.Resultado Then Exit Sub
    
    Set objAI = objfrm.AccionCorrectiva
    
    Call objCol.Replace(objAI.getID_ACCION, objAI)
    .TextMatrix(fila, GridCols.ACC_CORRECTIVAS_ESTADO) = objAI.getESTADO
    .TextMatrix(fila, GridCols.ACC_CORRECTIVAS_RESPONSABLE) = objAI.getRESPONSABLE
    .TextMatrix(fila, GridCols.ACC_CORRECTIVAS_TITULO) = objAI.getTITULO
    
    Set mvarobjPnc.AccionesCorrectoras = objCol
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstAccionesCorrectivas_DblClick"
    Exit Sub
lstAccionesCorrectivas_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstAccionesCorrectivas_DblClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstAccionesCorrectivas_DblClick of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub lstAccionesInmediatas_DblClick()
Dim intCont As Integer
Dim objfrm As New frmProcNC_AccInmediatas
Dim fila As Long, objAI As clsProcNcAccionInmediata
Dim objCol As clsGenericCollection

On Error GoTo lstAccionesInmediatas_DblClick_Error

With lstAccionesInmediatas
    
    If .RowSel <= 0 Then Exit Sub
    fila = .RowSel
    Set objCol = mvarobjPnc.AccionesInmediatas
    Set objAI = objCol.Item(.TextMatrix(fila, GridCols.ID))
    
    Set objfrm.AccionInmediata = objAI
    If cmdAnadirAccionInmediata.Enabled Then
        objfrm.TipoEdicion = EDICION
    Else
        objfrm.TipoEdicion = visualizar
    End If
    
    objfrm.Show vbModal
    
    If Not objfrm.Resultado Then Exit Sub
    
    Set objAI = objfrm.AccionInmediata
    
    Call objCol.Replace(objAI.getID_ACCION_INMEDIATA, objAI)
    .TextMatrix(fila, GridCols.ACC_INMEDIATAS_DESC) = objAI.getDESCRIPCION
    
    Set mvarobjPnc.AccionesInmediatas = objCol
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstAccionesInmediatas_DblClick"
    Exit Sub
lstAccionesInmediatas_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstAccionesInmediatas_DblClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstAccionesInmediatas_DblClick of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub lstCuestionesIdentProblema_DblClick()
Dim objfrm As New frmProcNC_CuestionesIdentProblema
Dim lngid As Long, fila As Long
Dim objPR As New clsProcNcPreguntaRespuesta

On Error GoTo lstCuestionesIdentProblema_DblClick_Error

With lstCuestionesIdentProblema
    If .RowSel <= 0 Then Exit Sub
    fila = .RowSel
    lngid = .TextMatrix(fila, GridCols.ID)
    Set objPR = mvarobjPnc.PreguntasRespuestas.Item(CStr(lngid))
    Set objfrm.PreguntaRespuesta = objPR
    
    If mvarobjPnc.getESTADO_ID < C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL Then
        If cmdAnadirPregunta.Enabled Or mvarNivelAcceso >= 2 Then
            objfrm.TipoEdicion = EDICION
        Else
            objfrm.TipoEdicion = visualizar
        End If
    Else
        objfrm.TipoEdicion = visualizar
    End If
    
    objfrm.Show vbModal
    
    If objfrm.Resultado Then
        Set objPR = objfrm.PreguntaRespuesta
        'Call mvarobjPnc.PreguntasRespuestas.Remove(CStr(lngid))
        'Call mvarobjPnc.PreguntasRespuestas.Add(objPR, CStr(objPR.ID_PREGUNTA_RESPUESTA))
        ' Lo Reemplaza en la Colección
        Call mvarobjPnc.PreguntasRespuestas.Replace(CStr(lngid), objPR)
        
        Call GridBooleanCell(lstCuestionesIdentProblema, fila, GridCols.IDPROBLEMA_REQUERIDA, objPR.getREQUERIDA)
        .TextMatrix(fila, GridCols.IDPROBLEMA_PREGUNTA) = objPR.getPREGUNTA
        .TextMatrix(fila, GridCols.IDPROBLEMA_TIPORESPUESTA) = objPR.getTIPO_PREGUNTA_RESPUESTA
        If objPR.getTIPO_PREGUNTA_RESPUESTA = RESP_SINO Then
            Call GridBooleanCell(lstCuestionesIdentProblema, fila, GridCols.IDPROBLEMA_RESPUESTA, (objPR.getRESPUESTA = "1"))
        Else
            Call GridBooleanCell_cambiar_a_no_booleano(lstCuestionesIdentProblema, fila, GridCols.IDPROBLEMA_RESPUESTA, fila, GridCols.IDPROBLEMA_PREGUNTA)
            .TextMatrix(fila, GridCols.IDPROBLEMA_RESPUESTA) = objPR.getRESPUESTA
        End If
    End If
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstCuestionesIdentProblema_DblClick"
    Exit Sub
lstCuestionesIdentProblema_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstCuestionesIdentProblema_DblClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstCuestionesIdentProblema_DblClick of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub




Private Sub lstDocumentacionIdentificacion_DblClick()
Dim fila As Long
Dim destino As String
    
On Error GoTo lstDocumentacionIdentificacion_DblClick_Error

    With lstDocumentacionIdentificacion
        If .RowSel <= 0 Then Exit Sub
        fila = .RowSel
        Set objCol = mvarobjPnc.AdjuntosIndentificacionProblemas
        Set objAI = objCol.Item(.TextMatrix(fila, GridCols.ID))
    End With
    
    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
        destino = objAI.getRUTA_TEMPORAL
    Else
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\PROCNC\" & objAI.getPROCNC_ID & "\IDENT_PROBLEMA\" & objAI.getRUTA
    End If
        
        
On Error GoTo fallo
    
    If fso.FileExists(destino) Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If

Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstDocumentacionIdentificacion_DblClick"
    Exit Sub
lstDocumentacionIdentificacion_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstDocumentacionIdentificacion_DblClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstDocumentacionIdentificacion_DblClick of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub lstDocumentacionIncidencia_DblClick()
    Dim fila As Long
    Dim destino As String
    
On Error GoTo lstDocumentacionIncidencia_DblClick_Error

    With lstDocumentacionIncidencia
        If .RowSel <= 0 Then Exit Sub
        fila = .RowSel
        Set objCol = mvarobjPnc.AdjuntosIncidencia
        Set objAI = objCol.Item(.TextMatrix(fila, GridCols.ID))
    End With
    
    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
        destino = objAI.getRUTA_TEMPORAL
    Else
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\PROCNC\" & objAI.getPROCNC_ID & "\DOC_INCIDENCIA\" & objAI.getRUTA
    End If
           
    If fso.FileExists(destino) Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If
            

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstDocumentacionIncidencia_DblClick"
    Exit Sub
lstDocumentacionIncidencia_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstDocumentacionIncidencia_DblClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstDocumentacionIncidencia_DblClick of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub


Private Sub lstDocumentacionRecoleccionDatos_Click()
Dim fila As Long
Dim destino As String
    
On Error GoTo lstDocumentacionRecoleccionDatos_Click_Error

    With lstDocumentacionRecoleccionDatos
        If .RowSel <= 0 Then Exit Sub
        fila = .RowSel
        Set objCol = mvarobjPnc.AdjuntosRecoleccionDatos
        Set objAI = objCol.Item(.TextMatrix(fila, GridCols.ID))
    End With
    
    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
        destino = objAI.getRUTA_TEMPORAL
    Else
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\PROCNC\" & objAI.getPROCNC_ID & "\RECOL_DATOS\" & objAI.getRUTA
    End If
        
            
    If fso.FileExists(destino) Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If

Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstDocumentacionRecoleccionDatos_Click"
    Exit Sub
lstDocumentacionRecoleccionDatos_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstDocumentacionRecoleccionDatos_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstDocumentacionRecoleccionDatos_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub lstEquipoHumano_DblClick()

Dim fila As Long, intCont As Long
On Error GoTo lstEquipoHumano_DblClick_Error

fila = lstEquipoHumano.RowSel

If fila <= 0 Then Exit Sub

If Not GridBooleanCell_Estado(lstEquipoHumano, fila, GridCols.EQUIPOS_RESPONSABLE) Then
    Call GridBooleanCell(lstEquipoHumano, fila, GridCols.EQUIPOS_RESPONSABLE, True)
    mvarlngIdJefeEquipo = CLng(lstEquipoHumano.TextMatrix(fila, 0))
End If


For intCont = 1 To lstEquipoHumano.Rows - 1
    If intCont <> fila Then
        Call GridBooleanCell(lstEquipoHumano, intCont, GridCols.EQUIPOS_RESPONSABLE, False)
    End If
Next intCont

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstEquipoHumano_DblClick"
    Exit Sub
lstEquipoHumano_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.lstEquipoHumano_DblClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lstEquipoHumano_DblClick of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub optDepartamento_Click()

On Error GoTo optDepartamento_Click_Error

If optDepartamento.value Then
    cmbDptoEquipo.Visible = True
    cmbGrupoPersonas.Visible = False
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.optDepartamento_Click"
    Exit Sub
optDepartamento_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.optDepartamento_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure optDepartamento_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub optGrupo_Click()

On Error GoTo optGrupo_Click_Error

If optGrupo.value Then
    cmbDptoEquipo.Visible = False
    cmbGrupoPersonas.Visible = True
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.optGrupo_Click"
    Exit Sub
optGrupo_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.optGrupo_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure optGrupo_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub tabDatosNoConformidad_Click(PreviousTab As Integer)

On Error GoTo tabDatosNoConformidad_Click_Error

    If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
        If PreviousTab = 0 Then
            strres = ComprobarDatosTab(0)
            If Trim(strres) <> "" Then
                MsgBox "Se han encontrado los siguientes Errores: " & strres
                tabDatosNoConformidad.Tab = 0
            End If
        End If
    Else
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.tabDatosNoConformidad_Click"
    Exit Sub
tabDatosNoConformidad_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.tabDatosNoConformidad_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure tabDatosNoConformidad_Click of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
On Error GoTo txtdatos_GotFocus_Error

    txtDatos(Index).BackColor = &H80C0FF

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.txtdatos_GotFocus"
    Exit Sub
txtdatos_GotFocus_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.txtdatos_GotFocus"
    error_grave Err.Number & " (" & Err.Description & ") in procedure txtdatos_GotFocus of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
On Error GoTo txtdatos_LostFocus_Error

    txtDatos(Index).BackColor = vbWhite

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.txtdatos_LostFocus"
    Exit Sub
txtdatos_LostFocus_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.txtdatos_LostFocus"
    error_grave Err.Number & " (" & Err.Description & ") in procedure txtdatos_LostFocus of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Function validar() As Boolean
On Error GoTo validar_Error

    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una descripción.", vbExclamation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(10)) = "" Then
        MsgBox "Debe rellenar las acciones inmediatas.", vbExclamation, App.Title
        txtDatos(10).SetFocus
        validar = False
        Exit Function
    End If
    
    If cmborigen.BoundText = "" Then
        MsgBox "Debe asignar un origen.", vbExclamation, App.Title
        cmborigen.SetFocus
        validar = False
        Exit Function
    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe asignar un estado.", vbExclamation, App.Title
        cmbEstado.SetFocus
        validar = False
        Exit Function
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.validar"
    Exit Function
validar_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.validar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure validar of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    
On Error GoTo cargar_combos_Error

    oDeco.Cargar_ComboBox cmbDptoResponsableApertura, decodificadora.PROCNC_DEPARTAMENTOS
    cmbDptoResponsableApertura.RemoveItem (0)
    
    oDeco.Cargar_ComboBox cmbDptoEquipo, decodificadora.PROCNC_DEPARTAMENTOS
    Cargar_ComboBox cmbGrupoPersonas, objUsuarios
    Cargar_ComboBox cmbPersonalImplicadoIncidencia, objUsuarios

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cargar_combos"
    Exit Sub
cargar_combos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.cargar_combos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_combos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub
Private Sub enviar_mensaje(tipo As Integer)
    ' Enviar aviso
    Dim oMensaje As New clsMensajes
    Dim ASUNTO As String
    Dim texto As String
    Dim mens As Integer
On Error GoTo enviar_mensaje_Error

    With oMensaje
        If tipo = 1 Then
            ASUNTO = "Alta incidencia, nº: " & txtDatos(2)
            texto = texto & "El usuario " & cmbUsuario.Text & " ha dado de alta una incidencia. " & vbNewLine & vbNewLine
        Else
            ASUNTO = "Modificación incidencia, nº: " & txtDatos(2)
            texto = texto & "El usuario " & cmbUsuario.Text & " ha modificado una incidencia. " & vbNewLine & vbNewLine
        End If
        texto = texto & "Fecha de Alta : " & Format(fecha, "dd-mm-yyyy") & vbNewLine & vbNewLine
        texto = texto & "Descripción : " & txtDatos(0) & vbNewLine & vbNewLine
        texto = texto & "Acc.Inmediata : " & txtDatos(10) & vbNewLine & vbNewLine
        texto = texto & "Origen : " & cmborigen.Text & vbNewLine
        
        .setASUNTO = ASUNTO
        .setTEXTO = texto
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(fecha.value, "yyyy-mm-dd")
        .setFECHA_FIN = Format(fecha.value + 7, "yyyy-mm-dd")
        .setACCION = "frmProcNC_Detalle;" & Trim(txtDatos(2))
        mens = .Insertar
        If mens > 0 Then
            Dim omu As New clsMensajes_usuarios
            Dim i As Integer
            Dim usuarios As New clsUsuarios
            Dim rs As ADODB.RecordSet
            Set rs = usuarios.Listado
            If rs.RecordCount > 0 Then
                Do
                    If rs("PER_NC") = 1 And rs("ID_EMPLEADO") <> USUARIO.getID_EMPLEADO Then
                        omu.setEMPLEADO_ID = rs("ID_EMPLEADO")
                        omu.setMENSAJE_ID = mens
                        omu.Insertar
                    End If
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        End If
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.enviar_mensaje"
    Exit Sub
enviar_mensaje_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.enviar_mensaje"
    error_grave Err.Number & " (" & Err.Description & ") in procedure enviar_mensaje of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub txtPncOrigen_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub


Private Sub ConfiguraCuestionesIdentificacionProblema()
    Dim objPR As clsProcNcPreguntaRespuesta
    
On Error GoTo ConfiguraCuestionesIdentificacionProblema_Error

    With lstCuestionesIdentProblema
        .Rows = 1
        
        For Each objPR In mvarobjPnc.PreguntasRespuestas.Iterator
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, GridCols.ID) = objPR.getID_PREGUNTA_RESPUESTA
            Call GridBooleanCell(lstCuestionesIdentProblema, .Rows - 1, GridCols.IDPROBLEMA_REQUERIDA, objPR.getREQUERIDA)
            .TextMatrix(.Rows - 1, GridCols.IDPROBLEMA_PREGUNTA) = objPR.getPREGUNTA
            
            If objPR.getTIPO_PREGUNTA_RESPUESTA = enumTipoPreguntaRespuesta.RESP_SINO Then
                Call GridBooleanCell(lstCuestionesIdentProblema, .Rows - 1, GridCols.IDPROBLEMA_RESPUESTA, (objPR.getRESPUESTA = "1"))
            Else
                .TextMatrix(.Rows - 1, GridCols.IDPROBLEMA_RESPUESTA) = objPR.getRESPUESTA
            End If
            
            .TextMatrix(.Rows - 1, GridCols.IDPROBLEMA_TIPORESPUESTA) = objPR.getTIPO_PREGUNTA_RESPUESTA
        Next objPR
        
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfiguraCuestionesIdentificacionProblema"
    Exit Sub
ConfiguraCuestionesIdentificacionProblema_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfiguraCuestionesIdentificacionProblema"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ConfiguraCuestionesIdentificacionProblema of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub RellenarPreguntasIdentificacionProblemaInicial()
Dim objPR As New clsProcNcPreguntaRespuesta

On Error GoTo RellenarPreguntasIdentificacionProblemaInicial_Error

mvarlngIndexPreguntasRespuestas = 0

Set objPR = New clsProcNcPreguntaRespuesta
objPR.setRESPUESTA = ""
objPR.setID_AUX = -2
objPR.setREQUERIDA = True
objPR.setPREGUNTA = "¿Cuál es el Problema?"
objPR.setTIPO_PREGUNTA_RESPUESTA = RESP_ALFANUMERICO
Call mvarobjPnc.PreguntasRespuestas.Add(objPR)
'
Set objPR = New clsProcNcPreguntaRespuesta
objPR.setID_AUX = -2
objPR.setRESPUESTA = ""
objPR.setREQUERIDA = True
objPR.setPREGUNTA = "¿Se Entiende?"
objPR.setTIPO_PREGUNTA_RESPUESTA = RESP_SINO
Call mvarobjPnc.PreguntasRespuestas.Add(objPR)
'
Set objPR = New clsProcNcPreguntaRespuesta
objPR.setID_AUX = -2
objPR.setRESPUESTA = ""
objPR.setREQUERIDA = True
objPR.setPREGUNTA = "¿Hay más de uno?"
objPR.setTIPO_PREGUNTA_RESPUESTA = RESP_ALFANUMERICO
Call mvarobjPnc.PreguntasRespuestas.Add(objPR)
'
Set objPR = New clsProcNcPreguntaRespuesta
objPR.setID_AUX = -2
objPR.setRESPUESTA = ""
objPR.setREQUERIDA = True
objPR.setPREGUNTA = "¿Cúal es el alcance?"
objPR.setTIPO_PREGUNTA_RESPUESTA = RESP_ALFANUMERICO
Call mvarobjPnc.PreguntasRespuestas.Add(objPR)
'
Set objPR = New clsProcNcPreguntaRespuesta
objPR.setID_AUX = -2
objPR.setRESPUESTA = ""
objPR.setREQUERIDA = True
objPR.setPREGUNTA = "¿A qué y a quién afecta?"
objPR.setTIPO_PREGUNTA_RESPUESTA = RESP_ALFANUMERICO
Call mvarobjPnc.PreguntasRespuestas.Add(objPR)
'
Set objPR = New clsProcNcPreguntaRespuesta
objPR.setID_AUX = -2
objPR.setRESPUESTA = ""
objPR.setREQUERIDA = True
objPR.setID_PREGUNTA_RESPUESTA = 6
objPR.setPREGUNTA = "¿Cual es el impacto en el Laboratorio?"
objPR.setTIPO_PREGUNTA_RESPUESTA = RESP_ALFANUMERICO
Call mvarobjPnc.PreguntasRespuestas.Add(objPR, CStr(objPR.getID_PREGUNTA_RESPUESTA))

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.RellenarPreguntasIdentificacionProblemaInicial"
    Exit Sub
RellenarPreguntasIdentificacionProblemaInicial_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.RellenarPreguntasIdentificacionProblemaInicial"
    error_grave Err.Number & " (" & Err.Description & ") in procedure RellenarPreguntasIdentificacionProblemaInicial of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub configurarProcNC_Abierto()

On Error GoTo configurarProcNC_Abierto_Error

    For intCont = 2 To 9
        tabDatosNoConformidad.TabVisible(intCont) = False
    Next intCont
    
    ' Boton
    cmdCambiarEstado.Caption = "Solic. Vº Bº Calidad"
    
    ' Texto de estado
    txtestado.Text = "P.N.C. Abierto"

   On Error GoTo 0
   Exit Sub


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_Abierto"
    Exit Sub
configurarProcNC_Abierto_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_Abierto"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_Abierto of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos()


On Error GoTo PresentarDatos_Error

    Call PresentarDatos_PorDefecto
    
    ' Dependiendo del estado, presentará unos datos y otros

    ' Por Defecto, se presentan los datos cuando es abierto, eso siempre
    If mvarobjPnc.getESTADO_ID >= C_PROCNC_ESTADOS.ABIERTA Then
        Call PresentarDatos_PncAbierto
    End If

    If mvarobjPnc.getESTADO_ID >= C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
        Call PresentarDatos_PncTramitacion
    End If

    If mvarobjPnc.getESTADO_ID >= C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
        Call PresentarDatos_PncPlanAccionesCorrectivas
    End If


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos"
    Exit Sub
PresentarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncAbierto()

Dim objUsuario As New clsUsuarios
Dim obj As Object

    

    ' Para la pestaña 1
On Error GoTo PresentarDatos_PncAbierto_Error

    lblResponsableAperturaIncidencia.Caption = mvarobjPnc.getRESPONSABLE_NOMBRE_APELLIDOS
    
    ' Si es la misma persona la que abrió la incidencia que la que está editando, lo puede cambiar, siempre que esté abierto.
        
    Call objUsuario.CARGAR(mvarobjPnc.getRESPONSABLE_ID)

    ' Elimina los departamentos de los que un usuario no es responsable
    For intCont = cmbDptoResponsableApertura.ListCount - 1 To 0 Step -1
        If objUsuario.getRESPONSABLE_DEPARTAMENTOS(cmbDptoResponsableApertura.ItemData(intCont)) = 0 Then
            Call cmbDptoResponsableApertura.RemoveItem(intCont)
        End If
    Next intCont
        
    If mvarenumTipoEdicion = ALTA Then
        ' Al crear nuevo pone uno por defecto
        If cmbDptoResponsableApertura.ListCount > 0 Then
            cmbDptoResponsableApertura.ListIndex = 0
        End If
        txtNGeneral.Text = mvarobjPnc.NuevoIDTemporal
        ' al terminar aqui, como es un alta, sale y no establece nada más
        Exit Sub
    End If
    
    txtNGeneral.Text = mvarobjPnc.getID_PROCNC
    txtNumeroMovimientos.Text = mvarobjPnc.getTOTAL_MOVIMIENTOS
    'Al editar Elige de entre los que hay, el que se señaló al guardar
    For intCont = 0 To cmbDptoResponsableApertura.ListCount - 1
        If cmbDptoResponsableApertura.ItemData(intCont) = mvarobjPnc.getRESPONSABLE_ID_DEPARTAMENTO Then
            cmbDptoResponsableApertura.ListIndex = intCont
            Exit For
        End If
    Next intCont
    
    ' Datos sobre Origen de la Incidencia
    For Each obj In mvarobjPnc.OrigenesIncidencia.Iterator
        chkOrigenNoConformidad(obj.getID).value = vbChecked
        If obj.getID = 5 Then
            txtPncOrigen.Text = obj.getDESCRIPCION
        ElseIf obj.getID = 15 Then
            ' Las observaciones a cualquier de origenes de incidencia, van a la clase generica como Descripcion
            txtOtros.Text = mvarobjPnc.getORIGEN_OTROS
        End If
    Next obj
    
    ' Datos sobre Departamentos implicados
    For Each obj In mvarobjPnc.DepartamentosImplicados.Iterator
        chkDpto(obj.getID).value = vbChecked
    Next obj
    
    ' De la petaña 2.- Descripción y documentacion
    txtResumen.Text = mvarobjPnc.getRESUMEN
    txtDescripcionIncidencia.Text = mvarobjPnc.getDESCRIPCION_INCIDENCIA
    
    
    Call PresentarDatos_PncAbierto_AccionesInmediatas
    
    Call PresentarDatos_PncAbierto_AdjuntosIncidencia

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto"
    Exit Sub
PresentarDatos_PncAbierto_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncAbierto of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub OpcionesEdicion()
Dim rs As ADODB.RecordSet
Dim x As Integer
On Error GoTo OpcionesEdicion_Error

    If mvarobjPnc.getRESPONSABLE_ID <> USUARIO.getID_EMPLEADO Then
        cmbDptoResponsableApertura.Enabled = False
    End If
    
    If mvarobjPnc.getESTADO_ID <> C_PROCNC_ESTADOS.ABIERTA Then
        txtFechaAlta.Enabled = False
    End If
    
    'MsgBox USUARIO.getRESPONSABLE_DEPARTAMENTOS(5)

    ' Estable los permisos que tenga cada uno
    If USUARIO.getID_EMPLEADO = mvarobjPnc.getRESPONSABLE_ID Then
        If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then
            mvarNivelAcceso = 1
        Else
            mvarNivelAcceso = 0
        End If
    Else
        mvarNivelAcceso = 0
    End If
    
    ' Mira si es responsable miembro de algun equipo
    Set rs = datos_bd("SELECT * FROM procnc_equipohumano where id_procnc = " & mvarobjPnc.getID_PROCNC)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            If USUARIO.getID_EMPLEADO = CInt(rs("id_usuario")) Then
                mvarNivelAcceso = 2
                If CInt(rs("jefe_equipo")) = 1 Then
                    If mvarobjPnc.getESTADO_ID >= C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
                        mvarNivelAcceso = 3
                    Else
                        mvarNivelAcceso = 2
                    End If
                End If
            End If

            rs.MoveNext
        Wend
        
    End If
    
    
    ' En caso de ser Responsable de argún departamento, da el nivel 3
    For x = 3 To enumDPTO.TOTAL_DEPARTAMENTOS
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(x) = 1 Then
            ' es responsable de calidad. Tiene acceso a todo.
            mvarNivelAcceso = 3
        End If
    Next x
    
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(1) = 1 Or USUARIO.getRESPONSABLE_DEPARTAMENTOS(2) = 1 Then
        ' es responsable de calidad. Tiene acceso a todo.
        mvarNivelAcceso = 4
    End If
    

    
    
    Call ConfigurarNivelesAcceso

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.OpcionesEdicion"
    Exit Sub
OpcionesEdicion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.OpcionesEdicion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure OpcionesEdicion of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncAbierto_AdjuntosIncidencia()
Dim obj As clsProcNcAdjuntos

On Error GoTo PresentarDatos_PncAbierto_AdjuntosIncidencia_Error

With lstDocumentacionIncidencia
    .Rows = 1
    
    If mvarobjPnc.AdjuntosIncidencia.Count = 0 Then Exit Sub
    
    For Each obj In mvarobjPnc.AdjuntosIncidencia.Iterator
        If obj.getID_AUX <> -1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = obj.getID_ADJUNTO
            .TextMatrix(.Rows - 1, GridCols.DOC_NOMBRE) = obj.getRUTA
            .TextMatrix(.Rows - 1, GridCols.DOC_OBSERVACIONES) = obj.getOBSERVACIONES
        End If
    Next obj
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto_AdjuntosIncidencia"
    Exit Sub
PresentarDatos_PncAbierto_AdjuntosIncidencia_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto_AdjuntosIncidencia"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncAbierto_AdjuntosIncidencia of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas()
Dim obj As clsProcNcAdjuntos

On Error GoTo PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas_Error

With lstDocumentacionIdentificacion
    .Rows = 1
    
    If mvarobjPnc.AdjuntosIndentificacionProblemas.Count = 0 Then Exit Sub
    
    For Each obj In mvarobjPnc.AdjuntosIndentificacionProblemas.Iterator
        If obj.getID_AUX <> -1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = obj.getID_ADJUNTO
            .TextMatrix(.Rows - 1, GridCols.DOC_NOMBRE) = obj.getRUTA
            .TextMatrix(.Rows - 1, GridCols.DOC_OBSERVACIONES) = obj.getOBSERVACIONES
        End If
    Next obj
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas"
    Exit Sub
PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncAbierto_AdjuntosIdentificacionProblemas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncAbierto_AdjuntosRecoleccionDatos()
Dim obj As clsProcNcAdjuntos

With lstDocumentacionRecoleccionDatos
    .Rows = 1
    
    If mvarobjPnc.AdjuntosRecoleccionDatos.Count = 0 Then Exit Sub
    
    For Each obj In mvarobjPnc.AdjuntosRecoleccionDatos.Iterator
        If obj.getID_AUX <> -1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = obj.getID_ADJUNTO
            .TextMatrix(.Rows - 1, GridCols.DOC_NOMBRE) = obj.getRUTA
            .TextMatrix(.Rows - 1, GridCols.DOC_OBSERVACIONES) = obj.getOBSERVACIONES
        End If
    Next obj
End With

End Sub

Private Sub GuardarDatos()
Dim chk As CheckBox
Dim objCol As clsGenericCollection
Dim objItem As clsGenericClass

On Error GoTo GuardarDatos_Error

With mvarobjPnc
    .setRESPONSABLE_ID_DEPARTAMENTO = cmbDptoResponsableApertura.ItemData(cmbDptoResponsableApertura.ListIndex)
    ' El id del responsable, si se establece, es al principio.
    
    Set objCol = New clsGenericCollection
    objCol.KeyName = "setID"
    For Each chk In chkOrigenNoConformidad
        If chk.value = vbChecked Then
            Set objItem = New clsGenericClass
            objItem.setDESCRIPCION = chk.Caption
            Call objCol.Add(objItem, chk.Index)
        End If
    Next chk
    Set .OrigenesIncidencia = objCol
    
    Set objCol = New clsGenericCollection
    objCol.KeyName = "setID"
    For Each chk In chkDpto
        If chk.value = vbChecked Then
            Set objItem = New clsGenericClass
            objItem.setDESCRIPCION = chk.Caption
            Call objCol.Add(objItem, chk.Index)
        End If
        
    Next chk
    Set .DepartamentosImplicados = objCol
    
    .setORIGEN_OTROS = txtOtros.Text
    .setRESUMEN = txtResumen.Text
    .setDESCRIPCION_INCIDENCIA = txtDescripcionIncidencia.Text
    ' Las Acciones Inmediatas se establecen directamente en el objeto
    
    'De Momento se quita, pero hay que volver a ponerlo
    'If .getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Then Exit Sub
    
    Call GuardarDatos_EquipoHumano
    
    Call GuardarDatos_AnalisisEscena
    
    Call GuardarDatos_AnalisisEscena_PersonalImplicado
    
    Call GuardarDatos_Causas
        
    .setRECOLECCION_DATOS = txtRecoleccionDatos.Text
    
    Call GuardarDatosEvaluacion
        
End With



On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos"
    Exit Sub
GuardarDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub GuardarDatosEvaluacion()

On Error GoTo GuardarDatosEvaluacion_Error

    With mvarobjPnc
        If Not (optEvidencias(1).value Or optEvidencias(2).value) Then
            .setEVIDENCIAS_EN_PLAZO = -1
        Else
            .setEVIDENCIAS_EN_PLAZO = IIf(optEvidencias(1).value, 1, 0)
        End If
        
        If Not (optEvidencias(3).value Or optEvidencias(4).value) Then
            .setEVIDENCIAS_EFECTIVAS = -1
        Else
            .setEVIDENCIAS_EFECTIVAS = IIf(optEvidencias(3).value, 1, 0)
        End If
        
        If Not (optEvidencias(5).value Or optEvidencias(6).value) Then
            .setEVIDENCIAS_EVIDENCIAS = -1
        Else
            .setEVIDENCIAS_EVIDENCIAS = IIf(optEvidencias(5).value, 1, 0)
        End If
        
        If Not (optEvidencias(7).value Or optEvidencias(8).value) Then
            .setEVIDENCIAS_COMUNICADO_MODIFICACIONES = -1
        Else
            .setEVIDENCIAS_COMUNICADO_MODIFICACIONES = IIf(optEvidencias(7).value, 1, 0)
        End If
        
        If Not (optEval_res_incidencia.value Or optEval_res_nc.value) Then
            .setES_NO_CONFORMIDAD = -1
        Else
            .setES_NO_CONFORMIDAD = IIf(optEval_res_nc.value, 1, 0)
        End If
        
        If Not (optEval_res_si.value Or optEval_res_no.value) Then
            .setES_SOLUCION_ACEPTABLE = -1
        Else
            .setES_SOLUCION_ACEPTABLE = IIf(optEval_res_si.value, 1, 0)
        End If
        
        .setOBSERVACIONES_RESULTADO = txtObservaciones.Text
        
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatosEvaluacion"
    Exit Sub
GuardarDatosEvaluacion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatosEvaluacion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatosEvaluacion of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Public Property Get Resultado() As Boolean

    Resultado = mvarblnResultado

End Property

Public Property Let Resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub CambiarEstado_a_VoBo()
    
On Error GoTo CambiarEstado_a_VoBo_Error

    If MsgBox("Para Proceder a enviar la Incidencia al Responsable de Calidad (VºBº), se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Tramitar P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    If mvarenumTipoEdicion = ALTA Then
        Call mvarobjPnc.Insertar
    Else
        Call mvarobjPnc.Modificar
    End If
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.PTE_VISTO_BUENO)
    
    mvarblnResultado = True
    Me.Hide


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_VoBo"
    Exit Sub
CambiarEstado_a_VoBo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_VoBo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_VoBo of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub CambiarEstado_a_Tramitacion()
    
On Error GoTo CambiarEstado_a_Tramitacion_Error

    If MsgBox("Para Proceder a Tramitar la Incidencia , se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Tramitar P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.EN_TRAMITACION)
    
    mvarblnResultado = True
    Me.Hide


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_Tramitacion"
    Exit Sub
CambiarEstado_a_Tramitacion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_Tramitacion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_Tramitacion of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub CambiarEstado_a_PteConfirmacionCalidad()
    
On Error GoTo CambiarEstado_a_PteConfirmacionCalidad_Error

    If MsgBox("Para Proceder a Solicitar la Confirmacion por parte de Calidad de la Incidencia , se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Solicitar Confirmación P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD)
    
    mvarblnResultado = True
    Me.Hide


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_PteConfirmacionCalidad"
    Exit Sub
CambiarEstado_a_PteConfirmacionCalidad_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_PteConfirmacionCalidad"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_PteConfirmacionCalidad of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub AnadirPersonaEquipo()
Dim strID_EMPLEADO As String, strNOMBRE_EMPLEADO  As String
Dim ya_existentes As String
Dim intCont As Long

On Error GoTo AnadirPersonaEquipo_Error

blnYaExisteJefeEquipo = False
ya_existentes = ""

If cmbGrupoPersonas.ListIndex < 0 Then
    Exit Sub
ElseIf CLng(cmbGrupoPersonas.ItemData(cmbGrupoPersonas.ListIndex)) = 0 Then
    Exit Sub
End If
    

' Recoge los que ya hay
For intCont = 1 To lstEquipoHumano.Rows - 1
    ya_existentes = ya_existentes & ";" & lstEquipoHumano.TextMatrix(intCont, GridCols.ID)
    If GridBooleanCell_Estado(lstEquipoHumano, intCont, GridCols.EQUIPOS_RESPONSABLE) Then blnYaExisteJefeEquipo = True
Next intCont


'lstEquipoHumano.Rows = 1
strID_EMPLEADO = cmbGrupoPersonas.ItemData(cmbGrupoPersonas.ListIndex)
strNOMBRE_EMPLEADO = cmbGrupoPersonas.List(cmbGrupoPersonas.ListIndex)

If InStr(1, ya_existentes, ";" & strID_EMPLEADO) <= 0 Then
    lstEquipoHumano.Rows = lstEquipoHumano.Rows + 1
    lstEquipoHumano.TextMatrix(lstEquipoHumano.Rows - 1, GridCols.ID) = strID_EMPLEADO
    lstEquipoHumano.TextMatrix(lstEquipoHumano.Rows - 1, GridCols.EQUIPOS_NOMBRE) = strNOMBRE_EMPLEADO
    Call GridBooleanCell(lstEquipoHumano, lstEquipoHumano.Rows - 1, GridCols.EQUIPOS_RESPONSABLE, False)
    
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.AnadirPersonaEquipo"
    Exit Sub
AnadirPersonaEquipo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.AnadirPersonaEquipo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure AnadirPersonaEquipo of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub AnadirDepartamentoEquipo()
Dim colDep As clsGenericCollection
Dim rs As ADODB.RecordSet
Dim responsables_dpto As String
Dim blnYaExisteJefeEquipo As Boolean
Dim ya_existentes As String
Dim intCont As Long

On Error GoTo AnadirDepartamentoEquipo_Error

blnYaExisteJefeEquipo = False
ya_existentes = ""


If cmbDptoEquipo.ListIndex < 0 Then
    Exit Sub
ElseIf CLng(cmbDptoEquipo.ItemData(cmbDptoEquipo.ListIndex)) = 0 Then
    Exit Sub
End If
    


' Recoge los que ya hay
For intCont = 1 To lstEquipoHumano.Rows - 1
    ya_existentes = ya_existentes & ";" & lstEquipoHumano.TextMatrix(intCont, GridCols.ID)
    If GridBooleanCell_Estado(lstEquipoHumano, intCont, GridCols.EQUIPOS_RESPONSABLE) Then blnYaExisteJefeEquipo = True
Next intCont


Set rs = objUsuarios.Listado_por_departamento(cmbDptoEquipo.ItemData(cmbDptoEquipo.ListIndex))
responsables_dpto = objUsuarios.Listado_responsables_departamento(cmbDptoEquipo.ItemData(cmbDptoEquipo.ListIndex))

'lstEquipoHumano.Rows = 1
If rs.RecordCount <> 0 Then
    With rs
        .MoveFirst
        While Not .EOF
            If InStr(1, ya_existentes, ";" & CStr(!ID_EMPLEADO)) <= 0 Then
                lstEquipoHumano.Rows = lstEquipoHumano.Rows + 1
                lstEquipoHumano.TextMatrix(lstEquipoHumano.Rows - 1, GridCols.ID) = !ID_EMPLEADO
                lstEquipoHumano.TextMatrix(lstEquipoHumano.Rows - 1, GridCols.EQUIPOS_NOMBRE) = !nombre & " " & !APELLIDOS
                If (InStr(1, responsables_dpto, CStr(!ID_EMPLEADO)) > 0) And (Not blnYaExisteJefeEquipo) Then
                    Call GridBooleanCell(lstEquipoHumano, lstEquipoHumano.Rows - 1, GridCols.EQUIPOS_RESPONSABLE, True)
                    mvarlngIdJefeEquipo = !ID_EMPLEADO
                    blnYaExisteJefeEquipo = True
                Else
                    Call GridBooleanCell(lstEquipoHumano, lstEquipoHumano.Rows - 1, GridCols.EQUIPOS_RESPONSABLE, False)
                End If
            End If
            .MoveNext
        Wend
    End With
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.AnadirDepartamentoEquipo"
    Exit Sub
AnadirDepartamentoEquipo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.AnadirDepartamentoEquipo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure AnadirDepartamentoEquipo of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub PresentarDatos_PncTramitacion_EquipoHumano()
On Error GoTo PresentarDatos_PncTramitacion_EquipoHumano_Error

Set objCol = mvarobjPnc.EquipoHumano

If objCol.Count = 0 Then
    lstEquipoHumano.Rows = 1
    Exit Sub
End If

With lstEquipoHumano
    For Each objEq In objCol.Iterator
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, GridCols.ID) = objEq.getID
        .TextMatrix(.Rows - 1, GridCols.EQUIPOS_NOMBRE) = objEq.getDESCRIPCION
        If (CLng(Trim(objEq.getOBSERVACIONES)) = 1) Then
            mvarlngIdJefeEquipo = CLng(objEq.getID)
            Call GridBooleanCell(lstEquipoHumano, .Rows - 1, GridCols.EQUIPOS_RESPONSABLE, True)
        Else
            Call GridBooleanCell(lstEquipoHumano, .Rows - 1, GridCols.EQUIPOS_RESPONSABLE, False)
        End If
    Next objEq
End With


' si no están pendiente del visto bueno, o no es Responsable de Calidad, no se podrá modificar el tema
If (mvarobjPnc.getESTADO_ID <> C_PROCNC_ESTADOS.PTE_VISTO_BUENO) And (USUARIO.getRESPONSABLE_DEPARTAMENTOS(2) = True) Then
    cmdAnadirAlEquipo.Enabled = False
    cmdRetirarDelEquipo.Enabled = False
    optDepartamento.Enabled = False
    optGrupo.Enabled = False
End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_EquipoHumano"
    Exit Sub
PresentarDatos_PncTramitacion_EquipoHumano_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_EquipoHumano"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion_EquipoHumano of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncTramitacion_PersonalImplicado()
Dim objEq As clsGenericClass

On Error GoTo PresentarDatos_PncTramitacion_PersonalImplicado_Error

Set objCol = mvarobjPnc.PersonalImplicado


If objCol.Count = 0 Then
    lstPersonalImplicado.Rows = 1
    Exit Sub
End If

With lstPersonalImplicado
    For Each objEq In objCol.Iterator
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, GridCols.ID) = objEq.getID
        .TextMatrix(.Rows - 1, GridCols.PERSONAL_NOMBRE) = objEq.getDESCRIPCION
    Next objEq
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_PersonalImplicado"
    Exit Sub
PresentarDatos_PncTramitacion_PersonalImplicado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_PersonalImplicado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion_PersonalImplicado of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncTramitacion_DatosAnalisisEscena()

On Error GoTo PresentarDatos_PncTramitacion_DatosAnalisisEscena_Error

With mvarobjPnc
    txtAnalisisEscenaLocalizacion.Text = .getANALISIS_ESCENA_LOCALIZACION
    If CLng(.getANALISIS_ESCENA_FECHA) > 0 Then
        txtAnalisisEscenaFecha.value = .getANALISIS_ESCENA_FECHA
    End If
    cmbAnalisisEscenaHora.ListIndex = .getANALISIS_ESCENA_HORA
    cmbAnalisisEscenaMinutos.ListIndex = .getANALISIS_ESCENA_MINUTOS
    
    txtAnalisisEscenaCambiosRecientes.Text = .getANALISIS_ESCENA_CAMBIOS_RECIENTES
    txtAnalisisEscenaCondicionesAmbientales.Text = .getANALISIS_ESCENA_CONDICIONES_AMBIENTALES
    txtAnalisisEscenaCondicionesOperacion.Text = .getANALISIS_ESCENA_CONDICIONES_OPERACION
    txtAnalisisEscenaEquiposImplicados.Text = .getANALISIS_ESCENA_EQUIPOSIMPLICADOS
    txtAnalisisEscenaExperiencia.Text = .getANALISIS_ESCENA_EXPERIENCIA
    txtAnalisisEscenaFormacion.Text = .getANALISIS_ESCENA_FORMACION
    txtAnalisisEscenaSecuencia.Text = .getANALISIS_ESCENA_SECUENCIA
    txtAnalisisEscenaComunicacion.Text = .getANALISIS_ESCENA_COMUNICACION
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_DatosAnalisisEscena"
    Exit Sub
PresentarDatos_PncTramitacion_DatosAnalisisEscena_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_DatosAnalisisEscena"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion_DatosAnalisisEscena of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncTramitacion_RecoleccionDatos()

    
On Error GoTo PresentarDatos_PncTramitacion_RecoleccionDatos_Error

    txtRecoleccionDatos.Text = mvarobjPnc.getRECOLECCION_DATOS
    
    
    PresentarDatos_PncAbierto_AdjuntosRecoleccionDatos

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_RecoleccionDatos"
    Exit Sub
PresentarDatos_PncTramitacion_RecoleccionDatos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_RecoleccionDatos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion_RecoleccionDatos of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub GuardarDatos_EquipoHumano()

    Dim intCont As Long
    Dim objCol As New clsGenericCollection
    Dim objEq As clsGenericClass
On Error GoTo GuardarDatos_EquipoHumano_Error

    objCol.KeyName = "setID"
    
    For intCont = 1 To lstEquipoHumano.Rows - 1
        Set objEq = New clsGenericClass
        objEq.setID = lstEquipoHumano.TextMatrix(intCont, GridCols.ID)
        If GridBooleanCell_Estado(lstEquipoHumano, intCont, GridCols.EQUIPOS_RESPONSABLE) Then
            objEq.setOBSERVACIONES = "1"
        Else
            objEq.setOBSERVACIONES = "0"
        End If
        
        Call objCol.Add(objEq, objEq.getID)
        
    Next intCont

    Set mvarobjPnc.EquipoHumano = objCol

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_EquipoHumano"
    Exit Sub
GuardarDatos_EquipoHumano_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_EquipoHumano"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatos_EquipoHumano of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub GuardarDatos_AnalisisEscena_PersonalImplicado()

    Dim intCont As Long
    Dim objCol As New clsGenericCollection
    Dim objPersonal As clsGenericClass
    
On Error GoTo GuardarDatos_AnalisisEscena_PersonalImplicado_Error

    objCol.KeyName = "setID"
    
    For intCont = 1 To lstPersonalImplicado.Rows - 1
        Set objPersonal = New clsGenericClass
        objPersonal.setID = lstPersonalImplicado.TextMatrix(intCont, GridCols.ID)
        
        Call objCol.Add(objPersonal, objPersonal.getID)
        
    Next intCont

    Set mvarobjPnc.PersonalImplicado = objCol

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_AnalisisEscena_PersonalImplicado"
    Exit Sub
GuardarDatos_AnalisisEscena_PersonalImplicado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_AnalisisEscena_PersonalImplicado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatos_AnalisisEscena_PersonalImplicado of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub GuardarDatos_AnalisisEscena()
On Error GoTo GuardarDatos_AnalisisEscena_Error

With mvarobjPnc
    .setANALISIS_ESCENA_CAMBIOS_RECIENTES = txtAnalisisEscenaCambiosRecientes.Text
    .setANALISIS_ESCENA_COMUNICACION = txtAnalisisEscenaComunicacion.Text
    
    .setANALISIS_ESCENA_CONDICIONES_AMBIENTALES = txtAnalisisEscenaCondicionesAmbientales.Text
    .setANALISIS_ESCENA_CONDICIONES_OPERACION = txtAnalisisEscenaCondicionesOperacion.Text
    .setANALISIS_ESCENA_EQUIPOSIMPLICADOS = txtAnalisisEscenaEquiposImplicados.Text
    .setANALISIS_ESCENA_EXPERIENCIA = txtAnalisisEscenaExperiencia.Text
    .setANALISIS_ESCENA_FECHA = txtAnalisisEscenaFecha.value
    .setANALISIS_ESCENA_FORMACION = txtAnalisisEscenaFormacion.Text
    If cmbAnalisisEscenaHora.ListIndex >= 0 Then
        .setANALISIS_ESCENA_HORA = cmbAnalisisEscenaHora.ListIndex
    Else
        .setANALISIS_ESCENA_HORA = 0
    End If
    If cmbAnalisisEscenaMinutos.ListIndex >= 0 Then
        .setANALISIS_ESCENA_MINUTOS = cmbAnalisisEscenaMinutos.ListIndex
    Else
        .setANALISIS_ESCENA_MINUTOS = 0
    End If
    .setANALISIS_ESCENA_LOCALIZACION = txtAnalisisEscenaLocalizacion.Text
    .setANALISIS_ESCENA_SECUENCIA = txtAnalisisEscenaSecuencia.Text
    

End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_AnalisisEscena"
    Exit Sub
GuardarDatos_AnalisisEscena_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_AnalisisEscena"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatos_AnalisisEscena of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub PresentarDatos_PorDefecto_ListadoCausas()
Dim rs As New ADODB.RecordSet

On Error GoTo PresentarDatos_PorDefecto_ListadoCausas_Error

    Set rs = mvarobjPnc.Listado_Causas
    
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    
    While Not rs.EOF
        
        Call lstCausas(rs!id_tipocausa).AddItem(rs!DESCRIPCION)
        lstCausas(rs!id_tipocausa).ItemData(lstCausas(rs!id_tipocausa).ListCount - 1) = rs!id_causa
        
        rs.MoveNext
    Wend



On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PorDefecto_ListadoCausas"
    Exit Sub
PresentarDatos_PorDefecto_ListadoCausas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PorDefecto_ListadoCausas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PorDefecto_ListadoCausas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub PresentarDatos_PncTramitacion_Causas()

Dim intLista As Integer, intCont As Integer

On Error GoTo PresentarDatos_PncTramitacion_Causas_Error

For intLista = 1 To 5
    For intCount = 0 To lstCausas(intLista).ListCount - 1
        If Not mvarobjPnc.Causas.Item(CStr(lstCausas(intLista).ItemData(intCount))) Is Nothing Then
            lstCausas(intLista).Selected(intCount) = True
        End If
    Next intCount
Next intLista


With mvarobjPnc
    If .getPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS Then
        OptProblemasHumanos_herramientas_adecuadas_si.value = True
    Else
        OptProblemasHumanos_herramientas_adecuadas_no.value = True
    End If
    
    If .getPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS Then
        OptProblemasHumanos_instrucciones_incompletas_si.value = True
    Else
        OptProblemasHumanos_instrucciones_incompletas_no.value = True
    End If
    
    If .getPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE Then
        OptProblemasHumanos_objetivos_marcados_claramente_si.value = True
    Else
        OptProblemasHumanos_objetivos_marcados_claramente_no.value = True
    End If
    
    If .getPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO Then
        OptProblemasHumanos_operador_sustituido_si.value = True
    Else
        OptProblemasHumanos_operador_sustituido_no.value = True
    End If
    
    If .getPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO Then
        OptProblemasHumanos_proceso_inusual_complejo_si.value = True
    Else
        OptProblemasHumanos_proceso_inusual_complejo_no.value = True
    End If
    
    If .getPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION Then
        OptProblemasHumanos_formacion_suficiente_si.value = True
    Else
        OptProblemasHumanos_formacion_suficiente_no.value = True
    End If
    
    txtCausaRaiz.Text = .getCAUSA_RAIZ
    txtCausaDirecta.Text = .getCAUSA_DIRECTA
    txtResumenCausas.Text = .getRESUMEN_CAUSAS
    txtCC(1).Text = .getCAUSA_CONTRIBUTIVA_1
    txtCC_Desc(1).Text = .getCAUSA_CONTRIBUTIVA_1_DESCRIPCION
    txtCC(2).Text = .getCAUSA_CONTRIBUTIVA_2
    txtCC_Desc(2).Text = .getCAUSA_CONTRIBUTIVA_2_DESCRIPCION
    txtCC(3).Text = .getCAUSA_CONTRIBUTIVA_3
    txtCC_Desc(3).Text = .getCAUSA_CONTRIBUTIVA_3_DESCRIPCION
    txtCC(4).Text = .getCAUSA_CONTRIBUTIVA_4
    txtCC_Desc(4).Text = .getCAUSA_CONTRIBUTIVA_4_DESCRIPCION
    txtCC(5).Text = .getCAUSA_CONTRIBUTIVA_5
    txtCC_Desc(5).Text = .getCAUSA_CONTRIBUTIVA_5_DESCRIPCION
    
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_Causas"
    Exit Sub
PresentarDatos_PncTramitacion_Causas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_Causas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion_Causas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub GuardarDatos_Causas()
Dim intLista As Integer, intCont As Integer
Dim objCol As New clsGenericCollection, objItem As clsGenericClass
On Error GoTo GuardarDatos_Causas_Error

objCol.KeyName = "setID"

For intLista = 1 To 5
    For intCount = 0 To lstCausas(intLista).ListCount - 1
        If lstCausas(intLista).Selected(intCount) Then
        
            Set objItem = New clsGenericClass
            objItem.setID = CStr(lstCausas(intLista).ItemData(intCount))
            objItem.setOBSERVACIONES = CStr(intLista)
            objItem.setDESCRIPCION = CStr(lstCausas(intLista).List(intCount))
            Call objCol.Add(objItem, objItem.getID)
        End If
    Next intCount
Next intLista


Set mvarobjPnc.Causas = objCol



With mvarobjPnc
    .setPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS = OptProblemasHumanos_herramientas_adecuadas_si.value
    .setPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS = OptProblemasHumanos_instrucciones_incompletas_si.value
    .setPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE = OptProblemasHumanos_objetivos_marcados_claramente_si.value
    .setPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO = OptProblemasHumanos_operador_sustituido_si.value
    .setPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO = OptProblemasHumanos_proceso_inusual_complejo_si.value
    .setPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION = OptProblemasHumanos_formacion_suficiente_si.value
    
    .setCAUSA_RAIZ = txtCausaRaiz.Text
    .setCAUSA_DIRECTA = txtCausaDirecta.Text
    .setRESUMEN_CAUSAS = txtResumenCausas.Text
    .setCAUSA_CONTRIBUTIVA_1 = txtCC(1).Text
    .setCAUSA_CONTRIBUTIVA_1_DESCRIPCION = txtCC_Desc(1).Text
    .setCAUSA_CONTRIBUTIVA_2 = txtCC(2).Text
    .setCAUSA_CONTRIBUTIVA_2_DESCRIPCION = txtCC_Desc(2).Text
    .setCAUSA_CONTRIBUTIVA_3 = txtCC(3).Text
    .setCAUSA_CONTRIBUTIVA_3_DESCRIPCION = txtCC_Desc(3).Text
    .setCAUSA_CONTRIBUTIVA_4 = txtCC(4).Text
    .setCAUSA_CONTRIBUTIVA_4_DESCRIPCION = txtCC_Desc(4).Text
    .setCAUSA_CONTRIBUTIVA_5 = txtCC(5).Text
    .setCAUSA_CONTRIBUTIVA_5_DESCRIPCION = txtCC_Desc(5).Text
    
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_Causas"
    Exit Sub
GuardarDatos_Causas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.GuardarDatos_Causas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure GuardarDatos_Causas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub configurarProcNC_EnTramitacion()

On Error GoTo configurarProcNC_EnTramitacion_Error

    For intCont = 8 To 9
        tabDatosNoConformidad.TabVisible(intCont) = False
    Next intCont
    
    
    ' Botones
    cmdCambiarEstado.Caption = "Solicitar Confirmación P.N.C."
    cmdRechazarEstado.Visible = False
    'cmdRechazarEstado.Caption = "Rechazar Vº Bº"
    
    ' Texto de estado
    txtestado.Text = "P.N.C. En Tramitación"


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_EnTramitacion"
    Exit Sub
configurarProcNC_EnTramitacion_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_EnTramitacion"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_EnTramitacion of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub configurarProcNC_PteVoBo()

On Error GoTo configurarProcNC_PteVoBo_Error

    For intCont = 3 To 9
        tabDatosNoConformidad.TabVisible(intCont) = False
    Next intCont
    
    
    ' Botones
    cmdCambiarEstado.Caption = "Tramitar P.N.C."
    cmdRechazarEstado.Visible = True
    cmdRechazarEstado.Caption = "Rechazar Vº Bº"
    
    ' Texto de estado
    txtestado.Text = "P.N.C. Pte. Vº Bº"

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_PteVoBo"
    Exit Sub
configurarProcNC_PteVoBo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_PteVoBo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_PteVoBo of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub PresentarDatos_PncTramitacion_AccionesCorrectivas()

On Error GoTo PresentarDatos_PncTramitacion_AccionesCorrectivas_Error

With lstAccionesCorrectivas
    .Rows = 1
    For Each obj In mvarobjPnc.AccionesCorrectoras.Iterator
        If obj.getID_AUX <> -1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GridCols.ID) = obj.getID_ACCION
            .TextMatrix(.Rows - 1, GridCols.ACC_CORRECTIVAS_ESTADO) = obj.getESTADO
            .TextMatrix(.Rows - 1, GridCols.ACC_CORRECTIVAS_TITULO) = obj.getTITULO
            .TextMatrix(.Rows - 1, GridCols.ACC_CORRECTIVAS_RESPONSABLE) = obj.getRESPONSABLE
        End If
    Next obj
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_AccionesCorrectivas"
    Exit Sub
PresentarDatos_PncTramitacion_AccionesCorrectivas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.PresentarDatos_PncTramitacion_AccionesCorrectivas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_PncTramitacion_AccionesCorrectivas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub CambiarEstado_a_PtePlanAccionesCorrectivas()
On Error GoTo CambiarEstado_a_PtePlanAccionesCorrectivas_Error

    If MsgBox("Para Proceder a Solicitar el Plan de Acciones Correctivas de la Incidencia , se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Solicitar Plan de Acciones Correctivas P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS)
    
    mvarblnResultado = True
    Me.Hide

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_PtePlanAccionesCorrectivas"
    Exit Sub
CambiarEstado_a_PtePlanAccionesCorrectivas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_PtePlanAccionesCorrectivas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_PtePlanAccionesCorrectivas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub CambiarEstado_a_PteCierre()
On Error GoTo CambiarEstado_a_PteCierre_Error

    If MsgBox("Para Proceder a Solicitar el Cierre de la Incidencia por parte de Calidad, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Solicitar Cierre P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.pte_cierre)
    
    mvarblnResultado = True
    Me.Hide

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_PteCierre"
    Exit Sub
CambiarEstado_a_PteCierre_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_PteCierre"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_PteCierre of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub CambiarEstado_a_CerradaParcial()
On Error GoTo CambiarEstado_a_CerradaParcial_Error

    If MsgBox("Para Proceder a Cerrar Parcialmente la Incidencia, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Cierre Parcial P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL)
    
    mvarblnResultado = True
    Me.Hide

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_CerradaParcial"
    Exit Sub
CambiarEstado_a_CerradaParcial_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_CerradaParcial"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_CerradaParcial of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub CambiarEstado_a_CerradaTotal()
On Error GoTo CambiarEstado_a_CerradaTotal_Error

    If MsgBox("Para Proceder a Cerrar Finalmente la Incidencia, se guardarán los datos acutales. ¿Desea Continuar?", vbInformation + vbYesNo, "Cierre Parcial P.N.C.") = vbNo Then Exit Sub
    
    If Not ComprobarDatos() Then Exit Sub

    If Not ComprobarCierreTotalPosible() Then Exit Sub
    
    Call GuardarDatos

    Call mvarobjPnc.Modificar
    
    Call mvarobjPnc.cambiarEstado(C_PROCNC_ESTADOS.CERRADA)
    
    mvarblnResultado = True
    Me.Hide

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_CerradaTotal"
    Exit Sub
CambiarEstado_a_CerradaTotal_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.CambiarEstado_a_CerradaTotal"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CambiarEstado_a_CerradaTotal of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub configurarProcNC_PtePlanAccionesCorrectivas()
    
On Error GoTo configurarProcNC_PtePlanAccionesCorrectivas_Error

    tabDatosNoConformidad.TabVisible(9) = False
    
        
    ' Botones
    cmdCambiarEstado.Caption = "Enviar Plan Acc. Correct."
    cmdRechazarEstado.Visible = False
    'cmdRechazarEstado.Caption = "Devolver a Tramitación"
    ' Texto de estado
    
    txtestado.Text = "P.N.C. Pte. Plan Acciones Correctivas"

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_PtePlanAccionesCorrectivas"
    Exit Sub
configurarProcNC_PtePlanAccionesCorrectivas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_PtePlanAccionesCorrectivas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_PtePlanAccionesCorrectivas of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub configurarProcNC_PteCierre()
On Error GoTo configurarProcNC_PteCierre_Error

    tabDatosNoConformidad.TabVisible(9) = False
        
    ' Botones
    cmdCambiarEstado.Caption = "Cerrar P.N.C. (Pte. Acciones Correctivas)"
    cmdRechazarEstado.Visible = False
    'cmdRechazarEstado.Caption = "Devolver a Tramitación"
    ' Texto de estado
    
    txtestado.Text = "P.N.C. Pte. Plan Acciones Correctivas"

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_PteCierre"
    Exit Sub
configurarProcNC_PteCierre_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_PteCierre"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_PteCierre of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub configurarProcNC_CierreParcialEval()
        
    ' Botones
On Error GoTo configurarProcNC_CierreParcialEval_Error

    cmdCambiarEstado.Caption = "Cerrar P.N.C."
    cmdRechazarEstado.Visible = False
    'cmdRechazarEstado.Caption = "Devolver a Tramitación"
    ' Texto de estado
    
    txtestado.Text = "P.N.C. Pte. Plan Acciones Correctivas"

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_CierreParcialEval"
    Exit Sub
configurarProcNC_CierreParcialEval_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.configurarProcNC_CierreParcialEval"
    error_grave Err.Number & " (" & Err.Description & ") in procedure configurarProcNC_CierreParcialEval of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub configurarProcNC_CierreTotal()
    ' Poner todas las pestañas a solo visualizar
    ' incluso las de los PNC
End Sub

Private Function ComprobarCierreTotalPosible() As Boolean

    Dim objAcc As clsProcNcAccionCorrectora
    Dim blnRes As Boolean
    Dim strMensaje As String
    
On Error GoTo ComprobarCierreTotalPosible_Error

    If mvarobjPnc.AccionesCorrectoras.Count = 0 Then
        ComprobarCierreTotalPosible = True
        Exit Function
    End If
    
    For Each objAcc In mvarobjPnc.AccionesCorrectoras.Iterator
        If objAcc.getESTADO_ID = C_PROCNC_ESTADOS.ABIERTA Or objAcc.getESTADO_ID = C_PROCNC_ESTADOS.EN_TRAMITACION Then
            MsgBox "Existen algunas Acciones Correctivas que aun no han sido Cerradas, por lo que no se puede cerrar totalmente esta Incidencia", vbInformation, "Cerrar Incidencia"
            ComprobarCierreTotalPosible = False
            Exit Function
        End If
    Next objAcc
    
    strMensaje = ""
    
    If Not (optEvidencias(1).value Or optEvidencias(2).value) Then
        strMensaje = vbCrLf & " - Debe Señalar las Evidencias."
    End If
    If Not (optEvidencias(3).value Or optEvidencias(4).value) Then
        strMensaje = vbCrLf & " - Debe Señalar las Evidencias."
    End If
    If Not (optEvidencias(5).value Or optEvidencias(6).value) Then
        strMensaje = vbCrLf & " - Debe Señalar las Evidencias."
    End If
    If Not (optEvidencias(7).value Or optEvidencias(8).value) Then
        strMensaje = vbCrLf & " - Debe Señalar las Evidencias."
    End If
    
    If Not (optEval_res_incidencia.value Or optEval_res_nc.value) Then
            strMensaje = strMensaje & vbCrLf & " - Debe Señalar el Resultado"
    End If
    If Not (optEval_res_si.value Or optEval_res_no.value) Then
            strMensaje = strMensaje & vbCrLf & " - Debe Señalar si la Solución es Aceptable"
    End If
    
    
    If Trim(strMensaje) <> "" Then
        MsgBox "Debe Responder a las siguiente cuestiones:" & strMensaje, vbInformation, "Cerrar Incidencia"
        ComprobarCierreTotalPosible = False
        Exit Function
    End If

    ComprobarCierreTotalPosible = True

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ComprobarCierreTotalPosible"
    Exit Function
ComprobarCierreTotalPosible_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ComprobarCierreTotalPosible"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ComprobarCierreTotalPosible of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Private Sub ConfigurarNivelesAcceso()


On Error GoTo ConfigurarNivelesAcceso_Error

    mvarTipoEdicionNivelAcceso = enumTipoEdicion.EDICION

    ' evalua según el estado, modifica el nivel de acceso .
    ' cuando está en uno de los siguientes estados, solo los responsables de calidad pueden acceder
    If mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
        If mvarNivelAcceso <= 3 Then mvarNivelAcceso = 0
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.pte_cierre Then
        If mvarNivelAcceso <= 3 Then mvarNivelAcceso = 0
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD Then
        If mvarNivelAcceso <= 3 Then mvarNivelAcceso = 0
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS Then
        If mvarNivelAcceso <= 2 Then mvarNivelAcceso = 0
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL Then
        If mvarNivelAcceso <= 3 Then mvarNivelAcceso = 0
    ElseIf mvarobjPnc.getESTADO_ID = C_PROCNC_ESTADOS.CERRADA Then
        mvarNivelAcceso = 0
    End If


    If mvarNivelAcceso = 0 Then
        mvarTipoEdicionNivelAcceso = enumTipoEdicion.visualizar
    Else
        mvarTipoEdicionNivelAcceso = enumTipoEdicion.EDICION
    End If
    
    'Botones
    cmdCambiarEstado.Enabled = (mvarNivelAcceso >= 3)
    cmdRechazarEstado.Enabled = (mvarNivelAcceso = 4)
    cmdok.Enabled = (mvarNivelAcceso > 0)
                
    ' Pestaña 1
    If mvarobjPnc.getESTADO_ID >= C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
        fraResponsableAperturaIncidencia.Enabled = (mvarNivelAcceso = 4)
        fraOrigen.Enabled = (mvarNivelAcceso = 4)
        fraDptoImplicados.Enabled = (mvarNivelAcceso = 4)
    Else
        fraResponsableAperturaIncidencia.Enabled = (mvarNivelAcceso = 1)
        fraOrigen.Enabled = (mvarNivelAcceso > 0)
        fraDptoImplicados.Enabled = (mvarNivelAcceso > 0)
    End If
    
    ' Pestaña 2
    If mvarobjPnc.getESTADO_ID >= C_PROCNC_ESTADOS.PTE_VISTO_BUENO Then
        txtResumen.Enabled = (mvarNivelAcceso = 4)
        txtDescripcionIncidencia.Enabled = (mvarNivelAcceso = 4)
        cmdAnadirAccionInmediata.Enabled = (mvarNivelAcceso = 4)
        cmdEliminarSelAccInmediatas.Enabled = (mvarNivelAcceso = 4)
        cmdAsociarDocumentacionIncidencia.Enabled = (mvarNivelAcceso = 4)
    Else
        txtResumen.Enabled = (mvarNivelAcceso > 0)
        txtDescripcionIncidencia.Enabled = (mvarNivelAcceso > 0)
        cmdAnadirAccionInmediata.Enabled = (mvarNivelAcceso > 0)
        cmdEliminarSelAccInmediatas.Enabled = (mvarNivelAcceso > 0)
        cmdAsociarDocumentacionIncidencia.Enabled = (mvarNivelAcceso > 0)
    End If
    
    'Pestaña 3
    optDepartamento.Enabled = (mvarNivelAcceso = 4)
    optGrupo.Enabled = (mvarNivelAcceso = 4)
    cmbDptoEquipo.Enabled = (mvarNivelAcceso = 4)
    cmbGrupoPersonas.Enabled = (mvarNivelAcceso = 4)
    cmdRetirarDelEquipo.Enabled = (mvarNivelAcceso = 4)
    cmdAnadirAlEquipo.Enabled = (mvarNivelAcceso = 4)
    lstEquipoHumano.Enabled = (mvarNivelAcceso = 4)
    
    'Pestaña 4
    cmdEliminarPregunta.Enabled = (mvarNivelAcceso >= 3)
    cmdAnadirPregunta.Enabled = (mvarNivelAcceso >= 3)
    
    cmdAsociarDocumentacionIdentificacion.Enabled = (mvarNivelAcceso >= 2)
    
    ' Pestaña 5
    cmbPersonalImplicadoIncidencia.Enabled = (mvarNivelAcceso >= 2)
    cmdAnadirPersonalImplicadoIncidencia.Enabled = (mvarNivelAcceso >= 2)
    cmdEliminarPersonalImplicado.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaFecha.Enabled = (mvarNivelAcceso >= 2)
    cmbAnalisisEscenaHora.Enabled = (mvarNivelAcceso >= 2)
    cmbAnalisisEscenaMinutos.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaLocalizacion.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaCondicionesOperacion.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaCondicionesAmbientales.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaComunicacion.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaSecuencia.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaEquiposImplicados.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaCambiosRecientes.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaFormacion.Enabled = (mvarNivelAcceso >= 2)
    txtAnalisisEscenaExperiencia.Enabled = (mvarNivelAcceso >= 2)
    
    ' Pestaña 6
    
    txtRecoleccionDatos.Enabled = (mvarNivelAcceso >= 2)
    cmdAdjuntarRecoleccionDatos.Enabled = (mvarNivelAcceso >= 2)
    
    ' Pestaña 7
    lstCausas(1).Enabled = (mvarNivelAcceso >= 2)
    lstCausas(2).Enabled = (mvarNivelAcceso >= 2)
    lstCausas(3).Enabled = (mvarNivelAcceso >= 2)
    lstCausas(4).Enabled = (mvarNivelAcceso >= 2)
    lstCausas(5).Enabled = (mvarNivelAcceso >= 2)
    fraProblemasHumanos.Enabled = (mvarNivelAcceso >= 2)
        
    ' Pestaña 8
    txtResumenCausas.Enabled = (mvarNivelAcceso >= 2)
    fraCausaDirecta.Enabled = (mvarNivelAcceso >= 2)
    fraCausasContributibas.Enabled = (mvarNivelAcceso >= 2)
    fraCausaRaiz.Enabled = (mvarNivelAcceso >= 2)
    
    ' Pestaña 9
    cmdAnadirAccionCorrectivas.Enabled = (mvarNivelAcceso >= 3)
    cmdEliminarAccionCorrectiva.Enabled = (mvarNivelAcceso > 3)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfigurarNivelesAcceso"
    Exit Sub
ConfigurarNivelesAcceso_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Detalle.ConfigurarNivelesAcceso"
    error_grave Err.Number & " (" & Err.Description & ") in procedure ConfigurarNivelesAcceso of Formulario frmProcNC_Detalle" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub
