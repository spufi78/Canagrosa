VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#34.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9810
   ClientLeft      =   1500
   ClientTop       =   930
   ClientWidth     =   12570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEquipos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEquipos_Detalle.frx":000C
   ScaleHeight     =   9810
   ScaleWidth      =   12570
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   3960
      Picture         =   "frmEquipos_Detalle.frx":0C4E
      Style           =   1  'Graphical
      TabIndex        =   145
      ToolTipText     =   "Generar etiqueta"
      Top             =   8865
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mantenimiento"
      Enabled         =   0   'False
      Height          =   870
      Index           =   2
      Left            =   2655
      Picture         =   "frmEquipos_Detalle.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   144
      ToolTipText     =   "Planes de mantenimiento del equipo"
      Top             =   8865
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Verificación"
      Enabled         =   0   'False
      Height          =   870
      Index           =   1
      Left            =   1350
      Picture         =   "frmEquipos_Detalle.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   8865
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calibración"
      Enabled         =   0   'False
      Height          =   870
      Index           =   0
      Left            =   45
      Picture         =   "frmEquipos_Detalle.frx":3F0E
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   8865
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de trabajo"
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
      Height          =   2070
      Left            =   45
      TabIndex        =   112
      Top             =   5040
      Width           =   12480
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   32
         Left            =   6030
         TabIndex        =   26
         Top             =   1665
         Width           =   3585
      End
      Begin VB.CommandButton cmdAbrirDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   11970
         Picture         =   "frmEquipos_Detalle.frx":4218
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Ver documento"
         Top             =   1575
         Width           =   420
      End
      Begin MSComctlLib.ListView lstDocumentacion 
         Height          =   1230
         Left            =   6030
         TabIndex        =   131
         Top             =   225
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ruta"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdExplorarDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   10350
         Picture         =   "frmEquipos_Detalle.frx":446D
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   "Buscar documento"
         Top             =   1575
         Width           =   465
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   31
         Left            =   12330
         TabIndex        =   42
         Top             =   585
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.CommandButton cmdAnadirDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   10890
         Picture         =   "frmEquipos_Detalle.frx":46DE
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Añadir documento"
         Top             =   1575
         Width           =   420
      End
      Begin VB.CommandButton cmdEliminarDocumento 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   11430
         Picture         =   "frmEquipos_Detalle.frx":4903
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Eliminar documento"
         Top             =   1575
         Width           =   420
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   25
         Left            =   4155
         MaxLength       =   100
         TabIndex        =   24
         Top             =   1305
         Width           =   1455
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   26
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   23
         Top             =   1305
         Width           =   1500
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   28
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   20
         Top             =   585
         Width           =   4485
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   19
         Left            =   4095
         MaxLength       =   100
         TabIndex        =   18
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   20
         Left            =   4905
         MaxLength       =   100
         TabIndex        =   19
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   18
         Left            =   2925
         MaxLength       =   100
         TabIndex        =   17
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   29
         Left            =   1125
         MaxLength       =   255
         TabIndex        =   25
         Top             =   1665
         Width           =   4500
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   23
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   21
         Top             =   945
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   24
         Left            =   1925
         MaxLength       =   100
         TabIndex        =   22
         Top             =   945
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   17
         Left            =   2115
         MaxLength       =   100
         TabIndex        =   16
         Top             =   225
         Width           =   690
      End
      Begin MSDataListLib.DataCombo cmbCAmb 
         Height          =   315
         Left            =   1110
         TabIndex        =   15
         Top             =   225
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documentación"
         Height          =   240
         Index           =   45
         Left            =   6030
         TabIndex        =   129
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   44
         Left            =   2790
         TabIndex        =   125
         Top             =   990
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incert. Máx. Adm."
         Height          =   195
         Index           =   31
         Left            =   2805
         TabIndex        =   124
         Top             =   1395
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Toler. Máx."
         Height          =   195
         Index           =   32
         Left            =   120
         TabIndex        =   123
         Top             =   1350
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otras"
         Height          =   195
         Index           =   37
         Left            =   120
         TabIndex        =   121
         Top             =   630
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "% Hr"
         Height          =   195
         Index           =   23
         Left            =   3690
         TabIndex        =   120
         Top             =   315
         Width           =   330
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   28
         Left            =   4815
         TabIndex        =   119
         Top             =   270
         Width           =   45
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lim. Uso"
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   118
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "R. Trabajo"
         Height          =   195
         Index           =   25
         Left            =   120
         TabIndex        =   117
         Top             =   990
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   29
         Left            =   1845
         TabIndex        =   116
         Top             =   990
         Width           =   45
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cond. Amb."
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   115
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "º C"
         Height          =   195
         Index           =   22
         Left            =   1845
         TabIndex        =   114
         Top             =   270
         Width           =   210
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   26
         Left            =   2835
         TabIndex        =   113
         Top             =   315
         Width           =   45
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mantenimiento"
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
      Height          =   1695
      Left            =   8370
      TabIndex        =   109
      Top             =   7110
      Width           =   4155
      Begin VB.CheckBox chkCon_Mantenimiento 
         Enabled         =   0   'False
         Height          =   195
         Left            =   3600
         TabIndex        =   35
         Top             =   0
         Width           =   195
      End
      Begin MSDataListLib.DataCombo cmbPeriMantenimiento 
         Height          =   315
         Left            =   1110
         TabIndex        =   36
         Top             =   180
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker fechaProximoMantenimiento 
         Height          =   315
         Left            =   1125
         TabIndex        =   37
         Top             =   540
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   68812801
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoMantenimiento 
         Height          =   315
         Left            =   1125
         TabIndex        =   153
         Top             =   900
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   1125
         TabIndex        =   155
         Top             =   1260
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbResponsable_interno 
         Height          =   330
         Left            =   1125
         TabIndex        =   157
         Top             =   1260
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   54
         Left            =   135
         TabIndex        =   156
         Top             =   1305
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   53
         Left            =   135
         TabIndex        =   154
         Top             =   945
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próxima"
         Height          =   195
         Index           =   43
         Left            =   135
         TabIndex        =   111
         Top             =   585
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   42
         Left            =   135
         TabIndex        =   110
         Top             =   225
         Width           =   870
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Verificación"
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
      Height          =   1695
      Left            =   4185
      TabIndex        =   106
      Top             =   7110
      Width           =   4155
      Begin VB.CheckBox chkCon_Verificacion 
         Enabled         =   0   'False
         Height          =   240
         Left            =   3600
         TabIndex        =   32
         Top             =   0
         Width           =   195
      End
      Begin MSDataListLib.DataCombo cmbPeriVerificacion 
         Height          =   315
         Left            =   1155
         TabIndex        =   33
         Top             =   180
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker fechaProximaVerificacion 
         Height          =   315
         Left            =   1170
         TabIndex        =   34
         Top             =   540
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   68812801
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoVerificacion 
         Height          =   315
         Left            =   1170
         TabIndex        =   148
         Top             =   900
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbVerificador 
         Height          =   330
         Left            =   1170
         TabIndex        =   150
         Top             =   1260
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVerificador_interno 
         Height          =   330
         Left            =   1170
         TabIndex        =   151
         Top             =   1260
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Verificador"
         Height          =   195
         Index           =   52
         Left            =   225
         TabIndex        =   152
         Top             =   1305
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   51
         Left            =   225
         TabIndex        =   149
         Top             =   945
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   41
         Left            =   225
         TabIndex        =   108
         Top             =   225
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próxima"
         Height          =   195
         Index           =   40
         Left            =   225
         TabIndex        =   107
         Top             =   585
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Calibración"
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
      Height          =   1695
      Left            =   45
      TabIndex        =   103
      Top             =   7110
      Width           =   4110
      Begin VB.CheckBox chkCon_Calibracion 
         Enabled         =   0   'False
         Height          =   195
         Left            =   3555
         TabIndex        =   27
         Top             =   0
         Width           =   195
      End
      Begin MSDataListLib.DataCombo cmbPeriCalibracion 
         Height          =   315
         Left            =   990
         TabIndex        =   28
         Top             =   225
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker FechaCalibracion 
         Height          =   315
         Left            =   990
         TabIndex        =   29
         Top             =   585
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   68812801
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoCalibracion 
         Height          =   315
         Left            =   990
         TabIndex        =   30
         Top             =   945
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbCalibrador 
         Height          =   330
         Left            =   990
         TabIndex        =   31
         Top             =   1305
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCalibrador_interno 
         Height          =   330
         Left            =   990
         TabIndex        =   140
         Top             =   1305
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calibrador"
         Height          =   195
         Index           =   48
         Left            =   90
         TabIndex        =   139
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   47
         Left            =   90
         TabIndex        =   138
         Top             =   990
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Próxima"
         Height          =   195
         Index           =   39
         Left            =   90
         TabIndex        =   105
         Top             =   630
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   38
         Left            =   90
         TabIndex        =   104
         Top             =   315
         Width           =   870
      End
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha. Técn-Histórica"
      Height          =   870
      Index           =   4
      Left            =   6570
      Picture         =   "frmEquipos_Detalle.frx":4A97
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   8865
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documentación del Equipo"
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
      Height          =   930
      Left            =   3870
      TabIndex        =   82
      Top             =   9675
      Visible         =   0   'False
      Width           =   4785
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Height          =   330
         Index           =   2
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   990
         Width           =   780
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Height          =   330
         Index           =   1
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   630
         Width           =   780
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver"
         Height          =   330
         Index           =   0
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   270
         Width           =   780
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   330
         Index           =   2
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   990
         Width           =   780
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   330
         Index           =   1
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   630
         Width           =   780
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   330
         Index           =   0
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   270
         Width           =   780
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   14
         Left            =   1620
         MaxLength       =   255
         TabIndex        =   53
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   13
         Left            =   1620
         MaxLength       =   255
         TabIndex        =   52
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   12
         Left            =   1620
         MaxLength       =   255
         TabIndex        =   51
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Protocolo M.Preven."
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   85
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Op.mant.Preventivo"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   84
         Top             =   675
         Width           =   1410
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Instrucciones"
         Height          =   195
         Index           =   15
         Left            =   90
         TabIndex        =   83
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reg.Puesta Servicio"
      Height          =   870
      Index           =   3
      Left            =   5265
      Picture         =   "frmEquipos_Detalle.frx":5361
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8865
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones del Equipo"
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
      Left            =   5940
      TabIndex        =   71
      Top             =   9945
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   9
         Left            =   1350
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   675
         Width           =   525
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mantenimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   2070
         TabIndex        =   56
         Top             =   945
         Width           =   1995
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Verificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   55
         Top             =   675
         Width           =   1635
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calibración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   2070
         TabIndex        =   54
         Top             =   405
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Notas Técnicas"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   78
         Top             =   855
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   8865
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8865
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Características Generales del Equipo"
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
      Height          =   4425
      Left            =   45
      TabIndex        =   62
      Top             =   585
      Width           =   12465
      Begin VB.TextBox txtNotas 
         Height          =   1005
         Left            =   6030
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   146
         Top             =   1755
         Width           =   6315
      End
      Begin pryCombo.miCombo cmbUnidad 
         Height          =   330
         Left            =   3465
         TabIndex        =   10
         Top             =   2115
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdAbrirDocNorma 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   5040
         Picture         =   "frmEquipos_Detalle.frx":5D4B
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Ver norma"
         Top             =   3870
         Width           =   465
      End
      Begin VB.CheckBox chkNoAplica_Normas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Aplica"
         Height          =   240
         Left            =   4545
         TabIndex        =   136
         Top             =   2565
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmdEliminarNorma 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   4455
         Picture         =   "frmEquipos_Detalle.frx":5FA0
         Style           =   1  'Graphical
         TabIndex        =   135
         ToolTipText     =   "Eliminar norma"
         Top             =   3870
         Width           =   465
      End
      Begin VB.CommandButton cmdAnadirNorma 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   3915
         Picture         =   "frmEquipos_Detalle.frx":6134
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Añadir norma"
         Top             =   3870
         Width           =   420
      End
      Begin VB.CommandButton cmdEliminarAccesorio 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   11925
         Picture         =   "frmEquipos_Detalle.frx":6359
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Eliminar accesorio"
         Top             =   3870
         Width           =   420
      End
      Begin MSDataListLib.DataCombo cmbProveedor0 
         Height          =   315
         Left            =   7155
         TabIndex        =   39
         Top             =   4500
         Visible         =   0   'False
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdAnadirAccesorio 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   11385
         Picture         =   "frmEquipos_Detalle.frx":64ED
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Añadir accesorio"
         Top             =   3870
         Width           =   420
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   30
         Left            =   6030
         MaxLength       =   250
         TabIndex        =   14
         Top             =   3960
         Width           =   5025
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   27
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2475
         Width           =   1510
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   22
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2115
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   21
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2115
         Width           =   690
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   16
         Left            =   6030
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1035
         Width           =   3225
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   15
         Left            =   9675
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1035
         Width           =   2685
      End
      Begin VB.CommandButton cmdbaja 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar"
         Height          =   780
         Left            =   11655
         Picture         =   "frmEquipos_Detalle.frx":6712
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   225
         Width           =   720
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   11
         Left            =   11115
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   5175
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   10
         Left            =   4770
         MaxLength       =   255
         TabIndex        =   41
         Top             =   2475
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   8
         Left            =   6570
         MaxLength       =   255
         TabIndex        =   49
         Top             =   4590
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   7
         Left            =   11775
         MaxLength       =   255
         TabIndex        =   48
         Top             =   4590
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   6
         Left            =   11280
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   4905
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   375
         Index           =   5
         Left            =   10170
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   630
         Width           =   1380
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   1110
         MaxLength       =   250
         TabIndex        =   2
         Top             =   1035
         Width           =   4350
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   11640
         MaxLength       =   255
         TabIndex        =   47
         Top             =   5265
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   11820
         MaxLength       =   255
         TabIndex        =   46
         Top             =   5265
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   375
         Index           =   0
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   61
         Top             =   315
         Width           =   1110
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Index           =   1
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   0
         Top             =   720
         Width           =   2685
      End
      Begin MSComCtl2.DTPicker fpuesta 
         Height          =   345
         Left            =   7875
         TabIndex        =   44
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
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
         Format          =   68812801
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker frecepcion 
         Height          =   345
         Left            =   4050
         TabIndex        =   43
         Top             =   315
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
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
         Format          =   68812801
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbFamilia 
         Height          =   315
         Left            =   1110
         TabIndex        =   5
         Top             =   1395
         Width           =   4350
         _ExtentX        =   7673
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbSituacion 
         Height          =   315
         Left            =   6030
         TabIndex        =   6
         Top             =   1395
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbUnidad0 
         Height          =   315
         Left            =   6840
         TabIndex        =   40
         Top             =   4500
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbEsNadcap 
         Height          =   315
         Left            =   7875
         TabIndex        =   1
         Top             =   675
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSComctlLib.ListView lstAccesorios 
         Height          =   960
         Left            =   6030
         TabIndex        =   13
         Top             =   2790
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   1693
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accesorio"
            Object.Width           =   10585
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstNormas 
         Height          =   915
         Left            =   120
         TabIndex        =   12
         Top             =   2835
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   1614
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID_NORMA"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Norma"
            Object.Width           =   6350
         EndProperty
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   330
         Left            =   1110
         TabIndex        =   7
         Top             =   1755
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbRevisado 
         Height          =   315
         Left            =   10170
         TabIndex        =   38
         Top             =   0
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisado"
         Height          =   195
         Index           =   50
         Left            =   9405
         TabIndex        =   147
         Top             =   45
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   49
         Left            =   2745
         TabIndex        =   141
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accesorio"
         Height          =   195
         Index           =   35
         Left            =   6045
         TabIndex        =   102
         Top             =   3735
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normas"
         Height          =   195
         Index           =   46
         Left            =   120
         TabIndex        =   133
         Top             =   2835
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         Height          =   195
         Index           =   34
         Left            =   7155
         TabIndex        =   101
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precisión"
         Height          =   195
         Index           =   33
         Left            =   120
         TabIndex        =   100
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   30
         Left            =   2745
         TabIndex        =   99
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   27
         Left            =   1845
         TabIndex        =   98
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "R. Medida"
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   97
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabric."
         Height          =   195
         Index           =   20
         Left            =   5535
         TabIndex        =   96
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mod"
         Height          =   195
         Index           =   19
         Left            =   9315
         TabIndex        =   95
         Top             =   1125
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   94
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modo de Empleo"
         Height          =   195
         Index           =   14
         Left            =   9585
         TabIndex        =   81
         Top             =   5355
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C. Normas"
         Height          =   195
         Index           =   13
         Left            =   3945
         TabIndex        =   80
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Recepción"
         Height          =   195
         Index           =   12
         Left            =   3105
         TabIndex        =   79
         Top             =   405
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Puesta Servicio"
         Height          =   195
         Index           =   10
         Left            =   6525
         TabIndex        =   77
         Top             =   405
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   76
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Situac."
         Height          =   195
         Index           =   8
         Left            =   5535
         TabIndex        =   75
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cond. Ambientales"
         Height          =   195
         Index           =   5
         Left            =   10260
         TabIndex        =   74
         Top             =   4635
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alta / Baja"
         Height          =   195
         Index           =   1
         Left            =   10665
         TabIndex        =   72
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rango de Trabajo"
         Height          =   195
         Index           =   7
         Left            =   10080
         TabIndex        =   68
         Top             =   5310
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rango de Medida"
         Height          =   195
         Index           =   6
         Left            =   10305
         TabIndex        =   67
         Top             =   4545
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   4
         Left            =   10395
         TabIndex        =   66
         Top             =   4905
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Equipo"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Serie"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   765
         Width           =   585
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   10125
      Top             =   9810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Equipo"
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
      TabIndex        =   70
      Top             =   45
      Width           =   750
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11970
      Picture         =   "frmEquipos_Detalle.frx":6FDC
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de medición y Ensayo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   69
      Top             =   315
      Width           =   2085
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12555
   End
End
Attribute VB_Name = "frmEquipos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Public ES_AVISO As Boolean
'E0074-I

Private Sub chkCon_Calibracion_Click()
    Call estado_frame_calibracion
End Sub

Private Sub chkCon_Verificacion_Click()
    Call estado_frame_verificacion
End Sub

Private Sub chkCon_Mantenimiento_Click()
    Call estado_frame_mantenimiento
End Sub

Private Sub estado_frame_calibracion()
    If chkCon_Calibracion.value = Checked Then
        cmd(0).Enabled = True
    Else
        cmbPeriCalibracion.Text = ""
        FechaCalibracion.value = "01-01-1900"
        cmbTipoCalibracion.Text = ""
        cmbCalibrador.Limpiar
        cmbCalibrador.desactivar
        cmbCalibrador_interno.Limpiar
        cmbCalibrador_interno.desactivar
        cmd(0).Enabled = False
    End If
End Sub

Private Sub estado_frame_verificacion()
    If chkCon_Verificacion.value = Checked Then
        cmd(1).Enabled = True
    Else
        cmbPeriVerificacion.Text = ""
        fechaProximaVerificacion.value = "01-01-1900"
        cmbTipoVerificacion.Text = ""
        cmbVerificador.Limpiar
        cmbVerificador.desactivar
        cmbVerificador_interno.Limpiar
        cmbVerificador_interno.desactivar
        cmd(1).Enabled = False
    End If
End Sub

Private Function estado_frame_mantenimiento()
    If chkCon_Mantenimiento.value = Checked Then
        cmd(2).Enabled = True
        'cmdHco(2).Enabled = True
    Else
        cmbPeriMantenimiento.Text = ""
        cmbPeriMantenimiento.Enabled = False
        fechaProximoMantenimiento.value = "01-01-1900"
        fechaProximoMantenimiento.Enabled = False
        cmbTipoMantenimiento.Enabled = False
        cmbResponsable.Limpiar
        cmbResponsable.desactivar
        cmbResponsable_interno.Limpiar
        cmbResponsable_interno.desactivar
        cmd(2).Enabled = False
        'cmdHco(2).Enabled = False
    End If
End Function

'E0074-F
Private Sub cmbCAmb_Change()
'E0023-I
' Desbloquear los controles de Tempepratura, Humedad y Otras, sólo si C. Amb. está a SI
    If cmbCAmb.Text = "SI" Then
        txtDatos(17).Enabled = True
        txtDatos(17).BackColor = vbWhite
        txtDatos(18).Enabled = True
        txtDatos(18).BackColor = vbWhite
        txtDatos(19).Enabled = True
        txtDatos(19).BackColor = vbWhite
        txtDatos(20).Enabled = True
        txtDatos(20).BackColor = vbWhite
        txtDatos(28).Enabled = True
        txtDatos(28).BackColor = vbWhite
    Else
        txtDatos(17).Enabled = False
        txtDatos(17).Text = ""
        txtDatos(17).BackColor = &HE0E0E0
        txtDatos(18).Enabled = False
        txtDatos(18).Text = ""
        txtDatos(18).BackColor = &HE0E0E0
        txtDatos(19).Enabled = False
        txtDatos(19).Text = ""
        txtDatos(19).BackColor = &HE0E0E0
        txtDatos(20).Enabled = False
        txtDatos(20).Text = ""
        txtDatos(20).BackColor = &HE0E0E0
        txtDatos(28).Enabled = False
        txtDatos(28).Text = ""
        txtDatos(28).BackColor = &HE0E0E0
    End If
'E0023-I
End Sub

Private Sub cmbProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbUnidad_Click(AREA As Integer)
    lblCampos(44).Caption = cmbUnidad.getTEXTO ' En la etiqueta de abajo debe aparecer lo mismo que en el combo
    lblCampos(49).Caption = cmbUnidad.getTEXTO ' En la etiqueta de abajo debe aparecer lo mismo que en el combo
End Sub

Private Sub cmbUnidad_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbUnidad_change()
    lblCampos(44) = cmbUnidad.getTEXTO
    lblCampos(49) = cmbUnidad.getTEXTO
End Sub

Private Sub cmdAbrirDocumento_Click()
    Dim destino As String
    Dim r As Long

On Error GoTo fallo

    If lstDocumentacion.ListItems.Count > 0 Then
        If Len(Replace(lstDocumentacion.SelectedItem.SubItems(2), "/", "\")) = 0 Then
            MsgBox "Debe seleccionar un documento de la lista.", vbCritical, App.Title
        Else
            destino = Replace(lstDocumentacion.SelectedItem.SubItems(2), "/", "\")
            If destino = "" Then
                Exit Sub
            End If
            If Dir(destino) <> "" Then
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
            Else
                MsgBox "El documento se ha eliminado o movido de la ruta almacenada.", vbCritical, App.Title
            End If
        End If
    Else
        MsgBox "Debe seleccionar un documento de la lista.", vbCritical, App.Title
    End If

    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
    Case 0
        With frmEquipos_Detalle_Calibracion
         .PK = PK
         .Show 1
        End With
    Case 1
        With frmEquipos_Detalle_Verificacion
         .PK = PK
         .Show 1
        End With
    Case 2
        With frmEquipos_Detalle_Mantenimiento
        'E0052-I
        ' Se parte de un formulario nuevo
        'With frmEquipos_Detalle_Mto
        'E0052-F
         .PK = PK
         .Show 1
        End With

        
    End Select
End Sub

Private Sub cmdbaja_Click()
    If txtDatos(5) = "ALTA" Then
        txtDatos(5) = "BAJA"
    Else
        txtDatos(5) = "ALTA"
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEtiqueta_Click()
    Dim prnPrinter As Printer
    
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros
    If Not oParametro.Carga(parametros.IMPRESORA_ETIQUETAS, USUARIO.getUSO) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Sub
    End If
    log ("Comienzo impresion de etiquetas")
    For Each prnPrinter In Printers
        If prnPrinter.DeviceName = oParametro.getVALOR Then
            Set Printer = prnPrinter
            Exit For
        End If
    Next
    
    With frmReport
        Firmas.copiar_firma_responsable_tecnico
        .iniciar
        .informe = "rptEquipos_Etiqueta"
        .criterio = "{equipos.ID_EQUIPO}= " & PK
        .imprimir = True
        .generar
        .Visible = False
    End With
    
    log ("Final impresion de etiquetas")
    
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir la etiqueta.", vbCritical, Err.Description
End Sub

Private Sub cmdok_Click()
    If validar = True Then
        Dim oEquipo As New clsEquipos
        Dim Equipo As Long
        With oEquipo
            .setSERIE = txtDatos(1)
            .setNOMBRE = txtDatos(2)
            If cmbProveedor.getTEXTO = "" Then
                .setPROVEEDOR_ID = 0
            Else
                .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
            End If
            
            If Format(fpuesta, "dd-mm-yyyy") <> "01-01-1900" Then
                .setFECHA_SERVICIO = Format(fpuesta, "dd-mm-yyyy")
            Else
                .setFECHA_SERVICIO = ""
            End If
            If Format(frecepcion, "dd-mm-yyyy") <> "01-01-1900" Then
                .setFECHA_RECEPCION = Format(frecepcion, "dd-mm-yyyy")
            Else
                .setFECHA_RECEPCION = ""
            End If
            .setNOTAS = txtDatos(9)
            If txtDatos(5) = "ALTA" Then
                .setALTA_BAJA = 0
            Else
                .setALTA_BAJA = 1
            End If
            
            ' Características generales
            .setFAMILIA_ID = IIf(cmbFamilia.BoundText = "", 0, cmbFamilia.BoundText)

            .setFABRICANTE = txtDatos(16)
            .setMODELO = txtDatos(15)
            .setES_NADCAP = IIf(cmbEsNadcap.BoundText = "", 2, cmbEsNadcap.BoundText)
            .setSITUACION_ID = IIf(cmbSituacion.BoundText = "", 0, cmbSituacion.BoundText)
            .setRANGO_MEDIDA_MIN = txtDatos(21)
            .setRANGO_MEDIDA_MAX = txtDatos(22)
            .setUNIDAD_ID = IIf(Len(cmbUnidad.getTEXTO) = 0, 0, cmbUnidad.getPK_SALIDA)
            .setPRECISIONN = txtDatos(27)
            
            ' Datos de trabajo
            .setCONDICIONES_AMBIENTALES = IIf(cmbCAmb.BoundText = "", 2, cmbCAmb.BoundText)
            .setTEMPERATURA_MIN = txtDatos(17)
            .setTEMPERATURA_MAX = txtDatos(18)
            .setHUMEDAD_MIN = txtDatos(19)
            .setHUMEDAD_MAX = txtDatos(20)
            .setCOND_AMBIENTALES_OTRAS = txtDatos(28)
            .setRANGO_TRABAJO_MIN = txtDatos(23)
            .setRANGO_TRABAJO_MAX = txtDatos(24)
            .setTOLERANCIA_MAXIMA = txtDatos(26)
            .setINCERTIDUMBRE_MAXIMA = txtDatos(25)
            .setLIMITACIONES_USO = txtDatos(29)
            
            ' calibración
            .setCON_CALIBRACION = chkCon_Calibracion.value
            
            ' verificación
            .setCON_VERIFICACION = chkCon_Verificacion.value
            
            ' Mantenimiento
            .setCON_MANTENIMIENTO = chkCon_Mantenimiento.value
'            .setPERIODICIDAD_MANTENIMIENTO_ID = IIf(cmbPeriMantenimiento.BoundText = "", 0, cmbPeriMantenimiento.BoundText)
'            If Format(fechaMantenimiento, "dd-mm-yyyy") <> "01-01-1900" Then
'                .setFECHA_PROX_MANTENIMIENTO = Format(fechaMantenimiento, "dd-mm-yyyy")
'            Else
'                .setFECHA_PROX_MANTENIMIENTO = ""
'            End If
            
            .setREVISADO = cmbRevisado.BoundText
        End With
        
        If PK = 0 Then
            If MsgBox("Va a introducir un nuevo equipo. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Equipo = oEquipo.Insertar
                MsgBox "El Equipo se ha introducido correctamente. " & vbCrLf & _
                       "Ahora puede informar Calibración, Verificación y Mantenimiento." & vbCrLf & _
                       "También puede añadirle accesorios, documentos y normas.", vbOKOnly + vbInformation, App.Title
                PK = Equipo
                CARGAR
            Else
                Exit Sub
            End If
        Else
            If MsgBox("Va a modificar el equipo. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                
                Dim oEQC As New clsEquipos_calibracion
                If chkCon_Calibracion.value = Unchecked Then ' Si no tiene con_calibración
                    oEQC.Eliminar (PK) ' se borra la que pudiera tener de antes
                Else
                    oEQC.marcar_como_efectiva (PK) ' se marca como efectiva
                End If
                Set oEQC = Nothing
                ' -----------------------------------------------------------------
                
                Dim oEQV As New clsEquipos_verificacion
                If chkCon_Verificacion.value = Unchecked Then ' Si no tiene con_verificación
                    oEQV.eliminar_todas_verificaciones (PK) ' se borran las que pudiera tener de antes
                Else
                    oEQV.marcar_no_efectivas_como_efectivas (PK)
                End If
                Set oEQV = Nothing
                ' ------------------------------------------------------------------
                
                oEquipo.Modificar (PK)
                
                Equipo = PK
                MsgBox "El Equipo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
                Unload Me
            Else
                Exit Sub
            End If
        End If
        
        'E0502-I
        If ES_AVISO Then            ' si se ha abierto el detalle desde la lista de avisos
            ES_AVISO = False
            frmEquipos_Listado_Avisos.cargar_lista      ' se recarga la lista de avisos
        Else
            frmEquipos_Listado.cargar_lista             ' se recarga la lista de equipos
        End If
        'E0502-F
        
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    'E0001-I
    'Se cargan los combos nuevos
    llenar_combo cmbProveedor, New clsProveedor, 0, frmProveedores, ""
    llenar_combo cmbCalibrador, New clsProveedor, 0, frmProveedores, ""
    llenar_combo cmbCalibrador_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbVerificador, New clsProveedor, 0, frmProveedores, ""
    llenar_combo cmbVerificador_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbResponsable, New clsProveedor, 0, frmProveedores, ""
    llenar_combo cmbResponsable_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbUnidad, New clsUnidades, 0, Me, ""
    Call cargar_combos
    Call estado_frame_calibracion
    Call estado_frame_verificacion
    Call estado_frame_mantenimiento
    'E0001F
    
    Dim titulo As String
    If PK <> 0 Then
        CARGAR
    Else
        cmdbaja.Enabled = False
        lbltitulo = "Alta de Equipo de Medición y Ensayo"
        Me.Caption = lbltitulo
        cmd(0).Enabled = False
        cmd(1).Enabled = False
        cmd(2).Enabled = False
        'E0018-I
        'Se deshabilitan los botones de añadir/eliminar accesorios y documentos.
        cambiar_estado_botones (False)
        'E0018-F
        Dim oEquipo As New clsEquipos
        oEquipo.CrearID
        txtDatos(0) = oEquipo.getID_EQUIPO
        txtDatos(5) = "ALTA"
        fpuesta = Date
        frecepcion = Date
'        Frame2.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call descargar_formulario
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

'E0044-I
Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27 ' Aquellos campos que deben ser numéricos
            If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = Asc(".") Or KeyAscii = 8 Or KeyAscii = Asc("+") Or KeyAscii = Asc("-")) Then ' Si no es un número o el "." no se permite
                KeyAscii = 0
            End If
            If InStr(1, txtDatos(Index), ".") > 0 And KeyAscii = Asc(".") Then ' Si ya hay un "." no se permite poner otro
                KeyAscii = 0
            End If
        Case 32 ' inserción de documentación
            KeyAscii = 0
            MsgBox "La introducción de documentación debe hacerse " & vbCrLf & _
                   "mediante el botón 'Buscar documento'.", vbInformation + vbOKOnly, App.Title
    End Select
End Sub

'E0044-F
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub CARGAR()
    Dim oEquipo As New clsEquipos
    If oEquipo.Carga(PK) = True Then
        lbltitulo = "Modificación de Equipo : " & oEquipo.getNOMBRE
        Me.Caption = lbltitulo
        
        'cmd(2).Enabled = False '' ' mantenimineto
        
        With oEquipo
            txtDatos(0) = .getID_EQUIPO
            txtDatos(1) = .getSERIE
            If .getALTA_BAJA = 0 Then
                txtDatos(5) = "ALTA"
            Else
                txtDatos(5) = "BAJA"
            End If
            txtDatos(2) = .getNOMBRE
            cmbProveedor.MostrarElemento .getPROVEEDOR_ID
            If IsDate(.getFECHA_SERVICIO) Then ' <> "0000-00-00" Then
                fpuesta = .getFECHA_SERVICIO
            End If
            If IsDate(.getFECHA_RECEPCION) Then  ' <> "0000-00-00" Then
                frecepcion = .getFECHA_RECEPCION
            End If
            txtDatos(9) = .getNOTAS
            
            ' Características generales
            cmbFamilia.BoundText = .getFAMILIA_ID
            txtDatos(16) = .getFABRICANTE
            txtDatos(15) = .getMODELO
            cmbEsNadcap.BoundText = .getES_NADCAP
            cmbSituacion.BoundText = .getSITUACION_ID
            txtDatos(21) = .getRANGO_MEDIDA_MIN
            txtDatos(22) = .getRANGO_MEDIDA_MAX
            cmbUnidad.MostrarElemento .getUNIDAD_ID
            txtDatos(27) = .getPRECISIONN
            
            ' Datos de trabajo
            cmbCAmb.BoundText = .getCONDICIONES_AMBIENTALES
            txtDatos(17) = .getTEMPERATURA_MIN
            txtDatos(18) = .getTEMPERATURA_MAX
            txtDatos(19) = .getHUMEDAD_MIN
            txtDatos(20) = .getHUMEDAD_MAX
            txtDatos(28) = .getCOND_AMBIENTALES_OTRAS
            txtDatos(23) = .getRANGO_TRABAJO_MIN
            txtDatos(24) = .getRANGO_TRABAJO_MAX
            txtDatos(26) = .getTOLERANCIA_MAXIMA
            txtDatos(25) = .getINCERTIDUMBRE_MAXIMA
            txtDatos(29) = .getLIMITACIONES_USO
            lblCampos(44).Caption = cmbUnidad.getTEXTO ' Debe aparecer lo mismo que lo seleccionado en el combo unidades
            lblCampos(49).Caption = cmbUnidad.getTEXTO ' Debe aparecer lo mismo que lo seleccionado en el combo unidades
            
            ' Calibración
            chkCon_Calibracion.Enabled = True
            chkCon_Calibracion.value = .getCON_CALIBRACION
            If chkCon_Calibracion.value = Checked Then
                Call datos_calibracion(PK)
            End If
            ' -----------
            
            ' verificación
'            chkCon_Verificacion.Enabled = True
'            chkCon_Verificacion.value = .getCON_VERIFICACION
'            If chkCon_Verificacion.value = Checked Then
'                Call datos_verificacion(PK)
'            End If
' ---------------------------
            chkCon_Verificacion.Enabled = True
            chkCon_Verificacion.value = .getCON_VERIFICACION
            If chkCon_Verificacion.value = Checked Then
                Call datos_verificacion(PK)
            End If
            ' --------------------
            
'            ' mantenimiento anterior
'            chkCon_Mantenimiento.value = .getCON_MANTENIMIENTO
'            Call estado_frame_mantenimiento
'            cmbPeriMantenimiento.BoundText = .getPERIODICIDAD_MANTENIMIENTO_ID
'            If IsDate(.getFECHA_PROX_MANTENIMIENTO) Then ' <> "0000-00-00" Then
'                fechaMantenimiento = .getFECHA_PROX_MANTENIMIENTO
'            End If
'            ' --------------------

            ' mantenimiento nuevo
            chkCon_Mantenimiento.Enabled = True
            chkCon_Mantenimiento.value = .getCON_MANTENIMIENTO
            If chkCon_Mantenimiento.value = Checked Then
                Call datos_mantenimiento(PK)
            End If
            ' --------------------
            
            txtNotas.Text = .getNOTAS
            
            cambiar_estado_botones (True) ' Se habilitan los botones de Normas, documentos y Accesorios
            Call cargar_accesorios
            Call cargar_documentos
            Call cargar_normas
            
            cmbRevisado.BoundText = .getREVISADO
        End With
    End If
    Set oEquipo = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    'E0030-I
    ' Se añaden las condiciones de obligatoriedad
    If Trim(txtDatos(1)) = "" Then ' Número de serie
        MsgBox "Debe darle un número de serie al equipo.", vbInformation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(2)) = "" Then
        MsgBox "Debe darle un nombre al equipo.", vbInformation, App.Title
        txtDatos(2).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(16)) = "" Then ' Fabricante
        MsgBox "Debe darle un fabricante al equipo.", vbInformation, App.Title
        txtDatos(16).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(15)) = "" Then ' Modelo
        MsgBox "Debe darle un modelo al equipo.", vbInformation, App.Title
        txtDatos(15).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(cmbEsNadcap.Text) = "" Then ' Es NADCAP
        MsgBox "Debe especificar si es NADCAP.", vbInformation, App.Title
        cmbEsNadcap.SetFocus
        validar = False
        Exit Function
    End If
    If Trim(cmbRevisado.Text) = "" Then ' revisado
        MsgBox "Debe especificar si están revisados los datos.", vbInformation, App.Title
        cmbRevisado.SetFocus
        validar = False
        Exit Function
    End If
    
    ' Se añaden las condiciones de que sean numéricos
    Dim campo_no_numerico As Long
    campo_no_numerico = 0
    campo_no_numerico = campos_son_numericos()
    If campo_no_numerico <> 0 Then
        MsgBox "Debe introducir un valor numérico.", vbInformation, App.Title
        txtDatos(campo_no_numerico).SetFocus
        validar = False
        Exit Function
    End If
    
    ' Se añaden las condiciones de V. Mí. < V. Máx.
    If txtDatos(21) <> "" And txtDatos(22) <> "" Then ' R. Medida
        If CDbl(Replace(txtDatos(21), ".", ",")) >= CDbl(Replace(txtDatos(22), ".", ",")) Then
                MsgBox "El Rango de medida no es correcto.", vbInformation, App.Title
                txtDatos(21).SetFocus
                validar = False
                Exit Function
        End If
    End If
    If txtDatos(23) <> "" And txtDatos(24) <> "" Then ' R. Trabajo
        If CDbl(Replace(txtDatos(23), ".", ",")) >= CDbl(Replace(txtDatos(24), ".", ",")) Then
                MsgBox "El Rango de trabajo no es correcto.", vbInformation, App.Title
                txtDatos(23).SetFocus
                validar = False
                Exit Function
        End If
    End If
    If txtDatos(17) <> "" And txtDatos(18) <> "" Then ' ºC
        If CDbl(Replace(txtDatos(17), ".", ",")) >= CDbl(Replace(txtDatos(18), ".", ",")) Then
                MsgBox "El rango de ºC no es correcto.", vbInformation, App.Title
                txtDatos(17).SetFocus
                validar = False
                Exit Function
        End If
    End If
    If txtDatos(19) <> "" And txtDatos(20) <> "" Then ' %Hr
        If CDbl(Replace(txtDatos(19), ".", ",")) >= CDbl(Replace(txtDatos(20), ".", ",")) Then
                MsgBox "El rango de % Hr no es correcto.", vbInformation, App.Title
                txtDatos(19).SetFocus
                validar = False
                Exit Function
        End If
    End If
    
    ' Se añaden las condiciones de que R. Trabajo C= R. Medida
    If txtDatos(23) <> "" And txtDatos(24) <> "" Then ' Si hay algo en R. Trabajo
        If txtDatos(21) <> "" And txtDatos(22) <> "" Then ' Y hay algo en R. Medida
            If (CDbl(Replace(txtDatos(21), ".", ",")) <= CDbl(Replace(txtDatos(23), ".", ","))) And (CDbl(Replace(txtDatos(23), ".", ",")) < CDbl(Replace(txtDatos(24), ".", ","))) And (CDbl(Replace(txtDatos(24), ".", ",")) <= CDbl(Replace(txtDatos(22), ".", ","))) Then
            Else
                MsgBox "El Rango de Trabajo no está contenido en el Rango de Medida.", vbInformation, App.Title
                validar = False
                Exit Function
            End If
        End If
    End If
    
    'E0030-F
End Function
' Función que comprueba que los campos numéricos del formulario tienen un valor numérico
Private Function campos_son_numericos() As Long
    Dim i As Long, lngResultado As Long
    
    lngResultado = 0
    For i = 17 To 27
        If Len(txtDatos(i)) <> 0 Then
            If Not IsNumeric(txtDatos(i)) Then
                lngResultado = i
            End If
        
        End If
    Next i
    campos_son_numericos = lngResultado
    
End Function
'E0002
Private Sub cargar_combos()
    Dim oDECO As New clsDecodificadora
    
    'oDECO.Cargar_Combo cmbProveedor, decodificadora.EQ_PROVEEDORES
    oDECO.Cargar_Combo cmbFamilia, decodificadora.EQ_FAMILIAS
    oDECO.Cargar_Combo cmbSituacion, decodificadora.EQ_SITUACIONES
    'oDECO.Cargar_Combo cmbUnidad, decodificadora.EQ_UNIDADES
    oDECO.Cargar_Combo cmbEsNadcap, decodificadora.EQ_SINO
    oDECO.Cargar_Combo cmbCAmb, decodificadora.EQ_SINO
    oDECO.Cargar_Combo cmbPeriCalibracion, decodificadora.EQ_periodicidad
    oDECO.Cargar_Combo cmbPeriVerificacion, decodificadora.EQ_periodicidad
    oDECO.Cargar_Combo cmbPeriMantenimiento, decodificadora.EQ_periodicidad
    oDECO.Cargar_Combo cmbTipoCalibracion, decodificadora.EQ_TIPO_CALIBRACION
    oDECO.Cargar_Combo cmbTipoVerificacion, decodificadora.EQ_TIPO_CALIBRACION
    oDECO.Cargar_Combo cmbTipoMantenimiento, decodificadora.EQ_TIPO_CALIBRACION
    
    oDECO.Cargar_Combo cmbRevisado, decodificadora.EQ_SINO
    'Cargar_Combo cmbUsuario, New clsUsuarios
End Sub
'E0002F

'E0024-I
Private Sub cmdExplorarDocumento_Click()
    'Se añade código para abrir el cuadro de diálogo
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtDatos(31) = cd.FileName ' Campo oculto para guardar ruta en la BD
        txtDatos(32) = cd.FileTitle ' Campo visible para mostrar en el formulario
    End If
End Sub
'E0024-I

'E0026-I
' Botón que añade el accesorio a la lista de accesorios del equipo
Private Sub cmdAnadirAccesorio_Click()
    If Len(Trim(txtDatos(30))) = 0 Then
        MsgBox "Debe escribir un accesorio a añadir.", vbInformation, App.Title
    Else
        ' Se añade el documento a la lista
        Dim oEquipo_acc As New clsEquipos_Accesorios
        With oEquipo_acc
            .setEQUIPO_ID = PK
            .setNOMBRE = txtDatos(30)
        End With
        oEquipo_acc.Insertar
        
        cargar_accesorios
        txtDatos(30) = ""
        txtDatos(30).SetFocus
    End If
End Sub

' Botón que elimina el accesorio de la lista de accesorios del equipo
Private Sub cmdEliminarAccesorio_Click()
    If lstAccesorios.ListItems.Count > 0 Then
        If lstAccesorios.SelectedItem.Text = "" Then
            MsgBox "Debe seleccionar el accesorio que quiere eliminar.", vbInformation, App.Title
        Else
            If MsgBox("Va a desvincular el siguiente accesorio del equipo: " & vbCrLf & lstAccesorios.SelectedItem.Text & vbCrLf & _
                      "¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Dim oEquipo_acc As New clsEquipos_Accesorios
                oEquipo_acc.Eliminar (Replace(lstAccesorios.SelectedItem.Key, "'", ""))
                
                cargar_accesorios
                Set oEquipo_acc = Nothing
            End If
        End If

    Else
        MsgBox "Debe seleccionar el accesorio que quiere eliminar.", vbInformation, App.Title
    End If
End Sub
'E0026-F

'E0025-I
' Botón que añade el documento a la lista de documentos del equipo
Private Sub cmdAnadirDocumento_Click()
    If txtDatos(32) = "" Then
        MsgBox "Debe seleccionar un documento a añadir.", vbInformation, App.Title
    Else
        ' Se añade el documento a la lista
        Dim oEquipo_doc As New clsEquipos_Documentacion
        With oEquipo_doc
            .setEQUIPO_ID = PK
            .setRUTA_DOCUMENTO = txtDatos(31)
            .setNOMBRE_DOCUMENTO = txtDatos(32)
        End With
        oEquipo_doc.Insertar
        
        cargar_documentos
        txtDatos(31) = ""
        txtDatos(32) = ""
        txtDatos(32).SetFocus
    End If
End Sub

' Botón que elimina el documento de la lista de documentos del equipo
Private Sub cmdEliminarDocumento_Click()
    If lstDocumentacion.ListItems.Count > 0 Then
        If lstDocumentacion.SelectedItem.Text = "" Then
            MsgBox "Debe seleccionar el documento que quiere quitar.", vbInformation, App.Title
        Else
            If MsgBox("Va a desvincular el siguiente documento del equipo: " & vbCrLf & lstDocumentacion.SelectedItem.SubItems(1) & vbCrLf & _
                      "¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Dim oEquipo_doc As New clsEquipos_Documentacion
                With oEquipo_doc
                    .Eliminar (lstDocumentacion.SelectedItem)
                End With
                cargar_documentos
                Set oEquipo_doc = Nothing
            End If
        End If

    Else
        MsgBox "Debe seleccionar el documento que quiere quitar.", vbInformation, App.Title
    End If
End Sub
'E0025-F
'E0027-I
Public Sub cargar_documentos()
    Dim oEquipo_doc As New clsEquipos_Documentacion
    Dim rs As ADODB.RecordSet
    
    lstDocumentacion.ListItems.Clear
    Set rs = oEquipo_doc.lista_por_equipos(PK)
    If rs.RecordCount > 0 Then
        Do
            With lstDocumentacion.ListItems.Add(, , rs("ID_DOCUMENTO"))
                .SubItems(1) = rs("NOMBRE_DOCUMENTO")
                .SubItems(2) = rs("RUTA_DOCUMENTO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEquipo_doc = Nothing
End Sub
Public Sub cargar_normas()
    Dim oEquipo_norm As New clsEquipos_Normas
    Dim rs As ADODB.RecordSet
    Dim i As Long
    
    lstNormas.ListItems.Clear
    Set rs = oEquipo_norm.lista_por_equipos(PK)
    i = 1
    If rs.RecordCount > 0 Then
        Do
            With lstNormas.ListItems.Add(, "'" & rs.Fields("ID_NORMA") & "'", rs.Fields("ID_NORMA"))
                .SubItems(1) = rs.Fields("CODIGO")
                .SubItems(2) = rs.Fields("NOMBRE")
            End With
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEquipo_norm = Nothing
End Sub


Public Sub cargar_accesorios()
    Dim oEquipo_acc As New clsEquipos_Accesorios
    Dim rs As ADODB.RecordSet
    
    lstAccesorios.ListItems.Clear
    Set rs = oEquipo_acc.lista_por_equipos(PK)
    If rs.RecordCount > 0 Then
        Do
            lstAccesorios.ListItems.Add , "'" & rs(1) & "'", rs(0)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEquipo_acc = Nothing
End Sub
'E0027-F

'E0031-I
' Procedimiento que establece el estado de los botones al pasado por parámetro
Public Sub cambiar_estado_botones(booEstado As Boolean)
    cmdAnadirAccesorio.Enabled = booEstado
    cmdEliminarAccesorio.Enabled = booEstado
    cmdExplorarDocumento.Enabled = booEstado
    cmdAnadirDocumento.Enabled = booEstado
    cmdEliminarDocumento.Enabled = booEstado
    cmdAbrirDocumento.Enabled = booEstado
    cmdAnadirNorma.Enabled = booEstado
    cmdEliminarNorma.Enabled = booEstado
    cmdAbrirDocNorma.Enabled = booEstado
End Sub
'E0031-F

'E0032-I
' Botón que abre el listado de normas y permite seleccionar una para vincularla con el equipo
Private Sub cmdAnadirNorma_Click()
'    frmEquipos_CA_Listado_Normas.FK_EQUIPO = PK ' Se le pasa la pk como fk al formulario
'    frmEquipos_CA_Listado_Normas.Show 1 'Se abre el formulario del listado de normas nuevo

    ' Se usa el formulario frmCA_Listado_Normas
'    frmCA_Listado_Normas.FK_EQUIPO = PK ' Se le pasa la pk como fk al formulario
    gID = 0
    frmCA_Listado_Normas.VINCULAR = True
    frmCA_Listado_Normas.Show 1 'Se abre el formulario del listado de normas
    If gID <> 0 Then
        Dim oEquipos_Normas As New clsEquipos_Normas
        oEquipos_Normas.setNORMA_ID = gID
        oEquipos_Normas.setEQUIPO_ID = PK
        oEquipos_Normas.Insertar
    End If
    cargar_normas
End Sub

Private Sub cmdEliminarNorma_Click()
    If lstNormas.ListItems.Count > 0 Then
        If lstNormas.SelectedItem.Text = "" Then
            MsgBox "Debe seleccionar la norma que quiere eliminar.", vbInformation, App.Title
        Else
            If MsgBox("Va a desvincular la siguiente norma del equipo: " & vbCrLf & _
                      "Código: " & lstNormas.ListItems(lstNormas.SelectedItem.Index).ListSubItems(1).Text & vbCrLf & _
                      "Norma: " & lstNormas.ListItems(lstNormas.SelectedItem.Index).ListSubItems(2).Text & vbCrLf & _
                      "¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Dim oEquipo_norm As New clsEquipos_Normas
                Call oEquipo_norm.Eliminar(Replace(lstNormas.SelectedItem.Key, "'", ""), PK)
                
                cargar_normas
                Set oEquipo_norm = Nothing
            End If
        End If

    Else
        MsgBox "Debe seleccionar la norma que quiere eliminar.", vbInformation, App.Title
    End If
End Sub

'E0043-I
' Botón que abre el documento de a la norma asociado al equipo
Private Sub cmdAbrirDocNorma_Click()
   On Error GoTo cmdAbrirDocNorma_Click_Error

    If lstNormas.ListItems.Count = 0 Then
        MsgBox "Debe seleccionar la norma que quiere ver.", vbInformation, App.Title
        Exit Sub
    End If
    Dim oNorma As New clsCa_normas
    oNorma.mostrar lstNormas.ListItems(lstNormas.SelectedItem.Index).Text
'    oNorma.Carga (lstNormas.ListItems(lstNormas.SelectedItem.Index).Text)
'    If oNorma.getRUTA = "" Then
'        MsgBox "La NORMA no tiene asignado documento.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    Dim destino As String
'    destino = Replace(oNorma.getRUTA, "/", "\")
'    On Error GoTo fallo
'    If Dir(destino) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If
'    Exit Sub
'fallo:
'    MsgBox "Error al abrir el documento.", vbCritical, App.Title

   On Error GoTo 0
   Exit Sub

cmdAbrirDocNorma_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAbrirDocNorma_Click of Formulario frmEquipos_Detalle"
End Sub
'E0043-F

' procedimiento que carga los datos de la calibración
Public Sub datos_calibracion(lngEquipo As Long)
    Dim oEQC As New clsEquipos_calibracion
    
    If chkCon_Calibracion.value = Checked Then
        cmd(0).Enabled = True ' se activa el botón para ver las calibraciones
    Else
        cmd(0).Enabled = False
    End If
    If oEQC.total_calibraciones(lngEquipo) = 1 Then
        oEQC.Carga (lngEquipo)
        With oEQC
            'chkCon_Calibracion.value = Checked
            cmbPeriCalibracion.BoundText = .getPERIODICIDAD_ID
            FechaCalibracion = .getFECHA_PROXIMA
            cmbTipoCalibracion.BoundText = .getMODALIDAD_ID
            cmbCalibrador_interno.MostrarElemento .getCALIBRADOR_INTERNO_ID
            cmbCalibrador.MostrarElemento .getCALIBRADOR_EXTERNO_ID
        End With
        If UCase(cmbTipoCalibracion.Text) = "EXTERNA" Then
            cmbCalibrador.Visible = True
            cmbCalibrador_interno.Visible = False
        ElseIf UCase(cmbTipoCalibracion.Text) = "INTERNA" Then
            cmbCalibrador.Visible = False
            cmbCalibrador_interno.Visible = True
        Else
            cmbCalibrador.desactivar
            cmbCalibrador_interno.desactivar
        End If
    End If
    
    Set oEQC = Nothing

End Sub

' procedimiento que carga los datos de la verificación
Public Sub datos_verificacion(lngEquipo As Long)
    Dim oEQV As New clsEquipos_verificacion
    Dim lngID_Verificacion As Long
    
    chkCon_Verificacion.Enabled = True
    If chkCon_Verificacion.value = Checked Then
        cmd(1).Enabled = True ' se activa el botón para ver las verificaciones
    Else
        cmd(1).Enabled = False
    End If
    If oEQV.total_verificaciones(lngEquipo) > 0 Then ' si tiene verificaciones
        ' se obtiene aquella con fecha próxima de verificación más cercana del equipo
        lngID_Verificacion = oEQV.Verificacion_mas_cercana(lngEquipo)
        
        oEQV.Carga (lngID_Verificacion) ' y se carga
        With oEQV
            cmbPeriVerificacion.BoundText = .getPERIODICIDAD_ID
            fechaProximaVerificacion = .getFECHA_PROXIMA
            cmbTipoVerificacion.BoundText = .getMODALIDAD_ID
            cmbVerificador_interno.MostrarElemento .getVERIFICADOR_INTERNO_ID
            cmbVerificador.MostrarElemento .getVERIFICADOR_EXTERNO_ID
        End With
        If UCase(cmbTipoVerificacion.Text) = "EXTERNA" Then
            cmbVerificador.Visible = True
            cmbVerificador_interno.Visible = False
        ElseIf UCase(cmbTipoVerificacion.Text) = "INTERNA" Then
            cmbVerificador.Visible = False
            cmbVerificador_interno.Visible = True
        Else
            cmbVerificador.desactivar
            cmbVerificador_interno.desactivar
        End If
        
    Else ' si no tiene verificaciones
        cmbPeriVerificacion.BoundText = 0
        fechaProximaVerificacion.value = "01-01-1900"
        cmbTipoVerificacion.BoundText = 0
        cmbVerificador.Limpiar
        cmbVerificador.desactivar
        cmbVerificador_interno.Limpiar
        cmbVerificador_interno.desactivar
    End If
    
    Set oEQV = Nothing

End Sub

' procedimiento que carga los datos del mantenimiento
Public Sub datos_mantenimiento(lngEquipo As Long)

    chkCon_Mantenimiento.Enabled = True
    If chkCon_Mantenimiento.value = Checked Then
        cmd(2).Enabled = True ' se activa el botón para ver el mantenimiento
    Else
        cmd(2).Enabled = False
    End If
        
    Call cargar_datos_mantenimiento_mas_proximo
        
    If UCase(cmbTipoMantenimiento.Text) = "EXTERNA" Then
        cmbResponsable.Visible = True
        cmbResponsable_interno.Visible = False
    ElseIf UCase(cmbTipoMantenimiento.Text) = "INTERNA" Then
        cmbResponsable.Visible = False
        cmbResponsable_interno.Visible = True
    End If
    cmbResponsable.desactivar
    cmbResponsable_interno.desactivar
        
End Sub

' procedimiento que elimina las calibraciónes y verificaciónes que no se hayan aceptado
' desde el formulario detalle del equipo
Private Sub descargar_formulario()
    Dim oEQ As New clsEquipos
    oEQ.Carga (PK)
    
    ' calibración
    Dim oEQC As New clsEquipos_calibracion
    If oEQ.getCON_CALIBRACION = 0 Then ' Si no tenía con_calibración
        oEQC.Eliminar_no_efectiva (PK) ' se elimina la no efectiva
    End If
    Set oEQC = Nothing
    
    ' verificación
    Dim oEQV As New clsEquipos_verificacion
    If oEQ.getCON_VERIFICACION = 0 Then ' Si no tenía con_verificacion
        oEQV.Eliminar_no_efectivas (PK) ' se eliminan las no efectivas
    End If
    Set oEQV = Nothing
    
    Set oEQ = Nothing
End Sub

' procedimiento que obtiene los datos del mantenimiento más próximo
Private Sub cargar_datos_mantenimiento_mas_proximo()
    Dim oEQM As New clsEquipos_mantenimiento
    Dim lngPeriodicidad As Long, lngModalidad As Long
    Dim strFecha_proxima As String
    Dim lngResponsable_interno As Long, lngResponsable_externo As Long
    Dim strMto As String
    
    If oEQM.Carga(PK) Then ' si tiene mantenimiento
        With oEQM
            
            If Format(.getSEMANAL_FECHA, "yyyy-mm-dd") = "1900-01-01" Then
                If Format(.getMENSUAL_FECHA, "yyyy-mm-dd") = "1900-01-01" Then
                    strMto = "NINGUNO"
                Else
                    strMto = "MENSUAL"
                End If
            Else
                If Format(.getMENSUAL_FECHA, "yyyy-mm-dd") = "1900-01-01" Then
                    strMto = "SEMANAL"
                Else
                    If Format(.getSEMANAL_FECHA, "yyyy-mm-dd") <= Format(.getMENSUAL_FECHA, "yyyy-mm-dd") Then
                        strMto = "SEMANAL"
                    Else
                        strMto = "MENSUAL"
                    End If
                End If
            End If
            
            Select Case strMto
                Case "SEMANAL"
                    cmbPeriMantenimiento.BoundText = 2
                    fechaProximoMantenimiento = .getSEMANAL_FECHA
                    cmbTipoMantenimiento.BoundText = .getSEMANAL_MODALIDAD_ID
                    cmbResponsable_interno.MostrarElemento .getSEMANAL_RESPONSABLE_INTERNO_ID
                    cmbResponsable.MostrarElemento .getSEMANAL_RESPONSABLE_EXTERNO_ID
                    
                Case "MENSUAL"
                    cmbPeriMantenimiento.BoundText = 4
                    fechaProximoMantenimiento = .getMENSUAL_FECHA
                    cmbTipoMantenimiento.BoundText = .getMENSUAL_MODALIDAD_ID
                    cmbResponsable_interno.MostrarElemento .getMENSUAL_RESPONSABLE_INTERNO_ID
                    cmbResponsable.MostrarElemento .getMENSUAL_RESPONSABLE_EXTERNO_ID
                    
                Case "NINGUNO"
                    cmbPeriMantenimiento.BoundText = 0
                    fechaProximoMantenimiento = Format("1900-01-01", "yyyy-mm-dd")
                    cmbTipoMantenimiento.BoundText = 0
                    cmbResponsable_interno.MostrarElemento 0
                    cmbResponsable.MostrarElemento 0

            End Select
            
        End With
        
    Else ' si no tiene mantenimiento
        cmbPeriMantenimiento.BoundText = 0
        fechaProximoMantenimiento.value = Format("1900-01-01", "yyyy-mm-dd")
        cmbTipoMantenimiento.BoundText = 0
        cmbResponsable.Limpiar
        cmbResponsable.desactivar
        cmbResponsable_interno.Limpiar
        cmbResponsable_interno.desactivar
    End If

    Set oEQM = Nothing
    
End Sub
