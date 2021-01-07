VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmEquiposAsignacionRapida 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Asignación Rápida para Equipos de Ensayo"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15465
   Icon            =   "frmEquiposAsignacionRapida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEsNadcap 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Es NADCAP Asignado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11175
      TabIndex        =   60
      Top             =   1650
      Width           =   2775
   End
   Begin VB.CheckBox chkVer_alguno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Verificación marcada, pero algun dato en blanco"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11175
      TabIndex        =   47
      Top             =   1410
      Width           =   5115
   End
   Begin VB.CheckBox chkCal_alguno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con calibracion marcada, pero algun dato en blanco"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11175
      TabIndex        =   46
      Top             =   1170
      Width           =   5115
   End
   Begin VB.CheckBox chkTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Tipo Asignado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11175
      TabIndex        =   45
      Top             =   930
      Width           =   2115
   End
   Begin VB.CheckBox chkFamilia 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Familia asignada"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11175
      TabIndex        =   44
      Top             =   690
      Width           =   2115
   End
   Begin VB.CheckBox chkProcedencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Procedencia Asignada"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   59
      Top             =   1890
      Width           =   2805
   End
   Begin VB.CheckBox chkVer_Tipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Verificacion, Sin Tipo (Int/Ext)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8310
      TabIndex        =   58
      Top             =   1890
      Width           =   3405
   End
   Begin VB.CheckBox chkVer_Responsable 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Verificacion, Sin Responsable"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8310
      TabIndex        =   57
      Top             =   1650
      Width           =   3405
   End
   Begin VB.CheckBox chkVer_periodo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Verificacion, Sin Periodicidad"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8310
      TabIndex        =   56
      Top             =   1410
      Width           =   2850
   End
   Begin VB.CheckBox chkCal_Responsable 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Calibracion, Sin Responsable"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8310
      TabIndex        =   54
      Top             =   930
      Width           =   3405
   End
   Begin VB.CheckBox chkCal_periodo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Calibracion, Sin Periodicidad"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8310
      TabIndex        =   53
      Top             =   690
      Width           =   3435
   End
   Begin VB.CheckBox chkLocalizacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Localizacion Asignada"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   52
      Top             =   1650
      Width           =   2925
   End
   Begin VB.CheckBox chkSiAccesorios 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Que Pueden Ser Accesorios"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   51
      Top             =   1410
      Width           =   2985
   End
   Begin VB.CheckBox chkNoAccesorios 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Que NO pueden ser accesorios"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   50
      Top             =   1170
      Width           =   2985
   End
   Begin VB.CheckBox chkProveedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Proveedor Asignado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   49
      Top             =   930
      Width           =   2685
   End
   Begin VB.CheckBox chkFabricante 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin Fabricante"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   48
      Top             =   690
      Width           =   1995
   End
   Begin VB.CommandButton cmdAplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      Height          =   870
      Left            =   7260
      Picture         =   "frmEquiposAsignacionRapida.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5940
      Width           =   960
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   780
      Left            =   14310
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8910
      Width           =   1050
   End
   Begin VB.Frame fraPropiedad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Propiedades a Asignar a los Equipos en grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   30
      TabIndex        =   9
      Top             =   6870
      Width           =   15405
      Begin VB.TextBox txtFabricante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   62
         Top             =   1890
         Width           =   3825
      End
      Begin VB.CheckBox chkno_es_accesorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO ES ACCESORIO"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   37
         Top             =   2460
         Width           =   1875
      End
      Begin VB.CheckBox chkSin_Verificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quitar 'Con Verificación'"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11340
         TabIndex        =   36
         Top             =   150
         Width           =   2745
      End
      Begin VB.CheckBox chkCon_Verificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Con Verificación"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11340
         TabIndex        =   29
         Top             =   360
         Width           =   1995
      End
      Begin VB.CheckBox chkSin_Calibracion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quitar 'Con Calibración'"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6270
         TabIndex        =   28
         Top             =   180
         Width           =   2445
      End
      Begin VB.CheckBox chkCon_Calibracion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Con Calibración"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6270
         TabIndex        =   21
         Top             =   390
         Width           =   1785
      End
      Begin VB.CheckBox chkes_accesorio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ES ACCESORIO"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   14
         Top             =   2220
         Width           =   1875
      End
      Begin MSDataListLib.DataCombo cmbFamilia 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbTipoEquipo 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   570
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbNadCap 
         Height          =   315
         Left            =   1200
         TabIndex        =   15
         Top             =   1560
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbResponsable 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   900
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProcedencia 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   1230
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCal_Periodo 
         Height          =   315
         Left            =   6270
         TabIndex        =   22
         Top             =   600
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCal_Tipo 
         Height          =   315
         Left            =   6270
         TabIndex        =   23
         Top             =   930
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCal_Responsable 
         Height          =   315
         Left            =   6270
         TabIndex        =   24
         Top             =   1260
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbVer_Periodo 
         Height          =   315
         Left            =   11340
         TabIndex        =   30
         Top             =   570
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbVer_Tipo 
         Height          =   315
         Left            =   11340
         TabIndex        =   31
         Top             =   900
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbVer_Responsable 
         Height          =   315
         Left            =   11340
         TabIndex        =   32
         Top             =   1230
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbCal_Procedimiento 
         Height          =   330
         Left            =   6270
         TabIndex        =   38
         Top             =   1590
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVer_Procedimiento 
         Height          =   330
         Left            =   11340
         TabIndex        =   40
         Top             =   1560
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbLocalizacion 
         Height          =   315
         Left            =   1200
         TabIndex        =   63
         Top             =   2190
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProveedor 
         Height          =   315
         Left            =   1200
         TabIndex        =   65
         Top             =   2520
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   66
         Top             =   2625
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Localizacion"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   64
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricante"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   61
         Top             =   1950
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   3
         Left            =   10260
         TabIndex        =   41
         Top             =   1635
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   37
         Left            =   5190
         TabIndex        =   39
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   2
         Left            =   10230
         TabIndex        =   35
         Top             =   645
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   10230
         TabIndex        =   34
         Top             =   960
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Interno"
         Height          =   195
         Index           =   0
         Left            =   10230
         TabIndex        =   33
         Top             =   1290
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   38
         Left            =   5190
         TabIndex        =   27
         Top             =   675
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   47
         Left            =   5160
         TabIndex        =   26
         Top             =   990
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Interno"
         Height          =   195
         Index           =   48
         Left            =   5190
         TabIndex        =   25
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedencia"
         Height          =   195
         Index           =   71
         Left            =   120
         TabIndex        =   20
         Top             =   1335
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   70
         Left            =   120
         TabIndex        =   19
         Top             =   975
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         Height          =   195
         Index           =   34
         Left            =   120
         TabIndex        =   18
         Top             =   1650
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Área"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Width           =   330
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Equipo"
         Height          =   195
         Index           =   36
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdQuitar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      Height          =   885
      Left            =   7260
      Picture         =   "frmEquiposAsignacionRapida.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3390
      Width           =   975
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   915
      Left            =   7260
      Picture         =   "frmEquiposAsignacionRapida.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2460
      Width           =   975
   End
   Begin VB.TextBox txtFiltro 
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
      Height          =   315
      Left            =   1035
      TabIndex        =   5
      Top             =   1125
      Width           =   4065
   End
   Begin MSComctlLib.ListView origen 
      Height          =   4395
      Left            =   30
      TabIndex        =   2
      Top             =   2430
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4395
      Left            =   8250
      TabIndex        =   3
      Top             =   2430
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13230796
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CheckBox chkCal_Tipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Con Calibracion, Sin Tipo (Int/Ext)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8310
      TabIndex        =   55
      Top             =   1170
      Width           =   2820
   End
   Begin MSDataListLib.DataCombo cmbResponsable2 
      Height          =   315
      Left            =   1035
      TabIndex        =   69
      Top             =   765
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbFamilia2 
      Height          =   315
      Left            =   1035
      TabIndex        =   71
      Top             =   1485
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin XtremeSuiteControls.PushButton cmdAnadirCalibracion 
      Height          =   255
      Left            =   30
      TabIndex        =   73
      Top             =   2160
      Width           =   1545
      _Version        =   851970
      _ExtentX        =   2725
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Marcar Todos"
      Appearance      =   5
      Picture         =   "frmEquiposAsignacionRapida.frx":2328
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   255
      Left            =   8250
      TabIndex        =   74
      Top             =   2160
      Width           =   1545
      _Version        =   851970
      _ExtentX        =   2725
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Marcar Todos"
      Appearance      =   5
      Picture         =   "frmEquiposAsignacionRapida.frx":8B8A
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Área"
      Height          =   195
      Index           =   8
      Left            =   45
      TabIndex        =   72
      Top             =   1530
      Width           =   330
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsable"
      Height          =   195
      Index           =   7
      Left            =   45
      TabIndex        =   70
      Top             =   795
      Width           =   930
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7260
      TabIndex        =   68
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encontrados"
      Height          =   255
      Left            =   7260
      TabIndex        =   67
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "Equipos a los que asignar las mismas características"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9855
      TabIndex        =   6
      Top             =   2160
      Width           =   5610
   End
   Begin VB.Label lblCap 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   4
      Top             =   1170
      Width           =   705
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Asignación Rápida para Equipos de Ensayo"
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
      TabIndex        =   1
      Top             =   0
      Width           =   4620
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ventana para asignar propiedades a un conjunto de equipos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   4275
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   15480
   End
End
Attribute VB_Name = "frmEquiposAsignacionRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs As ADODB.Recordset
Private oEq As New clsEquipos

Private lista_seleccionados As String
Private Sub cabecera()
    With origen.ColumnHeaders
        .Add , , "Nº Eq.", 900, lvwColumnLeft
        .Add , , "Equipo", 5900, lvwColumnLeft
    End With
    With lista.ColumnHeaders
        .Add , , "Nº Eq.", 900, lvwColumnLeft
        .Add , , "Equipo", 5900, lvwColumnLeft
    End With
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora

    oDeco.cargar_combo cmbFamilia, DECODIFICADORA.EQ_FAMILIAS
    oDeco.cargar_combo cmbFamilia2, DECODIFICADORA.EQ_FAMILIAS
    oDeco.cargar_combo cmbTipoEquipo, DECODIFICADORA.EQ_TIPOS_EQUIPO
    cargar_combo cmbResponsable, New clsUsuarios
    cargar_combo cmbResponsable2, New clsUsuarios
    oDeco.cargar_combo cmbProcedencia, DECODIFICADORA.EQ_PROCEDENCIA_EQUIPOS
    oDeco.cargar_combo cmbNadCap, DECODIFICADORA.EQ_SINO
    
    oDeco.cargar_combo cmbLocalizacion, DECODIFICADORA.EQ_SITUACIONES
    cargar_combo cmbProveedor, New clsProveedor
    oDeco.cargar_combo cmbCal_Periodo, DECODIFICADORA.EQ_periodicidad
    oDeco.cargar_combo cmbVer_Periodo, DECODIFICADORA.EQ_periodicidad
    
    oDeco.cargar_combo cmbCal_Tipo, DECODIFICADORA.EQ_TIPO_CALIBRACION
    oDeco.cargar_combo cmbVer_Tipo, DECODIFICADORA.EQ_TIPO_CALIBRACION
    
    cargar_combo cmbVer_Responsable, New clsUsuarios
    cargar_combo cmbCal_Responsable, New clsUsuarios
    
    llenar_combo cmbCal_Procedimiento, New clsCa_documentos, 0, frmCA_Documento, " codigo like '%PNT C%'"
    llenar_combo cmbVer_Procedimiento, New clsCa_documentos, 0, frmCA_Documento, " codigo like '%PNT C%'"
       
    cmbCal_Procedimiento.desactivar
    cmbVer_Procedimiento.desactivar

End Sub

Private Sub CARGAR_ORIGEN()


    Set rs = oEq.Listado_asignacion_rapida(cmbResponsable2.BoundText, cmbFamilia2.BoundText, txtFiltro.Text, chkFamilia.Value, chkTipo.Value, chkCal_alguno.Value, chkVer_alguno.Value, chkEsNadcap.Value, chkFabricante.Value, chkProveedor.Value, chkSiAccesorios.Value, chkNoAccesorios.Value, chkLocalizacion.Value, chkProcedencia.Value, chkCal_periodo.Value, chkCal_Responsable.Value, chkCal_Tipo.Value, chkVer_periodo.Value, chkVer_Responsable.Value, chkVer_Tipo.Value)

    origen.ListItems.Clear
    lblCap(1).Caption = "0"
    
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    
    While Not rs.EOF
        With origen.ListItems.Add(, , rs("id_equipo"))
            .SubItems(1) = rs("nombre")
        End With
        rs.MoveNext
    Wend

    lblCap(1).Caption = origen.ListItems.Count

End Sub


Private Function comprobar_cal_ver_completo() As Boolean
comprobar_cal_ver_completo = False

If chkCon_Calibracion.Value = vbChecked Then
    If getDataComboSel(cmbCal_Periodo) = -1 Or cmbCal_Procedimiento.getPK_SALIDA <= 0 Or getDataComboSel(cmbCal_Responsable) = -1 Or getDataComboSel(cmbCal_Tipo) = -1 Then
        MsgBox "Para asignar datos Con Calibración, debe rellenar todos los datos para calibracion."
        Exit Function
    End If
End If

If chkCon_Verificacion.Value = vbChecked Then
    If getDataComboSel(cmbVer_Periodo) = -1 Or cmbVer_Procedimiento.getPK_SALIDA <= 0 Or getDataComboSel(cmbVer_Responsable) = -1 Or getDataComboSel(cmbVer_Tipo) = -1 Then
        MsgBox "Para asignar datos Con verificación, debe rellenar todos los datos para verificacion."
        Exit Function
    End If
End If

comprobar_cal_ver_completo = True

End Function

Private Sub des_activar_calibracion()

    Dim val As Boolean
    
    val = (chkCon_Calibracion.Value = vbChecked)
    
    cmbCal_Procedimiento.activar
    
    If Not val Then
        cmbCal_Procedimiento.limpiar
        cmbCal_Procedimiento.desactivar
        
        cmbCal_Periodo.BoundText = ""
        cmbCal_Responsable.BoundText = ""
        cmbCal_Tipo.BoundText = ""
    End If
    
    cmbCal_Periodo.Enabled = val
    cmbCal_Responsable.Enabled = val
    cmbCal_Tipo.Enabled = val


End Sub

Private Sub des_activar_verificacion()
Dim val As Boolean
    
    val = (chkCon_Verificacion.Value = vbChecked)
    
    cmbVer_Procedimiento.activar
    
    If Not val Then
        cmbVer_Procedimiento.limpiar
        cmbVer_Procedimiento.desactivar
        
        cmbVer_Periodo.BoundText = ""
        cmbVer_Responsable.BoundText = ""
        cmbVer_Tipo.BoundText = ""
    End If
    
    cmbVer_Periodo.Enabled = val
    cmbVer_Responsable.Enabled = val
    cmbVer_Tipo.Enabled = val

End Sub

Private Sub cmbFamilia2_Change()
    CARGAR_ORIGEN
End Sub
Private Sub cmbResponsable2_Change()
    CARGAR_ORIGEN
End Sub
Private Sub cmdAnadirCalibracion_Click()
    Dim i As Integer
    For i = 1 To origen.ListItems.Count
        origen.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdAplicar_Click()
    Dim familia As Long, tipo As Long, responsable As Long, Procedencia As Long, nadcap As Long
    Dim accesorio As Integer, calibracion As Integer, verificacion As Integer
    Dim Cal_periodo As Long, Cal_Tipo As Long, Cal_Responsable As Long, cal_procedimiento As Long
    Dim Ver_periodo As Long, ver_tipo As Long, Ver_Responsable As Long, ver_procedimiento As Long
    Dim proveedor As Long, localizacion As Long
        
    Dim res As Boolean
    
    If Not comprobar_cal_ver_completo() Then Exit Sub
    
    'INICIALIZA POR DEFECTO
    familia = -1
    tipo = -1
    responsable = -1
    Procedencia = -1
    nadcap = -1
    localizacion = -1
    proveedor = -1
    
    ' recoge los datos para modificarlos
    If Trim(cmbFamilia.BoundText) <> "" Then familia = CLng(cmbFamilia.BoundText)
    If Trim(cmbTipoEquipo.BoundText) <> "" Then tipo = CLng(cmbTipoEquipo.BoundText)
    If Trim(cmbResponsable.BoundText) <> "" Then responsable = CLng(cmbResponsable.BoundText)
    If Trim(cmbProcedencia.BoundText) <> "" Then Procedencia = CLng(cmbProcedencia.BoundText)
    If Trim(cmbNadCap.BoundText) <> "" Then nadcap = CLng(cmbNadCap.BoundText)
    If Trim(cmbProveedor.BoundText) <> "" Then proveedor = CLng(cmbProveedor.BoundText)
    If Trim(cmbLocalizacion.BoundText) <> "" Then localizacion = CLng(cmbLocalizacion.BoundText)
       
    accesorio = IIf(chkes_accesorio.Value + chkno_es_accesorio.Value = 0, -1, chkes_accesorio.Value)
    calibracion = IIf(chkCon_Calibracion.Value + chkSin_Calibracion.Value = 0, -1, chkCon_Calibracion.Value)
    verificacion = IIf(chkCon_Verificacion.Value + chkSin_Verificacion.Value = 0, -1, chkCon_Verificacion.Value)
    
    'sobre calibraciones
    Cal_periodo = IIf(Trim(cmbCal_Periodo.BoundText) = "", -1, cmbCal_Periodo.BoundText)
    cal_procedimiento = IIf(cmbCal_Procedimiento.getPK_SALIDA = 0, -1, CLng(cmbCal_Procedimiento.getPK_SALIDA))
    Cal_Responsable = IIf(Trim(cmbCal_Responsable.BoundText) = "", -1, cmbCal_Responsable.BoundText)
    Cal_Tipo = IIf(Trim(cmbCal_Tipo.BoundText) = "", -1, cmbCal_Tipo.BoundText)
    
    'sobre verificaciones
    Ver_periodo = IIf(Trim(cmbVer_Periodo.BoundText) = "", -1, cmbVer_Periodo.BoundText)
    ver_procedimiento = IIf(cmbVer_Procedimiento.getPK_SALIDA = 0, -1, cmbVer_Procedimiento.getPK_SALIDA)
    Ver_Responsable = IIf(Trim(cmbVer_Responsable.BoundText) = "", -1, cmbVer_Responsable.BoundText)
    ver_tipo = IIf(Trim(cmbVer_Tipo.BoundText) = "", -1, cmbVer_Tipo.BoundText)
        
    
    res = oEq.AplicarAsignacionRapida(lista_seleccionados, familia, tipo, responsable, Procedencia, nadcap, accesorio, calibracion, verificacion, _
    Cal_periodo, cal_procedimiento, Cal_Responsable, Cal_Tipo, _
    Ver_periodo, ver_procedimiento, Ver_Responsable, ver_tipo, Trim(txtFabricante.Text), localizacion, proveedor)


    If res Then MsgBox "Fin del proceso de Asignación"

End Sub

Private Sub cmdQuitar_Click()
    Dim x As Long
    Dim total As Long
    If lista.ListItems.Count = 0 Then Exit Sub
    
    total = lista.ListItems.Count
    x = 1
    While x <= total
        If lista.ListItems(x).Checked Then
            ' lo quita de la lista de seleccionados
            lista_seleccionados = Replace(lista_seleccionados, ":" & lista.ListItems(x) & ":", "")
            With origen.ListItems.Add(, , lista.ListItems(x))
                .SubItems(1) = lista.ListItems(x).SubItems(1)
            End With
            lista.ListItems.Remove x
            total = total - 1
        Else
            x = x + 1
        End If
    Wend
End Sub

Private Sub Modificar()
    
    On Error GoTo Modificar_Error
    
    Dim objfrm As New frmEquipoEdicion
    Dim lngid As Long
    Dim objEquipo As New clsEquipos
    
    lngid = CLng(origen.ListItems(origen.selectedItem.Index))
    If lngid <= 0 Then Exit Sub
    
    Call objEquipo.Carga(lngid)
    
    Set objfrm.EQUIPO = objEquipo
    
    objfrm.TipoEdicion = visualizar
    
    'If objEquipo.getALTA_BAJA = 1 Then
       
    'Else
    '    objfrm.TipoEdicion = EDICION
    'End If
    
    objfrm.Show vbModal
    
    'If objfrm.Resultado Then
        'Call cargar_lista
        'cargar_linea_lista lngid, lista.SelectedItem.Index
    'End If
    
    Unload objfrm
    Set objfrm = Nothing
    
    On Error GoTo 0
        Exit Sub
Modificar_Error:
        'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmEquipoListado"

End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdAnadir_Click()
Dim x As Long, ID As String
Dim total As Long

If origen.ListItems.Count = 0 Then Exit Sub

total = origen.ListItems.Count

x = 1
While x <= total
    If origen.ListItems(x).Checked Then
        ID = origen.ListItems(x).Text
        If InStr(1, lista_seleccionados, ":" & ID & ":") = 0 Then
            lista_seleccionados = lista_seleccionados & ":" & ID & ":"
            With lista.ListItems.Add(, , ID)
                .SubItems(1) = origen.ListItems(x).SubItems(1)
            End With
            origen.ListItems.Remove x
            total = total - 1
        End If
    Else
        x = x + 1
    End If
    
Wend




End Sub

Private Sub chkCal_alguno_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkCal_periodo_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkCal_Responsable_Click()
CARGAR_ORIGEN
End Sub


Private Sub chkCal_Tipo_Click()
CARGAR_ORIGEN
End Sub


Private Sub chkCon_Calibracion_Click()
    
    If chkCon_Calibracion.Value = vbChecked Then chkSin_Calibracion.Value = vbUnchecked

    des_activar_calibracion
    
    
End Sub

Private Sub chkCon_Verificacion_Click()

    If chkSin_Verificacion.Value = vbChecked Then chkCon_Verificacion.Value = vbUnchecked
    
    des_activar_verificacion
    
End Sub
Private Sub chkes_accesorio_Click()
If chkes_accesorio.Value = vbChecked Then chkno_es_accesorio.Value = Unchecked
End Sub

Private Sub chkEsNadcap_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkFabricante_Click()
CARGAR_ORIGEN
End Sub


Private Sub chkFamilia_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkLocalizacion_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkno_es_accesorio_Click()
If chkno_es_accesorio.Value = vbChecked Then chkes_accesorio.Value = Unchecked
End Sub


Private Sub chkNoAccesorios_Click()
If chkNoAccesorios.Value = vbChecked Then chkSiAccesorios.Value = vbUnchecked


CARGAR_ORIGEN
End Sub

Private Sub chkProcedencia_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkProveedor_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkSiAccesorios_Click()
If chkSiAccesorios.Value = vbChecked Then chkNoAccesorios.Value = vbUnchecked
CARGAR_ORIGEN
End Sub

Private Sub chkSin_Calibracion_Click()

    If chkSin_Calibracion.Value = vbChecked Then chkCon_Calibracion.Value = vbUnchecked
    
    des_activar_calibracion
    
    

End Sub

Private Sub chkSin_Verificacion_Click()

    If chkSin_Verificacion.Value = vbChecked Then chkCon_Verificacion.Value = vbUnchecked
    
    des_activar_verificacion
    
End Sub


Private Sub chkTipo_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkVer_alguno_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkVer_periodo_Click()
CARGAR_ORIGEN
End Sub

Private Sub chkVer_Responsable_Click()
CARGAR_ORIGEN
End Sub


Private Sub chkVer_Tipo_Click()
CARGAR_ORIGEN
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    CARGAR_ORIGEN
    cargar_combos
End Sub


Private Sub origen_DblClick()
Modificar
End Sub


Private Sub PushButton1_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub txtfiltro_Change()
    CARGAR_ORIGEN
End Sub

