VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEquipoPlanMtoEdicion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Equipos - Planes de Mantenimiento - Detalle"
   ClientHeight    =   9675
   ClientLeft      =   1245
   ClientTop       =   2055
   ClientWidth     =   14535
   Icon            =   "frmEquipoPlanMtoEdicion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabPlanif 
      Height          =   4245
      Left            =   30
      TabIndex        =   400
      Top             =   2070
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "    Plan Mensual/Semanal/Diario    "
      TabPicture(0)   =   "frmEquipoPlanMtoEdicion.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "    Plan Anual    "
      TabPicture(1)   =   "frmEquipoPlanMtoEdicion.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmPlanificacion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "    Simulador Fechas Mantenimiento Anual    "
      TabPicture(2)   =   "frmEquipoPlanMtoEdicion.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAno"
      Tab(2).Control(1)=   "lstFechas"
      Tab(2).Control(2)=   "cmdGenerarFechasAno"
      Tab(2).Control(3)=   "lblCampos(3)"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtAno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68640
         MaxLength       =   4
         TabIndex        =   461
         Text            =   "2010"
         Top             =   690
         Width           =   615
      End
      Begin VB.ListBox lstFechas 
         Columns         =   4
         Height          =   2985
         ItemData        =   "frmEquipoPlanMtoEdicion.frx":05DE
         Left            =   -74970
         List            =   "frmEquipoPlanMtoEdicion.frx":05E0
         TabIndex        =   459
         Top             =   1170
         Width           =   14415
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   3900
         Left            =   0
         TabIndex        =   402
         Top             =   330
         Width           =   14505
         Begin VB.Frame fraSabDomExep 
            BackColor       =   &H00C0C0C0&
            Height          =   825
            Left            =   10470
            TabIndex        =   469
            Top             =   90
            Width           =   3945
            Begin VB.OptionButton optEvitarExcep 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Evitar Sábados, Domingos y Excepciones"
               Height          =   255
               Left            =   90
               TabIndex        =   471
               Top             =   180
               Width           =   3765
            End
            Begin VB.OptionButton optTrasladarExcep 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Trasladar Sábados, Domingos y Excepciones"
               Height          =   255
               Left            =   90
               TabIndex        =   470
               Top             =   480
               Value           =   -1  'True
               Width           =   3525
            End
         End
         Begin VB.OptionButton optTipoPlan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Calendario"
            Height          =   255
            Index           =   4
            Left            =   420
            TabIndex        =   468
            Top             =   3420
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmdEliminarExcepcion 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   14040
            Picture         =   "frmEquipoPlanMtoEdicion.frx":05E2
            Style           =   1  'Graphical
            TabIndex        =   467
            ToolTipText     =   "Eliminar accesorio"
            Top             =   1080
            Width           =   285
         End
         Begin VB.CommandButton cmdAnadirExcepcion 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   13710
            Picture         =   "frmEquipoPlanMtoEdicion.frx":0776
            Style           =   1  'Graphical
            TabIndex        =   466
            ToolTipText     =   "Añadir accesorio"
            Top             =   1080
            Width           =   285
         End
         Begin VB.ComboBox cmbMes 
            Height          =   315
            ItemData        =   "frmEquipoPlanMtoEdicion.frx":099B
            Left            =   11430
            List            =   "frmEquipoPlanMtoEdicion.frx":09C6
            Style           =   2  'Dropdown List
            TabIndex        =   465
            Top             =   1050
            Width           =   2175
         End
         Begin VB.ComboBox cmbDia 
            Height          =   315
            ItemData        =   "frmEquipoPlanMtoEdicion.frx":0A2F
            Left            =   10500
            List            =   "frmEquipoPlanMtoEdicion.frx":0AA6
            Style           =   2  'Dropdown List
            TabIndex        =   464
            Top             =   1050
            Width           =   885
         End
         Begin VB.ListBox lstExcepciones 
            Columns         =   2
            Height          =   2310
            ItemData        =   "frmEquipoPlanMtoEdicion.frx":0B1D
            Left            =   10500
            List            =   "frmEquipoPlanMtoEdicion.frx":0B1F
            Style           =   1  'Checkbox
            TabIndex        =   463
            Top             =   1410
            Width           =   3855
         End
         Begin VB.OptionButton optTipoPlan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Semanal"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   458
            Top             =   1830
            Width           =   945
         End
         Begin VB.Frame fraPorSemanas 
            BackColor       =   &H00C0C0C0&
            Caption         =   "De cada Semana, los días"
            Height          =   915
            Left            =   1980
            TabIndex        =   450
            Top             =   2850
            Visible         =   0   'False
            Width           =   8385
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "LUNES"
               Height          =   435
               Index           =   1
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   457
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "MARTES"
               Height          =   435
               Index           =   2
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   456
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "MIERCOLES"
               Height          =   435
               Index           =   3
               Left            =   3420
               Style           =   1  'Graphical
               TabIndex        =   455
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "JUEVES"
               Height          =   435
               Index           =   4
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   454
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "VIERNES"
               Height          =   435
               Index           =   5
               Left            =   6660
               Style           =   1  'Graphical
               TabIndex        =   453
               Top             =   360
               Width           =   1575
            End
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "SABADO"
               Height          =   435
               Index           =   6
               Left            =   5970
               Style           =   1  'Graphical
               TabIndex        =   452
               Top             =   840
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.CheckBox chkPorSemana 
               Caption         =   "DOMINGO"
               Height          =   435
               Index           =   7
               Left            =   7140
               Style           =   1  'Graphical
               TabIndex        =   451
               Top             =   840
               Visible         =   0   'False
               Width           =   1125
            End
         End
         Begin VB.OptionButton optTipoPlan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Diario"
            Height          =   255
            Index           =   3
            Left            =   420
            TabIndex        =   449
            Top             =   2970
            Width           =   945
         End
         Begin VB.OptionButton optTipoPlan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mensual"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   448
            Top             =   720
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.Frame fraPorMesDias 
            BackColor       =   &H00C0C0C0&
            Caption         =   "De cada Mes, los días"
            Height          =   2715
            Left            =   6510
            TabIndex        =   416
            Top             =   90
            Width           =   3855
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "31"
               Height          =   435
               Index           =   31
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   447
               Top             =   2190
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "30"
               Height          =   435
               Index           =   30
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   446
               Top             =   2190
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "29"
               Height          =   435
               Index           =   29
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   445
               Top             =   2190
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "28"
               Height          =   435
               Index           =   28
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   444
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "27"
               Height          =   435
               Index           =   27
               Left            =   2730
               Style           =   1  'Graphical
               TabIndex        =   443
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "26"
               Height          =   435
               Index           =   26
               Left            =   2220
               Style           =   1  'Graphical
               TabIndex        =   442
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "25"
               Height          =   435
               Index           =   25
               Left            =   1710
               Style           =   1  'Graphical
               TabIndex        =   441
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "24"
               Height          =   435
               Index           =   24
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   440
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "23"
               Height          =   435
               Index           =   23
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   439
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "22"
               Height          =   435
               Index           =   22
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   438
               Top             =   1710
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "21"
               Height          =   435
               Index           =   21
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   437
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "20"
               Height          =   435
               Index           =   20
               Left            =   2730
               Style           =   1  'Graphical
               TabIndex        =   436
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "19"
               Height          =   435
               Index           =   19
               Left            =   2220
               Style           =   1  'Graphical
               TabIndex        =   435
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "18"
               Height          =   435
               Index           =   18
               Left            =   1710
               Style           =   1  'Graphical
               TabIndex        =   434
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "17"
               Height          =   435
               Index           =   17
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   433
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "16"
               Height          =   435
               Index           =   16
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   432
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "15"
               Height          =   435
               Index           =   15
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   431
               Top             =   1230
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "14"
               Height          =   435
               Index           =   14
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   430
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "13"
               Height          =   435
               Index           =   13
               Left            =   2730
               Style           =   1  'Graphical
               TabIndex        =   429
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "12"
               Height          =   435
               Index           =   12
               Left            =   2220
               Style           =   1  'Graphical
               TabIndex        =   428
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "11"
               Height          =   435
               Index           =   11
               Left            =   1710
               Style           =   1  'Graphical
               TabIndex        =   427
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "10"
               Height          =   435
               Index           =   10
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   426
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "9"
               Height          =   435
               Index           =   9
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   425
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "8"
               Height          =   435
               Index           =   8
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   424
               Top             =   750
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "7"
               Height          =   435
               Index           =   7
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   423
               Top             =   270
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "6"
               Height          =   435
               Index           =   6
               Left            =   2730
               Style           =   1  'Graphical
               TabIndex        =   422
               Top             =   270
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "5"
               Height          =   435
               Index           =   5
               Left            =   2220
               Style           =   1  'Graphical
               TabIndex        =   421
               Top             =   270
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "4"
               Height          =   435
               Index           =   4
               Left            =   1710
               Style           =   1  'Graphical
               TabIndex        =   420
               Top             =   270
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "3"
               Height          =   435
               Index           =   3
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   419
               Top             =   270
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "2"
               Height          =   435
               Index           =   2
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   418
               Top             =   270
               Width           =   435
            End
            Begin VB.CheckBox chkDiaPorMes 
               Caption         =   "1"
               Height          =   435
               Index           =   1
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   417
               Top             =   270
               Width           =   435
            End
         End
         Begin VB.Frame fraPorMeses 
            BackColor       =   &H00C0C0C0&
            Caption         =   "En los Meses"
            Height          =   2715
            Left            =   1980
            TabIndex        =   403
            Top             =   90
            Width           =   4275
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "DICIEMBRE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   2340
               TabIndex        =   415
               Top             =   2310
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "OCTUBRE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   2340
               TabIndex        =   414
               Top             =   1470
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "AGOSTO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   2340
               TabIndex        =   413
               Top             =   630
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "JUNIO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   412
               Top             =   2310
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ABRIL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   411
               Top             =   1470
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "FEBRERO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   410
               Top             =   630
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "NOVIEMBRE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   2340
               TabIndex        =   409
               Top             =   1890
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "SEPTIEMBRE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   2340
               TabIndex        =   408
               Top             =   1050
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "JULIO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   2340
               TabIndex        =   407
               Top             =   210
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "MAYO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   406
               Top             =   1890
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "MARZO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   405
               Top             =   1050
               Width           =   1605
            End
            Begin VB.CheckBox chkPorMes 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ENERO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   404
               Top             =   240
               Width           =   1095
            End
         End
      End
      Begin VB.Frame frmPlanificacion 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   3870
         Left            =   -75000
         TabIndex        =   401
         Top             =   330
         Width           =   14475
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ENERO"
            Height          =   1905
            Index           =   1
            Left            =   60
            TabIndex        =   367
            Top             =   30
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   1
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   2
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   2
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   3
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   4
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   5
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   6
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   7
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   8
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   9
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   10
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   11
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   12
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   13
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   14
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   15
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   16
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   17
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   18
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   19
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   20
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   21
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   22
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   23
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   24
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   25
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   26
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   27
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   28
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   29
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   30
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   31
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "FEBRERO"
            Height          =   1905
            Index           =   2
            Left            =   2460
            TabIndex        =   368
            Top             =   30
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   60
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   59
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   58
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   57
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   56
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   55
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   54
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   53
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   52
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   51
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   50
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   49
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   48
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   47
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   46
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   45
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   44
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   43
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   42
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   41
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   40
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   39
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   38
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   37
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   36
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   35
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   34
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   33
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   32
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   210
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "MARZO"
            Height          =   1905
            Index           =   3
            Left            =   4860
            TabIndex        =   369
            Top             =   30
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   61
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   62
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   63
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   64
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   64
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   65
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   65
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   66
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   67
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   68
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   68
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   69
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   69
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   70
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   71
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   71
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   72
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   73
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   74
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   74
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   75
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   76
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   77
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   77
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   79
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   79
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   78
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   78
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   80
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   81
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   82
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   83
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   84
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   85
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   86
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   87
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   87
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   88
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   88
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   89
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   89
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   90
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   90
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   91
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   91
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ABRIL"
            Height          =   1905
            Index           =   4
            Left            =   7260
            TabIndex        =   370
            Top             =   30
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   92
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   92
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   93
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   93
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   94
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   95
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   95
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   96
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   96
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   97
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   97
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   98
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   98
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   99
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   99
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   100
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   100
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   101
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   101
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   102
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   102
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   103
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   103
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   104
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   104
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   105
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   106
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   107
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   108
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   108
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   109
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   109
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   110
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   110
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   111
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   111
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   112
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   112
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   113
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   113
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   114
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   114
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   115
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   115
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   116
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   116
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   117
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   117
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   118
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   118
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   119
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   119
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   120
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   120
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   121
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   121
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "MAYO"
            Height          =   1905
            Index           =   5
            Left            =   9660
            TabIndex        =   371
            Top             =   30
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   122
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   122
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   123
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   123
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   124
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   124
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   125
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   125
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   126
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   126
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   127
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   128
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   128
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   129
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   129
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   130
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   130
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   131
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   132
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   133
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   133
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   134
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   134
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   135
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   135
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   136
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   136
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   137
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   137
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   138
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   138
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   139
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   139
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   140
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   140
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   141
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   141
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   142
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   142
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   143
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   143
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   144
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   144
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   145
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   145
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   146
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   146
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   147
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   147
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   148
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   148
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   149
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   149
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   150
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   150
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   151
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   151
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   152
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   152
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "JUNIO"
            Height          =   1905
            Index           =   6
            Left            =   12060
            TabIndex        =   372
            Top             =   30
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   153
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   153
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   154
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   154
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   155
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   155
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   156
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   156
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   157
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   157
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   158
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   158
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   159
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   159
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   160
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   160
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   161
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   161
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   162
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   162
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   163
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   163
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   164
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   164
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   165
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   165
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   166
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   166
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   167
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   167
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   168
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   168
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   169
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   169
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   170
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   170
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   171
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   171
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   172
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   172
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   173
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   173
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   174
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   174
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   175
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   175
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   176
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   177
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   177
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   178
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   178
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   179
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   179
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   180
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   180
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   181
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   181
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   182
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   182
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "JULIO"
            Height          =   1905
            Index           =   7
            Left            =   60
            TabIndex        =   373
            Top             =   1950
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   212
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   212
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   211
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   211
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   210
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   210
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   209
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   209
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   208
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   208
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   207
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   207
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   206
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   206
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   205
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   205
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   204
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   204
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   203
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   203
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   202
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   202
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   201
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   201
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   200
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   200
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   199
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   199
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   198
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   198
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   197
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   197
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   196
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   196
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   195
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   195
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   194
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   194
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   193
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   193
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   192
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   192
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   191
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   191
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   190
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   190
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   189
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   189
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   188
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   188
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   187
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   187
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   186
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   186
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   185
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   185
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   184
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   184
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   183
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   183
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   213
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   213
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AGOSTO"
            Height          =   1905
            Index           =   8
            Left            =   2460
            TabIndex        =   374
            Top             =   1950
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   214
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   214
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   215
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   215
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   216
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   216
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   217
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   217
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   218
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   218
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   219
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   219
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   220
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   220
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   221
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   221
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   222
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   222
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   223
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   223
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   224
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   224
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   225
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   225
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   226
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   226
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   227
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   227
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   228
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   228
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   229
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   229
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   230
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   230
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   231
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   231
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   232
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   232
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   233
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   233
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   234
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   234
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   235
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   235
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   236
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   236
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   237
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   237
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   238
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   238
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   239
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   239
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   240
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   240
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   241
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   241
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   242
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   242
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   243
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   243
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   244
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   244
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SEPTIEMBRE"
            Height          =   1905
            Index           =   9
            Left            =   4860
            TabIndex        =   375
            Top             =   1950
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   274
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   274
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   273
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   273
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   272
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   272
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   271
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   271
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   270
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   270
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   269
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   269
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   268
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   268
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   267
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   267
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   266
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   266
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   265
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   265
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   264
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   264
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   263
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   263
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   262
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   262
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   261
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   261
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   260
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   260
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   259
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   259
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   258
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   258
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   257
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   257
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   256
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   256
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   255
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   255
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   254
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   254
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   253
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   253
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   252
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   252
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   251
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   251
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   250
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   250
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   249
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   249
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   248
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   248
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   247
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   247
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   246
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   246
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   245
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   245
               Top             =   210
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OCTUBRE"
            Height          =   1905
            Index           =   10
            Left            =   7260
            TabIndex        =   376
            Top             =   1950
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   304
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   304
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   303
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   303
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   302
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   302
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   301
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   301
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   300
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   300
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   299
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   299
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   298
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   298
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   297
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   297
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   296
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   296
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   295
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   295
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   294
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   294
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   293
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   293
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   292
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   292
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   291
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   291
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   290
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   290
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   289
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   289
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   288
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   288
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   287
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   287
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   286
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   286
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   285
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   285
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   284
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   284
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   283
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   283
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   282
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   282
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   281
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   281
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   280
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   280
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   279
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   279
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   278
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   278
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   277
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   277
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   276
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   276
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   275
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   275
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   305
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   305
               Top             =   1530
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NOVIEMBRE"
            Height          =   1905
            Index           =   11
            Left            =   9660
            TabIndex        =   377
            Top             =   1950
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   335
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   335
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   334
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   334
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   333
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   333
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   332
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   332
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   331
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   331
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   330
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   330
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   329
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   329
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   328
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   328
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   327
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   327
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   326
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   326
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   325
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   325
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   324
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   324
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   323
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   323
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   322
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   322
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   321
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   321
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   320
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   320
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   319
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   319
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   318
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   318
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   317
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   317
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   316
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   316
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   315
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   315
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   314
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   314
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   313
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   313
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   312
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   312
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   311
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   311
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   310
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   310
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   309
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   309
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   308
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   308
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   307
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   307
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   306
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   306
               Top             =   210
               Width           =   315
            End
         End
         Begin VB.Frame frmMes 
            BackColor       =   &H00C0C0C0&
            Caption         =   "DICIEMBRE"
            Height          =   1905
            Index           =   12
            Left            =   12060
            TabIndex        =   378
            Top             =   1950
            Width           =   2385
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "31"
               Height          =   315
               Index           =   366
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   366
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "30"
               Height          =   315
               Index           =   365
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   365
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "29"
               Height          =   315
               Index           =   364
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   364
               Top             =   1530
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "28"
               Height          =   315
               Index           =   363
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   363
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "27"
               Height          =   315
               Index           =   362
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   362
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "26"
               Height          =   315
               Index           =   361
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   361
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "25"
               Height          =   315
               Index           =   360
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   360
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "24"
               Height          =   315
               Index           =   359
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   359
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "23"
               Height          =   315
               Index           =   358
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   358
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "22"
               Height          =   315
               Index           =   357
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   357
               Top             =   1200
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "21"
               Height          =   315
               Index           =   356
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   356
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "20"
               Height          =   315
               Index           =   355
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   355
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "19"
               Height          =   315
               Index           =   354
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   354
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "18"
               Height          =   315
               Index           =   353
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   353
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "17"
               Height          =   315
               Index           =   352
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   352
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "16"
               Height          =   315
               Index           =   351
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   351
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "15"
               Height          =   315
               Index           =   350
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   350
               Top             =   870
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "14"
               Height          =   315
               Index           =   349
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   349
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "13"
               Height          =   315
               Index           =   348
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   348
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "12"
               Height          =   315
               Index           =   347
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   347
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "11"
               Height          =   315
               Index           =   346
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   346
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "10"
               Height          =   315
               Index           =   345
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   345
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "9"
               Height          =   315
               Index           =   344
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   344
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "8"
               Height          =   315
               Index           =   343
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   343
               Top             =   540
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "7"
               Height          =   315
               Index           =   342
               Left            =   2010
               Style           =   1  'Graphical
               TabIndex        =   342
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "6"
               Height          =   315
               Index           =   341
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   341
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "5"
               Height          =   315
               Index           =   340
               Left            =   1350
               Style           =   1  'Graphical
               TabIndex        =   340
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "4"
               Height          =   315
               Index           =   339
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   339
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "3"
               Height          =   315
               Index           =   338
               Left            =   690
               Style           =   1  'Graphical
               TabIndex        =   338
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "2"
               Height          =   315
               Index           =   337
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   337
               Top             =   210
               Width           =   315
            End
            Begin VB.CheckBox chkDiaMes 
               Caption         =   "1"
               Height          =   315
               Index           =   336
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   336
               Top             =   210
               Width           =   315
            End
         End
      End
      Begin VB.CommandButton cmdGenerarFechasAno 
         Caption         =   "Generar Fechas"
         Height          =   315
         Left            =   -67950
         TabIndex        =   462
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         Caption         =   "Indique el año para el que quiere comprobar las fechas de este Plan de mantenimiento"
         Height          =   225
         Index           =   3
         Left            =   -74880
         TabIndex        =   460
         Top             =   720
         Width           =   6405
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selección de acciones"
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
      Height          =   1140
      Left            =   30
      TabIndex        =   389
      Top             =   8520
      Width           =   11055
      Begin VB.CommandButton cmdCrearAccion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear acción"
         Height          =   870
         Left            =   7680
         Picture         =   "frmEquipoPlanMtoEdicion.frx":0B21
         Style           =   1  'Graphical
         TabIndex        =   399
         ToolTipText     =   "Crear nueva acción"
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar acción"
         Enabled         =   0   'False
         Height          =   870
         Left            =   9915
         Style           =   1  'Graphical
         TabIndex        =   395
         ToolTipText     =   "Eliminar acción del plan de mantenimiento"
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir acción"
         Enabled         =   0   'False
         Height          =   870
         Left            =   8790
         Style           =   1  'Graphical
         TabIndex        =   394
         ToolTipText     =   "Añadir acción al plan de mantenimiento"
         Top             =   180
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cmbFamiliaAcc 
         Height          =   315
         Left            =   750
         TabIndex        =   396
         Top             =   270
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbAcciones 
         Height          =   315
         Left            =   750
         TabIndex        =   398
         Top             =   630
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   240
         Index           =   4
         Left            =   180
         TabIndex        =   397
         Top             =   315
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Acción"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   392
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Plan de mantenimiento"
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
      Height          =   1380
      Left            =   0
      TabIndex        =   383
      Top             =   630
      Width           =   14505
      Begin VB.TextBox txtTiempo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6195
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   382
         Text            =   "0"
         Top             =   255
         Width           =   600
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   379
         Top             =   270
         Width           =   3930
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   1110
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   381
         Top             =   585
         Width           =   6855
      End
      Begin MSDataListLib.DataCombo cmbFrecuencia 
         Height          =   315
         Left            =   9240
         TabIndex        =   380
         Top             =   585
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbProtocolo 
         Height          =   315
         Left            =   9240
         TabIndex        =   472
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   5
         Left            =   8070
         TabIndex        =   473
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Minutos"
         Height          =   195
         Index           =   0
         Left            =   6885
         TabIndex        =   393
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   42
         Left            =   8070
         TabIndex        =   391
         Top             =   660
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Previsto"
         Height          =   195
         Index           =   4
         Left            =   5265
         TabIndex        =   390
         Top             =   270
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   385
         Top             =   630
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   384
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8775
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2160
      Left            =   30
      TabIndex        =   386
      Top             =   6330
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   3810
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   870
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   388
      Top             =   8775
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   13950
      Picture         =   "frmEquipoPlanMtoEdicion.frx":1192
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modificación de Plan de mantenimiento"
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
      Height          =   585
      Index           =   0
      Left            =   0
      TabIndex        =   387
      Top             =   0
      Width           =   14505
   End
End
Attribute VB_Name = "frmEquipoPlanMtoEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarlngPK As Long
Private iCont As Long

Private mvarlngIdTipoPlan As Long
Private mvarintMeses() As Integer
Private mvarintDiasDeMeses() As Integer
Private mvarintSemanas() As Integer
Private mvarintCalendario() As Integer

Private mvarintExcepciones() As Integer
Private mvarintIndiceExcepciones As Integer

Private mvarobjPlanMto As New clsPlanMantenimiento
Private mvarenuTipoEdicion As enumTipoEdicion
Private mvarblnResultado As Boolean
Private mvarAccionesPlan As New clsGenericCollection
Private mvarlngID_PLAN_MTO As Long

Private Sub cmdAceptar_Click()
   
   If Not ComprobarDatos Then Exit Sub
   
   Call RecogerDatos
   
   If MsgBox("¿Está seguro que desea guardar las modificaciones en este Plan de Mantenimiento de Equipos?", vbInformation + vbYesNo, "Guadar Cambios en Plan de Mantenimiento de Equipos") = vbNo Then
    Exit Sub
   End If
   
   If mvarenuTipoEdicion = Alta Then
    mvarlngID_PLAN_MTO = mvarobjPlanMto.Insertar
   Else
    mvarobjPlanMto.Modificar (mvarlngID_PLAN_MTO)
   End If

    cmdAnadir.Enabled = True
    cmdEliminar.Enabled = True
    
    mvarblnResultado = True
    
    mvarenuTipoEdicion = EDICION
    Form_Load
End Sub


Private Sub cmdAnadirExcepcion_Click()
Dim strCad As String
Dim intCodigo As Integer
Dim iCont As Integer
    If cmbDia.ListIndex >= 0 And cmbMes.ListIndex >= 0 Then
        
        intCodigo = CalcularCodigoFecha(cmbDia.ItemData(cmbDia.ListIndex), cmbMes.ItemData(cmbMes.ListIndex))
        
        ' Busca si existe el código
        For iCont = 0 To mvarintIndiceExcepciones - 1
            If mvarintExcepciones(iCont) = intCodigo Then
                Exit Sub
            End If
        Next iCont
        
        ReDim Preserve mvarintExcepciones(mvarintIndiceExcepciones)
        mvarintExcepciones(mvarintIndiceExcepciones) = intCodigo
        mvarintIndiceExcepciones = mvarintIndiceExcepciones + 1
        
        Call PresentarListaExcepciones
        
    End If
End Sub

Private Sub cmdEliminarExcepcion_Click()

Dim iCont As Integer
Dim intTotalEliminados As Integer

intTotalEliminados = 0

    For iCont = 0 To lstExcepciones.ListCount - 1
        If lstExcepciones.Selected(iCont) Then
            mvarintIndiceExcepciones = mvarintIndiceExcepciones - 1
            Call EliminarElementoArrayInt(mvarintExcepciones(), iCont - intTotalEliminados, mvarintIndiceExcepciones)
            intTotalEliminados = intTotalEliminados + 1
        End If
    Next iCont
    
    Call PresentarListaExcepciones
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjPlanMto = Nothing

End Sub


Private Sub cmbAcciones_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbFamiliaAcc_Change()
    Call llenar_combo_filtrado
End Sub

Private Sub cmdGenerarFechasAno_Click()

Dim objPlan As New clsPlanMantenimiento
Dim objCol As clsGenericCollection
Dim objM As clsEquipoMantenimiento

    objPlan.setID_PLAN_MTO = mvarobjPlanMto.getID_PLAN_MTO
    objPlan.setID_TIPO_PLAN = mvarlngIdTipoPlan
    objPlan.setExcepciones = objPlan.DevolverExcepcionesCodificado(mvarintExcepciones, mvarintIndiceExcepciones)
    objPlan.setEVITAR_EXCEPCIONES = IIf(optEvitarExcep.value, 1, 0)
    
    Select Case mvarlngIdTipoPlan
        Case 1
            Call objPlan.setArrMeses(mvarintMeses)
            Call objPlan.setArrDiasMeses(mvarintDiasDeMeses)
        Case 2
            Call objPlan.setArrSemanas(mvarintSemanas)
        Case 4
            Call objPlan.setArrCalendario(mvarintCalendario)
    End Select
    
    Set objCol = objPlan.generarFechasPlanMto(CInt(txtAno.Text))
    
    
    lstFechas.Clear
    
    For Each objM In objCol.Iterator
        lstFechas.AddItem Format(CDate(objM.getFECHA_ACTUAL), "dd \d\e mmmm/yyyy")
    Next
    
    'strCodigoFechas = mvarobjPlanMto.DevolverCalendarioCodificado(mvarlngIdTipoPlan, mvarintMeses(), mvarintDiasDeMeses(), mvarintSemanas(), mvarintCalendario(), mvarintExcepciones(), optEvitarExcep.value)

End Sub

Private Sub cmdCrearAccion_Click()
    frmEquipos_Planes_Acciones.Show 1
    Call llenar_combo_filtrado ' Acciones
End Sub

Private Sub chkDiaMes_Click(Index As Integer)
If Not optTipoPlan(4).value Then
    If MsgBox("Si establece las fechas del plan a través del calendario, esta sustituirá la configuración Mensual/Semanal/Diaria previa si la hubiera. ¿Desea continuar?", vbYesNo, "Planes de Mantenimiento") = vbNo Then
        chkDiaMes(Index).value = vbUnchecked
    Else
        optTipoPlan(4).value = True
        Call optTipoPlan_Click(4)
    End If
End If

mvarintCalendario(Index) = chkDiaMes(Index).value


End Sub

Private Sub chkDiaPorMes_Click(Index As Integer)
    mvarintDiasDeMeses(Index) = chkDiaPorMes(Index).value
End Sub

Private Sub chkPorMes_Click(Index As Integer)
        mvarintMeses(Index) = chkPorMes(Index).value
End Sub

Private Sub chkPorSemana_Click(Index As Integer)
    mvarintSemanas(Index) = chkPorSemana(Index).value
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cabecera
    cargar_botones Me
    cargar_combos
    
    mvarAccionesPlan.KeyName = "setID_ACCION"
    
    inicializar_arrays_fechas_plan
    
    mvarlngIdTipoPlan = 1 ' Por defecto, Mensual
    
    PresentarListaExcepciones
    
   If mvarenuTipoEdicion = EDICION Then
        ' si es edicion, permite añadir acciones
        cmdAnadir.Enabled = True
        cmdEliminar.Enabled = True
        
        If mvarlngID_PLAN_MTO <> 0 Then
            Call mvarobjPlanMto.Carga(mvarlngID_PLAN_MTO)
        End If
         PresentarDatos
   End If
    
    
End Sub

' Botón que permite añadir acciones al detalle del plan
Private Sub cmdAnadir_Click()
    Dim oAccion As New clsEquipos_planes_Acciones
    
    If Len(Trim(cmbAcciones.BoundText)) = 0 Then
        MsgBox "Debe seleccionar la acción que desea añadir.", vbInformation, App.Title
        Exit Sub
    ElseIf Len(Trim(cmbFamiliaAcc.BoundText)) = 0 Then
        MsgBox "Debe seleccionar la familia de la acción que desea añadir.", vbInformation, App.Title
        Exit Sub
    Else
        
        ' lo anexa al plan
        If oAccion.Insertar_en_plan(CLng(cmbAcciones.BoundText), mvarlngID_PLAN_MTO, lista.ListItems.Count + 1) Then
        
            oAccion.Carga (CLng(cmbAcciones.BoundText)) ' Se carga la acción para mostrar en la lista su detalle
            With lista.ListItems.Add(, , oAccion.getID_ACCION)
                .SubItems(1) = lista.ListItems.Count
                .SubItems(2) = cmbFamiliaAcc ' la descripción de la familia está aquí
                .SubItems(3) = oAccion.getNOMBRE
                .SubItems(4) = oAccion.getT_PREVISTO
                txttiempo.Text = CLng(txttiempo.Text) + CLng(.SubItems(4))
            End With
            lista.ListItems(lista.ListItems.Count).EnsureVisible
            cmbAcciones.BoundText = ""
        End If
    End If
End Sub

' Botón que permite eliminar acciones del detalle
Private Sub cmdEliminar_Click()
    Dim i As Long
    Dim oAccion As New clsEquipos_planes_Acciones
    
    If lista.ListItems.Count > 0 Then
        txttiempo.Text = CLng(txttiempo.Text) - CLng(lista.ListItems(lista.selectedItem.Index).SubItems(4))
        For i = lista.selectedItem.Index + 1 To lista.ListItems.Count ' Cambiar la numeración de los que le siguen (-1)
            lista.ListItems(i).SubItems(1) = lista.ListItems(i).SubItems(1) - 1
        Next i
        oAccion.eliminar_de_plan CStr(lista.ListItems(lista.selectedItem.Index)), mvarlngID_PLAN_MTO, lista.ListItems(lista.selectedItem.Index).SubItems(1)
        lista.ListItems.Remove lista.selectedItem.Index
        cmbAcciones.BoundText = ""
    Else
        MsgBox "Debe seleccionar la acción que desea eliminar.", vbInformation, App.Title
        Exit Sub
    End If
End Sub

Private Sub cmdcancel_Click()
   Me.Hide
End Sub

Private Sub optTipoPlan_Click(Index As Integer)
    
    mvarlngIdTipoPlan = Index
    
    If Index = 1 Then
        fraPorMeses.Visible = True
        fraPorMesDias.Visible = True
        fraPorSemanas.Visible = False
    ElseIf Index = 2 Then
        fraPorMeses.Visible = False
        fraPorMesDias.Visible = False
        fraPorSemanas.Visible = True
    ElseIf Index = 3 Then
        fraPorMeses.Visible = False
        fraPorMesDias.Visible = False
        fraPorSemanas.Visible = False
    ElseIf Index = 4 Then
        fraPorMeses.Visible = False
        fraPorMesDias.Visible = False
        fraPorSemanas.Visible = False
    End If
End Sub


Private Sub cabecera()
    lista.ColumnHeaders.Clear

    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Orden", 650, lvwColumnLeft
        .Add , , "Familia", 3500, lvwColumnLeft
        .Add , , "Acción", 7070, lvwColumnLeft
        .Add , , "Minutos", 1000, lvwColumnRight
    End With
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbFrecuencia, DECODIFICADORA.EQ_periodicidad
    oDeco.cargar_combo cmbFamiliaAcc, DECODIFICADORA.EQ_FAMILIAS_ACCIONES_PLANES_MTO
    llenar_combo cmbProtocolo, New clsCa_documentos, 0, frmCA_Documento, ""
    
End Sub

' ------------------------------------
' si no hay ninguna familia seleccionada el combo de acciones no estará cargado
' si hay alguna familia, se cargarán sólo las acciones de esa familia.
Private Sub llenar_combo_filtrado()
    Dim oAcciones As New clsEquipos_planes_Acciones
    
    cmbAcciones.BoundText = ""
    
    If cmbFamiliaAcc = "" Then
        Set cmbAcciones.RowSource = Nothing
    Else
        Set cmbAcciones.RowSource = oAcciones.Listado_por_Familia(CLng(cmbFamiliaAcc.BoundText))
    End If
    cmbAcciones.ListField = "NOMBRE"
    cmbAcciones.BoundColumn = "ID_ACCION"

    Set oAcciones = Nothing
End Sub

Private Sub inicializar_arrays_fechas_plan()
Dim cont As Integer

    For cont = 1 To 12
        ReDim Preserve mvarintMeses(cont)
        mvarintMeses(cont) = 0
    Next cont
    
    For cont = 1 To 7
        ReDim Preserve mvarintSemanas(cont)
        mvarintSemanas(cont) = 0
    Next cont
    
    For cont = 1 To 31
        ReDim Preserve mvarintDiasDeMeses(cont)
        mvarintDiasDeMeses(cont) = 0
    Next cont
    
    For cont = 1 To 366
        ReDim Preserve mvarintCalendario(cont)
        mvarintCalendario(cont) = 0
    Next cont
    
    ' Las Por Defecto que si
    mvarintIndiceExcepciones = 6
    ReDim Preserve mvarintExcepciones(mvarintIndiceExcepciones - 1)
    
    mvarintExcepciones(0) = 1 ' 1 Enero
    mvarintExcepciones(1) = 6 ' 6 Enero
    mvarintExcepciones(2) = 59 ' 28 Febrero
    mvarintExcepciones(3) = 122 ' 1 Mayo
    mvarintExcepciones(4) = 286 ' 12 Octubre
    mvarintExcepciones(5) = 360 ' 25 Diciembre
    
End Sub

Public Property Get PlanMto() As clsPlanMantenimiento

    Set PlanMto = mvarobjPlanMto

End Property

Public Property Set PlanMto(objPlanMto As clsPlanMantenimiento)

    Set mvarobjPlanMto = objPlanMto

End Property

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub PresentarListaExcepciones()
Dim iCont As Integer, dtmFecha As Date


    mvarintExcepciones = OrdernarValoresArrayInt(mvarintExcepciones, 0, mvarintIndiceExcepciones - 1)

    With lstExcepciones
        
        .Clear
        
        For iCont = 0 To mvarintIndiceExcepciones - 1
            dtmFecha = CalcularFechaPorCodigo(mvarintExcepciones(iCont))
            .AddItem Format(dtmFecha, "dd/mmmm")
            .ItemData(.ListCount - 1) = mvarintExcepciones(iCont)
        Next iCont
    End With
    
End Sub

Private Function ComprobarDatos() As Boolean

   Dim blnRes As Boolean
   Dim strCad As String
   Dim intres As Integer
On Error GoTo ComprobarDatos_Error

   ComprobarDatos = False
   strCad = ""
      
    If Trim(txtNombre.Text) = "" Then
        strCad = strCad & vbCrLf & " - Debe indicar un nómbre válido para el Plan"
    End If
      
    If getDataComboSel(cmbFrecuencia) < 0 Then
        strCad = strCad & vbCrLf & " - Debe indicar la Periodicidad del Plan"
    End If
    
    If cmbProtocolo.getPK_SALIDA <= 0 Then
        strCad = strCad & vbCrLf & " - Debe indicar el Protocolo del Plan de Mantenimiento"
    End If
    
    
    If optTipoPlan(1).value Then ' Mensual
        'Meses
        intres = 0
        For iCont = 1 To 12
            intres = intres + chkPorMes(iCont).value
        Next iCont
        If intres = 0 Then
            strCad = strCad & vbCrLf & " - Cuando el Plan es Mensual, debe indicar algún mes"
        End If
        
        ' Dias
        intres = 0
        For iCont = 1 To 31
            intres = intres + chkDiaPorMes(iCont).value
        Next iCont
        If intres = 0 Then
            strCad = strCad & vbCrLf & " - Cuando el Plan es Mensual, debe indicar algún día del mes"
        End If
    
    
    ElseIf optTipoPlan(2).value Then ' Semanal
        intres = 0
        For iCont = 1 To 7
            intres = intres + chkPorSemana(iCont).value
        Next iCont
        If intres = 0 Then
            strCad = strCad & vbCrLf & " - Cuando el Plan es Semanal, debe indicar algún día de la semana"
        End If

    ElseIf optTipoPlan(4).value Then ' Calendario
        intres = 0
        For iCont = 1 To 366
            intres = intres + chkDiaMes(iCont).value
        Next iCont
        If intres = 0 Then
            strCad = strCad & vbCrLf & " - Cuando el Plan es Anual, debe indicar algún día en el calendario"
        End If

    End If
    
    If strCad <> "" Then
        MsgBox "Se han encontrado los siguientes Errores: " & strCad, vbInformation, "Plan de Mantenimiento de Equipos"
        Exit Function
    End If
   
   
   ComprobarDatos = True





On Error GoTo 0
Exit Function
ComprobarDatos_Error:
    
MsgBox Err.Description


End Function

Private Sub PresentarDatos()
Dim x() As Integer

On Error GoTo PresentarDatos_Error

   With mvarobjPlanMto
      ' Preliminares
   
      mvarlngIdTipoPlan = .getID_TIPO_PLAN
      mvarintMeses = SplitString2IntArr(.getPLANIF_MESES)
      mvarintDiasDeMeses = SplitString2IntArr(.getPLANIF_DIAS_MES)
      mvarintSemanas = SplitString2IntArr(.getPLANIF_DIAS_SEMANA)
      mvarintCalendario = SplitString2IntArr(.getCALENDARIO)
      mvarintExcepciones = .getArrayExcepciones()
      
      
      ' Pestaña 1
      optEvitarExcep.value = .getEVITAR_EXCEPCIONES
      optTrasladarExcep.value = Not .getEVITAR_EXCEPCIONES
      optTipoPlan(.getID_TIPO_PLAN).value = True
      txtNombre.Text = .getNOMBRE
      cmbFrecuencia.BoundText = .getFRECUENCIA_ID
      txtDescripcion.Text = .getDESCRIPCION
      cmbProtocolo.MostrarElemento .getPROTOCOLO_ID
      
      Call PresentarDatos_Acciones
            
      Call optTipoPlan_Click(.getID_TIPO_PLAN)
      
      Select Case .getID_TIPO_PLAN
         Case 1 ' mensual
            Call PresentarDatos_Mensual
         Case 2 ' Semanal
            Call PresentarDatos_Semanal
         Case 4 ' Calendario
            Call PresentarDatos_Calendario
      End Select
      
   End With
   
   
   'Call PresentarDatos_Acciones
    
   

On Error GoTo 0
Exit Sub
PresentarDatos_Error:
    
    

End Sub

Private Sub PresentarDatos_Mensual()
On Error GoTo PresentarDatos_Mensual_Error


    For iCont = 1 To 12
        chkPorMes(iCont).value = mvarintMeses(iCont)
    Next iCont
   
    For iCont = 1 To 31
        chkDiaPorMes(iCont).value = mvarintDiasDeMeses(iCont)
    Next iCont
   

On Error GoTo 0
Exit Sub
PresentarDatos_Mensual_Error:
    

End Sub

Private Sub PresentarDatos_Semanal()

On Error GoTo PresentarDatos_Semanal_Error

    For iCont = 1 To 7
        chkPorSemana(iCont).value = mvarintSemanas(iCont)
    Next iCont

On Error GoTo 0
Exit Sub
PresentarDatos_Semanal_Error:
    

End Sub

Private Sub PresentarDatos_Calendario()

On Error GoTo PresentarDatos_Calendario_Error
    
    For iCont = 1 To 366
        chkDiaMes(iCont).value = mvarintCalendario(iCont)
    Next iCont
    

On Error GoTo 0
Exit Sub
PresentarDatos_Calendario_Error:
    

End Sub

Private Sub PresentarDatos_Acciones()
On Error GoTo PresentarDatos_Acciones_Error

    Dim oEQ_plan_detalle As New clsEquipos_PM_Detalle
    Dim rs As ADODB.Recordset
    Dim lngTotal_tiempo As Long
    lngTotal_tiempo = 0
    lista.ListItems.Clear
    Set rs = oEQ_plan_detalle.Listado(mvarlngID_PLAN_MTO)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(3) ' orden
                .SubItems(2) = rs(4) ' familia
                .SubItems(3) = rs(1) ' accion
                .SubItems(4) = rs(2) ' tiempo
                lngTotal_tiempo = lngTotal_tiempo + CLng(rs(2))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    txttiempo.Text = lngTotal_tiempo


On Error GoTo 0
Exit Sub
PresentarDatos_Acciones_Error:
    

End Sub

Private Sub RecogerDatos()
On Error GoTo RecogerDatos_Error

    
    With mvarobjPlanMto
        .setID_TIPO_PLAN = mvarlngIdTipoPlan
        .setNOMBRE = txtNombre.Text
        .setDESCRIPCION = txtDescripcion.Text
        .setFRECUENCIA_ID = getDataComboSel(cmbFrecuencia)
        .setExcepciones = .DevolverExcepcionesCodificado(mvarintExcepciones, mvarintIndiceExcepciones)
        .setEVITAR_EXCEPCIONES = IIf(optEvitarExcep.value, 1, 0)
        .setPROTOCOLO_ID = cmbProtocolo.getPK_SALIDA
        .setArrCalendario mvarintCalendario
        .setArrDiasMeses mvarintDiasDeMeses
        .setArrMeses mvarintMeses
        .setArrSemanas mvarintSemanas
                
    End With
    

On Error GoTo 0
Exit Sub
RecogerDatos_Error:
    

End Sub

Public Property Get ID_PLAN_MTO() As Long

On Error GoTo ID_PLAN_MTO_Error

    ID_PLAN_MTO = mvarlngID_PLAN_MTO

On Error GoTo 0
Exit Property
ID_PLAN_MTO_Error:
    


End Property

Public Property Let ID_PLAN_MTO(ByVal lngID_PLAN_MTO As Long)

On Error GoTo ID_PLAN_MTO_Error

    mvarlngID_PLAN_MTO = lngID_PLAN_MTO

On Error GoTo 0
Exit Property
ID_PLAN_MTO_Error:
    


End Property

Public Property Get PK() As Long

    PK = mvarlngPK

End Property

Public Property Let PK(ByVal lngPK As Long)

    mvarlngPK = lngPK
    
    mvarlngID_PLAN_MTO = mvarlngPK
    
    mvarenuTipoEdicion = EDICION

End Property


