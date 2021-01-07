VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmAlodine_Alodine 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lotes de Alodine"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16080
   Icon            =   "frmAlodine_Alodine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
   Begin Geslab.ControlPanelXP ControlPanelXP3 
      Height          =   6270
      Left            =   9405
      TabIndex        =   47
      Top             =   2475
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   11060
      Caption         =   "Muestras realizadas"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   6270
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   45
         TabIndex        =   51
         Top             =   5085
         Width           =   8715
         Begin VB.TextBox txtanno 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   5445
            TabIndex        =   53
            Top             =   150
            Width           =   465
         End
         Begin VB.TextBox txtmuestra 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3240
            TabIndex        =   52
            Top             =   150
            Width           =   1515
         End
         Begin MSComCtl2.UpDown cambiar 
            Height          =   330
            Left            =   5911
            TabIndex        =   54
            Top             =   150
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   582
            _Version        =   393216
            Value           =   2004
            OrigLeft        =   1590
            OrigTop         =   6570
            OrigRight       =   1830
            OrigBottom      =   6975
            Max             =   2015
            Min             =   2004
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Año"
            Height          =   225
            Index           =   0
            Left            =   4980
            TabIndex        =   56
            Top             =   210
            Width           =   435
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Número general de la muestra a añadir "
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   55
            Top             =   195
            Width           =   2775
         End
      End
      Begin MSComctlLib.ListView listaMuestras 
         Height          =   4590
         Left            =   45
         TabIndex        =   50
         Top             =   450
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   8096
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
      Begin XtremeSuiteControls.PushButton cmdInsertaMuestra 
         Height          =   435
         Left            =   6525
         TabIndex        =   49
         Top             =   5670
         Width           =   2145
         _Version        =   851970
         _ExtentX        =   3784
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Insertar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminaMuestra 
         Height          =   435
         Left            =   90
         TabIndex        =   48
         Top             =   5670
         Width           =   2115
         _Version        =   851970
         _ExtentX        =   3731
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":712C
      End
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Enabled         =   0   'False
      Height          =   870
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   8820
      Width           =   1410
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   9450
      TabIndex        =   57
      Top             =   315
      Width           =   8700
      Begin VB.CommandButton cmdAnadirCaducidad 
         Caption         =   "+"
         Height          =   345
         Left            =   8145
         TabIndex        =   12
         Top             =   1485
         Width           =   315
      End
      Begin VB.CheckBox chkFinalizado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Finalizado"
         Height          =   285
         Left            =   6930
         TabIndex        =   8
         Top             =   270
         Width           =   1155
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   6
         Left            =   1350
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker fecha_creacion 
         Height          =   330
         Left            =   1350
         TabIndex        =   6
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   60817409
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_terminacion 
         Height          =   330
         Left            =   4680
         TabIndex        =   7
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   60817409
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   1350
         TabIndex        =   9
         Top             =   675
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cmbCaducidad 
         Height          =   315
         Left            =   1350
         TabIndex        =   11
         Top             =   1500
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   62
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Creación"
         Height          =   240
         Index           =   9
         Left            =   90
         TabIndex        =   61
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Terminación"
         Height          =   240
         Index           =   10
         Left            =   3510
         TabIndex        =   60
         Top             =   315
         Width           =   1110
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preparado Por"
         Height          =   240
         Index           =   11
         Left            =   90
         TabIndex        =   59
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volumen"
         Height          =   240
         Index           =   12
         Left            =   90
         TabIndex        =   58
         Top             =   1125
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   30
      TabIndex        =   15
      Top             =   315
      Width           =   9375
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   7
         Left            =   1275
         TabIndex        =   0
         Top             =   225
         Width           =   1560
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   1275
         TabIndex        =   3
         Top             =   945
         Width           =   8010
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1275
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1665
         Width           =   8010
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1275
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1305
         Width           =   8010
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1275
         TabIndex        =   2
         Top             =   585
         Width           =   8010
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   5985
         TabIndex        =   1
         Top             =   225
         Width           =   3300
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NºLOTE"
         Height          =   240
         Index           =   13
         Left            =   90
         TabIndex        =   23
         Top             =   285
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote Componente"
         Height          =   465
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   240
         Index           =   6
         Left            =   105
         TabIndex        =   20
         Top             =   1710
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   4
         Left            =   105
         TabIndex        =   19
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Index           =   3
         Left            =   5265
         TabIndex        =   16
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clientes"
      Enabled         =   0   'False
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8820
      Width           =   1410
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8790
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   16050
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8790
      Width           =   1050
   End
   Begin Geslab.ControlPanelXP cpNormas 
      Height          =   3120
      Left            =   45
      TabIndex        =   24
      Top             =   2475
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   5503
      Caption         =   "Normas de Referencia"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3120
      Begin MSComctlLib.ListView listaNormas 
         Height          =   1545
         Left            =   45
         TabIndex        =   28
         Top             =   450
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
      Begin pryCombo.miCombo cmbNormas 
         Height          =   330
         Left            =   45
         TabIndex        =   27
         Top             =   2160
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirNorma 
         Height          =   435
         Left            =   7110
         TabIndex        =   26
         Top             =   2610
         Width           =   2145
         _Version        =   851970
         _ExtentX        =   3784
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Insertar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":D98E
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarNorma 
         Height          =   435
         Left            =   45
         TabIndex        =   25
         Top             =   2610
         Width           =   2115
         _Version        =   851970
         _ExtentX        =   3731
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":141F0
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   3165
      Left            =   1350
      TabIndex        =   29
      Top             =   6345
      Visible         =   0   'False
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   5583
      Caption         =   "Parámetros"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3165
      Begin MSComctlLib.ListView listaParametros 
         Height          =   1560
         Left            =   45
         TabIndex        =   37
         Top             =   450
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   2752
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
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   45
         TabIndex        =   32
         Top             =   2010
         Width           =   9210
         Begin VB.TextBox txtParametros 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   30
            TabIndex        =   35
            Top             =   180
            Width           =   3705
         End
         Begin VB.TextBox txtParametros 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   3750
            TabIndex        =   34
            Top             =   180
            Width           =   1545
         End
         Begin VB.TextBox txtParametros 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   5310
            TabIndex        =   33
            Top             =   180
            Width           =   1425
         End
         Begin MSDataListLib.DataCombo cmbUnidades 
            Height          =   315
            Left            =   6750
            TabIndex        =   36
            Top             =   180
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
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
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   435
         Left            =   45
         TabIndex        =   31
         Top             =   2655
         Width           =   2115
         _Version        =   851970
         _ExtentX        =   3731
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":1AA52
      End
      Begin XtremeSuiteControls.PushButton cmdInsertaParametro 
         Height          =   435
         Left            =   7110
         TabIndex        =   30
         Top             =   2655
         Width           =   2145
         _Version        =   851970
         _ExtentX        =   3784
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Insertar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":212B4
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   3120
      Left            =   45
      TabIndex        =   38
      Top             =   5625
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   5503
      Caption         =   "Listado de Reactivos"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3120
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   45
         TabIndex        =   42
         Top             =   1980
         Width           =   8985
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   7230
            TabIndex        =   43
            Top             =   180
            Width           =   1365
         End
         Begin pryCombo.miCombo cmbReactivos 
            Height          =   330
            Left            =   780
            TabIndex        =   44
            Top             =   180
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reactivo"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   46
            Top             =   210
            Width           =   645
         End
         Begin VB.Label lblCampos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Volumen/Peso"
            Height          =   240
            Index           =   7
            Left            =   6120
            TabIndex        =   45
            Top             =   240
            Width           =   1065
         End
      End
      Begin MSComctlLib.ListView listaReactivos 
         Height          =   1530
         Left            =   45
         TabIndex        =   41
         Top             =   450
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   2699
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
      Begin XtremeSuiteControls.PushButton cmdEliminaReactivo 
         Height          =   435
         Left            =   45
         TabIndex        =   40
         Top             =   2610
         Width           =   2115
         _Version        =   851970
         _ExtentX        =   3731
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":27B16
      End
      Begin XtremeSuiteControls.PushButton cmdInsertaReactivo 
         Height          =   435
         Left            =   6885
         TabIndex        =   39
         Top             =   2610
         Width           =   2145
         _Version        =   851970
         _ExtentX        =   3784
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Insertar"
         Appearance      =   5
         Picture         =   "frmAlodine_Alodine.frx":2E378
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Nuevo Tipo de Alodine"
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
      Height          =   330
      Index           =   2
      Left            =   30
      TabIndex        =   18
      Top             =   15
      Width           =   18930
   End
End
Attribute VB_Name = "frmAlodine_Alodine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
End Sub

Private Sub chkFinalizado_Click()
    If chkFinalizado.Value = Checked Then
        fecha_terminacion.Enabled = True
    Else
        fecha_terminacion.Enabled = False
    End If
End Sub

Private Sub cmdAnadirCaducidad_Click()
    frmTipos_caducidad.Show 1
    cargar_combo cmbCaducidad, New clsTipos_caducidad
End Sub

Private Sub cmdAnadirNorma_Click()
    If cmbNormas.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar una de entre las existentes", vbOK + vbExclamation, "Añadir Norma"
        Exit Sub
    End If
    Dim i As Integer
    For i = 1 To listaNormas.ListItems.Count
        If CLng(listaNormas.ListItems(i).Text) = CLng(cmbNormas.getPK_SALIDA) Then
            MsgBox "La norma ya se encuentra en la lista.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    
    With listaNormas.ListItems.Add(, , cmbNormas.getPK_SALIDA)
        .SubItems(1) = cmbNormas.getTEXTO
    End With
    cmbNormas.limpiar
End Sub

Private Sub cmdcancel_Click()
    gAlodine = 0
    Unload Me
End Sub

Private Sub cmdClientes_Click()
    frmAlodine_Clientes.Show 1
End Sub

Private Sub cmdEliminaMuestra_Click()
    If listaMuestras.ListItems.Count > 0 Then
        listaMuestras.ListItems.Remove listaMuestras.selectedItem.Index
        txtMuestra = ""
    End If
End Sub

Private Sub cmdEliminaParametro_Click()

End Sub
Private Sub cmdEliminaReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        txtDatos(5) = ""
    End If
End Sub

Private Sub cmdEliminarNorma_Click()
    If listaNormas.ListItems.Count > 0 Then
        listaNormas.ListItems.Remove listaNormas.selectedItem.Index
    End If
End Sub

Private Sub cmdetiqueta_Click()
    frmAlodine_Etiquetas.PK = gAlodine
    frmAlodine_Etiquetas.Show 1
End Sub

Private Sub cmdInsertaMuestra_Click()
    If txtMuestra <> "" Then
        If IsNumeric(txtMuestra) Then
            cargar_muestra_por_numero txtMuestra, txtanno
        End If
    End If
End Sub

Private Sub cmdInsertaParametro_Click()
    If txtParametros(0).Text = "" Then
        MsgBox "Introduzca una descripción para el parámetro.", vbCritical, App.Title
        txtParametros(0).SetFocus
        Exit Sub
    End If
    If txtParametros(1).Text = "" Then
        MsgBox "Introduzca un rango para el parámetro.", vbCritical, App.Title
        txtParametros(1).SetFocus
        Exit Sub
    End If
    With listaParametros.ListItems.Add(, , txtParametros(0))
        .SubItems(1) = txtParametros(1)
        .SubItems(2) = txtParametros(2)
        If cmbUnidades.BoundText = "" Then
            .SubItems(4) = 0
        Else
            .SubItems(3) = cmbUnidades.Text
            .SubItems(4) = cmbUnidades.BoundText
        End If
        .SubItems(5) = 0
    End With
    borrar_campos

End Sub

Private Sub cmdInsertaReactivo_Click()
    If cmbReactivos.getTEXTO = "" Then
        MsgBox "Seleccione el reactivo.", vbExclamation, App.Title
        cmbReactivos.SetFocus
        Exit Sub
    End If
    If txtDatos(5) = "" Then
        MsgBox "Introduzca el Volumen/Peso.", vbExclamation, App.Title
        txtDatos(5).SetFocus
        Exit Sub
    End If
        
    Dim oBote As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    oBote.CARGAR cmbReactivos.getPK_SALIDA
    oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
    oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
    With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
        .SubItems(1) = oTR.getNOMBRE
        .SubItems(2) = txtDatos(5)
        .SubItems(3) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
    End With
    listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    ' Limpiar Combos
    cmbReactivos.limpiar
    txtDatos(5) = ""
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      ' Alodine
      Dim alodine As Long
      alodine = gAlodine
      Dim oalodine As New clsAlodine
      Dim oAlodine_Parametros As New clsAlodine_parametros
      With oalodine
           .setPRODUCTO = txtDatos(0)
           .setCODIGO = txtDatos(1)
           .setLOTE = txtDatos(4)
           .setDESCRIPCION = txtDatos(2)
           .setPROCEDIMIENTO = txtDatos(3)
           .setTIPO_CADUCIDAD_ID = cmbCaducidad.BoundText
           .setFECHA_CREACION = Format(fecha_creacion, "yyyy-mm-dd")
           .setVOLUMEN = txtDatos(6)
           If chkFinalizado.Value = Checked Then
            .setFECHA_TERMINACION = Format(fecha_terminacion, "yyyy-mm-dd")
            .setTERMINADO = 1
           Else
            .setFECHA_TERMINACION = "0000-00-00"
            .setTERMINADO = 0
           End If
           .setUSUARIO_ID = cmbUsuario.BoundText
            ' Muestras
           Dim muestras As String
           For i = 1 To listaMuestras.ListItems.Count
               muestras = muestras & listaMuestras.ListItems(i).SubItems(6) & ","
           Next
           If muestras <> "" Then
             muestras = Left(muestras, Len(muestras) - 1)
           End If
            .setMUESTRAS = muestras
            .setMADRE = txtDatos(7)
      End With
      If gAlodine = 0 Then
        If MsgBox("Va a introducir un nuevo alodine. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            gAlodine = oalodine.Insertar
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el alodine. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If oalodine.Modificar(gAlodine) = False Then
                Exit Sub
            End If
            ' Eliminar parametros
            oAlodine_Parametros.Eliminar (gAlodine)
        Else
            Exit Sub
        End If
      End If
      ' Insertar parámetros
      For i = 1 To listaParametros.ListItems.Count
        With oAlodine_Parametros
            .setID_PARAMETRO = listaParametros.ListItems(i).SubItems(5)
            .setALODINE_ID = gAlodine
            .setPARAMETRO = listaParametros.ListItems(i).Text
            .setRANGO = listaParametros.ListItems(i).SubItems(1)
            .setPROCEDIMIENTO = listaParametros.ListItems(i).SubItems(2)
            .setUNIDAD_ID = listaParametros.ListItems(i).SubItems(4)
            If .Insertar = 0 Then
                Exit Sub
            End If
        End With
      Next
      ' Normas
      Dim oAN As New clsAlodine_normas
      oAN.Eliminar gAlodine
      For i = 1 To listaNormas.ListItems.Count
        With oAN
            .setALODINE_ID = gAlodine
            .setNORMA_ID = listaNormas.ListItems(i).Text
            .setORDEN = i
            .Insertar
        End With
      Next
      Set oAN = Nothing
      ' Reactivos
      Dim oAR As New clsAlodine_reactivos
      oAR.Eliminar gAlodine
      For i = 1 To listaReactivos.ListItems.Count
        With oAR
            .setALODINE_ID = gAlodine
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setVOLUMEN = listaReactivos.ListItems(i).SubItems(2)
            .setORDEN = i
            .Insertar
        End With
      Next
      Set oAR = Nothing
      If alodine = 0 Then
          MsgBox "El alodine se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          cmdClientes.Enabled = True
          cmdEtiqueta.Enabled = True
      Else
          MsgBox "El alodine se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    
    Exit Sub
fallo:
    error_grave ("Error al insertar el alodine : " & Err.Description)
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    If gAlodine <> 0 Then
        Label1(2) = "Modificación de Tipo de Alodine"
        cmdClientes.Enabled = True
        cmdEtiqueta.Enabled = True
        CARGAR
'        txtDatos(7).Enabled = True
    Else
        cmbUsuario.BoundText = USUARIO.getID_EMPLEADO
'        txtDatos(7).Enabled = False
        txtDatos(7) = "1"
        txtDatos(7).BackColor = &HE0E0E0
    End If
End Sub

Private Sub cabecera()
    ' Normas
    With listaNormas.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Nombre", listaNormas.Width - 200, lvwColumnLeft
    End With
    ' Parametros
    With listaParametros.ColumnHeaders
        .Add , , "Parámetro", 3705, lvwColumnLeft
        .Add , , "Rango", 1645, lvwColumnCenter
        .Add , , "Procedimiento", 1425, lvwColumnCenter
        .Add , , "Unidad", 1415, lvwColumnCenter
        .Add , , "ID_UNIDAD", 1, lvwColumnCenter
        .Add , , "ID_PARAMETRO", 1, lvwColumnCenter
    End With
    ' Reactivos
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 4800, lvwColumnLeft
        .Add , , "Volumen/Peso", 1500, lvwColumnCenter
        .Add , , "Caducidad", 1200, lvwColumnCenter
    End With
    ' Muestras
    With listaMuestras.ColumnHeaders
        .Add , , "Código", 900, lvwColumnLeft
        .Add , , "Cliente", 1800, lvwColumnLeft
        .Add , , "Tipo de Analisis/Solución", 2000, lvwColumnLeft
        .Add , , "Ref.Cliente", 2000, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "General", 800, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
    End With
    
End Sub

Private Sub listaMuestras_DblClick()
    If listaMuestras.ListItems.Count > 0 Then
        gmuestra = listaMuestras.ListItems(listaMuestras.selectedItem.Index).SubItems(6)
        frmVerMuestra.Show 1
    End If
End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub PushButton2_Click()
    If listaParametros.ListItems.Count > 0 Then
        listaParametros.ListItems.Remove listaParametros.selectedItem.Index
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub CARGAR()
    Dim oalodine As New clsAlodine
    With oalodine
    If .Carga(gAlodine) = True Then
        txtDatos(0) = .getPRODUCTO
        txtDatos(1) = .getCODIGO
        txtDatos(2) = .getDESCRIPCION
        txtDatos(3) = .getPROCEDIMIENTO
        txtDatos(4) = .getLOTE
        txtDatos(6) = .getVOLUMEN
        txtDatos(7) = .getMADRE
        If .getFECHA_CREACION <> "0000-00-00" Then
            fecha_creacion = .getFECHA_CREACION
        End If
        If .getFECHA_TERMINACION <> "0000-00-00" Then
            fecha_terminacion = .getFECHA_TERMINACION
        End If
        chkFinalizado.Value = .getTERMINADO
        cmbUsuario.BoundText = .getUSUARIO_ID
        ' Caducidad
        Dim oCaducidad As New clsTipos_caducidad
        oCaducidad.CARGAR (.getTIPO_CADUCIDAD_ID)
        cmbCaducidad = oCaducidad.getCADUCIDAD
        ' Parámetros
        Dim oAlodine_Parametros As New clsAlodine_parametros
        Dim rs As ADODB.Recordset
        Set rs = oAlodine_Parametros.Listado_Parametros(gAlodine)
        If rs.RecordCount <> 0 Then
            Do
                With listaParametros.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                    .SubItems(3) = rs(3)
                    .SubItems(4) = rs(4)
                    .SubItems(5) = rs(5)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        ' Reactivos
        Dim oAR As New clsAlodine_reactivos
        Set rs = oAR.Listado(gAlodine)
        If rs.RecordCount <> 0 Then
            Do
                With listaReactivos.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                    .SubItems(3) = Format(rs(3), "DD-MM-YYYY")
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        cargar_normas
        ' Muestras
        If .getMUESTRAS <> "" Then
            cargar_muestras .getMUESTRAS
        End If
    End If
    End With
    Set oalodine = Nothing
End Sub
Private Sub cargar_normas()
    Dim rs As ADODB.Recordset
    Dim oAN As New clsAlodine_normas
    Set rs = oAN.Listado(gAlodine)

    listaNormas.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaNormas.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub

Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe introducir una descripción en el producto.", vbInformation, App.Title
        validar = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If cmbCaducidad.BoundText = "" Then
        MsgBox "Seleccione un tipo de caducidad.", vbInformation, App.Title
        validar = False
        cmbCaducidad.SetFocus
        Exit Function
    End If
    If cmbUsuario.BoundText = "" Then
        MsgBox "Debe indicar la persona que prepara el Alodine.", vbInformation, App.Title
        cmbUsuario.SetFocus
        validar = False
        Exit Function
    End If
End Function

Private Sub cargar_combos()
    cargar_combo cmbCaducidad, New clsTipos_caducidad
    cargar_combo cmbUnidades, New clsUnidades
    cargar_combo cmbUsuario, New clsUsuarios
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
    
    Dim consulta As String
    Dim oParam As New clsParametros
    oParam.Carga parametros.PARAM_TIPOS_BOTE_ALODINE, ""
    
    consulta = "SELECT A.ID_BOTE_EX, CONCAT(C.NOMBRE, '  (', CONCAT('Num.', CAST(A.ID_BOTE_EX AS CHAR)), CONCAT('  Lote:',A.LOTE),')') " & _
               "  FROM BOTES_EX A,TIPOS_BOTE_EX B, TIPOS_REACTIVO_EX C " & _
               " Where A.TIPO_BOTE_EX_ID = B.ID_TIPO_BOTE_EX " & _
               "   AND B.TIPO_REACTIVO_EX_ID = C.ID_TIPO_REACTIVO_EX " & _
               "   AND A.TIPO_BOTE_EX_ID IN (" & oParam.getVALOR & ")" & _
               "   AND A.FINALIZADO = 0"
    
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbReactivos
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "BOTES_EX"
            .setDESCRIPCION = "Reactivos"
            .setPK = "ID_EQUIPO_EX"
            .setCAMPO = "CONCAT(C.NOMBRE, '  (',CONCAT('Num.',CAST(A.ID_BOTE_EX AS CHAR)), CONCAT('  Lote:',A.LOTE),')')"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = False
            Set .FORMULARIO = Nothing
        End With
    End If
    Set conn = Nothing
End Sub
Private Sub borrar_campos()
    txtParametros(0) = ""
    txtParametros(1) = ""
    txtParametros(2) = ""
    cmbUnidades.Text = ""
    txtParametros(0).SetFocus
End Sub
Private Sub cargar_muestras(muestras As String)
    Dim consulta As String
   On Error GoTo cargar_muestras_Error

    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                      " mu.id_muestra in (" & muestras & ")" & _
                      " order by mu.id_general desc"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Do
            With listaMuestras.ListItems.Add(, , rs(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
                If Not IsNull(rs.Fields(9)) Then
                   .SubItems(5) = Format(rs.Fields(9), "00000")
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = rs.Fields(6)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_muestras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestras of Formulario frmAlodine_Alodine"

End Sub
Private Sub cargar_muestra_por_numero(muestra As Long, ANNO As Long)
    Dim consulta As String
   On Error GoTo cargar_muestra_por_numero_Error

    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      " AND mu.id_general =" & muestra & _
                      " AND mu.anno = " & ANNO & _
                      " order by mu.id_general desc"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Do
            With listaMuestras.ListItems.Add(, , rs(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
                If Not IsNull(rs.Fields(9)) Then
                   .SubItems(5) = Format(rs.Fields(9), "00000")
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = rs.Fields(6)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_muestra_por_numero_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra_por_numero of Formulario frmAlodine_Alodine"

End Sub


