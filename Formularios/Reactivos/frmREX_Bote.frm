VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmREX_Bote 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Botes de Reactivos externos"
   ClientHeight    =   11835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13680
   Icon            =   "frmREX_Bote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11835
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmExistencias 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Existencias"
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
      Height          =   2985
      Left            =   45
      TabIndex        =   80
      Top             =   4185
      Width           =   13515
      Begin MSComctlLib.ListView listaExistencias 
         Height          =   2670
         Left            =   90
         TabIndex        =   81
         Top             =   225
         Width           =   13365
         _ExtentX        =   23574
         _ExtentY        =   4710
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
   End
   Begin VB.CommandButton cmdParametros 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisitos"
      Height          =   870
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   10860
      Width           =   1050
   End
   Begin VB.Frame frameMR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "M.R. / M.R.C."
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
      Height          =   3615
      Left            =   45
      TabIndex        =   47
      Top             =   7200
      Width           =   13515
      Begin Geslab.ControlPanelXP ControlPanelXP2 
         Height          =   2625
         Left            =   8955
         TabIndex        =   74
         Top             =   315
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   4630
         Caption         =   "Equipos Asociados al Reactivo"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   2625
         Begin VB.Frame frmEquipos 
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
            Height          =   2085
            Left            =   45
            TabIndex        =   75
            Top             =   450
            Width           =   4425
            Begin MSComctlLib.ListView listaEquipos 
               Height          =   1140
               Left            =   45
               TabIndex        =   76
               Top             =   0
               Width           =   4305
               _ExtentX        =   7594
               _ExtentY        =   2011
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
            Begin pryCombo.miCombo cmbEquipos 
               Height          =   330
               Left            =   90
               TabIndex        =   77
               Top             =   1215
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   582
            End
            Begin XtremeSuiteControls.PushButton cmdEliminarEquipo 
               Height          =   435
               Left            =   2295
               TabIndex        =   78
               Top             =   1575
               Width           =   1845
               _Version        =   851970
               _ExtentX        =   3254
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Eliminar Equipo"
               Appearance      =   5
               Picture         =   "frmREX_Bote.frx":08CA
            End
            Begin XtremeSuiteControls.PushButton cmdAnadirEquipo 
               Height          =   435
               Left            =   315
               TabIndex        =   79
               Top             =   1575
               Width           =   1875
               _Version        =   851970
               _ExtentX        =   3307
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Añadir Equipo"
               Appearance      =   5
               Picture         =   "frmREX_Bote.frx":712C
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Height          =   600
         Left            =   4815
         TabIndex        =   71
         Top             =   2295
         Width           =   3975
         Begin VB.CheckBox chkENAC 
            Caption         =   "Check1"
            Height          =   195
            Left            =   765
            TabIndex        =   72
            Top             =   225
            Width           =   195
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Afecta a ensayo ENAC"
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
            Height          =   195
            Index           =   26
            Left            =   1080
            TabIndex        =   73
            Top             =   225
            Width           =   1950
         End
      End
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   1710
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   3015
         Width           =   11670
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Uso previsto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   4815
         TabIndex        =   53
         Top             =   315
         Width           =   3975
         Begin VB.CheckBox chkTipo2 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   2205
            TabIndex        =   60
            Top             =   1125
            Width           =   195
         End
         Begin VB.CheckBox chkTipo2 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   2205
            TabIndex        =   59
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkTipo2 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   2205
            TabIndex        =   58
            Top             =   315
            Width           =   195
         End
         Begin VB.CheckBox chkTipo1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   57
            Top             =   1530
            Width           =   195
         End
         Begin VB.CheckBox chkTipo1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   56
            Top             =   1125
            Width           =   195
         End
         Begin VB.CheckBox chkTipo1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   55
            Top             =   720
            Width           =   195
         End
         Begin VB.CheckBox chkTipo1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   54
            Top             =   315
            Width           =   195
         End
         Begin VB.Label lblPeriodicidad 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            Height          =   195
            Index           =   2
            Left            =   2520
            TabIndex        =   67
            Top             =   1125
            Width           =   90
         End
         Begin VB.Label lblPeriodicidad 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   66
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblPeriodicidad 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   65
            Top             =   315
            Width           =   90
         End
         Begin VB.Label lblUso 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "3"
            Height          =   195
            Index           =   3
            Left            =   585
            TabIndex        =   64
            Top             =   1530
            Width           =   90
         End
         Begin VB.Label lblUso 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            Height          =   195
            Index           =   2
            Left            =   585
            TabIndex        =   63
            Top             =   1125
            Width           =   90
         End
         Begin VB.Label lblUso 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   62
            Top             =   720
            Width           =   90
         End
         Begin VB.Label lblUso 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   585
            TabIndex        =   61
            Top             =   315
            Width           =   90
         End
      End
      Begin Geslab.ControlPanelXP cpDocumentos 
         Height          =   2625
         Left            =   180
         TabIndex        =   48
         Top             =   315
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4630
         Caption         =   "Documentos asociados (PNT)"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   2625
         Begin XtremeSuiteControls.PushButton cmdEliminarPNT 
            Height          =   435
            Left            =   2295
            TabIndex        =   52
            Top             =   2025
            Width           =   1845
            _Version        =   851970
            _ExtentX        =   3254
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar PNT"
            Appearance      =   5
            Picture         =   "frmREX_Bote.frx":D98E
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirPNT 
            Height          =   435
            Left            =   315
            TabIndex        =   51
            Top             =   2025
            Width           =   1875
            _Version        =   851970
            _ExtentX        =   3307
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Añadir PNT"
            Appearance      =   5
            Picture         =   "frmREX_Bote.frx":141F0
         End
         Begin pryCombo.miCombo cmbDocumentos 
            Height          =   330
            Left            =   90
            TabIndex        =   50
            Top             =   1620
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   582
         End
         Begin MSComctlLib.ListView listaDocumentos 
            Height          =   1140
            Left            =   90
            TabIndex        =   49
            Top             =   405
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   2011
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
      End
      Begin VB.Label lblEquipos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones:"
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
         Left            =   315
         TabIndex        =   70
         Top             =   3195
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Condiciones de Almacenamiento"
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
      Height          =   1875
      Left            =   9000
      TabIndex        =   40
      Top             =   630
      Width           =   4590
      Begin VB.CommandButton cmdAnadirCaducidad 
         Caption         =   "+"
         Height          =   345
         Left            =   3675
         TabIndex        =   17
         Top             =   1440
         Width           =   315
      End
      Begin VB.CheckBox chkCaducidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   195
         Left            =   3150
         TabIndex        =   15
         Top             =   315
         Width           =   1185
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   6
         Left            =   1290
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   285
         Width           =   1575
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   7
         Left            =   1290
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   8
         Left            =   1290
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1050
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo cmbCaducidad 
         Height          =   315
         Left            =   660
         TabIndex        =   16
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Temperatura"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   43
         Top             =   345
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Humedad"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   690
         Width           =   690
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Productos Controlados"
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
      Height          =   1590
      Left            =   9000
      TabIndex        =   36
      Top             =   2520
      Width           =   4590
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   12
         Left            =   1755
         TabIndex        =   18
         Top             =   225
         Width           =   2610
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   11
         Left            =   1755
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   585
         Width           =   2610
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   270
         TabIndex        =   20
         Top             =   1170
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable:"
         Height          =   195
         Index           =   18
         Left            =   270
         TabIndex        =   46
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Ident. Lote"
         Height          =   195
         Index           =   16
         Left            =   270
         TabIndex        =   44
         Top             =   315
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto Certificado"
         Height          =   195
         Index           =   15
         Left            =   270
         TabIndex        =   38
         Top             =   630
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10860
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10860
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      Height          =   3495
      Left            =   45
      TabIndex        =   24
      Top             =   630
      Width           =   8895
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   13
         Left            =   6585
         TabIndex        =   11
         Top             =   2970
         Width           =   1470
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   9
         Left            =   6585
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2625
         Width           =   1470
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   10
         Left            =   6585
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2265
         Width           =   1470
      End
      Begin pryCombo.miCombo cmbSustancia 
         Height          =   330
         Left            =   1845
         TabIndex        =   1
         Top             =   540
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   582
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   5700
         TabIndex        =   3
         Top             =   900
         Width           =   2040
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   1845
         TabIndex        =   2
         Top             =   885
         Width           =   2040
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1845
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2625
         Width           =   1515
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1845
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2265
         Width           =   1515
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1845
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   6210
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1845
         TabIndex        =   5
         Top             =   1575
         Width           =   6210
      End
      Begin MSDataListLib.DataCombo cmbmat 
         Height          =   315
         Left            =   1845
         TabIndex        =   0
         Top             =   195
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   255
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbetiqueta 
         Height          =   315
         Left            =   3540
         TabIndex        =   23
         Top             =   2565
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
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
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   330
         Left            =   1845
         TabIndex        =   4
         Top             =   1230
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad que se recibe por Unidad de Pedido"
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
         Height          =   195
         Index           =   17
         Left            =   2610
         TabIndex        =   45
         Top             =   3030
         Width           =   3885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stock Mínimo"
         Height          =   195
         Index           =   12
         Left            =   5490
         TabIndex        =   39
         Top             =   2715
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Mínima Pedido"
         Height          =   195
         Index           =   14
         Left            =   4725
         TabIndex        =   37
         Top             =   2310
         Width           =   1740
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cód. Producto Prov."
         Height          =   195
         Index           =   8
         Left            =   4140
         TabIndex        =   34
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Canagrosa"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tam.Etiqueta"
         Height          =   195
         Index           =   6
         Left            =   3510
         TabIndex        =   32
         Top             =   2370
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad de Producto"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   2340
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calidad"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1995
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   28
         Top             =   1290
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sustancia&/Material"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   615
         Width           =   1335
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Restricciones"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1620
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio Compra"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   2700
         Width           =   1035
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alta/Modificación de Reactivos Externos/Productos Controlados"
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
      TabIndex        =   35
      Top             =   180
      Width           =   6660
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13050
      Top             =   30
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   13620
   End
End
Attribute VB_Name = "frmREX_Bote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
'M1166-I
Const uso As Integer = 3
Const PERIODICIDAD As Integer = 2
Const ALTO_MAX = 12210
Const ALTO_MIN = 8565
Const TOP_MIN = 7245
Const TOP_MAX = 10890
'M1166-F
 
Private Sub cabecera()
    With listaDocumentos.ColumnHeaders
        .Add , , "ID", 650, lvwColumnLeft
        .Add , , "Documento", 3300, lvwColumnLeft
    End With
    With listaEquipos.ColumnHeaders
        .Add , , "Nº", 500, lvwColumnLeft
        .Add , , "Nombre", 2600, lvwColumnLeft
        .Add , , "NºSerie", 1100, lvwColumnCenter
    End With
    With listaExistencias.ColumnHeaders
        .Add , , "General", 900, lvwColumnLeft
        .Add , , "Particular", 1050, lvwColumnCenter
        .Add , , "Cantidad", 1150, lvwColumnCenter
        .Add , , "Reactivo", 4000, lvwColumnLeft
        .Add , , "Recepción", 1050, lvwColumnCenter
        .Add , , "Apertura", 1050, lvwColumnCenter
        .Add , , "Terminado", 1, lvwColumnCenter
        .Add , , "Caducidad", 1050, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Lote", 1700, lvwColumnCenter
        .Add , , "Precio", 900, lvwColumnRight
        .Add , , "tipo_material_referencia", 1, lvwColumnRight
    End With
End Sub

Private Sub cmbmat_Change()
    If verificarMRC Then
        frameMR.visible = True
    Else
        frameMR.visible = False
    End If
    iniciarForm
End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim i As Integer
        For i = 1 To listaEquipos.ListItems.Count
            If listaEquipos.ListItems(i) = cmbEquipos.getPK_SALIDA Then
                MsgBox "El equipo ya se encuentra en la lista.", vbExclamation, App.Title
                Exit Sub
            End If
        Next
        Dim oEquipo As New clsEquipos
        oEquipo.Carga_Datos_Basicos cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
    End If
End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
    End If
End Sub

'M1166-F
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cargar_combos
    'M1166-I
    'inicializar_grid
    cabecera
    'M1166-F
    If PK <> 0 Then
        cargar_BoteReactivoEx
        cargarExistencias PK
        'M1166-I
        'JGMiniciarForm
        'M1166-F
    Else
        txtDatos(10) = "1"
        txtDatos(9) = "0"
        txtDatos(13) = "1"
        'JGM-I
        iniciarForm
        'JGM-F
    End If
End Sub
Private Sub cargarExistencias(TIPO_BOTE_ID As Long)
    Dim oBote As New clsBotes_ex
    Dim rs As ADODB.Recordset
   On Error GoTo cargarExistencias_Error

    listaExistencias.ListItems.Clear
    Set rs = oBote.listadoExistenciasPorTipo(TIPO_BOTE_ID)
    frmExistencias.Caption = " Existen un total de " & rs.RecordCount & " registros "
    If rs.RecordCount > 0 Then
        Do
            With listaExistencias.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                .SubItems(1) = rs(13) & "-" & Format(rs(12), "000") & "-" & Format(rs(3), "yy") ' Número particular
                .SubItems(2) = rs(10)
                .SubItems(3) = rs.Fields(2)
                If Not IsNull(rs.Fields(3)) Then
                    .SubItems(4) = rs.Fields(3)
                End If
                If Not IsNull(rs.Fields(4)) Then
                    .SubItems(5) = rs.Fields(4)
                End If
                If Not IsNull(rs.Fields(5)) Then
                    .SubItems(6) = rs.Fields(5)
                End If
                If IsNull(rs.Fields(6)) Then
                    .SubItems(7) = "N.A."
                Else
                    If rs(15) = 1 Then
                        .SubItems(7) = "N.A."
                    Else
                        .SubItems(7) = rs.Fields(6)
                    End If
                End If
                .SubItems(8) = rs(7)
                .SubItems(9) = rs(8)
                .SubItems(10) = Format(rs(9), "currency")
                .SubItems(11) = rs(11)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oBote = Nothing

   On Error GoTo 0
   Exit Sub

cargarExistencias_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarExistencias of Formulario frmREX_Bote"
End Sub
Private Sub cmdAnadirPNT_Click()
    If cmbDocumentos.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un documento", vbOK + vbExclamation, "Añadir documento"
        Exit Sub
    End If
    
    Dim i As Integer
    
    ' Verificar si existe el documento en la lista
    For i = 1 To listaDocumentos.ListItems.Count
        If CLng(listaDocumentos.ListItems(i).Text) = CLng(cmbDocumentos.getPK_SALIDA) Then
            MsgBox "El documento ya existe en la lista.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    
    With listaDocumentos.ListItems.Add(, , cmbDocumentos.getPK_SALIDA)
        .SubItems(1) = cmbDocumentos.getTEXTO
        
    End With
    cmbDocumentos.limpiar
End Sub

Private Sub cmdEliminarPNT_Click()
    If listaDocumentos.ListItems.Count > 0 Then
        listaDocumentos.ListItems.Remove listaDocumentos.selectedItem.Index
    End If
End Sub

Private Sub cmdParametros_Click()
    frmREX_Bote_Parametros.PK = PK
    frmREX_Bote_Parametros.Show 1
End Sub

Private Sub cmdAnadirCaducidad_Click()
    frmTipos_caducidad.Show 1
    cargar_combo cmbCaducidad, New clsTipos_caducidad
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If validar = True Then
      On Error GoTo fallo
      Dim obr As New clsTipos_bote_ex
'M1166-I
      Dim oMR As New clsTipos_bote_ex_mr
'M1166-F
      With obr
            .setTIPO_REACTIVO_EX_ID = cmbSustancia.getPK_SALIDA
            If cmbmat.BoundText = "" Then
                .setTIPO_M_REFERENCIA_ID = 1
            Else
                .setTIPO_M_REFERENCIA_ID = cmbmat.BoundText
            End If
            .setCODIGO_PROVEEDOR = txtDatos(5)
            .setRESTRICCIONES = txtDatos(0)
            .setCODIGO = txtDatos(4)
            .setCALIDAD = txtDatos(1)
            .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
            .setCANTIDAD = txtDatos(2)
            If txtDatos(3) <> "" Then
                .setPRECIO = moneda_bd(txtDatos(3))
            Else
                .setPRECIO = moneda_bd("0")
            End If
            If cmbEtiqueta.BoundText = "" Then
                .setTAMANO_ETIQUETA_ID = 1
            Else
                .setTAMANO_ETIQUETA_ID = cmbEtiqueta.BoundText
            End If
            If txtDatos(10) <> "" Then
                .setCANTIDAD_MINIMA_PEDIDO = txtDatos(10)
            End If
            If txtDatos(9) <> "" Then
                .setSTOCK_MINIMO = txtDatos(9)
            End If
            If cmbCaducidad.BoundText <> "" Then
                .setTIPO_CADUCIDAD_ID = cmbCaducidad.BoundText
            End If
            .setTEMPERATURA = txtDatos(6)
            .setHUMEDAD = txtDatos(7)
            .setOTROS = txtDatos(8)
            .setCANTIDAD_UNIDAD_PEDIDO = txtDatos(13)
            .setCODIGO_IDEN_LOTE = txtDatos(12)
            .setTEXTO_CERTIFICADO = txtDatos(11)
            .setRESPONSABLE_ID = cmbResponsable.getPK_SALIDA
      End With
'M1166-I
'TIPOS M.R. y M.R.C.
      If verificarMRC Then
         Dim indice As Integer
         Dim str As String
         With oMR
            .setENSAYO_ENAC = chkENAC.Value
            .setEQUIPO_ID = cmbEquipos.getPK_SALIDA
            .setOBSERVACIONES = Trim(txtObservaciones.Text)
            .setTIPO_BOTE_EX_ID = PK
            'USO
            str = ""
            For indice = 0 To uso
                If chkTipo1(indice).Value = 1 Then
                    str = str & CStr(indice) & ";"
                End If
            Next indice
            .setUSO_PREVISTO = str
            'PERIODICIDAD
            str = ""
            For indice = 0 To PERIODICIDAD
                If chkTipo2(indice).Value = 1 Then
                    str = str & CStr(indice) & ";"
                End If
            Next indice
            .setPERIODICIDAD = str
            .setPNTS = cargarPNTs()
         End With
         anadirEquipos
      End If
'M1166-F
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo Bote de Reactivo Externo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            PK = obr.Insertar
            If PK = 0 Then
                Exit Sub
            End If
'M1166-I
            If verificarMRC Then
               oMR.Insertar
            End If
'M1166-F
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar un Bote de Reactivo Externo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If obr.Modificar(PK) = False Then
                Exit Sub
            End If
'M1166-I
            If verificarMRC Then
                oMR.EliminarTipo PK
                oMR.Insertar
            End If
'M1166-F
        Else
            Exit Sub
        End If
      End If
      'insertar_parametros
      If PK = 0 Then
          MsgBox "El Bote de Reactivo Externo se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El Bote de Reactivo Externo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el bote. " & Err.Description)
End Sub
Private Function verificarMRC() As Boolean
    verificarMRC = False
    ' JGM
    If cmbmat.BoundText <> "" Then
        If CInt(cmbmat.BoundText) = 2 Or CInt(cmbmat.BoundText) = 3 Then
            verificarMRC = True
        End If
    End If
End Function
Private Sub cargarEquipos()
    Dim oEquipos As New clsTipos_bote_ex_mr_equipos
    Dim rs As ADODB.Recordset
    Set rs = oEquipos.Listado(PK)
    listaEquipos.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEquipos = Nothing
End Sub
Private Sub anadirEquipos()
' Insertar equipos en la tabla relacionada para el informe
    Dim i As Integer
    Dim oEquipos As New clsTipos_bote_ex_mr_equipos
    oEquipos.Eliminar PK
    For i = 1 To listaEquipos.ListItems.Count
        With oEquipos
            .setTIPO_BOTE_EX_ID = PK
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .Insertar
        End With
    Next i
End Sub
Private Sub chkCaducidad_Click()
    If chkCaducidad.Value = Checked Then
        cmbCaducidad.Enabled = True
    Else
        cmbCaducidad.Enabled = False
        cmbCaducidad.Text = ""
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
'M1166-I
Public Sub iniciarForm()
    'Carga los valores del control FRAME específico para tipos M.R.
    If verificarMRC Then
        redimensionar ALTO_MAX, TOP_MAX
        cargaCaption
        Dim oMR As New clsTipos_bote_ex_mr
        If oMR.CargaTipo(PK) Then
            Dim strPNTs() As String
            Dim strUso() As String
            Dim intCount As Integer
            Dim VALOR As Integer
            frameMR.visible = True
            
            'DOCUMENTANCIÓN ASOCIADA. LISTADO DE PNTS
            listaDocumentos.ListItems.Clear
            strPNTs = Split(oMR.getPNTS, ";")
            For intCount = LBound(strPNTs) To UBound(strPNTs)
                If strPNTs(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
                  VALOR = CInt(Solo_Numeros(strPNTs(intCount)))
                  Dim oDoc As New clsCa_documentos
                  oDoc.Carga CLng(VALOR)
                  With listaDocumentos.ListItems.Add(, , oDoc.getID_DOCUMENTO)
                    .SubItems(1) = oDoc.getNOMBRE
                  End With
                End If
            Next intCount
            
            Dim oDeco As New clsDecodificadora
            'USO PREVISTO
            strUso = Split(oMR.getUSO_PREVISTO, ";")
            For intCount = LBound(strUso) To UBound(strUso)
                If strUso(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
'                  VALOR = CInt(Solo_Numeros(strUso(intCount)))
'                  oDeco.Carga_valor 190, CLng(VALOR)
                  chkTipo1(strUso(intCount)).Value = 1
                End If
            Next intCount
            'PERIODICIDAD
            strUso = Split(oMR.getPERIODICIDAD, ";")
            For intCount = LBound(strUso) To UBound(strUso)
                If strUso(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
'                  VALOR = CInt(Solo_Numeros(strUso(intCount)))
'                  oDeco.Carga_valor 191, CLng(VALOR)
                  chkTipo2(strUso(intCount)).Value = 1
                End If
            Next intCount
            Set oDeco = Nothing
            Set oDoc = Nothing
            
            'INFORMACIÓN GENERAL MRC/MR
            chkENAC.Value = oMR.getENSAYO_ENAC
            cmbEquipos.MostrarElemento oMR.getEQUIPO_ID
            txtObservaciones.Text = oMR.getOBSERVACIONES
            
            'EQUIPOS
             cargarEquipos
        End If
    Else
        redimensionar ALTO_MIN, TOP_MIN
        frameMR.visible = False
    End If
End Sub
Private Sub redimensionar(ALTO_FORMULARIO As Long, TOP_BOTONES As Long)
    On Error Resume Next
        Me.Height = ALTO_FORMULARIO
        cmdParametros.top = TOP_BOTONES
        cmdok.top = TOP_BOTONES
        cmdcancel.top = TOP_BOTONES
        
        Me.top = Screen.Height / 2 - Me.Height / 2
End Sub
Private Sub cargaCaption()
    'Leyenda(caption) en checks (se cargan dinámicamente para que siempre haya correspondencia entre lista de valores y contenidos)
    Dim oDeco As New clsDecodificadora
    Dim i As Long
    'USO
    For i = 0 To uso
        If oDeco.Carga_valor(190, i) Then
            lblUso(i).Caption = oDeco.getDESCRIPCION
        End If
    Next i
    'PERIODICIDAD
    For i = 0 To PERIODICIDAD
        If oDeco.Carga_valor(191, i) Then
            lblPeriodicidad(i).Caption = oDeco.getDESCRIPCION
        End If
    Next i
End Sub
Private Function cargarPNTs() As String
    Dim cadena As String
    Dim i As Integer
    cadena = ""
    If listaDocumentos.ListItems.Count > 0 Then
       For i = 1 To listaDocumentos.ListItems.Count
           cadena = cadena & listaDocumentos.ListItems(i).Text & ";"
       Next i
    End If
    cargarPNTs = cadena
End Function
'M1166-F
Public Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbmat, DECODIFICADORA.REX_TIPOS
    cargar_combo cmbEtiqueta, New clsTamanos_etiqueta
    llenar_combo cmbProveedor, New clsProveedor, 0, frmProveedores_Detalle, ""
    If PK <> 0 Then
        llenar_combo cmbSustancia, New clsTipos_reactivo_ex, 0, frmREX_Reactivo, ""
    Else
        llenar_combo cmbSustancia, New clsTipos_reactivo_ex, 0, frmREX_Reactivo, " ANULADO = 0 "
    End If
    cargar_combo cmbCaducidad, New clsTipos_caducidad
    llenar_combo cmbResponsable, New clsUsuarios, 0, frmUsuarios, ""
'M1166-I
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbDocumentos, New clsCa_documentos, 0, frmCA_Documento, ""
'M1166-F
End Sub
Private Sub listaExistencias_DblClick()
    If listaExistencias.ListItems.Count > 0 Then
        frmREX_Bote_Modificacion.PK = CLng(listaExistencias.ListItems(listaExistencias.selectedItem.Index).Text)
        frmREX_Bote_Modificacion.Show 1
        cargarExistencias PK
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 3 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = Format(txtDatos(Index), "currency")
        End If
    End If
End Sub
Public Sub cargar_BoteReactivoEx()
    Dim obr As New clsTipos_bote_ex
    With obr
      .CARGAR (PK)
      cmbSustancia.MostrarElemento .getTIPO_REACTIVO_EX_ID
      cmbProveedor.MostrarElemento .getPROVEEDOR_ID
      txtDatos(0) = .getRESTRICCIONES
      txtDatos(1) = .getCALIDAD
      txtDatos(2) = .getCANTIDAD
      txtDatos(3) = Format(.getPRECIO, "currency")
      txtDatos(4) = .getCODIGO
      txtDatos(5) = .getCODIGO_PROVEEDOR
      cmbmat.BoundText = .getTIPO_M_REFERENCIA_ID
      cmbEtiqueta.BoundText = .getTAMANO_ETIQUETA_ID
      ' Nuevos campos
      txtDatos(6) = .getTEMPERATURA
      txtDatos(7) = .getHUMEDAD
      txtDatos(8) = .getOTROS
      txtDatos(9) = .getSTOCK_MINIMO
      txtDatos(10) = .getCANTIDAD_MINIMA_PEDIDO
      txtDatos(11) = .getTEXTO_CERTIFICADO
      txtDatos(13) = .getCANTIDAD_UNIDAD_PEDIDO
      txtDatos(12) = .getCODIGO_IDEN_LOTE
      
      cmbCaducidad.BoundText = .getTIPO_CADUCIDAD_ID
      If .getTIPO_CADUCIDAD_ID <> 0 Then
        chkCaducidad.Value = Checked
        cmbCaducidad.Enabled = True
      End If
      
      cmbResponsable.MostrarElemento .getRESPONSABLE_ID
    End With
    Set obr = Nothing
    ' Parametros
    'M1166-I
    'Dim RS As ADODB.Recordset
    'Dim oTBP As New clsTipos_bote_ex_parametros
    'Set RS = oTBP.Listado(PK)
    'If RS.RecordCount > 0 Then
    'Dim i As Integer
    'i = 0
    '    Do
    '        X(i, COLS.PARAMETRO) = CStr(RS(0))
    '        X(i, COLS.VALOR) = CStr(RS(1))
    '        X(i, COLS.UNIDADES) = CStr(RS(2))
    '        RS.MoveNext
    '        i = i + 1
    '    Loop Until RS.EOF
    'End If
    'Set oTBP = Nothing
    'Set RS = Nothing
    'M1166-F
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbmat.Text = "" Then
        MsgBox "Debe seleccionar el tipo de reactivo.", vbExclamation, App.Title
        cmbmat.SetFocus
        validar = False
        Exit Function
    End If
    If cmbSustancia.getTEXTO = "" Then
        MsgBox "Debe seleccionar un tipo de Sustancia/Material.", vbExclamation, App.Title
        cmbSustancia.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(4) = "" Then
        MsgBox "Debe introducir el código de reactivo.", vbExclamation, App.Title
        txtDatos(4).SetFocus
        validar = False
        Exit Function
    End If
    If cmbProveedor.getTEXTO = "" Then
        MsgBox "Debe seleccionar un proveedor.", vbExclamation, App.Title
        cmbProveedor.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(10) <> "" Then
        If Not IsNumeric(txtDatos(10)) Then
            MsgBox "La cantidad mínima de pedido debe ser numérica (Número de unidades mínima a pedir de la cantidad definida).", vbExclamation, App.Title
            txtDatos(10).SetFocus
            validar = False
            Exit Function
        Else
            If CInt(txtDatos(10)) < 1 Then
                MsgBox "La cantidad mínima de pedido debe ser mayor de 0.", vbExclamation, App.Title
                txtDatos(10).SetFocus
                validar = False
                Exit Function
            End If
        End If
    Else
        MsgBox "La cantidad mínima de pedido debe ser numérica (Número de unidades mínima a pedir de la cantidad definida).", vbExclamation, App.Title
        txtDatos(10).SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(13) <> "" Then
        If Not IsNumeric(txtDatos(13)) Then
            MsgBox "La cantidad por Unidad de pedido debe ser numérica (Número de unidades que componen cada unidad, por ejemplo, si 1 paquete trae 6 unidades, se pedirá 1 paquete, pero realmente vienen 6 unidades por lo que hay que indicar 6).", vbExclamation, App.Title
            txtDatos(13).SetFocus
            validar = False
            Exit Function
        Else
            If CInt(txtDatos(13)) < 1 Then
                MsgBox "La cantidad por Unidad de pedido de pedido debe ser mayor de 0.", vbExclamation, App.Title
                txtDatos(13).SetFocus
                validar = False
                Exit Function
            End If
        End If
    Else
        MsgBox "La cantidad por Unidad de pedido debe ser numérica (Número de unidades que componen cada unidad, por ejemplo, si 1 paquete trae 6 unidades, se pedirá 1 paquete, pero realmente vienen 6 unidades por lo que hay que indicar 6).", vbExclamation, App.Title
        txtDatos(13).SetFocus
        validar = False
        Exit Function
    End If
    
    If txtDatos(9) <> "" Then
        If Not IsNumeric(txtDatos(9)) Then
            MsgBox "El stock mínimo debe ser numérico.", vbExclamation, App.Title
            txtDatos(9).SetFocus
            validar = False
            Exit Function
        End If
    End If
End Function
'Private Sub inicializar_grid()
'    X.ReDim 0, filas, 0, Col
'    X.Clear
'    Set grid.Array = X
'    grid.Refresh
'End Sub

'Private Sub insertar_parametros()
'    ' Evidencias
'   Dim oTBP As New clsTipos_bote_ex_parametros
'   On Error GoTo insertar_parametros_Error''
'
'    oTBP.Eliminar PK
'    Dim i As Integer
'    For i = X.LowerBound(1) To X.UpperBound(1)
'        If Trim(X.value(i, COLS.PARAMETRO)) <> "" Then
'            With oTBP
'               .setTIPO_BOTE_EX_ID = PK
'               .setORDEN = i
'                .setPARAMETRO = X.value(i, COLS.PARAMETRO)
'                .setVALOR = X.value(i, COLS.VALOR)
'                .setUNIDADES = X.value(i, COLS.UNIDADES)
'                .Insertar
'            End With
'        End If
'    Next
'    Set oTBP = Nothing
'
'   On Error GoTo 0
'   Exit Sub
'
'insertar_parametros_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_parametros of Formulario frmREX_Bote"
'End Sub

'M1166-I
'Valor del USO PREVISTO. Lista de Valores.
Private Sub usoPrevisto()
   Dim i As Integer
   Dim valores As String
   For i = 0 To 6
       If chkTipo1(i).Value = True Then
          valores = valores & lblCampos(i).Caption & ";"
       End If
   Next i
   indice = indice + 1
End Sub
'M1166-F
