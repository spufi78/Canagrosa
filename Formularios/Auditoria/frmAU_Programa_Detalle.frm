VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmAU_Programa_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Programa de Auditorías"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15420
   Icon            =   "frmAU_Programa_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmAU_Programa_Detalle.frx":08CA
   ScaleHeight     =   9045
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8100
      Width           =   1725
   End
   Begin VB.Frame frmHistorico 
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
      Height          =   6900
      Left            =   2880
      TabIndex        =   14
      Top             =   900
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton cmdOcultarVersiones 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Minimizar"
         Height          =   870
         Left            =   8235
         Picture         =   "frmAU_Programa_Detalle.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5940
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker fAprobacion 
         Height          =   330
         Left            =   7965
         TabIndex        =   39
         Top             =   5445
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fCreacion 
         Height          =   330
         Left            =   7965
         TabIndex        =   37
         Top             =   4995
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1365
         Index           =   4
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   3555
         Width           =   9180
      End
      Begin MSComctlLib.ListView listaVersiones 
         Height          =   2370
         Left            =   135
         TabIndex        =   2
         Top             =   855
         Width           =   9180
         _ExtentX        =   16193
         _ExtentY        =   4180
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
      Begin pryCombo.miCombo cmbVersionCreador 
         Height          =   330
         Left            =   1035
         TabIndex        =   33
         Top             =   4995
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVersionAprobador 
         Height          =   330
         Left            =   1035
         TabIndex        =   35
         Top             =   5445
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   582
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   8865
         Picture         =   "frmAU_Programa_Detalle.frx":14D6
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lbltitulo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Histórico de Versiones del Programa"
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
         Index           =   2
         Left            =   135
         TabIndex        =   42
         Top             =   315
         Width           =   3855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   645
         Left            =   45
         Top             =   135
         Width           =   9405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Aprobación"
         Height          =   195
         Index           =   10
         Left            =   6885
         TabIndex        =   38
         Top             =   5490
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Creación"
         Height          =   195
         Index           =   9
         Left            =   6885
         TabIndex        =   36
         Top             =   5040
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobador"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   34
         Top             =   5490
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Creador"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   32
         Top             =   5040
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   31
         Top             =   3285
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdAprobar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aprobar"
      Enabled         =   0   'False
      Height          =   870
      Left            =   45
      Picture         =   "frmAU_Programa_Detalle.frx":1DA0
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8100
      Width           =   1725
   End
   Begin VB.CommandButton cmdNuevaVersion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nueva Versión"
      Enabled         =   0   'False
      Height          =   870
      Left            =   1800
      Picture         =   "frmAU_Programa_Detalle.frx":266A
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8100
      Width           =   1725
   End
   Begin VB.CommandButton cmdVersiones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consultar Versiones"
      Height          =   870
      Left            =   3555
      Picture         =   "frmAU_Programa_Detalle.frx":2F34
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8100
      Width           =   1725
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lista de Distribución"
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
      Height          =   2355
      Left            =   45
      TabIndex        =   17
      Top             =   5715
      Width           =   6885
      Begin VB.CommandButton cmdEliminaDistribucion 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Programa_Detalle.frx":37FE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   405
         Width           =   735
      End
      Begin VB.CommandButton cmdInsertaDistribucion 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Programa_Detalle.frx":40C8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1215
         Width           =   735
      End
      Begin MSComctlLib.ListView listaDistribucion 
         Height          =   1695
         Left            =   135
         TabIndex        =   20
         Top             =   225
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2990
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
      Begin pryCombo.miCombo cmbDistribucion 
         Height          =   330
         Left            =   135
         TabIndex        =   21
         Top             =   1935
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   582
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14310
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8100
      Width           =   1050
   End
   Begin VB.Frame frmanalisis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Áreas a auditar"
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
      Height          =   7305
      Left            =   7020
      TabIndex        =   12
      Top             =   765
      Width           =   8370
      Begin VB.CommandButton cmdModificarArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   7560
         Picture         =   "frmAU_Programa_Detalle.frx":4992
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modificar"
         Top             =   4815
         Width           =   735
      End
      Begin VB.CheckBox chkFechaRealizacion 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2610
         TabIndex        =   49
         Top             =   6795
         Width           =   195
      End
      Begin VB.CommandButton cmdEliminaArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   7560
         Picture         =   "frmAU_Programa_Detalle.frx":525C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminar"
         Top             =   4095
         Width           =   735
      End
      Begin VB.CommandButton cmdInsertaArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   7560
         Picture         =   "frmAU_Programa_Detalle.frx":5B26
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Añadir"
         Top             =   5535
         Width           =   735
      End
      Begin MSComctlLib.ListView listaAreas 
         Height          =   6015
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   10610
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
      Begin pryCombo.miCombo cmbArea 
         Height          =   330
         Left            =   945
         TabIndex        =   16
         Top             =   6345
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   945
         TabIndex        =   29
         Top             =   6750
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaRealizacion 
         Height          =   330
         Left            =   3960
         TabIndex        =   47
         Top             =   6750
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Realización"
         Height          =   195
         Index           =   13
         Left            =   2880
         TabIndex        =   48
         Top             =   6795
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Prevista"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   28
         Top             =   6795
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Área"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   27
         Top             =   6390
         Width           =   330
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Generales"
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
      Height          =   4875
      Index           =   1
      Left            =   45
      TabIndex        =   8
      Top             =   765
      Width           =   6885
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Height          =   330
         Index           =   5
         Left            =   4410
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Text            =   "PDTE. APROBACION"
         Top             =   315
         Width           =   2295
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Height          =   330
         Index           =   3
         Left            =   2655
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   315
         Width           =   810
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Height          =   330
         Index           =   2
         Left            =   675
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   315
         Width           =   900
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   0
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   1035
         Width           =   6660
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   2130
         Index           =   1
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1890
         Width           =   6660
      End
      Begin pryCombo.miCombo cmbAprobador 
         Height          =   330
         Left            =   135
         TabIndex        =   22
         Top             =   4320
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   11
         Left            =   3780
         TabIndex        =   45
         Top             =   405
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Versión"
         Height          =   195
         Index           =   4
         Left            =   1890
         TabIndex        =   26
         Top             =   405
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   405
         Width           =   285
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   810
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobador"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   13
         Top             =   4095
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1665
         Width           =   1065
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los datos necesarios para el Programa de Auditoría"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   360
      Width           =   4170
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14850
      Picture         =   "frmAU_Programa_Detalle.frx":63F0
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión del Programa de Auditorías"
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
      TabIndex        =   10
      Top             =   90
      Width           =   3720
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   -45
      Width           =   15390
   End
End
Attribute VB_Name = "frmAU_Programa_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Integer
Public NUEVA_VERSION As Boolean
Private Sub cmdAdjuntos_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_AU_PROGRAMA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub

Private Sub chkFechaRealizacion_Click()
    If chkFechaRealizacion.value = Checked Then
        fechaRealizacion.Enabled = True
    Else
        fechaRealizacion.Enabled = False
    End If
End Sub

Private Sub cmdAprobar_Click()
   On Error GoTo cmdAprobar_Click_Error

    If MsgBox("¿Desea aprobar el programa?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oPrograma As New clsAu_programa
        oPrograma.Aprobar PK
        Dim oVersion As New clsAu_programa_versiones
        oVersion.Aprobar PK, txtDatos(3), Format(Date, "yyyy-mm-dd")
        MsgBox "Programa aprobado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdAprobar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAprobar_Click of Formulario frmAU_Programa_Detalle"
End Sub

Private Sub cmdEliminaArea_Click()
    If listaAreas.ListItems.Count > 0 Then
       listaAreas.ListItems.Remove listaAreas.selectedItem.Index
    End If
End Sub

Private Sub cmdEliminaDistribucion_Click()
    If listaDistribucion.ListItems.Count > 0 Then
       listaDistribucion.ListItems.Remove listaDistribucion.selectedItem.Index
    End If
End Sub

Private Sub cmdInsertaArea_Click()
    If cmbarea.getTEXTO <> "" Then
        With listaAreas.ListItems.Add(, , Format(cmbarea.getPK_SALIDA, "000"))
            .SubItems(1) = cmbarea.getTEXTO
            .SubItems(2) = fecha
        End With
        cmbarea.Limpiar
    End If
End Sub

Private Sub cmdInsertaDistribucion_Click()
    If cmbDistribucion.getTEXTO <> "" Then
        With listaDistribucion.ListItems.Add(, , cmbDistribucion.getPK_SALIDA)
            .SubItems(1) = cmbDistribucion.getTEXTO
            Dim oUsuario As New clsUsuarios
            oUsuario.CARGAR cmbDistribucion.getPK_SALIDA
            .SubItems(2) = oUsuario.getEMAIL
        End With
        cmbDistribucion.Limpiar
    End If
End Sub

Private Sub cmdModificarArea_Click()
    If listaAreas.ListItems.Count > 0 And PK > 0 Then
        Dim oAu As New clsAu_programa_areas
        With oAu
            If chkFechaRealizacion.value = Checked Then
                .setFECHA_REALIZACION = "'" & Format(fechaRealizacion, "yyyy-mm-dd") & "'"
            Else
                .setFECHA_REALIZACION = "NULL"
            End If
            .ModificarRealizacion PK, listaAreas.ListItems(listaAreas.selectedItem.Index).Text
            cargarAreas
        End With
    End If
End Sub

Private Sub cmdNuevaVersion_Click()
    cmdNuevaVersion.Enabled = False
    txtDatos(0).Locked = True
    txtDatos(1).Locked = False
    cmdok.Enabled = True
    cmdEliminaDistribucion.Enabled = True
    cmdInsertaDistribucion.Enabled = True
    cmdEliminaArea.Enabled = True
    cmdInsertaArea.Enabled = True
    txtDatos(3) = CInt(txtDatos(3)) + 1
    txtDatos(5) = "PDTE.CREACION"
    NUEVA_VERSION = True
    txtDatos(1) = ""
    txtDatos(1).SetFocus
End Sub

Private Sub cmdOcultarVersiones_Click()
    frmHistorico.Visible = False
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim PROGRAMA As Integer
      Dim oPrograma As New clsAu_programa
      With oPrograma
        .setANNO = txtDatos(2)
        .setVERSION = txtDatos(3)
        .setDESCRIPCION = txtDatos(0)
        .setAPROBADO = 0
        If PK = 0 Then
            PROGRAMA = .Insertar
        Else
            .Modificar PK
            PROGRAMA = PK
        End If
      End With
      ' Versión
      Dim oVersion As New clsAu_programa_versiones
      With oVersion
        .setPROGRAMA_ID = PROGRAMA
        .setVERSION = txtDatos(3)
        .setOBSERVACIONES = txtDatos(1)
        .setUSUARIO_CREACION = usuario.getID_EMPLEADO
        .setFECHA_CREACION = Format(Date, "yyyy-mm-dd")
        .setUSUARIO_APROBACION = cmbAprobador.getPK_SALIDA
        .setFECHA_APROBACION = Format(Date, "YYYY-MM-DD")
        If PK = 0 Or NUEVA_VERSION = True Then
            .Insertar
        Else
            .Modificar PK, txtDatos(3)
        End If
      End With
      ' Lista de Distribución
      Dim oDistribucion As New clsAu_programa_distribucion
      Dim i As Integer
      If PK <> 0 Then
        oDistribucion.Eliminar PK
      End If
      For i = 1 To listaDistribucion.ListItems.Count
        With oDistribucion
            .setPROGRAMA_ID = PROGRAMA
            .setUSUARIO_ID = listaDistribucion.ListItems(i).Text
            .setTIPO_USUARIO = 1
            .setORDEN = i
            .Insertar
        End With
      Next
      ' Lista de Áreas
      Dim oAreas As New clsAu_programa_areas
      If PK <> 0 Then
        oAreas.Eliminar PK
      End If
      For i = 1 To listaAreas.ListItems.Count
        With oAreas
            .setPROGRAMA_ID = PROGRAMA
            .setAREA_ID = listaAreas.ListItems(i).Text
            .setFECHA_PREVISTA = Format(listaAreas.ListItems(i).SubItems(2), "yyyy-mm-dd")
            If Trim(listaAreas.ListItems(i).SubItems(3)) = "" Then
                .setFECHA_REALIZACION = "NULL"
            Else
                .setFECHA_REALIZACION = "'" & Format(listaAreas.ListItems(i).SubItems(3), "yyyy-mm-dd") & "'"
            End If
            .setORDEN = i
            .Insertar
        End With
      Next
      ' Areas impresion
      Dim oAI As New clsAu_programa_areas_impresion
      
      listaAreas.Sorted = True
      listaAreas.SortKey = 1

      If PK <> 0 Then
        oAI.Eliminar CLng(PK)
      End If
      Dim aux_area As Integer
      aux_area = 0
      
         With oAI
            .setM1 = 0
            .setM2 = 0
            .setM3 = 0
            .setM4 = 0
            .setM5 = 0
            .setM6 = 0
            .setM7 = 0
            .setM8 = 0
            .setM9 = 0
            .setM10 = 0
            .setM11 = 0
            .setM12 = 0
        End With
        
      For i = 1 To listaAreas.ListItems.Count
         With oAI
            If (aux_area <> 0 And listaAreas.ListItems(i).Text <> aux_area) Then
                .Insertar
                .setM1 = 0
                .setM2 = 0
                .setM3 = 0
                .setM4 = 0
                .setM5 = 0
                .setM6 = 0
                .setM7 = 0
                .setM8 = 0
                .setM9 = 0
                .setM10 = 0
                .setM11 = 0
                .setM12 = 0
            End If
            .setPROGRAMA_ID = PROGRAMA
            .setAREA_ID = listaAreas.ListItems(i).Text
            Select Case CInt(Format(listaAreas.ListItems(i).SubItems(2), "mm"))
            Case 1
                .setM1 = 1
            Case 2
                .setM2 = 1
            Case 3
                .setM3 = 1
            Case 4
                .setM4 = 1
            Case 5
                .setM5 = 1
            Case 6
                .setM6 = 1
            Case 7
                .setM7 = 1
            Case 8
                .setM8 = 1
            Case 9
                .setM9 = 1
            Case 10
                .setM10 = 1
            Case 11
                .setM11 = 1
            Case 12
                .setM12 = 1
            End Select
            
            aux_area = listaAreas.ListItems(i).Text
        End With
      Next
      If listaAreas.ListItems.Count > 0 Then
        oAI.Insertar
      End If
      
      If PK = 0 Then
          If MsgBox("El Programa se ha introducido correctamente. ¿Desea enviar al aprobador?", vbYesNo + vbInformation, App.Title) = vbYes Then
            oPrograma.Enviar_Al_Aprobador PROGRAMA
          End If
      Else
          If MsgBox("El Programa se ha modificado correctamente. ¿Desea enviar al aprobador?", vbYesNo + vbInformation, App.Title) = vbYes Then
            oPrograma.Enviar_Al_Aprobador PROGRAMA
          End If
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmAU_Programa_Detalle"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVersiones_Click()
    frmHistorico.Visible = True
    listaVersiones_Click
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    fecha = Date
    fechaRealizacion = Date
    NUEVA_VERSION = False
    If PK = 0 Then
        txtDatos(2) = Year(Date)
        txtDatos(3) = "1"
        txtDatos(5) = "CREACION"
        cmdAdjuntos.Visible = False
    Else
        CARGAR
    End If
End Sub

Private Sub listaAreas_Click()
    If listaAreas.ListItems.Count > 0 Then
        cmbarea.MostrarElemento listaAreas.ListItems(listaAreas.selectedItem.Index).Text
        fecha = listaAreas.ListItems(listaAreas.selectedItem.Index).SubItems(2)
        If listaAreas.ListItems(listaAreas.selectedItem.Index).SubItems(3) = "" Then
            chkFechaRealizacion.value = Unchecked
            fechaRealizacion.Enabled = False
        Else
            chkFechaRealizacion.value = Checked
            fechaRealizacion.Enabled = True
            fechaRealizacion = listaAreas.ListItems(listaAreas.selectedItem.Index).SubItems(3)
        End If
    End If
End Sub

Private Sub listaVersiones_Click()
    If listaVersiones.ListItems.Count > 0 Then
        txtDatos(4) = listaVersiones.ListItems(listaVersiones.selectedItem.Index).SubItems(1)
        cmbVersionCreador.MostrarElemento listaVersiones.ListItems(listaVersiones.selectedItem.Index).SubItems(2)
        fCreacion = listaVersiones.ListItems(listaVersiones.selectedItem.Index).SubItems(3)
        cmbVersionAprobador.MostrarElemento listaVersiones.ListItems(listaVersiones.selectedItem.Index).SubItems(4)
        If listaVersiones.ListItems(listaVersiones.selectedItem.Index).SubItems(5) <> "" Then
            fAprobacion = listaVersiones.ListItems(listaVersiones.selectedItem.Index).SubItems(5)
            fAprobacion.Visible = True
        Else
            fAprobacion.Visible = False
        End If
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub CARGAR()
    Dim oPrograma As New clsAu_programa
   On Error GoTo cargar_Error

    With oPrograma
        .Carga PK
        txtDatos(2) = .getANNO
        txtDatos(3) = .getVERSION
        txtDatos(0) = .getDESCRIPCION
        Select Case .getAPROBADO
            Case 0 ' Pdte. Creación
                cmdNuevaVersion.Enabled = False
                txtDatos(5) = "PTE.CREACIÓN"
            Case 1 ' Aprobado
                txtDatos(5) = "APROBADO"
                cmdNuevaVersion.Enabled = True
                txtDatos(0).Locked = True
                txtDatos(1).Locked = True
                cmdok.Enabled = False
                cmdEliminaDistribucion.Enabled = False
                cmdInsertaDistribucion.Enabled = False
                cmdEliminaArea.Enabled = False
                cmdInsertaArea.Enabled = False
                cmbarea.desactivar
                fecha.Enabled = False
            Case 2 ' Pdte. Aprobación
                txtDatos(5) = "PDTE.APROBACIÓN"
                cmdNuevaVersion.Enabled = False
                txtDatos(0).Locked = True
                txtDatos(1).Locked = True
                cmdok.Enabled = False
                cmdEliminaDistribucion.Enabled = False
                cmdInsertaDistribucion.Enabled = False
                cmdEliminaArea.Enabled = False
                cmdInsertaArea.Enabled = False
                cmbarea.desactivar
                fecha.Enabled = False
        End Select
    End With
    ' Versiones
    Dim oVersion As New clsAu_programa_versiones
    Dim rs As ADODB.Recordset
    Set rs = oVersion.Listado(PK)
    If rs.RecordCount > 0 Then
        Dim i As Integer
        i = 1
        txtDatos(1) = rs(1) ' Observacion de la versión en vigor
        cmbAprobador.MostrarElemento rs(4)  ' Usuario aprobador
        ' Si esta Pdte. de Aprobación y el usuario es el logeado
        If oPrograma.getAPROBADO = 2 And usuario.getID_EMPLEADO = rs(4) Then
            cmdAprobar.Enabled = True
        End If
        Do
            With listaVersiones.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = rs(4)
                If i = 1 Then
                    If oPrograma.getAPROBADO = 1 Then
                        .SubItems(5) = rs(5)
                    Else
                        .SubItems(5) = ""
                    End If
                End If
                i = i + 1
            End With
                    
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Usuarios
    Dim oDistribucion As New clsAu_programa_distribucion
    Set rs = oDistribucion.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With listaDistribucion.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(3) ' Correo
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    cargarAreas

   On Error GoTo 0
   Exit Sub

cargar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar of Formulario frmAU_Programa_Detalle"
End Sub
Private Sub cargarAreas()
    ' Areas
    Dim oAreas As New clsAu_programa_areas
    Set rs = oAreas.Listado(PK)
    listaAreas.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaAreas.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1) ' Area
                .SubItems(2) = rs(2) ' Fecha
                If IsNull(rs(3)) Then
                    .SubItems(3) = ""
                Else
                    .SubItems(3) = rs(3) ' Fecha Realizacion
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cargar_combos()
    llenar_combo cmbAprobador, New clsUsuarios, 0, Me, ""
    llenar_combo cmbDistribucion, New clsUsuarios, 0, Me, ""
    llenar_combo cmbVersionCreador, New clsUsuarios, 0, Me, ""
    llenar_combo cmbVersionAprobador, New clsUsuarios, 0, Me, ""
    llenar_combo cmbarea, New clsAu_areas, 0, frmAU_Areas_Detalle, ""
End Sub
Private Sub cabecera()
    With listaDistribucion.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Usuario", 2900, lvwColumnLeft
        .Add , , "Correo", 2515, lvwColumnLeft
    End With
    With listaAreas.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Area", 4500, lvwColumnLeft
        .Add , , "F.Prevista", 1200, lvwColumnCenter
        .Add , , "F.Realización", 1200, lvwColumnCenter
    End With
    With listaVersiones.ColumnHeaders
        .Add , , "Versión", 1180, lvwColumnLeft
        .Add , , "Observaciones", 8000, lvwColumnLeft
        .Add , , "Usuario_creacion", 0, lvwColumnLeft
        .Add , , "Fecha_creacion", 0, lvwColumnLeft
        .Add , , "Usuario_aprobador", 0, lvwColumnLeft
        .Add , , "Fecha_aprobacion", 0, lvwColumnLeft
    End With
End Sub

Private Function validar() As Boolean
    validar = True
    If txtDatos(0) = "" Then
        MsgBox "Debe indicar una Descripción.", vbExclamation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(1) = "" Then
        MsgBox "Debe indicar una Observación.", vbExclamation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If cmbAprobador.getTEXTO = "" Then
        MsgBox "Debe indicar el usuario Aprobador.", vbExclamation, App.Title
        cmbAprobador.SetFocus
        validar = False
        Exit Function
    End If
    If listaDistribucion.ListItems.Count = 0 Then
        MsgBox "Debe añadir al menos un usuario en la lista de distribucion.", vbExclamation, App.Title
        listaDistribucion.SetFocus
        validar = False
        Exit Function
    End If
    If listaAreas.ListItems.Count = 0 Then
        MsgBox "Debe añadir al menos un Área.", vbExclamation, App.Title
        listaAreas.SetFocus
        validar = False
        Exit Function
    End If
End Function
