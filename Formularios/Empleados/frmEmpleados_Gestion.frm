VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmpleados_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FICHA DE PERSONAL"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14700
   Icon            =   "frmEmpleados_Gestion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   14700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Formación"
      Height          =   915
      Left            =   5310
      Picture         =   "frmEmpleados_Gestion.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   7740
      Width           =   1275
   End
   Begin VB.CommandButton cmdAnticipo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anticipos"
      Height          =   915
      Left            =   7020
      Picture         =   "frmEmpleados_Gestion.frx":12B4
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   8820
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame frmPrivado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   45
      TabIndex        =   67
      Top             =   8685
      Width           =   3975
      Begin VB.CommandButton cmdExpediente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contratos"
         Height          =   915
         Left            =   45
         Picture         =   "frmEmpleados_Gestion.frx":15BE
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdControl 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ausencias"
         Height          =   915
         Left            =   1350
         Picture         =   "frmEmpleados_Gestion.frx":1E88
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdNominas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nominas"
         Height          =   915
         Left            =   2655
         Picture         =   "frmEmpleados_Gestion.frx":2752
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Height          =   510
      Left            =   45
      TabIndex        =   60
      Top             =   4905
      Width           =   8250
      Begin VB.CheckBox chkExternal 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EXTERNO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   5715
         TabIndex        =   61
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo cmbempresas 
         Height          =   315
         Left            =   1800
         TabIndex        =   63
         Top             =   135
         Width           =   2790
         _ExtentX        =   4921
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empresa: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   675
         TabIndex        =   62
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   915
      Left            =   4005
      Picture         =   "frmEmpleados_Gestion.frx":301C
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   7740
      Width           =   1275
   End
   Begin VB.CommandButton cmdFormador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Formador"
      Height          =   915
      Left            =   2700
      Picture         =   "frmEmpleados_Gestion.frx":38E6
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   7740
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cualificaciones"
      Height          =   915
      Left            =   1395
      Picture         =   "frmEmpleados_Gestion.frx":41B0
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7740
      Width           =   1275
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PNTS en los que se puede formar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   10800
      TabIndex        =   54
      Top             =   2655
      Width           =   3840
      Begin MSComctlLib.ListView listapnt 
         Height          =   4665
         Left            =   90
         TabIndex        =   55
         Top             =   270
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Departamentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   8325
      TabIndex        =   43
      Top             =   2655
      Width           =   2400
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gestión Metrológica"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   135
         TabIndex        =   66
         Top             =   4680
         Width           =   2205
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compras"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   135
         TabIndex        =   65
         Top             =   4320
         Width           =   2205
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "RR.HH."
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   135
         TabIndex        =   64
         Top             =   3960
         Width           =   2205
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "I + D"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   135
         TabIndex        =   53
         Top             =   2880
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Logística"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   52
         Top             =   2475
         Width           =   2175
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Metrología"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   135
         TabIndex        =   51
         Top             =   2115
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Laborat. Aeronáutico"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   135
         TabIndex        =   50
         Top             =   1755
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Laborat. Agroalimentario"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   135
         TabIndex        =   49
         Top             =   1350
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Administación"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   48
         Top             =   1035
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calidad"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   47
         Top             =   630
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gerencia"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   46
         Top             =   270
         Width           =   2175
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Informática"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   135
         TabIndex        =   45
         Top             =   3240
         Width           =   2220
      End
      Begin VB.CheckBox chkDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepción"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   135
         TabIndex        =   44
         Top             =   3600
         Width           =   2205
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F.P.T."
      Height          =   915
      Left            =   90
      Picture         =   "frmEmpleados_Gestion.frx":4A7A
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7740
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   13350
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7785
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   12030
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7785
      Width           =   1275
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado"
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
      Left            =   8325
      TabIndex        =   27
      Top             =   675
      Width           =   4200
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   90
         TabIndex        =   28
         Top             =   315
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSComCtl2.DTPicker fechaIncorporacion 
         Height          =   360
         Left            =   2340
         TabIndex        =   36
         Top             =   855
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaBaja 
         Height          =   360
         Left            =   2340
         TabIndex        =   38
         Top             =   1350
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         CalendarForeColor=   255
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Baja"
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
         Index           =   12
         Left            =   135
         TabIndex        =   39
         Top             =   1395
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Incorporación"
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
         Left            =   135
         TabIndex        =   37
         Top             =   900
         Width           =   1845
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8145
      Top             =   7830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   1950
      Left            =   12555
      TabIndex        =   24
      Top             =   675
      Width           =   2100
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Height          =   375
         Index           =   11
         Left            =   45
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1485
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1710
         Left            =   315
         Picture         =   "frmEmpleados_Gestion.frx":5344
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apodo"
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
         TabIndex        =   26
         Top             =   1215
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Index           =   13
      Left            =   45
      TabIndex        =   23
      Top             =   5490
      Width           =   8250
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1815
         Index           =   9
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   7995
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Empleado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   15
      TabIndex        =   14
      Top             =   675
      Width           =   8280
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   13
         Left            =   1350
         MaxLength       =   25
         TabIndex        =   8
         Top             =   3285
         Width           =   6810
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   12
         Left            =   5805
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2385
         Width           =   2355
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   10
         Left            =   3915
         TabIndex        =   9
         Top             =   2835
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   345
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2835
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   5805
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1980
         Width           =   2340
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1980
         Width           =   2715
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   8
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2370
         Width           =   2715
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1140
         Width           =   960
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   1
         Top             =   735
         Width           =   6825
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   0
         Top             =   330
         Width           =   6810
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   3420
         TabIndex        =   33
         Top             =   1140
         Width           =   4725
         _ExtentX        =   8334
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
      Begin MSDataListLib.DataCombo cmbMunicipio 
         Height          =   315
         Left            =   1350
         TabIndex        =   34
         Top             =   1575
         Width           =   6795
         _ExtentX        =   11986
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
      Begin MSDataListLib.DataCombo cmbUsuario 
         Bindings        =   "frmEmpleados_Gestion.frx":90A6
         Height          =   360
         Left            =   1665
         TabIndex        =   11
         Top             =   3735
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSComCtl2.DTPicker fnacimiento 
         Height          =   405
         Left            =   1350
         TabIndex        =   7
         Top             =   2790
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Nacimiento"
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
         Index           =   13
         Left            =   135
         TabIndex        =   59
         Top             =   2880
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario Geslab"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   35
         Top             =   3780
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.C.C."
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
         Left            =   135
         TabIndex        =   30
         Top             =   3345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Interno"
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
         Index           =   8
         Left            =   4410
         TabIndex        =   29
         Top             =   2430
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imagen"
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
         Left            =   3105
         TabIndex        =   25
         Top             =   2880
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Movil"
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
         Left            =   5070
         TabIndex        =   21
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   19
         Top             =   2460
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
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
         Left            =   2460
         TabIndex        =   18
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
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
         Left            =   120
         TabIndex        =   17
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
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
         Left            =   135
         TabIndex        =   16
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   375
         Width           =   735
      End
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de ficha de Personal"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   42
      Top             =   315
      Width           =   1995
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14130
      Picture         =   "frmEmpleados_Gestion.frx":90EC
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FICHA DE PERSONAL"
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
      TabIndex        =   41
      Top             =   45
      Width           =   2325
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   14670
   End
End
Attribute VB_Name = "frmEmpleados_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmbestados_Change()
    If cmbestados.BoundText <> 0 Then
        fechabaja.Enabled = True
    Else
        fechabaja.Enabled = False
    End If
End Sub
Private Sub cmbProvincia_Change()
    If cmbProvincia.Text <> "" Then
        cargar_municipios (cmbProvincia.BoundText)
    End If
End Sub

Private Sub cmbUsuario_Change()
    If cmbUsuario.Text <> "" Then
        Dim oUsuario As New clsUsuarios
        oUsuario.CARGAR cmbUsuario.BoundText
        txtDatos(10) = oUsuario.getIMAGEN
        Set oUsuario = Nothing
    End If
End Sub

Private Sub cmdAnticipo_Click()
    frmEmpleados_Anticipo.Show 1
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdControl_Click()
    If PK <> 0 Then
        frmEmpleados_Control.PK = PK
        frmEmpleados_Control.Show 1
    End If
End Sub
Private Sub cmdExpediente_Click()
    If PK <> 0 Then
        frmEmpleados_Expediente.PK = PK
        frmEmpleados_Expediente.Show 1
    End If
End Sub
Private Sub cmdEXplorar_Click()
    cd.DialogTitle = "Abrir fichero de imagen"
    cd.InitDir = App.Path & "\recursos\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtDatos(10).Text = cd.FileName  ' cd.FileTitle
    End If
End Sub

Private Sub cmdFicha_Click()
    frmReport.iniciar
    frmReport.informe = "\Empleados\rptEmpleados_Ficha"
    frmReport.criterio = "{empleados.ID_EMPLEADO} =" & PK
    frmReport.imprimir = False
    frmReport.generar
    frmReport.Show 1
    Unload frmReport
End Sub

Private Sub cmdFormador_Click()
    frmEmpleados_Formador.PK = PK
    frmEmpleados_Formador.Show 1
End Sub

Private Sub cmdNominas_Click()
    frmEmpleados_Nominas.Show 1
End Sub

Private Sub cmdok_Click()
    If PK > 0 Then
        Modificar
    Else
        Insertar
    End If
End Sub

Private Sub Command1_Click()
    If PK <> 0 Then
        frmEmpleados_Categorias_Historia.PK = PK
        frmEmpleados_Categorias_Historia.Show 1
    End If
    
End Sub

Private Sub Command2_Click()
    If PK <> 0 Then
        frmEmpleados_Formacion.PK = PK
        frmEmpleados_Formacion.Show 1
    End If
End Sub

Private Sub Command3_Click()
    If PK <> 0 Then
        frmEmpleados_Cualificaciones.PK = PK
        frmEmpleados_Cualificaciones.Show 1
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_estados
    cargar_empresas
    cargar_lista_pnt
    cargar_combo cmbUsuario, New clsUsuarios
    cargar_combo cmbProvincia, New clsProvincias
    permisos
    If PK > 0 Then
        consulta
    Else
        fechaIncorporacion = Date
        fnacimiento = Date
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Gestion = Nothing
End Sub

Private Sub txtDatos_Change(Index As Integer)
    On Error Resume Next
    Set img.Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "no"))
    If Index = 10 And txtDatos(10) <> "" Then
        If Dir(txtDatos(10)) <> "" Then
            Set img.Picture = LoadPicture(txtDatos(10))
        End If
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 9 Then
        txtDatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtDatos(9).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 9 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Private Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 13
        ' LP005
        If i <> 4 And i <> 5 Then
            txtDatos(i) = ""
        End If
    Next
    txtDatos(1).SetFocus
End Sub

Private Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 13
        ' LP005
        If i <> 4 And i <> 5 Then
        txtDatos(i).Locked = True
        End If
    Next
End Sub

Private Sub Insertar()
   On Error GoTo Insertar_Error

    If valida_datos = False Then
        Exit Sub
    End If
    PREGUNTA = "Va a dar de alta el empleado. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set oempleado = mover_datos
        If oempleado.Insertar > 0 Then
            MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set oempleado = Nothing
    End If

   On Error GoTo 0
   Exit Sub

Insertar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Insertar of Formulario frmEmpleados_Gestion"
End Sub

Private Sub Modificar()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim Pos As Integer
    Dim operario As Integer
    PREGUNTA = "Va a modificar los datos del empleado. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oempleado = mover_datos
        If oempleado.Modificar(PK) = True Then
            MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set oempleado = Nothing
    End If
End Sub
Private Function valida_datos() As Boolean
    valida_datos = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre del empleado no puede estar en blanco.", vbCritical, "Error"
        txtDatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
'    If txtdatos(11) = "" Then
'        MsgBox "El apodo del empleado no puede estar en blanco.", vbCritical, "Error"
'        txtdatos(11).SetFocus
'        valida_datos = False
'        Exit Function
'    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe seleccionar un estado.", vbCritical, "Error"
        cmbestados.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbempresas.BoundText = "" Then
        MsgBox "Debe seleccionar una Empresa.", vbCritical, "Error"
        cmbempresas.SetFocus
        valida_datos = False
        Exit Function
    End If
'    If cmbUsuario.BoundText = "" Then
'        MsgBox "Debe seleccionar un usuario asignado a GESLAB.", vbCritical, "Error"
'        cmbUsuario.SetFocus
'        valida_datos = False
'        Exit Function
'    End If
    
End Function

Private Sub consulta()
   On Error GoTo consulta_Error

    On Error GoTo fallo
    Dim oempleado As New clsEmpleados
    oempleado.CARGAR (PK)
    With oempleado
        txtDatos(1) = .getNOMBRE
        txtDatos(2) = .getDIRECCION
        txtDatos(3) = .getCP
        cmbProvincia.BoundText = .getPROVINCIA_ID
        cargar_municipios .getPROVINCIA_ID
        cmbMunicipio.BoundText = .getMUNICIPIO_ID
        txtDatos(6) = .getTELEFONO
        txtDatos(7) = .getMOVIL
        txtDatos(8) = .getCIF
        txtDatos(9) = .getOBSERVACIONES
        txtDatos(10) = .getfoto
        txtDatos(11) = .getAPODO
        txtDatos(12) = .getCODIGO_INTERNO
        txtDatos(13) = .getCCC
        cmbestados.BoundText = .getESTADO_ID
        If .getUSUARIO_ID <> 0 Then
            cmbUsuario.BoundText = .getUSUARIO_ID
        End If
        
        'MANTIS-XXX-I
        cmbempresas.BoundText = .getEMPRESA_ID
        'MANTIS-XXX-F
        fnacimiento = .getFECHA_NACIMIENTO
        fechaIncorporacion = .getFECHA_INCORPORACION
        fechabaja = .getFECHA_BAJA
        'M0521-I
        chkExternal.Value = .getEXTERNAL
        'M0521-F
        Dim i As Integer
        Dim j As Integer
        ' Departamentos
        Dim d() As String
        If .getDEPARTAMENTOS <> "" Then
            d = Split(.getDEPARTAMENTOS, ";")
            For i = LBound(d) To UBound(d)
                If IsNumeric(d(i)) Then
                    chkDepartamento(d(i)).Value = Checked
                End If
            Next
        End If
        ' PNTS
        If .getPNTS <> "" Then
            d = Split(.getPNTS, ";")
            For i = LBound(d) To UBound(d)
                If IsNumeric(d(i)) Then
                    For j = 1 To listapnt.ListItems.Count
                        If CInt(listapnt.ListItems(j).Text) = CInt(d(i)) Then
                            listapnt.ListItems(j).Checked = True
                        End If
                    Next
                End If
            Next
        End If
    End With
    Set oempleado = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del empleado.", vbCritical, Err.Description

   On Error GoTo 0
   Exit Sub

consulta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure consulta of Formulario frmEmpleados_Gestion"
End Sub
Private Function mover_datos() As clsEmpleados
   On Error GoTo mover_datos_Error

    On Error GoTo fallo
    Dim oempleado As New clsEmpleados
    With oempleado
        .setNOMBRE = txtDatos(1)
        .setDIRECCION = txtDatos(2)
        If Trim(txtDatos(3)) <> "" Then
            .setCP = CLng(txtDatos(3).Text)
        Else
            .setCP = 0
        End If
        If cmbProvincia.Text = "" Then
            .setPROVINCIA_ID = 0
        Else
            .setPROVINCIA_ID = cmbProvincia.BoundText
        End If
        If cmbMunicipio.Text = "" Then
            .setMUNICIPIO_ID = 0
        Else
            .setMUNICIPIO_ID = cmbMunicipio.BoundText
        End If
        If cmbUsuario.BoundText <> "" Then
            .setUSUARIO_ID = cmbUsuario.BoundText
        End If
        .setCIF = txtDatos(8)
        .setTELEFONO = txtDatos(6)
        .setMOVIL = txtDatos(7)
        .setOBSERVACIONES = txtDatos(9)
        .setfoto = txtDatos(10)
        .setAPODO = txtDatos(11)
        .setESTADO_ID = cmbestados.BoundText
        .setCODIGO_INTERNO = txtDatos(12)
        .setCCC = txtDatos(13)
        .setFECHA_NACIMIENTO = Format(fnacimiento, "yyyy-mm-dd")
        .setFECHA_INCORPORACION = Format(fechaIncorporacion, "yyyy-mm-dd")
        .setFECHA_BAJA = Format(fechabaja, "yyyy-mm-dd")
        'M0521-I
        .setEXTERNAL = chkExternal.Value
        'M0521-F
        'MANTIS-XXX-I
        .setEMPRESA_ID = cmbempresas.BoundText
        'MANTIS-XXX-F
        Dim i As Integer
        Dim d As String
        ' Departamentos
        d = ""
        'M1377-I
        'For i = 1 To 10
        For i = 1 To 13
        'M1377-F
            If chkDepartamento(i).Value = Checked Then
                d = d & i & ";"
            End If
        Next
        .setDEPARTAMENTOS = d
        ' Pnts
        d = ""
        For i = 1 To listapnt.ListItems.Count
            If listapnt.ListItems(i).Checked = True Then
                d = d & listapnt.ListItems(i).Text & ";"
            End If
        Next
        .setPNTS = d
    End With
    Set mover_datos = oempleado
    Set oempleado = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del empleado.", vbCritical, Err.Description

   On Error GoTo 0
   Exit Function

mover_datos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mover_datos of Formulario frmEmpleados_Gestion"
End Function
Private Sub cargar_estados()
    Dim ooe As New clsEmpleados_Estados
    Set cmbestados.RowSource = ooe.Listado
    cmbestados.ListField = "nombre"
    cmbestados.BoundColumn = "id_estado"
    Set ooe = Nothing
End Sub
'MANTIS-XXX-I
Private Sub cargar_empresas()
    Dim ooe As New clsEmpleados_Empresas
    Set cmbempresas.RowSource = ooe.Listado
    cmbempresas.ListField = "descripcion"
    cmbempresas.BoundColumn = "id_empresa"
    Set ooe = Nothing
End Sub
'MANTIS-XXX-F
'M1382-I
Private Sub cargardepartamentos()
    Dim ndepartamentos As Integer
    Dim i As Long
    Dim oDecodificadora As New clsDecodificadora
    
    ndepartamentos = CInt(chkDepartamento.Count)
    For i = 1 To ndepartamentos
        oDecodificadora.Carga_valor DECODIFICADORA.PROCNC_DEPARTAMENTOS, i
        chkDepartamento(i).Caption = oDecodificadora.getDESCRIPCION
    Next i
    Set oDecodificadora = Nothing
End Sub
'M1382-F
Public Sub permisos()
    frmPrivado.visible = False
    If USUARIO.getPER_EMPLEADOS = 0 Then
        cmdControl.Enabled = False
        cmdExpediente.Enabled = False
        cmdAnticipo.Enabled = False
    End If
    If USUARIO.getUSO = "MARIBEL-PC" Or USUARIO.getUSO = "RRHH-PC" Or USUARIO.getUSO = "DES-JGM" Or USUARIO.getUSO = "JOSEPC" Then
        Me.Height = 10260
        frmPrivado.visible = True
    Else
        Me.Height = 9165
    End If
End Sub
Private Sub cargar_municipios(PROVINCIA As Long)
     If IsNumeric(PROVINCIA) Then
        Dim omuni As New clsMunicipios
        Set cmbMunicipio.RowSource = omuni.Listado(PROVINCIA)
        cmbMunicipio.ListField = "nombre" 'campo que veo
        cmbMunicipio.DataField = "nombre" 'campo asociado
        cmbMunicipio.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
End Sub

Private Sub cargar_lista_pnt()
    Dim RS As ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set RS = oDeco.Listado(DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS)
    If RS.RecordCount > 0 Then
        Do
            With listapnt.ListItems.Add(, , RS("VALOR"))
                .SubItems(1) = RS("DESCRIPCION")
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Set RS = Nothing
    
        
End Sub

Private Sub cabecera()
        With listapnt.ColumnHeaders
            .Add , , "ID", 300, lvwColumnLeft
            .Add , , "Descripcion", 3000, lvwColumnLeft
        End With
End Sub
