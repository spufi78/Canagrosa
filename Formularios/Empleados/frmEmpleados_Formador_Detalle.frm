VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmEmpleados_Formador_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formación de Empleados"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   Icon            =   "frmEmpleados_Formador_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkfechas 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fechas de obtención / autorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   52
      Top             =   7785
      Width           =   3480
   End
   Begin VB.CheckBox chkEnHistorico 
      BackColor       =   &H00C0C0C0&
      Caption         =   "En histórico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   51
      Top             =   8550
      Width           =   2310
   End
   Begin VB.Frame frmRecualificacion 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8280
      TabIndex        =   48
      Top             =   7830
      Width           =   2940
      Begin VB.CheckBox chkRecualificacion 
         BackColor       =   &H008080FF&
         Caption         =   "Ultima Recualificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   2310
      End
      Begin MSComCtl2.DTPicker fechacualificacion 
         Height          =   360
         Left            =   1170
         TabIndex        =   49
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   53608449
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha "
         Height          =   195
         Index           =   13
         Left            =   585
         TabIndex        =   50
         Top             =   405
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   1095
      Left            =   8820
      Picture         =   "frmEmpleados_Formador_Detalle.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   9045
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evidencias y adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   3
      Left            =   7560
      TabIndex        =   44
      Top             =   4365
      Width           =   3705
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   2700
         Picture         =   "frmEmpleados_Formador_Detalle.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2835
         Width           =   465
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   3
         Left            =   3210
         Picture         =   "frmEmpleados_Formador_Detalle.frx":79E6
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Eliminar accesorio"
         Top             =   2835
         Width           =   420
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2550
         Index           =   3
         Left            =   135
         TabIndex        =   46
         Top             =   225
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   4498
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
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entrada manual de ensayos"
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
      Height          =   1185
      Left            =   45
      TabIndex        =   31
      Top             =   8955
      Width           =   7575
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2745
         TabIndex        =   36
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   900
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin pryCombo.miCombo cmbTipoMuestra 
         Height          =   330
         Left            =   900
         TabIndex        =   32
         Top             =   315
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Index           =   11
         Left            =   2160
         TabIndex        =   37
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   35
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame frmFechas 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   45
      TabIndex        =   24
      Top             =   7830
      Width           =   8205
      Begin MSComCtl2.DTPicker fechaTecnico 
         Height          =   360
         Left            =   1305
         TabIndex        =   25
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   53608449
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaFormador 
         Height          =   360
         Left            =   4095
         TabIndex        =   27
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   53608449
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaDirector 
         Height          =   360
         Left            =   6750
         TabIndex        =   29
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   53608449
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Director"
         Height          =   195
         Index           =   8
         Left            =   5580
         TabIndex        =   30
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Formador"
         Height          =   195
         Index           =   4
         Left            =   2790
         TabIndex        =   28
         Top             =   450
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Técnico"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   26
         Top             =   450
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Formación Teórica"
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
      Left            =   45
      TabIndex        =   19
      Top             =   3015
      Width           =   11220
      Begin VB.TextBox txtobservacion 
         Height          =   960
         Left            =   4365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Text            =   "frmEmpleados_Formador_Detalle.frx":7B7A
         Top             =   225
         Width           =   6720
      End
      Begin MSComCtl2.DTPicker fechaFormacion 
         Height          =   360
         Left            =   1125
         TabIndex        =   20
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   53608449
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   0
         Left            =   3195
         TabIndex        =   22
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   21
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ensayos Duplicados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   2
      Left            =   5040
      TabIndex        =   17
      Top             =   4365
      Width           =   2490
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   2
         Left            =   1950
         Picture         =   "frmEmpleados_Formador_Detalle.frx":7BB2
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Eliminar accesorio"
         Top             =   2835
         Width           =   420
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   2
         Left            =   1485
         Picture         =   "frmEmpleados_Formador_Detalle.frx":7D46
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Añadir accesorio"
         Top             =   2835
         Width           =   420
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2550
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   225
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   4498
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ensayos bajo Observación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   1
      Left            =   2475
      TabIndex        =   15
      Top             =   4365
      Width           =   2490
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   1
         Left            =   1950
         Picture         =   "frmEmpleados_Formador_Detalle.frx":E598
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Eliminar accesorio"
         Top             =   2835
         Width           =   420
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   1
         Left            =   1485
         Picture         =   "frmEmpleados_Formador_Detalle.frx":E72C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Añadir accesorio"
         Top             =   2835
         Width           =   420
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2565
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   225
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   4524
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ensayos Observados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Index           =   0
      Left            =   45
      TabIndex        =   13
      Top             =   4365
      Width           =   2400
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   0
         Left            =   1905
         Picture         =   "frmEmpleados_Formador_Detalle.frx":14F7E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Eliminar accesorio"
         Top             =   2880
         Width           =   420
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   0
         Left            =   1440
         Picture         =   "frmEmpleados_Formador_Detalle.frx":15112
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Añadir accesorio"
         Top             =   2880
         Width           =   420
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2610
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   225
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   4604
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de la cualificación"
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
      Left            =   45
      TabIndex        =   3
      Top             =   765
      Width           =   11265
      Begin VB.OptionButton opModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externa"
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   6
         Top             =   1260
         Width           =   1410
      End
      Begin VB.OptionButton opModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interna"
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   1260
         Value           =   -1  'True
         Width           =   960
      End
      Begin pryCombo.miCombo cmbPNT 
         Height          =   330
         Left            =   1080
         TabIndex        =   7
         Top             =   405
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFormador 
         Height          =   330
         Left            =   1080
         TabIndex        =   9
         Top             =   1665
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTecnico 
         Height          =   330
         Left            =   1080
         TabIndex        =   11
         Top             =   810
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Técnico"
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   12
         Top             =   855
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formador"
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   10
         Top             =   1755
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.N.T."
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   8
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   4
         Top             =   1305
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   1095
      Left            =   10080
      Picture         =   "frmEmpleados_Formador_Detalle.frx":1B964
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9045
      Width           =   1185
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del tipo de análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   1830
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10710
      Picture         =   "frmEmpleados_Formador_Detalle.frx":1C22E
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cualificaciones del Empleado"
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
      TabIndex        =   0
      Top             =   45
      Width           =   3120
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmEmpleados_Formador_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EMPLEADO_ID As Long
Public ID_CUALIFICACION As Long

Private Sub chkfechas_Click()
    If chkfechas.value = Checked Then
        frmFechas.Enabled = True
        fechaTecnico.Enabled = True
        fechaFormador.Enabled = True
        fechaDirector.Enabled = True
    Else
        frmFechas.Enabled = False
        fechaTecnico.Enabled = False
        fechaFormador.Enabled = False
        fechaDirector.Enabled = False
    End If
End Sub

Private Sub chkRecualificacion_Click()
    If chkRecualificacion.value = Checked Then
        frmRecualificacion.Enabled = True
        fechacualificacion.Enabled = True
    Else
        frmRecualificacion.Enabled = False
        fechacualificacion.Enabled = False
    End If
End Sub

Private Sub cmdAnadir_Click(Index As Integer)
   On Error GoTo cmdAnadir_Click_Error

    Select Case Index
    Case 0, 1, 2
        If cmbTipoMuestra.getTEXTO = "" Then
            MsgBox "Seleccione el tipo de muestra.", vbCritical, App.Title
            cmbTipoMuestra.SetFocus
            Exit Sub
        End If
        If txtnumero = "" Then
            MsgBox "Indique el numero particular de la muestra.", vbCritical, App.Title
            txtnumero.SetFocus
            Exit Sub
        Else
            If Not IsNumeric(txtnumero) Then
                MsgBox "El numero de muestra debe ser numerico.", vbCritical, App.Title
                txtnumero.SetFocus
                Exit Sub
            End If
        End If
        If txtanno = "" Then
            MsgBox "Seleccione el anno de registro.", vbCritical, App.Title
            txtanno.SetFocus
            Exit Sub
        Else
            If Not IsNumeric(txtanno) Then
                MsgBox "El año debe ser numerico.", vbCritical, App.Title
                txtanno.SetFocus
                Exit Sub
            End If
        End If
        Dim rs As ADODB.RecordSet
        Dim c As String
        c = "SELECT ID_MUESTRA,CONCAT(TM.codigo,'-',CAST(ID_PARTICULAR AS CHAR)) AS CODIGO, ID_GENERAL " & _
            " FROM MUESTRAS M, tipos_muestra TM " & _
            " WHERE M.TIPO_MUESTRA_ID = " & cmbTipoMuestra.getPK_SALIDA & _
            "   AND M.ID_PARTICULAR = " & txtnumero & _
            "   AND M.ANNO = " & txtanno & _
            "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA "
        Set rs = datos_bd(c)
        If rs.RecordCount = 0 Then
            MsgBox "No existe la muestra indicada.", vbCritical, App.Title
        Else
            With lista(Index).ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
            End With
            lista(Index).ListItems(lista(Index).ListItems.Count).EnsureVisible
            txtnumero = ""
            txtnumero.SetFocus
        End If
    End Select

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmEmpleados_Formador_Detalle"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click(Index As Integer)
    If lista(Index).ListItems.Count > 0 Then
        lista(Index).ListItems.Remove lista(Index).SelectedItem.Index
    End If
End Sub

Private Sub cmdEscaner_Click()
   On Error GoTo cmdEscaner_Click_Error

    If ID_CUALIFICACION = 0 Then
        Dim c As String
        
        c = "Para añadir evidencias, es necesario primero añadir la cualificación."
        c = c & vbNewLine & " Pulse aceptar, para introducir la evidencia en el sistema y "
        c = c & vbNewLine & " posteriormente añada las evidencias que desee. "
        
        MsgBox c, vbInformation, App.Title
        Exit Sub
    End If
        
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            If Dir(documento_escaner) = "" Then
                MsgBox "El documento que intenta vincular no existe en la ruta.", vbExclamation, App.Title
                Exit Sub
            End If
            On Error Resume Next
            Dim RUTA As String
            RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "ca_evidencias")
            MkDir RUTA & "\" & CStr(ID_CUALIFICACION)
            FileCopy documento_escaner, RUTA & "\" & CStr(ID_CUALIFICACION) & "\" & nombreNuevo & ".pdf"
            With lista(3).ListItems.Add(, , lista(3).ListItems.Count + 1)
                .SubItems(1) = nombreNuevo
                .SubItems(2) = nombreNuevo & ".pdf"
            End With
            MsgBox "Evidencia vinculada correctamente.", vbInformation, App.Title
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmEmpleados_Formador_Detalle"
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar Then
        Dim oEC As New clsEmpleados_cualificaciones
        Dim Insertar As Boolean
        Insertar = False
        With oEC
            .setEMPLEADO_ID = EMPLEADO_ID
            .setDOCUMENTO_ID = cmbPNT.getPK_SALIDA
            .setEMPLEADO_ID_FORMADOR = cmbFormador.getPK_SALIDA
            .setFECHA_FORMACION_TEORICA = Format(fechaFormacion, "yyyy-mm-dd")
            .setTEXTO_FORMACION_TEORICA = txtobservacion
            If chkfechas.value = Checked Then
                .setFECHA_FIRMA_TECNICO = Format(fechaTecnico, "yyyy-mm-dd")
                .setFECHA_FIRMA_FORMADOR = Format(fechaFormador, "yyyy-mm-dd")
                .setFECHA_FIRMA_DIRECTOR = Format(fechaDirector, "yyyy-mm-dd")
            Else
                .setFECHA_FIRMA_TECNICO = "1900-01-01"
                .setFECHA_FIRMA_FORMADOR = "1900-01-01"
                .setFECHA_FIRMA_DIRECTOR = "1900-01-01"
            End If
            If chkRecualificacion.value = Checked Then
                .setFECHA_ULTIMA_RECUALIFICACION = Format(fechacualificacion, "yyyy-mm-dd")
            Else
                .setFECHA_ULTIMA_RECUALIFICACION = "1900-01-01"
            End If
            If opModalidad(0).value = True Then
                .setMODALIDAD = 0
            Else
                .setMODALIDAD = 1
            End If
            .setEN_HISTORICO = chkEnHistorico.value
            If ID_CUALIFICACION = 0 Then
                ID_CUALIFICACION = .Insertar
                Insertar = True
            Else
                .Modificar ID_CUALIFICACION
            End If
        End With
        ' Muestras
        Dim I As Integer
        Dim j As Integer
        Dim oECM As New clsEmpleados_cualificaciones_m
        oECM.Eliminar ID_CUALIFICACION
        For I = 0 To 2
            If lista(I).ListItems.Count > 0 Then
                For j = 1 To lista(I).ListItems.Count
                    With oECM
                        .setCUALIFICACION_ID = ID_CUALIFICACION
                        .setMUESTRA_ID = CLng(lista(I).ListItems(j).Text)
                        .setTIPO = I
                        .Insertar
                    End With
                Next
            End If
        Next
        ' Evidencias
        Dim OEVE As New clsEmpleados_cualificaciones_e
        OEVE.Eliminar ID_CUALIFICACION
        For I = 1 To lista(3).ListItems.Count
            With OEVE
                .setCUALIFICACION_ID = ID_CUALIFICACION
                .setDESCRIPCION = lista(3).ListItems(I).SubItems(1)
                .setRUTA = lista(3).ListItems(I).SubItems(2)
                .setORDEN = I
                .Insertar
            End With
        Next
        Set oECM = Nothing
        Set oEC = Nothing
        If Insertar Then
            If MsgBox("Cualificación almacenada correctamente. ¿Desea añadir evidencias?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Unload Me
            End If
        Else
            MsgBox "Cualificación almacenada correctamente.", vbInformation, App.Title
            Unload Me
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmEmpleados_Formador_Detalle"
End Sub

Private Sub fechaFormacion_Change()
    txtanno = Year(fechaFormacion)
End Sub

Private Sub Form_Load()
    log Me.Name
'    cargar_botones Me
    cargar_combos
    cabecera
    cmbTecnico.MostrarElemento EMPLEADO_ID
    cmbTecnico.desactivar
    fechaFormacion = Date
    fechaTecnico = Date
    fechaFormador = Date
    fechaDirector = Date
    fechacualificacion = Date
    If ID_CUALIFICACION <> 0 Then
        cargar_cualificacion
    End If
End Sub

Private Sub cargar_combos()
    llenar_combo cmbTecnico, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    llenar_combo cmbFormador, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, ""
    llenar_combo cmbTipoMuestra, New clsTipos_muestra, 0, frmTM_Detalle, ""
End Sub

Private Sub cabecera()
    Dim I As Integer
    For I = 0 To 2
        With lista(I).ColumnHeaders
            .Add , , "ID_MUESTRA", 1, lvwColumnLeft
            .Add , , "Num.Particular", 1000, lvwColumnCenter
            .Add , , "Num.General", 1000, lvwColumnCenter
        End With
    Next
    With lista(3).ColumnHeaders
         .Add , , "ORDEN", 1, lvwColumnLeft
         .Add , , "Descripción", 3200, lvwColumnLeft
         .Add , , "Ruta", 1, lvwColumnLeft
    End With
End Sub

Private Function validar() As Boolean
    validar = False
    If cmbPNT.getTEXTO = "" Then
        MsgBox "Debe indicar el PNT.", vbCritical, App.Title
        cmbPNT.SetFocus
        Exit Function
    End If
    If cmbFormador.getTEXTO = "" Then
        MsgBox "Debe indicar el formador.", vbCritical, App.Title
        cmbFormador.SetFocus
        Exit Function
    End If
    validar = True
End Function

Private Sub lista_DblClick(Index As Integer)
    If lista(Index).ListItems.Count = 0 Then
        Exit Sub
    End If
    Select Case Index
    Case 0, 1, 2
        gmuestra = CLng(lista(Index).ListItems(lista(Index).SelectedItem.Index).Text)
        frmVerMuestra.Show 1
    Case 3
        On Error GoTo fallo
        Dim RUTA As String
        RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "ca_evidencias")
        RUTA = RUTA & "\" & CStr(ID_CUALIFICACION)
        RUTA = RUTA & "\" & lista(Index).ListItems(lista(Index).SelectedItem.Index).SubItems(2)
        If RUTA <> "" Then
            If Dir(RUTA) <> "" Then
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & RUTA, vbMaximizedFocus)
            Else
                MsgBox "El documento vinculado no existe.", vbCritical, App.Title
            End If
        End If
    End Select
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub cargar_cualificacion()
    Dim oEC As New clsEmpleados_cualificaciones
    With oEC
        oEC.Carga ID_CUALIFICACION
        cmbPNT.MostrarElemento oEC.getDOCUMENTO_ID
        opModalidad(oEC.getMODALIDAD).value = True
        cmbFormador.MostrarElemento oEC.getEMPLEADO_ID_FORMADOR
        fechaFormacion = oEC.getFECHA_FORMACION_TEORICA
        txtanno = Year(oEC.getFECHA_FORMACION_TEORICA)
        txtobservacion = oEC.getTEXTO_FORMACION_TEORICA
        If Format(oEC.getFECHA_FIRMA_TECNICO, "yyyy-mm-dd") <> "1900-01-01" Then
            chkfechas.value = Checked
            fechaTecnico.Enabled = True
            fechaFormador.Enabled = True
            fechaDirector.Enabled = True
            fechaTecnico = oEC.getFECHA_FIRMA_TECNICO
            fechaFormador = oEC.getFECHA_FIRMA_FORMADOR
            fechaDirector = oEC.getFECHA_FIRMA_DIRECTOR
        Else
            chkfechas.value = Unchecked
            fechaTecnico.Enabled = False
            fechaFormador.Enabled = False
            fechaDirector.Enabled = False
        End If
        If Format(oEC.getFECHA_ULTIMA_RECUALIFICACION, "yyyy-mm-dd") <> "1900-01-01" Then
            chkRecualificacion.value = Checked
            fechacualificacion.Enabled = True
            fechacualificacion = oEC.getFECHA_ULTIMA_RECUALIFICACION
        Else
            chkRecualificacion.value = Unchecked
            fechacualificacion.Enabled = False
        End If
        chkEnHistorico.value = oEC.getEN_HISTORICO
    End With
    Set oEC = Nothing
    ' Muestras
    Dim rs As ADODB.RecordSet
    Dim oECM As New clsEmpleados_cualificaciones_m
    Set rs = oECM.Listado(ID_CUALIFICACION)
    If rs.RecordCount > 0 Then
        Do
            With lista(rs(0)).ListItems.Add(, , rs(1))
                .SubItems(1) = rs(2)
                .SubItems(2) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oECM = Nothing
    ' Evidencias
    Dim oECE As New clsEmpleados_cualificaciones_e
    Set rs = oECE.Listado(ID_CUALIFICACION)
    If rs.RecordCount > 0 Then
        Do
            With lista(3).ListItems.Add(, , rs("ORDEN"))
                .SubItems(1) = rs("DESCRIPCION")
                .SubItems(2) = rs("RUTA")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Set oECE = Nothing
    Set rs = Nothing
End Sub

Private Sub opModalidad_Click(Index As Integer)
    cmbFormador.Limpiar
    If Index = 0 Then
        llenar_combo cmbFormador, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    Else
        llenar_combo cmbFormador, New clsEmpleados_fexternos, 0, frmEmpleados_fexternos, ""
    End If
End Sub
