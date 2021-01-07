VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipos_Detalle_Mto 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Mantenimiento equipos - Detalle mantenimiento"
   ClientHeight    =   8580
   ClientLeft      =   2340
   ClientTop       =   1140
   ClientWidth     =   11400
   Icon            =   "frmEquipos_Detalle_Mto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Histórico"
      Height          =   870
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7695
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Operaciones generales de mto preventivo"
      Height          =   870
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7695
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle del plan de mantenimiento"
      Height          =   3750
      Left            =   0
      TabIndex        =   5
      Top             =   3870
      Width           =   11355
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos del mantenimiento"
         Height          =   2625
         Left            =   6480
         TabIndex        =   55
         Top             =   270
         Width           =   4740
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Guardar"
            Height          =   330
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2160
            Width           =   960
         End
         Begin VB.CheckBox chkConforme 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conforme"
            Height          =   240
            Left            =   3600
            TabIndex        =   60
            Top             =   0
            Width           =   1005
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   960
            Index           =   4
            Left            =   180
            MaxLength       =   255
            TabIndex        =   58
            Top             =   1170
            Width           =   4380
         End
         Begin pryCombo.miCombo miCombo1 
            Height          =   330
            Left            =   180
            TabIndex        =   56
            Top             =   540
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   59
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Operador"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   57
            Top             =   315
            Width           =   660
         End
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   14
         Top             =   0
         Width           =   1140
      End
      Begin MSComctlLib.ListView listaPlanMto_Detalle 
         Height          =   3330
         Left            =   135
         TabIndex        =   6
         Top             =   360
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   5874
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
      Begin MSComCtl2.DTPicker fechaMantenimiento 
         Height          =   315
         Left            =   2970
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   78643201
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Minutos"
         Height          =   195
         Index           =   0
         Left            =   5535
         TabIndex        =   16
         Top             =   45
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Previsto"
         Height          =   195
         Index           =   1
         Left            =   3465
         TabIndex        =   15
         Top             =   45
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   43
         Left            =   2610
         TabIndex        =   10
         Top             =   45
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planes de mantenimiento asignados"
      Height          =   3300
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   11355
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Planificación"
         Height          =   3030
         Left            =   6480
         TabIndex        =   18
         Top             =   135
         Width           =   4740
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   2
            Left            =   2745
            TabIndex        =   30
            Top             =   225
            Width           =   240
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   1
            Left            =   1305
            TabIndex        =   29
            Top             =   225
            Width           =   240
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   4
            Left            =   1305
            TabIndex        =   28
            Top             =   945
            Width           =   285
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   3
            Left            =   4230
            TabIndex        =   27
            Top             =   225
            Width           =   240
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   6
            Left            =   4230
            TabIndex        =   26
            Top             =   945
            Width           =   285
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   5
            Left            =   2745
            TabIndex        =   25
            Top             =   945
            Width           =   240
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   8
            Left            =   2745
            TabIndex        =   24
            Top             =   1665
            Width           =   285
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   7
            Left            =   1305
            TabIndex        =   23
            Top             =   1665
            Width           =   240
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   10
            Left            =   1305
            TabIndex        =   22
            Top             =   2385
            Width           =   285
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   9
            Left            =   4230
            TabIndex        =   21
            Top             =   1665
            Width           =   240
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   12
            Left            =   4230
            TabIndex        =   20
            Top             =   2385
            Width           =   285
         End
         Begin VB.CheckBox chkop 
            BackColor       =   &H00C0C0C0&
            Height          =   240
            Index           =   11
            Left            =   2745
            TabIndex        =   19
            Top             =   2385
            Width           =   240
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   31
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   2
            Left            =   1755
            TabIndex        =   32
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   3
            Left            =   3240
            TabIndex        =   33
            Top             =   450
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   4
            Left            =   315
            TabIndex        =   34
            Top             =   1170
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   5
            Left            =   1755
            TabIndex        =   35
            Top             =   1170
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   6
            Left            =   3240
            TabIndex        =   36
            Top             =   1170
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   7
            Left            =   315
            TabIndex        =   37
            Top             =   1890
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   8
            Left            =   1755
            TabIndex        =   38
            Top             =   1890
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   9
            Left            =   3240
            TabIndex        =   39
            Top             =   1890
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   10
            Left            =   315
            TabIndex        =   40
            Top             =   2610
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   11
            Left            =   1755
            TabIndex        =   41
            Top             =   2610
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker f1 
            Height          =   345
            Index           =   12
            Left            =   3240
            TabIndex        =   42
            Top             =   2610
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   78643201
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Febrero"
            Height          =   195
            Index           =   2
            Left            =   1755
            TabIndex        =   54
            Top             =   225
            Width           =   540
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Enero"
            Height          =   195
            Index           =   10
            Left            =   315
            TabIndex        =   53
            Top             =   225
            Width           =   420
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abril"
            Height          =   195
            Index           =   12
            Left            =   315
            TabIndex        =   52
            Top             =   945
            Width           =   300
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marzo"
            Height          =   195
            Index           =   13
            Left            =   3240
            TabIndex        =   51
            Top             =   225
            Width           =   435
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Junio"
            Height          =   195
            Index           =   14
            Left            =   3240
            TabIndex        =   50
            Top             =   945
            Width           =   375
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mayo"
            Height          =   195
            Index           =   15
            Left            =   1755
            TabIndex        =   49
            Top             =   945
            Width           =   390
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Agosto"
            Height          =   195
            Index           =   16
            Left            =   1755
            TabIndex        =   48
            Top             =   1665
            Width           =   495
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Julio"
            Height          =   195
            Index           =   17
            Left            =   315
            TabIndex        =   47
            Top             =   1665
            Width           =   315
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Octubre"
            Height          =   195
            Index           =   18
            Left            =   315
            TabIndex        =   46
            Top             =   2385
            Width           =   570
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Septiembre"
            Height          =   195
            Index           =   19
            Left            =   3240
            TabIndex        =   45
            Top             =   1665
            Width           =   795
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Diciembre"
            Height          =   195
            Index           =   20
            Left            =   3240
            TabIndex        =   44
            Top             =   2385
            Width           =   705
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Noviembre"
            Height          =   195
            Index           =   21
            Left            =   1755
            TabIndex        =   43
            Top             =   2385
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdOperador 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Operador"
         Height          =   870
         Left            =   4455
         Picture         =   "frmEquipos_Detalle_Mto.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cambiar operador del plan de mantenimiento"
         Top             =   540
         Width           =   1905
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar plan"
         Height          =   735
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar plan de mantenimiento del equipo"
         Top             =   2475
         Width           =   960
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir plan"
         Height          =   735
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Añadir plan de mantenimiento al equipo"
         Top             =   2475
         Width           =   960
      End
      Begin MSComctlLib.ListView listaPlanes_Mto 
         Height          =   945
         Left            =   90
         TabIndex        =   4
         Top             =   1485
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   1667
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
      Begin pryCombo.miCombo cmbOperador 
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   1080
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbTipoCalibracion 
         Height          =   315
         Left            =   90
         TabIndex        =   63
         Top             =   495
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   64
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Operador"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   13
         Top             =   855
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7695
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7695
      Width           =   1050
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10845
      Picture         =   "frmEquipos_Detalle_Mto.frx":0E54
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimientos asignados al equipo"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3825
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   11520
   End
End
Attribute VB_Name = "frmEquipos_Detalle_Mto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Public FK_PLAN_MTO As Long ' Clave foránea del plan de mantenimiento que se asociará al equipo

Private Sub Form_Load()
    log (Me.Name)
    cabecera
    cargar_botones Me
    llenar_combo cmbOperador, New clsUsuarios, 0, Me, ""
    Dim titulo As String
    If PK <> 0 Then
        cargar_lista
    End If
End Sub

' Bótón que asigna el plan de mto seleccionado al equipo
Private Sub cmdAnadir_Click()
    If cmbOperador.getPK_SALIDA = 0 Or Len(cmbOperador.getTEXTO) <= 1 Then
        MsgBox "Debe especificar el operador que realizará el mantenimiento.", vbCritical, App.Title
        Exit Sub
    End If
    frmEquipos_PlanesMantenimiento_listado.FK_EQUIPO = PK ' Se le pasa la pk como fk al formulario
    frmEquipos_PlanesMantenimiento_listado.Show 1 'Se abre el formulario del listado de planes de mantenimientos
    
    If FK_PLAN_MTO <> 0 Then ' Se ha seleccionado un plan
        Dim i As Long
        For i = 1 To listaPlanes_Mto.ListItems.Count ' Se comprueba si no está ya añadido ese plan
            If listaPlanes_Mto.ListItems(i) = FK_PLAN_MTO Then
                MsgBox "El equipo ya tiene asignado ese plan de mantenimiento.", vbCritical, App.Title
                Exit Sub
            End If
        Next i
        
        Dim oPlan As New clsEquipos_planes_Mantenimiento ' Se inserta en la lista el nuevo plan
        oPlan.Carga (FK_PLAN_MTO)
        With listaPlanes_Mto.ListItems.Add(, , oPlan.getID_PLAN_MTO)
             .SubItems(1) = oPlan.getNOMBRE
             .SubItems(2) = oPlan.getDESCRIPCION
             .SubItems(3) = oPlan.Frecuencia_descrip(oPlan.getFRECUENCIA_ID)
             .SubItems(4) = cmbOperador.getTEXTO
             .SubItems(5) = cmbOperador.getPK_SALIDA
        End With
        cmbOperador.Limpiar
    End If
End Sub
'Botón que desasigna el plan de mto seleccionado al equipo
Private Sub cmdEliminar_Click()
    If listaPlanes_Mto.ListItems.Count = 0 Then
        MsgBox "Debe selecionar un plan de mantenimiento.", vbCritical, App.Title
        Exit Sub
    Else
        If MsgBox("Va a eliminar el plan de mantenimiento del equipo, ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            listaPlanes_Mto.ListItems.Remove listaPlanes_Mto.SelectedItem.Index
            If listaPlanes_Mto.ListItems.Count = 0 Then
                listaPlanMto_Detalle.ListItems.Clear
            Else
                listaPlanes_Mto.ListItems(1).Selected = True
                listaPlanes_Mto_Click
            End If
        End If
    End If
End Sub

' Cambiar el operador asignado al mantenimiento
Private Sub cmdOperador_Click()
    If Len(cmbOperador.getTEXTO) = 0 Then
        MsgBox "Debe selecionar un operador.", vbCritical, App.Title
        Exit Sub
    ElseIf listaPlanes_Mto.ListItems.Count = 0 Then
        MsgBox "Debe selecionar un plan de mantenimiento.", vbCritical, App.Title
        Exit Sub
    Else
        listaPlanes_Mto.ListItems(listaPlanes_Mto.SelectedItem.Index).SubItems(4) = cmbOperador.getTEXTO
        listaPlanes_Mto.ListItems(listaPlanes_Mto.SelectedItem.Index).SubItems(5) = cmbOperador.getPK_SALIDA
        cmbOperador.Limpiar
        MsgBox "El operador se cambió correctamente.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub listaPlanes_Mto_Click()
    If listaPlanes_Mto.ListItems.Count = 0 Then
        Exit Sub
    End If
    listaPlanMto_Detalle.ListItems.Clear
    
    Dim rs As ADOdb.RecordSet
    Dim oPlan_detalle As New clsEquipos_PM_Detalle
    Set rs = oPlan_detalle.Listado(listaPlanes_Mto.ListItems(listaPlanes_Mto.SelectedItem.Index).Text)
    txtDatos(0) = 0
    If rs.RecordCount <> 0 Then
        Do
            With listaPlanMto_Detalle.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(3)
                .SubItems(2) = rs(1)
                .SubItems(3) = rs(2)
                txtDatos(0) = CLng(txtDatos(0)) + CLng(.SubItems(3))
            End With
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPlan_detalle = Nothing
End Sub

Private Sub cmdok_Click()
    On Error GoTo cmdok_Click_Error
    
    Dim oPlanAsig As New clsEquipos_Planes_mto_asignados
    Dim i As Long
    
    oPlanAsig.Eliminar_planes (PK) ' se eliminan todos los planes
    For i = 1 To listaPlanes_Mto.ListItems.Count ' se recorre la lista con los que se van a asignar
        oPlanAsig.setEQUIPO_ID = PK
        oPlanAsig.setPLAN_MTO_ID = listaPlanes_Mto.ListItems(i).Text
        oPlanAsig.setOPERADOR_ID = listaPlanes_Mto.ListItems(i).SubItems(5)
        oPlanAsig.Insertar ' se insertan los nuevos planes
    Next i
    MsgBox "Los planes de mantenimiento se asignaron correctamente al equipo.", vbExclamation, App.Title
    Unload Me
    Exit Sub

cmdok_Click_Error:
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmEquipos_Detalle_Mto")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' Funciones auxiliares del módulo del formulario
' ----------------------------------------------
' Procedimiento que establece las cabeceras de las listas
Private Sub cabecera()
    With listaPlanes_Mto.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Plan", 2500, lvwColumnLeft
        'E0305-I
        '.Add , , "Descripción", 4000, lvwColumnLeft
        'E0305-F
        .Add , , "Frecuencia", 1000, lvwColumnLeft
        .Add , , "Operador", 2500, lvwColumnLeft
        .Add , , "Operador_id", 0, lvwColumnLeft
        .Add , , "Orden", 0, lvwColumnLeft
    End With
    
    With listaPlanMto_Detalle.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        'E0308-I
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Acción", 4600, lvwColumnLeft
        .Add , , "T. Previsto (Min)", 1400, lvwColumnRight
        'E0308-F
    End With
End Sub
' Procedimiento que carga la lista
Public Sub cargar_lista()
    Dim rs As ADOdb.RecordSet
    Dim oEQ_Plan_Mto_asignados As New clsEquipos_Planes_mto_asignados

    listaPlanes_Mto.ListItems.Clear
    Set rs = oEQ_Plan_Mto_asignados.lista_por_equipos(PK)
    If rs.RecordCount <> 0 Then
        Do
            With listaPlanes_Mto.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4) ' Operador
             .SubItems(5) = rs(5) ' T Estimado
             'E0307-I
             '.SubItems(6) = rs(6)
             'E0307-F
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEQ_Plan_Mto_asignados = Nothing
    listaPlanMto_Detalle.ListItems.Clear
    listaPlanes_Mto_Click
End Sub

Private Sub listaPlanes_Mto_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If listaPlanes_Mto.ListItems.Count > 0 Then
     listaPlanes_Mto.SortKey = ColumnHeader.Index - 1
     If listaPlanes_Mto.SortOrder = 0 Then
        listaPlanes_Mto.SortOrder = 1
     Else
        listaPlanes_Mto.SortOrder = 0
     End If
     listaPlanes_Mto.Sorted = True
   End If
End Sub
