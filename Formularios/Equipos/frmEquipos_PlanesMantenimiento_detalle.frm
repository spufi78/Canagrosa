VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipos_PlanesMantenimiento_detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Equipos - Planes de Mantenimiento - Detalle"
   ClientHeight    =   8310
   ClientLeft      =   1335
   ClientTop       =   2010
   ClientWidth     =   12270
   Icon            =   "frmEquipos_PlanesMantenimiento_detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12270
   Begin VB.Frame frmPlanificacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planificación"
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
      Height          =   2400
      Left            =   0
      TabIndex        =   19
      Top             =   4995
      Width           =   12255
      Begin VB.Frame frmDiasMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Días del mes"
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   6705
         TabIndex        =   22
         Top             =   180
         Width           =   3345
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "31"
            Height          =   330
            Index           =   31
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   70
            Tag             =   "0"
            Top             =   1710
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "30"
            Height          =   330
            Index           =   30
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   69
            Tag             =   "0"
            Top             =   1710
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "29"
            Height          =   330
            Index           =   29
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   68
            Tag             =   "0"
            Top             =   1710
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "28"
            Height          =   330
            Index           =   28
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   67
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "27"
            Height          =   330
            Index           =   27
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   66
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "26"
            Height          =   330
            Index           =   26
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   65
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "25"
            Height          =   330
            Index           =   25
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   64
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "24"
            Height          =   330
            Index           =   24
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   63
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "23"
            Height          =   330
            Index           =   23
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   62
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "22"
            Height          =   330
            Index           =   22
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   61
            Tag             =   "0"
            Top             =   1350
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "21"
            Height          =   330
            Index           =   21
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   60
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "20"
            Height          =   330
            Index           =   20
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   59
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "19"
            Height          =   330
            Index           =   19
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   58
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "18"
            Height          =   330
            Index           =   18
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   57
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "17"
            Height          =   330
            Index           =   17
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   56
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "16"
            Height          =   330
            Index           =   16
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   55
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "15"
            Height          =   330
            Index           =   15
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   54
            Tag             =   "0"
            Top             =   990
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "14"
            Height          =   330
            Index           =   14
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   53
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "13"
            Height          =   330
            Index           =   13
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   52
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "12"
            Height          =   330
            Index           =   12
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   51
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "11"
            Height          =   330
            Index           =   11
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "10"
            Height          =   330
            Index           =   10
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "9"
            Height          =   330
            Index           =   9
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "8"
            Height          =   330
            Index           =   8
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "0"
            Top             =   630
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "7"
            Height          =   330
            Index           =   7
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "6"
            Height          =   330
            Index           =   6
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "5"
            Height          =   330
            Index           =   5
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "4"
            Height          =   330
            Index           =   4
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "3"
            Height          =   330
            Index           =   3
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2"
            Height          =   330
            Index           =   2
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdDiasMes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "1"
            Height          =   330
            Index           =   1
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
      End
      Begin VB.Frame frmDiasSemana 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Días de la semana"
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   3465
         TabIndex        =   21
         Top             =   180
         Width           =   2985
         Begin VB.CommandButton cmdDiasSemana 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Viernes"
            Height          =   330
            Index           =   5
            Left            =   1035
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "0"
            Top             =   675
            Width           =   870
         End
         Begin VB.CommandButton cmdDiasSemana 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Jueves"
            Height          =   330
            Index           =   4
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "0"
            Top             =   675
            Width           =   870
         End
         Begin VB.CommandButton cmdDiasSemana 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Miércoles"
            Height          =   330
            Index           =   3
            Left            =   1980
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "0"
            Top             =   270
            Width           =   870
         End
         Begin VB.CommandButton cmdDiasSemana 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Martes"
            Height          =   330
            Index           =   2
            Left            =   1035
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "0"
            Top             =   270
            Width           =   870
         End
         Begin VB.CommandButton cmdDiasSemana 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lunes"
            Height          =   330
            Index           =   1
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "0"
            Top             =   270
            Width           =   870
         End
      End
      Begin VB.Frame frmMeses 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Meses"
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   135
         TabIndex        =   20
         Top             =   180
         Width           =   3120
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "12"
            Height          =   330
            Index           =   12
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "0"
            Top             =   675
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "11"
            Height          =   330
            Index           =   11
            Left            =   2070
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "0"
            Top             =   675
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "10"
            Height          =   330
            Index           =   10
            Left            =   1575
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "0"
            Top             =   675
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "9"
            Height          =   330
            Index           =   9
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "0"
            Top             =   675
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "8"
            Height          =   330
            Index           =   8
            Left            =   585
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "0"
            Top             =   675
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "7"
            Height          =   330
            Index           =   7
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "0"
            Top             =   675
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "6"
            Height          =   330
            Index           =   6
            Left            =   2565
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "5"
            Height          =   330
            Index           =   5
            Left            =   2070
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "4"
            Height          =   330
            Index           =   4
            Left            =   1575
            Style           =   1  'Graphical
            TabIndex        =   26
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "3"
            Height          =   330
            Index           =   3
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "2"
            Height          =   330
            Index           =   2
            Left            =   585
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
         Begin VB.CommandButton cmdMeses 
            BackColor       =   &H00E0E0E0&
            Caption         =   "1"
            Height          =   330
            Index           =   1
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "0"
            Top             =   270
            Width           =   420
         End
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear acción"
      Height          =   870
      Left            =   10
      Picture         =   "frmEquipos_PlanesMantenimiento_detalle.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Crear nueva acción"
      Top             =   7425
      Width           =   1050
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
      Left            =   0
      TabIndex        =   11
      Top             =   3825
      Width           =   12255
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar acción"
         Height          =   870
         Left            =   11115
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminar acción del plan de mantenimiento"
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir acción"
         Height          =   870
         Left            =   9990
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Añadir acción al plan de mantenimiento"
         Top             =   180
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cmbFamiliaAcc 
         Height          =   315
         Left            =   750
         TabIndex        =   71
         Top             =   270
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbAcciones 
         Height          =   315
         Left            =   750
         TabIndex        =   73
         Top             =   630
         Width           =   9060
         _ExtentX        =   15981
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
         TabIndex        =   72
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
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7425
      Width           =   1050
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
      Height          =   1050
      Left            =   0
      TabIndex        =   5
      Top             =   585
      Width           =   12255
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   9705
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   4
         Text            =   "0"
         Top             =   225
         Width           =   600
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   1
         Top             =   270
         Width           =   3930
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   3
         Top             =   585
         Width           =   7575
      End
      Begin MSDataListLib.DataCombo cmbFrecuencia 
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   225
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Minutos"
         Height          =   195
         Index           =   0
         Left            =   10395
         TabIndex        =   15
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Frecuencia"
         Height          =   195
         Index           =   42
         Left            =   5310
         TabIndex        =   13
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Previsto"
         Height          =   195
         Index           =   4
         Left            =   8775
         TabIndex        =   12
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11205
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7425
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2160
      Left            =   0
      TabIndex        =   8
      Top             =   1665
      Width           =   12255
      _ExtentX        =   21616
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   11745
      Picture         =   "frmEquipos_PlanesMantenimiento_detalle.frx":0BFB
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
      TabIndex        =   9
      Top             =   0
      Width           =   12285
   End
End
Attribute VB_Name = "frmEquipos_PlanesMantenimiento_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub cmbAcciones_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbFamiliaAcc_Change()
    Call llenar_combo_filtrado
End Sub

Private Sub Command1_Click()
    frmEquipos_Planes_Acciones.Show 1
    Call llenar_combo_filtrado ' Acciones
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cabecera
    cargar_botones Me
    cargar_combos
    
    If PK <> 0 Then
        lbltitulo(0) = "Modificación de un Plan de Mantenimiento"
        cargar_plan
    Else
        lbltitulo(0) = "Creación de un Plan de Mantenimiento"
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
        oAccion.Carga (CLng(cmbAcciones.BoundText)) ' Se carga la acción para mostrar en la lista su detalle
        With lista.ListItems.Add(, , oAccion.getID_ACCION)
            .SubItems(1) = lista.ListItems.Count
            .SubItems(2) = cmbFamiliaAcc ' la descripción de la familia está aquí
            .SubItems(3) = oAccion.getNOMBRE
            .SubItems(4) = oAccion.getT_PREVISTO
            txtdatos(4) = CLng(txtdatos(4)) + CLng(.SubItems(4))
        End With
        lista.ListItems(lista.ListItems.Count).EnsureVisible
        cmbAcciones.BoundText = ""
    End If
End Sub

' Botón que permite eliminar acciones del detalle
Private Sub cmdEliminar_Click()
    Dim i As Long
    
    If lista.ListItems.Count > 0 Then
        txtdatos(4) = CLng(txtdatos(4)) - CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(4))
        For i = lista.SelectedItem.Index + 1 To lista.ListItems.Count ' Cambiar la numeración de los que le siguen (-1)
            lista.ListItems(i).SubItems(1) = lista.ListItems(i).SubItems(1) - 1
        Next i
        lista.ListItems.Remove lista.SelectedItem.Index
        cmbAcciones.BoundText = ""
    Else
        MsgBox "Debe seleccionar la acción que desea eliminar.", vbInformation, App.Title
        Exit Sub
    End If
End Sub

' Botón que da de alta un plan de mantenimiento
Private Sub cmdAceptar_Click()
    On Error GoTo trataError

    If datos_correctos Then
        ' Plan
        Dim oEQ_Mto As New clsEquipos_planes_Mantenimiento
        Dim oEQ_Mto_Detalle As New clsEquipos_PM_Detalle
        With oEQ_Mto
            .setNOMBRE = txtdatos(0)
            .setDESCRIPCION = txtdatos(1)
            .setFRECUENCIA_ID = cmbFrecuencia.BoundText
            ' planificación
            .setPLANIF_MESES = codificar_planif_meses()
            .setPLANIF_DIAS_SEMANA = codificar_planif_dias_semana()
            .setPLANIF_DIAS_MES = codificar_planif_dias_mes()
        End With
        
        Dim lngPlan As Long
        If PK = 0 Then ' Alta de nuevo plan
            lngPlan = oEQ_Mto.Insertar ' El plan no se ha creado correctamente
            If lngPlan = 0 Then
                Exit Sub
            End If
        Else ' Modificación de plan
            If Not oEQ_Mto.Modificar(PK) Then ' El plan no se ha modificado correctamente
                Exit Sub
            End If
            oEQ_Mto_Detalle.Eliminar (PK)
            lngPlan = PK
        End If
        
        ' Detalle
        Dim lngOrden As Integer
        For lngOrden = 1 To lista.ListItems.Count
            With oEQ_Mto_Detalle
                .setPLAN_MTO_ID = lngPlan
                .setACCION_ID = lista.ListItems(lngOrden)
                .setORDEN = lngOrden
                .Insertar
            End With
        Next
        
        MsgBox "El plan de mantenimiento ha sido almacenado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

trataError:
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_click of Formulario frmEquipos_PlanesMantenimiento_detalle")
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

' Funciones auxiliares del módulo del formulario
' ----------------------------------------------
' Procedimiento que carga el plan de mto seleccionado
Private Sub cargar_plan()
    If PK > 0 Then
        ' Plan
        Dim oEQ_Plan As New clsEquipos_planes_Mantenimiento
        With oEQ_Plan
            .Carga (PK)
            txtdatos(0) = .getNOMBRE
            txtdatos(1) = .getDESCRIPCION
            cmbFrecuencia.BoundText = .getFRECUENCIA_ID
            ' planificación
            Call decodificar_planif_meses(.getPLANIF_MESES)
            Call decodificar_planif_dias_semana(.getPLANIF_DIAS_SEMANA)
            Call decodificar_planif_dias_mes(.getPLANIF_DIAS_MES)
        End With
        
        ' Detalle
        Dim oEQ_plan_detalle As New clsEquipos_PM_Detalle
        Dim rs As ADODB.RecordSet
        Dim lngTotal_tiempo As Long
        lngTotal_tiempo = 0
        lista.ListItems.Clear
        Set rs = oEQ_plan_detalle.Listado(PK)
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
        txtdatos(4) = lngTotal_tiempo
    End If
End Sub

' Función que valida los datos introducidos en el formulario
Private Function datos_correctos() As Boolean
    datos_correctos = True
    
    If Len(Trim(txtdatos(0))) = 0 Then
        MsgBox "Introduzca un nombre para el plan de mantenimiento.", vbCritical, App.Title
        datos_correctos = False
        txtdatos(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txtdatos(1))) = 0 Then
        MsgBox "Introduzca una descripción para el plan de mantenimiento.", vbCritical, App.Title
        datos_correctos = False
        txtdatos(1).SetFocus
        Exit Function
    End If
    If Len(Trim(cmbFrecuencia.BoundText)) = 0 Then
        MsgBox "Introduzca la frecuencia del plan de mantenimiento.", vbCritical, App.Title
        datos_correctos = False
        cmbFrecuencia.SetFocus
        Exit Function
    End If
End Function

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Orden", 650, lvwColumnLeft
        .Add , , "Familia", 3500, lvwColumnLeft
        .Add , , "Acción", 7070, lvwColumnLeft
        .Add , , "Minutos", 1000, lvwColumnRight
    End With
End Sub

Private Sub cargar_combos()
    Dim oDECO As New clsDecodificadora
    oDECO.Cargar_Combo cmbFrecuencia, decodificadora.EQ_periodicidad
    oDECO.Cargar_Combo cmbFamiliaAcc, decodificadora.EQ_FAMILIAS_ACCIONES_PLANES_MTO
End Sub

' PLANIFICACIÓN ----------------------
Private Sub cmdMeses_Click(Index As Integer)
    If cmdMeses(Index).Tag = 0 Then
        cmdMeses(Index).Tag = 1
'        cmdMeses(Index).FontBold = True
'        SetButtonForecolor cmdMeses(Index).hWnd, RGB(255, 0, 0)
        cmdMeses(Index).BackColor = vbRed
    Else
        cmdMeses(Index).Tag = 0
'        cmdMeses(Index).FontBold = False
'        SetButtonForecolor cmdMeses(Index).hWnd, RGB(0, 0, 0)
        cmdMeses(Index).BackColor = &HE0E0E0
    End If
End Sub

Private Sub cmdDiasSemana_Click(Index As Integer)
    If cmdDiasSemana(Index).Tag = 0 Then
        cmdDiasSemana(Index).Tag = 1
'        cmdDiasSemana(Index).FontBold = True
'        SetButtonForecolor cmdDiasSemana(Index).hWnd, RGB(255, 0, 0)
        cmdDiasSemana(Index).BackColor = vbRed
    Else
        cmdDiasSemana(Index).Tag = 0
'        cmdDiasSemana(Index).FontBold = False
'        SetButtonForecolor cmdDiasSemana(Index).hWnd, RGB(0, 0, 0)
        cmdDiasSemana(Index).BackColor = &HE0E0E0
    End If
End Sub

Private Sub cmdDiasMes_Click(Index As Integer)
'    SetButtonForecolor cmdDiasMes(Index).hWnd, RGB(255, 0, 0)
    If cmdDiasMes(Index).Tag = 0 Then
        cmdDiasMes(Index).Tag = 1
'        cmdDiasMes(Index).FontBold = True
'        SetButtonForecolor cmdDiasMes(Index).hWnd, RGB(255, 0, 0)
        cmdDiasMes(Index).BackColor = vbRed
    Else
        cmdDiasMes(Index).Tag = 0
'        cmdDiasMes(Index).FontBold = False
'        SetButtonForecolor cmdDiasMes(Index).hWnd, RGB(0, 0, 0)
        cmdDiasMes(Index).BackColor = &HE0E0E0
    End If
End Sub

Private Function codificar_planif_meses() As String
    Dim lngConta As Long
    Dim strResul As String
    
    strResul = ""
    For lngConta = 1 To 12
        If cmdMeses(lngConta).Tag = 1 Then
            strResul = strResul & lngConta & ","
        End If
    Next lngConta
    If Len(strResul) > 0 Then
        strResul = Left(strResul, Len(strResul) - 1) ' Se le quita la coma final
    End If
    codificar_planif_meses = strResul
End Function

Private Sub decodificar_planif_meses(strPlanif As String)
    Dim vector() As Integer
    Dim lngConta As Long
    
    vector = SplitString2IntArr(strPlanif)
    For lngConta = 0 To UBound(vector)
        cmdMeses(lngConta).Tag = vector(lngConta)
'        cmdMeses(vector(lngConta)).FontBold = True
'        SetButtonForecolor cmdMeses(vector(lngConta)).hWnd, RGB(255, 0, 0)
        'cmdMeses(lngConta).BackColor = vbRed
    Next lngConta
End Sub
Private Sub decodificar_planif_dias_semana(strPlanif As String)
    Dim vector() As String
    Dim lngConta As Long
    
    vector = SplitString2IntArr(strPlanif)
    For lngConta = 0 To UBound(vector)
        cmdDiasSemana(vector(lngConta)).Tag = 1
'        cmdDiasSemana(vector(lngConta)).FontBold = True
'        SetButtonForecolor cmdDiasSemana(vector(lngConta)).hWnd, RGB(255, 0, 0)
        cmdDiasSemana(vector(lngConta)).BackColor = vbRed
    Next lngConta
End Sub
Private Sub decodificar_planif_dias_mes(strPlanif As String)
    Dim vector() As String
    Dim lngConta As Long
    
    vector = SplitString2IntArr(strPlanif, ",")
    For lngConta = 0 To UBound(vector)
        cmdDiasMes(lngConta).Tag = vector(lngConta)
'        cmdDiasMes(vector(lngConta)).FontBold = True
'        SetButtonForecolor cmdDiasMes(vector(lngConta)).hWnd, RGB(255, 0, 0)
        'cmdDiasMes(vector(lngConta)).BackColor = vbRed
    Next lngConta
End Sub

Private Function codificar_planif_dias_semana() As String
    Dim lngConta As Long
    Dim strResul As String
    
    strResul = ""
    For lngConta = 1 To 5
        If cmdDiasSemana(lngConta).Tag = 1 Then
            strResul = strResul & lngConta & ","
        End If
    Next lngConta
    If Len(strResul) > 0 Then
        strResul = Left(strResul, Len(strResul) - 1) ' Se le quita la coma final
    End If
    codificar_planif_dias_semana = strResul
End Function

Private Function codificar_planif_dias_mes() As String
    Dim lngConta As Long
    Dim strResul As String
    
    strResul = ""
    For lngConta = 1 To 31
        If cmdDiasMes(lngConta).Tag = 1 Then
            strResul = strResul & lngConta & ","
        End If
    Next lngConta
    If Len(strResul) > 0 Then
        strResul = Left(strResul, Len(strResul) - 1) ' Se le quita la coma final
    End If
    codificar_planif_dias_mes = strResul
End Function
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
