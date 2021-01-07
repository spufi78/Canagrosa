VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmFluidos_Resultados 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultados de la contaminación por partículas"
   ClientHeight    =   10260
   ClientLeft      =   4080
   ClientTop       =   1485
   ClientWidth     =   13635
   Icon            =   "frmFluidos_Resultados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   13635
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reactivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   45
      TabIndex        =   91
      Top             =   7965
      Width           =   6945
      Begin VB.CommandButton cmdAnadirReactivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   810
         Left            =   5805
         Picture         =   "frmFluidos_Resultados.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   93
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   1170
         Width           =   915
      End
      Begin VB.CommandButton cmdEliminarReactivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   795
         Left            =   5805
         Picture         =   "frmFluidos_Resultados.frx":1B3C
         Style           =   1  'Graphical
         TabIndex        =   92
         Tag             =   "Elimina el campo seleccionado"
         Top             =   270
         Width           =   915
      End
      Begin MSComctlLib.ListView listaReactivos 
         Height          =   1155
         Left            =   90
         TabIndex        =   94
         Top             =   270
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   2037
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
      Begin pryCombo.miCombo cmbReactivosExternos 
         Height          =   330
         Left            =   840
         TabIndex        =   95
         Top             =   1485
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbReactivosInternos 
         Height          =   330
         Left            =   840
         TabIndex        =   96
         Top             =   1830
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interno"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   98
         Top             =   1875
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externos"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   97
         Top             =   1530
         Width           =   615
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Contador de Partículas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   45
      TabIndex        =   82
      Top             =   6030
      Width           =   6945
      Begin VB.TextBox txtFlujo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1305
         TabIndex        =   83
         Top             =   1440
         Width           =   3300
      End
      Begin pryCombo.miCombo cmbContador 
         Height          =   345
         Left            =   1305
         TabIndex        =   84
         Top             =   315
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbSensor 
         Height          =   345
         Left            =   1305
         TabIndex        =   87
         Top             =   1080
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fechaUltimaCalibracion 
         Height          =   345
         Left            =   1305
         TabIndex        =   89
         Top             =   675
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
         Format          =   51970049
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Ult.Calibración"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   90
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sensor"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   88
         Top             =   1125
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flujo"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   86
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contador"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   85
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.TextBox txtcampo 
      Height          =   375
      Left            =   7515
      TabIndex        =   63
      Text            =   "Text1"
      Top             =   9675
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Peso de los Grados"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7200
      TabIndex        =   55
      Top             =   8460
      Visible         =   0   'False
      Width           =   6270
      Begin VB.TextBox txtpesogrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   0
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox txtpesogrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   1
         Left            =   1215
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox txtpesogrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   2
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox txtpesogrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   3
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   270
         Width           =   1245
      End
      Begin VB.TextBox txtpesogrado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   4
         Left            =   4815
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Frame frmEquipos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipos Utilizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   45
      TabIndex        =   35
      Top             =   3870
      Width           =   6945
      Begin VB.CommandButton cmdEliminarEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   810
         Left            =   5805
         Picture         =   "frmFluidos_Resultados.frx":2406
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "Elimina el campo seleccionado"
         Top             =   270
         Width           =   915
      End
      Begin VB.CommandButton cmdAnadirEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   765
         Left            =   5805
         Picture         =   "frmFluidos_Resultados.frx":2CD0
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   1215
         Width           =   915
      End
      Begin MSComctlLib.ListView listaEquipos 
         Height          =   1425
         Left            =   90
         TabIndex        =   38
         Top             =   270
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   2514
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
         TabIndex        =   39
         Top             =   1710
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   582
      End
   End
   Begin VB.Frame frmTipo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle del Ensayo"
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
      TabIndex        =   5
      Top             =   585
      Width           =   6945
      Begin VB.TextBox txtparametro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   4920
      End
      Begin VB.TextBox txtparametro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   585
         Width           =   4920
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "APC"
         Height          =   240
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   990
         Width           =   915
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Microscope"
         Height          =   240
         Index           =   1
         Left            =   1935
         TabIndex        =   6
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Equipo"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   990
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parámetro Analizado"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   30
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma de Clasificación"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   630
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Análisis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7755
      Left            =   7020
      TabIndex        =   19
      Top             =   585
      Width           =   6540
      Begin VB.CheckBox chkFRangoContaminacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Rango"
         Height          =   285
         Left            =   4725
         TabIndex        =   104
         Top             =   7380
         Width           =   1140
      End
      Begin VB.Frame frmTotales 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totales"
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   180
         TabIndex        =   77
         Top             =   6030
         Visible         =   0   'False
         Width           =   6270
         Begin VB.CheckBox chkTotal 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Rango"
            Height          =   285
            Left            =   4545
            TabIndex        =   105
            Top             =   540
            Width           =   1140
         End
         Begin VB.TextBox txttotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   1
            Left            =   4545
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   180
            Width           =   1470
         End
         Begin VB.TextBox txttotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   0
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   180
            Width           =   1200
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Clasificación"
            Height          =   195
            Left            =   3555
            TabIndex        =   81
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "T.Particulas"
            Height          =   195
            Left            =   180
            TabIndex        =   79
            Top             =   270
            Width           =   840
         End
      End
      Begin VB.OptionButton opCalibracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NAVAIR"
         Height          =   240
         Index           =   3
         Left            =   4365
         TabIndex        =   76
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "% de Diferencia calculada"
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   180
         TabIndex        =   70
         Top             =   4365
         Width           =   6270
         Begin VB.TextBox txtdifres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   4
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtdifres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtdifres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   2
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtdifres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   1
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtdifres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   0
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   225
            Width           =   1200
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "% de Diferencia máxima permitida"
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   180
         TabIndex        =   64
         Top             =   3690
         Width           =   6270
         Begin VB.TextBox txtdif 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   0
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtdif 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   1
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtdif 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   2
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtdif 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtdif 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   4
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   225
            Width           =   1200
         End
      End
      Begin VB.TextBox txtdifmax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   7020
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   2745
         TabIndex        =   52
         Top             =   180
         Width           =   2760
         Begin VB.OptionButton opIP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Conforme"
            Height          =   240
            Index           =   1
            Left            =   1305
            TabIndex        =   54
            Top             =   135
            Width           =   1275
         End
         Begin VB.OptionButton opIP 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Conforme"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   53
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.CheckBox chkPC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No utilizar el primer conteo"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1800
         TabIndex        =   51
         Top             =   2970
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.TextBox txtgradocontaminacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   7020
         Width           =   1500
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grados"
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   180
         TabIndex        =   43
         Top             =   5040
         Width           =   6270
         Begin VB.CheckBox chkFRango 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Rango"
            Height          =   285
            Index           =   4
            Left            =   4815
            TabIndex        =   103
            Top             =   585
            Width           =   1140
         End
         Begin VB.CheckBox chkFRango 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Rango"
            Height          =   285
            Index           =   3
            Left            =   3645
            TabIndex        =   102
            Top             =   585
            Width           =   1140
         End
         Begin VB.CheckBox chkFRango 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Rango"
            Height          =   285
            Index           =   2
            Left            =   2430
            TabIndex        =   101
            Top             =   585
            Width           =   1140
         End
         Begin VB.CheckBox chkFRango 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Rango"
            Height          =   285
            Index           =   1
            Left            =   1260
            TabIndex        =   100
            Top             =   585
            Width           =   1140
         End
         Begin VB.CheckBox chkFRango 
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Rango"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   99
            Top             =   585
            Width           =   1140
         End
         Begin VB.TextBox txtgrado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   4
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtgrado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtgrado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   2
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtgrado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   1
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   225
            Width           =   1200
         End
         Begin VB.TextBox txtgrado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   0
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   225
            Width           =   1200
         End
      End
      Begin VB.OptionButton opCalibracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ACFTD"
         Height          =   240
         Index           =   1
         Left            =   1710
         TabIndex        =   41
         Top             =   1080
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.TextBox txtvaloranalizado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2025
         TabIndex        =   34
         Top             =   630
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Media de conteos"
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   180
         TabIndex        =   21
         Top             =   2970
         Width           =   6270
         Begin VB.TextBox txtmedia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   4
            Left            =   4815
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   270
            Width           =   1200
         End
         Begin VB.TextBox txtmedia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   270
            Width           =   1245
         End
         Begin VB.TextBox txtmedia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   2
            Left            =   2385
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   270
            Width           =   1245
         End
         Begin VB.TextBox txtmedia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   1
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   270
            Width           =   1200
         End
         Begin VB.TextBox txtmedia 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   0
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   270
            Width           =   1200
         End
      End
      Begin TrueDBGrid80.TDBGrid gridR 
         Height          =   1515
         Left            =   180
         TabIndex        =   20
         Top             =   1350
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   2672
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "(5 - 15) um"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "General Number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "(16 - 25) um"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "(26 - 50) um"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "General Number"
         Columns(2).ExternalEditor=   "TDBDate1"
         Columns(2).ExternalEditor.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "(51 - 100) um"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "General Number"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "> 100 um"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "General Number"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).PartialRightColumn=   0   'False
         Splits(0).MarqueeStyle=   1
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=2117"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2037"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(1).DropDownList=1"
         Splits(0)._ColumnProps(16)=   "Column(2).Width=2117"
         Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=2037"
         Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(23)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(24)=   "Column(3).Width=2117"
         Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2037"
         Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(3).AllowSizing=0"
         Splits(0)._ColumnProps(29)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(31)=   "Column(4).Width=556"
         Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=476"
         Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(36)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         EditDropDown    =   0   'False
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Nº DE PARTICULAS POR 100 ml EN CADA RANGO"
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         MultiSelect     =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
         _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H8000000A&,.bold=0"
         _StyleDefs(14)  =   ":id=3,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
         _StyleDefs(23)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
         _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=11,.alignment=2,.locked=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=36,.parent=11,.alignment=2,.fgcolor=&H0&,.bold=0"
         _StyleDefs(41)  =   ":id=36,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(42)  =   ":id=36,.fontname=MS Sans Serif"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=33,.parent=12"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=34,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=35,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=58,.parent=11,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=12"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=12"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=15"
         _StyleDefs(58)  =   "Named:id=37:Normal"
         _StyleDefs(59)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(60)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(61)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(62)  =   "Named:id=38:Heading"
         _StyleDefs(63)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(66)  =   ":id=38,.fontname=MS Sans Serif"
         _StyleDefs(67)  =   "Named:id=39:Footing"
         _StyleDefs(68)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=40:Selected"
         _StyleDefs(70)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(71)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(72)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(73)  =   "Named:id=41:Caption"
         _StyleDefs(74)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(75)  =   "Named:id=42:HighlightRow"
         _StyleDefs(76)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(77)  =   "Named:id=43:EvenRow"
         _StyleDefs(78)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(79)  =   "Named:id=44:OddRow"
         _StyleDefs(80)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(81)  =   "Named:id=47:RecordSelector"
         _StyleDefs(82)  =   ":id=47,.parent=38"
         _StyleDefs(83)  =   "Named:id=50:FilterBar"
         _StyleDefs(84)  =   ":id=50,.parent=37"
      End
      Begin VB.OptionButton opCalibracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ISO 11170"
         Height          =   240
         Index           =   2
         Left            =   3015
         TabIndex        =   40
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%.Dif. Maximo"
         Height          =   195
         Left            =   225
         TabIndex        =   62
         Top             =   7110
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grado Contaminación"
         Height          =   195
         Left            =   3105
         TabIndex        =   50
         Top             =   7110
         Width           =   1530
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de calibración"
         Height          =   195
         Left            =   225
         TabIndex        =   42
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volumen analizado (mL)"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resultado de la inspección previa:"
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Top             =   315
         Width           =   2445
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limpieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   45
      TabIndex        =   9
      Top             =   1935
      Width           =   6945
      Begin VB.OptionButton opLimpieza 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Conforme"
         Height          =   240
         Index           =   1
         Left            =   5400
         TabIndex        =   18
         Top             =   945
         Width           =   1275
      End
      Begin VB.OptionButton opLimpieza 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conforme"
         Height          =   240
         Index           =   0
         Left            =   4185
         TabIndex        =   17
         Top             =   945
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.TextBox txtob 
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1305
         Width           =   6630
      End
      Begin VB.TextBox txtparticulas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2475
         TabIndex        =   12
         Top             =   675
         Width           =   1455
      End
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   345
         Left            =   1170
         TabIndex        =   10
         Top             =   315
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   609
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Criterio menor de 1500 particulas*mL"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   4185
         TabIndex        =   16
         Top             =   720
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número de partículas > 4 um"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   765
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      Height          =   870
      Left            =   10290
      Picture         =   "frmFluidos_Resultados.frx":359A
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   9315
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11430
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9315
      Width           =   1050
   End
   Begin VB.TextBox txtnorma 
      Height          =   375
      Left            =   8460
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   9675
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9315
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados de la contaminación por partículas"
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
      Top             =   30
      Width           =   4830
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13005
      Picture         =   "frmFluidos_Resultados.frx":3E64
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Introduzca los resultados para la contaminación por partículas"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   285
      Width           =   4830
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13680
   End
End
Attribute VB_Name = "frmFluidos_Resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public muestra As Long
Public determinacion_id As Long

Dim xR As New XArrayDB
Const filasR As Integer = 100
Const ColR As Integer = 4
Private Enum ColsR
    r1 = 0
    R2 = 1
    R3 = 2
    R4 = 3
    R5 = 4
End Enum

Private Sub chkFRango_Click(Index As Integer)
    If chkFRango(Index).Value = Checked Then
        txtgrado(Index).ForeColor = vbRed
    Else
        txtgrado(Index).ForeColor = vbBlack
    End If
End Sub

Private Sub chkFRangoContaminacion_Click()
    If chkFRangoContaminacion.Value = Checked Then
        txtgradocontaminacion.ForeColor = vbRed
    Else
        txtgradocontaminacion.ForeColor = vbBlack
    End If
End Sub

Private Sub chkPC_Click()
    calcular_media
End Sub

Private Sub cmbContador_change()
   On Error GoTo cmbContador_change_Error
    If cmbContador.getTEXTO = "" Then
        Exit Sub
    End If
    Dim oEquipo As New clsEquipos
    Dim calibracion As Long
    If cmbContador.getPK_SALIDA = 1302 Or cmbContador.getPK_SALIDA = 1304 Then
        calibracion = oEquipo.devolver_id_ult_verificacion(cmbContador.getPK_SALIDA)
        If calibracion > 0 Then
            Dim oEV As New clsEquipoVerificacion
            oEV.Carga calibracion
            fechaUltimaCalibracion = oEV.getFECHA_ACTUAL
            Set oEV = Nothing
        End If
    
    Else
        calibracion = oEquipo.devolver_id_ult_calibracion(cmbContador.getPK_SALIDA)
        If calibracion > 0 Then
            Dim oEC As New clsEquipoCalibracion
            oEC.Carga calibracion
            fechaUltimaCalibracion = oEC.getFECHA_ACTUAL
            Set oEC = Nothing
        End If
    End If
    Set oEquipo = Nothing

   On Error GoTo 0
   Exit Sub

cmbContador_change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbContador_change of Formulario frmFluidos_Resultados"
End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim oEquipo As New clsEquipos
        oEquipo.Carga_Datos_Basicos cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
    End If
End Sub

Private Sub cmdAnadirReactivo_Click()
    ' Externo (E)
    If cmbReactivosExternos.getTEXTO <> "" Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
        Dim oTR As New clsTipos_reactivo_ex
        oBote.CARGAR cmbReactivosExternos.getPK_SALIDA
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "E"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Interno (I)
    If cmbReactivosInternos.getTEXTO <> "" Then
        Dim oRPR As New clsRpr_botes
        Dim oTRPR As New clsRPR_Tipos
        oRPR.Carga cmbReactivosInternos.getPK_SALIDA
        oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
        With listaReactivos.ListItems.Add(, , oRPR.getID_BOTE_PR)
            .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    cmbReactivosExternos.limpiar
    cmbReactivosInternos.limpiar
    almacenar_reactivos
End Sub

Private Sub almacenar_reactivos()
      ' Equipos
      Dim OTDE As New clsDeterminaciones_reactivos
      Dim i As Integer
     ' Dim oDeterminacion As New clsDeterminaciones
      
   On Error GoTo almacenar_reactivos_Error
   
    'oDeterminacion.CargarDeterminacion (DETERMINACION_ID)

      OTDE.Eliminar determinacion_id
      For i = 1 To listaReactivos.ListItems.Count
        With OTDE
            .setDETERMINACION_ID = determinacion_id
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setTIPO = listaReactivos.ListItems(i).SubItems(3)
            .setORDEN = i
            .Insertar_Determinacion
        End With
      Next

   On Error GoTo 0
   Exit Sub

almacenar_reactivos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure almacenar_reactivos of Formulario frmDeterminaciones"

End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
    End If

End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub



Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivosExternos.limpiar
        cmbReactivosInternos.limpiar
        almacenar_reactivos
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    If Not validar Then
        Exit Sub
    End If
    ' Almacenar datos de inspección previa
    Dim oIP As New clsFluidos_ip
    With oIP
        .Eliminar muestra
        .setMUESTRA_ID = muestra
        If opTipo(1).Value = True Then
            .setTIPO_EQUIPO = 1
        Else
            .setTIPO_EQUIPO = 2
        End If
        .setBOTE_EX_ID = cmbReactivos.getPK_SALIDA
        .setNUMERO_PARTICULAS = CLng(txtparticulas)
        If opLimpieza(0).Value = True Then
            .setLIMPIEZA_CONFORME = 0
        Else
            .setLIMPIEZA_CONFORME = 1
        End If
        .setLIMPIEZA_OB = txtob
        If opIP(0).Value = True Then
            .setINSPECCION_CONFORME = 0
        Else
            .setINSPECCION_CONFORME = 1
        End If
        .setVOLUMEN = txtvaloranalizado
        If opCalibracion(1).Value = True Then
            .setTIPO_CALIBRACION = 1
        ElseIf opCalibracion(2).Value = True Then
            .setTIPO_CALIBRACION = 2
        ElseIf opCalibracion(3).Value = True Then
            .setTIPO_CALIBRACION = 3
        End If
        .setPRIMER_CONTEO = chkPC.Value
        ' Contador de particulas
        .setCONTADOR_ID = cmbContador.getPK_SALIDA
        .setFECHA_ULT_CALIBRACION = fechaUltimaCalibracion.Value
        .setSENSOR_ID = cmbSensor.getPK_SALIDA
        .setFLUJO = txtFlujo
        
        .Insertar
    End With
    ' Almacenar los datos de los equipos
    almacenar_equipos
    ' Almacenar los resultados del fluido
    Dim oFR As New clsFluidos_resultados
    Dim i As Integer
    Dim j As Integer
    With oFR
        .Eliminar muestra
        ' Almacenamos en ORDEN 0 el resultado
        For i = 1 To 5
            .setMUESTRA_ID = muestra
            .setTAMANO = i
            If txtMedia(i - 1) = "" Then
                .setRESULTADO = 0
            Else
                .setRESULTADO = txtMedia(i - 1)
            End If
            .setCLASIFICACION = txtgrado(i - 1)
            .setFUERA_RANGO = chkFRango(i - 1).Value
            .Insertar
        Next
        ' Si es NAVAIR almacenamos el TOTAL
        If opCalibracion(3).Value = True Then
            .setTAMANO = 6
            If txttotal(0) = "" Then
                .setRESULTADO = 0
            Else
                .setRESULTADO = txttotal(0)
            End If
            .setCLASIFICACION = txttotal(1)
            .setFUERA_RANGO = chkTotal.Value
            .Insertar
        End If
    End With
    ' Almacenamos los resultados del  grid
    Dim oFRG As New clsFluidos_resultados_grid
    With oFRG
        oFRG.Eliminar muestra
        For i = 0 To filasR
            For j = 0 To ColR
                If Not IsEmpty(xR(i, j)) Then
                    If IsNumeric(xR(i, j)) Then
                        .setMUESTRA_ID = muestra
                        .setORDEN = i
                        .setTAMANO = j + 1
                        .setRESULTADO = xR(i, j)
                        .Insertar
                    End If
                End If
            Next
        Next
    End With
    Set oFRG = Nothing
    ' Almacenar el resultado de la determinacion
    If txtgradocontaminacion <> "" Then
        Dim oDeter As New clsDeterminaciones
        With oDeter
            .setRESULTADO = txtgradocontaminacion
            .setDIF_DUPLICADOS = txtdifmax
            .setFECHA = Format(Date, "yyyy-mm-dd")
            .setHORA = Format(Time, "hh:mm")
            .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
            
'            .InsertarSolucion (CLng(txtdeter))
            .InsertarSolucion determinacion_id
                        
            ' Indicar en Situacion si esta o no en grado
            If chkFRangoContaminacion.Value = Checked Then
                execute_bd "UPDATE DETERMINACIONES SET SITUACION = " & C_SITUACION.S_FUERA_RANGO & " WHERE ID_DETERMINACION = " & determinacion_id, False
            Else
                execute_bd "UPDATE DETERMINACIONES SET SITUACION = " & C_SITUACION.S_EN_RANGO & " WHERE ID_DETERMINACION = " & determinacion_id, False
            End If
        End With
        Set oDeter = Nothing
        Dim odd As New clsDatos_determinaciones
        With odd
'            .setDETERMINACION_ID = CLng(txtdeter)
            .setDETERMINACION_ID = determinacion_id
            .setCAMPO_ID = CLng(txtcampo)
            .setVALOR_1 = txtgradocontaminacion
            .Insertar_Valores
        End With
        Set odd = Nothing

    End If
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmFluidos_Resultados"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbContador, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbSensor, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbReactivosExternos, New clsBotes_ex, 0, Me, " AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
    
'    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, "AND ABIERTO = 1"
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, "AND ABIERTO = 1 AND FINALIZADO = 0"
    inicializar_grid filasR
    cargar_muestra
    cargar_resultados
    cargar_reactivos
End Sub
Private Sub cabecera()
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 800, lvwColumnLeft
        .Add , , "Nombre", 3200, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "ID", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
End Sub
Private Sub gridR_KeyUp(KeyCode As Integer, Shift As Integer)
    calcular_media
End Sub
Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipos_Detalle.PK = listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text
        frmEquipos_Detalle.Show 1
    End If
End Sub

Private Sub opCalibracion_Click(Index As Integer)
    frmTotales.visible = False
    Select Case Index
    Case 1 ' ACFTD
        gridR.Columns(ColsR.r1).Caption = "(5 - 15) um"
        gridR.Columns(ColsR.R2).Caption = "(16 - 25) um"
        gridR.Columns(ColsR.R3).Caption = "(26 - 50) um"
        gridR.Columns(ColsR.R4).Caption = "(51 - 100) um"
        gridR.Columns(ColsR.R5).Caption = ">100 um"
    Case 2 ' ISO 11170
        gridR.Columns(ColsR.r1).Caption = "(6 - 14) um"
        gridR.Columns(ColsR.R2).Caption = "(15 - 21) um"
        gridR.Columns(ColsR.R3).Caption = "(22 - 38) um"
        gridR.Columns(ColsR.R4).Caption = "(39 - 70) um"
        gridR.Columns(ColsR.R5).Caption = ">70 um"
    Case 3 ' NAVAIR
        gridR.Columns(ColsR.r1).Caption = "(5 - 10) um"
        gridR.Columns(ColsR.R2).Caption = "(10 - 25) um"
        gridR.Columns(ColsR.R3).Caption = "(25 - 50) um"
        gridR.Columns(ColsR.R4).Caption = "(50 - 100) um"
        gridR.Columns(ColsR.R5).Caption = ">100 um"
        frmTotales.visible = True
    End Select
    txtnorma = Index
    calcular_media
End Sub

Private Sub opLimpieza_Click(Index As Integer)
    calcular_criterio
End Sub

'Private Sub resultados_Click()
'    If resultados.ListItems.Count > 0 Then
'        txtparametro(2) = resultados.ListItems(resultados.selectedItem.Index).SubItems(1)
'        txtparametro(3) = resultados.ListItems(resultados.selectedItem.Index).SubItems(2)
'        txtparametro(4) = resultados.ListItems(resultados.selectedItem.Index).SubItems(3)
'        txtparametro(3).SetFocus
'    End If
'End Sub
Private Function calcular_grado(NORMA As Integer, TAMANO As Integer, resultado As Long) As String
    Dim rango As String
    Dim oFN As New clsFluidos_normas_valores
    rango = oFN.Calcula_Grado(NORMA, TAMANO, resultado)
'    If Trim(rango) = "" Then
'        rango = "N/A"
'    End If
    calcular_grado = rango
End Function

Private Sub txtparticulas_Change()
    calcular_criterio
End Sub

Private Sub calcular_criterio()
    If txtparticulas.Text <> "" Then
        If IsNumeric(txtparticulas) Then
            If CLng(txtparticulas) >= 1500 Then
                opLimpieza(0).Value = False
            Else
                opLimpieza(0).Value = True
            End If
        End If
    End If
End Sub
Private Sub inicializar_grid(filas As Integer)
   On Error GoTo inicializar_grid_Error

    gridR.Col = 0
    gridR.Row = 0
    xR.Clear
    xR.ReDim 0, filas, 0, ColR
    Set gridR.Array = xR
    gridR.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub
Private Sub cargar_muestra()
    Dim oMuestra As New clsMuestra
    Dim oFluido As New clsFluidos_ficha
    Dim oFN As New clsFluidos_normas
    
    oMuestra.CargaMuestra (muestra)
    If oMuestra.getCERRADA = 1 Then
        cmdok.Enabled = False
    End If
    oFluido.Carga_por_BANO (oMuestra.getBANO_ID)
    oFN.Carga (oFluido.getNORMA_ID)
    
    txtparametro(1) = oFN.getNOMBRE
'    cargar_lista (oFluido.getNORMA_ID)
    txtnorma = oFluido.getNORMA_ID
    Set oMuestra = Nothing
    Set oFluido = Nothing
    Set oFN = Nothing
    ' Campo
    Dim odd As New clsDatos_determinaciones
    Dim rs As ADODB.Recordset
'    Set rs = oDD.cargar_determinacion(CLng(txtdeter))
    Set rs = odd.cargar_determinacion(determinacion_id)
    If rs.RecordCount > 0 Then
        txtcampo = rs("CAMPO_ID")
    End If
    ' Equipos
    listaEquipos.ListItems.Clear
    Dim OTDEQUIPOS As New clsDeterminaciones_equipos
    Set rs = OTDEQUIPOS.Listado(determinacion_id)
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Cargar datos primera inspeccion
    Dim oFI As New clsFluidos_ip
    If oFI.Carga(muestra) Then
        With oFI
            If .getTIPO_EQUIPO <> 0 Then
                opTipo(.getTIPO_EQUIPO).Value = True
            End If
            If .getBOTE_EX_ID <> 0 Then
                llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, ""
                cmbReactivos.MostrarElemento .getBOTE_EX_ID
            End If
            txtparticulas = .getNUMERO_PARTICULAS
            txtob = .getLIMPIEZA_OB
            opIP(.getINSPECCION_CONFORME).Value = True
            txtvaloranalizado = .getVOLUMEN
            opCalibracion(.getTIPO_CALIBRACION).Value = True
            chkPC.Value = .getPRIMER_CONTEO
            ' Contador de partículas
            cmbContador.MostrarElemento .getCONTADOR_ID
            If .getFECHA_ULT_CALIBRACION <> "" And .getFECHA_ULT_CALIBRACION <> "0000-00-00" Then
                fechaUltimaCalibracion = .getFECHA_ULT_CALIBRACION
            End If
            cmbSensor.MostrarElemento .getSENSOR_ID
            txtFlujo = .getFLUJO
            ' Mari Cruz 14-11-2020
            If txtFlujo = "" Then
                txtFlujo = "60 mL/min"
            End If
        End With
    Else
        txtFlujo = "60 mL/min"
    End If
    Set oFI = Nothing
End Sub

Private Sub calcular_media()
    Dim i As Integer
    Dim media As Long
    Dim inicio As Integer
    Dim valores As Integer
    Dim GRADO() As String
    Dim g As String
   On Error GoTo calcular_media_Error

    txtgradocontaminacion = ""
    Dim resultado As Integer
    Dim dif As Single
    Dim mediad As Single
    Dim dif_media As Single
    resultado = 0
    inicio = 0
    If chkPC.Value = Checked Then
        inicio = 1
    End If
    Dim menor As Long
    Dim mayor As Long
  
    Dim VALOR As Long
    
    For j = 0 To ColR
        media = 0
        valores = 0
        menor = 0
        mayor = 0
        For i = inicio To filasR
            If Not IsEmpty(xR(i, j)) Then
                If IsNumeric(xR(i, j)) Then
                    media = media + CLng(xR(i, j))
                    valores = valores + 1
                    ' Buscamos el menor y mayor valor
                    If menor > CLng(xR(i, j)) Or menor = 0 Then
                        menor = CLng(xR(i, j))
                    End If
                    If mayor < CLng(xR(i, j)) Then
                        mayor = CLng(xR(i, j))
                    End If
                End If
            End If
        Next
        If valores > 0 Then
            txtMedia(j) = CLng(media / valores)
            g = calcular_grado(CInt(txtnorma), j + 1, CLng(media / valores))
            If g <> "" Then
                GRADO = Split(g, ";")
                txtgrado(j) = GRADO(0)
                txtpesogrado(j) = GRADO(1)
            End If
            ' Dif
            ' JGM
'            g = calcular_grado(100, 1, CLng(media / valores))
            If IsNumeric(txtvaloranalizado) Then
                VALOR = ((CLng(media / valores) * txtvaloranalizado)) / 100
            Else
                VALOR = CLng(media / valores)
            End If
            g = calcular_grado(100, 1, VALOR)
            If g <> "" Then
                txtdif(j) = g
            End If
            ' Calcular diferencia entre resultados
            If (menor <> 0 Or mayor <> 0) And valores > 1 Then
                dif = Abs((CSng(mayor) - CSng(menor)))
                mediad = (CSng(mayor) + CSng(menor)) / 2
                dif_media = (dif / mediad) * 100
                
                txtdifres(j) = Format(dif_media, "#0.0")
            End If
        Else
            txtMedia(j) = 0
            txtgrado(j) = ""
        End If
        If txtpesogrado(j) <> "" Then
            If CInt(txtpesogrado(j)) >= resultado Then
                txtgradocontaminacion = txtgrado(j)
                resultado = CInt(txtpesogrado(j))
            End If
        End If
    Next
    ' Calcular el total para los NAVAIR
    If opCalibracion(3).Value = True Then
        Dim total As Long
        For i = 0 To 4
            If IsNumeric(txtMedia(i)) Then
                total = total + CLng(txtMedia(i))
            End If
        Next
        txttotal(0) = total
        ' Buscar la clase del resultado Total
        g = calcular_grado(CInt(txtnorma), 6, total)
        If g <> "" Then
            GRADO = Split(g, ";")
            txttotal(1) = GRADO(0)
       End If
    End If
    ' Dif max
'    Dim txtmax As Single
'    txtmax = 0
'    For i = 0 To 4
'        If IsNumeric(txtdif(i)) Then
'            If CSng(Replace(txtdif(i), ".", ",")) > CSng(Replace(txtmax, ".", ",")) Then
'                txtmax = Replace(txtdif(i), ".", ",")
'            End If
'        End If
'    Next
'    If txtmax = 0 Then
'        txtdifmax = "N.A."
'    Else
'        txtdifmax = Replace(txtmax, ",", ".")
'    End If
    colorear

   On Error GoTo 0
   Exit Sub

calcular_media_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcular_media of Formulario frmFluidos_Resultados"
End Sub
Private Sub almacenar_equipos()
      ' Equipos
      Dim OTDE As New clsDeterminaciones_equipos
   On Error GoTo almacenar_equipos_Error
      Dim i As Integer
      OTDE.Eliminar determinacion_id
      For i = 1 To listaEquipos.ListItems.Count
        With OTDE
            .setDETERMINACION_ID = determinacion_id
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setORDEN = i
            .Insertar_Determinacion
        End With
      Next

   On Error GoTo 0
   Exit Sub

almacenar_equipos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure almacenar_equipos of Formulario frmDeterminaciones"
End Sub


Private Function validar() As Boolean
    validar = True
    If opTipo(1).Value = False And opTipo(2).Value = False Then
        MsgBox "Indique un tipo de equipo (Microscope o APC)", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbReactivos.getTEXTO = "" Then
        MsgBox "Indique el reactivo utilizado.", vbExclamation, App.Title
        cmbReactivos.SetFocus
        validar = False
        Exit Function
    End If
    If txtparticulas = "" Then
        MsgBox "Indique el numero de particulas.", vbExclamation, App.Title
        txtparticulas.SetFocus
        validar = False
        Exit Function
    Else
        If Not IsNumeric(txtparticulas) Then
            MsgBox "El numero de particulas debe ser numérico.", vbExclamation, App.Title
            txtparticulas.SetFocus
            validar = False
            Exit Function
        End If
    End If
    If opLimpieza(0).Value = False And opLimpieza(1).Value = False Then
        MsgBox "Indique si la limpieza es o no conforme.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If opIP(0).Value = False And opIP(1).Value = False Then
        MsgBox "Indique el resultado de la inspección previa.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    

End Function

Private Sub cargar_resultados()
    ' Cargar datos de resultados
    Dim oFR As New clsFluidos_resultados_grid
    Dim rs As ADODB.Recordset
    Set rs = oFR.Listado(muestra)
    If rs.RecordCount <> 0 Then
        Do
            xR(rs("orden"), rs("tamano") - 1) = CStr(rs("resultado"))
            rs.MoveNext
        Loop Until rs.EOF
        calcular_media
    End If
    Set oFR = Nothing
    ' Carga de los resultados F.RANGO
    Dim oFResultados As New clsFluidos_resultados
    Set rs = oFResultados.Listado(muestra)
    If rs.RecordCount > 0 Then
        Do
            If rs("TAMANO") = 6 Then
                chkTotal.Value = rs("FUERA_RANGO")
            Else
                chkFRango(rs("TAMANO") - 1).Value = rs("FUERA_RANGO")
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Carga determinacion
    Dim oDeterminacion As New clsDeterminaciones
    oDeterminacion.CargarDeterminacion determinacion_id
    chkFRangoContaminacion.Value = oDeterminacion.getSITUACION
    Set oDeterminacion = Nothing
    Set rs = Nothing
End Sub
Private Sub cargar_reactivos()
    ' Reactivos
    listaReactivos.ListItems.Clear
    Dim OTDR As New clsDeterminaciones_reactivos
    Dim oReactivo As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    Dim oRPR As New clsRpr_botes
    Dim oTRPR As New clsRPR_Tipos
    
    Set rs = OTDR.Listado(determinacion_id)
    If rs.RecordCount > 0 Then
        Do
            If rs(1) = "E" Then
               oReactivo.CARGAR CLng(rs(0))
               oTb.CARGAR oReactivo.getTIPO_BOTE_EX_ID
               oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
               With listaReactivos.ListItems.Add(, , rs(0))
                  .SubItems(1) = oTR.getNOMBRE
                  .SubItems(2) = Format(oReactivo.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                  .SubItems(3) = "E"
               End With
            Else
                oRPR.Carga CLng(rs(0))
                oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
                With listaReactivos.ListItems.Add(, , rs(0))
                    .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
                    .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                    .SubItems(3) = "I"
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub colorear()
    Dim i As Integer
    For i = 0 To 4
        If txtgrado(i).Text <> "" Then
           If (txtgrado(i).Text = "GRADO III" Or _
              txtgrado(i).Text = "GRADO IV" Or _
              txtgrado(i).Text = "GRADO V" Or _
              txtgrado(i).Text = "GRADO 9" Or _
              txtgrado(i).Text = "GRADO 10" Or _
              txtgrado(i).Text = "GRADO 11" Or _
              txtgrado(i).Text = "GRADO 12" Or _
              txtgrado(i).Text = "CLASE III" Or _
              txtgrado(i).Text = "CLASE IV" Or _
              txtgrado(i).Text = "CLASE V" Or _
              txtgrado(i).Text = "CLASE 9" Or _
              txtgrado(i).Text = "CLASE 10" Or _
              txtgrado(i).Text = "CLASE 11" Or _
              txtgrado(i).Text = "CLASE 12") Then
                txtgrado(i).ForeColor = vbRed
                chkFRango(i).Value = Checked
            Else
                txtgrado(i).ForeColor = vbBlack
                chkFRango(i).Value = Unchecked
            End If
        End If
        If txtdifres(i) <> "" And txtdif(i) <> "" Then
            If IsNumeric(txtdifres(i)) And IsNumeric(txtdif(i)) Then
                If CSng(Replace(txtdifres(i), ",", ".")) > CSng(txtdif(i)) Then
                    txtdifres(i).ForeColor = vbRed
                Else
                    txtdifres(i).ForeColor = vbBlack
                End If
            Else
                txtdifres(i).ForeColor = vbBlack
                txtdifres(i) = "N.A."
            End If
        End If
    Next
    If txtgradocontaminacion.Text <> "" Then
       If (txtgradocontaminacion.Text = "GRADO III" Or _
          txtgradocontaminacion.Text = "GRADO IV" Or _
          txtgradocontaminacion.Text = "GRADO V" Or _
          txtgradocontaminacion.Text = "GRADO 9" Or _
          txtgradocontaminacion.Text = "GRADO 10" Or _
          txtgradocontaminacion.Text = "GRADO 11" Or _
          txtgradocontaminacion.Text = "GRADO 12" Or _
          txtgradocontaminacion.Text = "CLASE III" Or _
          txtgradocontaminacion.Text = "CLASE IV" Or _
          txtgradocontaminacion.Text = "CLASE V" Or _
          txtgradocontaminacion.Text = "CLASE 9" Or _
          txtgradocontaminacion.Text = "CLASE 10" Or _
          txtgradocontaminacion.Text = "CLASE 11" Or _
          txtgradocontaminacion.Text = "CLASE 12") Then
            txtgradocontaminacion.ForeColor = vbRed
            chkFRangoContaminacion.Value = Checked
        Else
            txtgradocontaminacion.ForeColor = vbBlack
            chkFRangoContaminacion.Value = Unchecked
        End If
    End If
End Sub

Private Sub txtvaloranalizado_Change()
    calcular_media
End Sub
