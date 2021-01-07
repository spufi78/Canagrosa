VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmRecepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion de muestras"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13770
   Icon            =   "frmRecepcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   13770
   Begin VB.TextBox txtdias 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3240
      TabIndex        =   61
      Text            =   "0"
      Top             =   8775
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdNuevoCliente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo Cliente"
      Height          =   870
      Left            =   90
      Picture         =   "frmRecepcion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8550
      Width           =   1365
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12645
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8550
      Width           =   1050
   End
   Begin VB.CommandButton cmdDeterminaciones 
      Caption         =   "Determinaciones"
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
      Height          =   390
      Left            =   7065
      TabIndex        =   52
      Top             =   9015
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdBannos 
      Caption         =   "Baños"
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
      Height          =   390
      Left            =   5445
      TabIndex        =   51
      Top             =   9015
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2790
      Left            =   7785
      TabIndex        =   42
      Top             =   5685
      Width           =   5955
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1845
         Index           =   11
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   840
         Width           =   5745
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   10
         Left            =   4455
         TabIndex        =   21
         ToolTipText     =   "NO USE LA COMA COMO SEPARADOR DE DECIMALES ,USE EL PUNTO -- EJEMPLO: 6020.85 --"
         Top             =   180
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker DTfechaPrevistaEntrega 
         Height          =   330
         Left            =   1755
         TabIndex        =   20
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   61669377
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha entrega"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   45
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incidencias"
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   44
         Top             =   585
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   195
         Index           =   19
         Left            =   3825
         TabIndex        =   43
         Top             =   225
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   3960
      Left            =   60
      TabIndex        =   37
      Top             =   4515
      Width           =   7710
      Begin VB.CheckBox chkOpcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Repetición"
         Height          =   240
         Index           =   5
         Left            =   2070
         TabIndex        =   81
         Top             =   1800
         Width           =   1320
      End
      Begin VB.CheckBox chkFechaSolicitud 
         Caption         =   "Check1"
         Height          =   195
         Left            =   135
         TabIndex        =   73
         Top             =   1440
         Width           =   240
      End
      Begin VB.CheckBox chkOpcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestra No Rutinaria"
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   72
         Top             =   1800
         Width           =   2040
      End
      Begin VB.CheckBox chkFechaSolicitudNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Aplica"
         Height          =   195
         Left            =   4140
         TabIndex        =   68
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Frame frmTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   6075
         TabIndex        =   63
         Top             =   405
         Width           =   1500
         Begin VB.CheckBox chkOpcion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "A/C-Vuelo"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   67
            Top             =   585
            Width           =   1095
         End
         Begin VB.CheckBox chkOpcion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "In Situ"
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   66
            Top             =   900
            Width           =   1095
         End
         Begin VB.CheckBox chkOpcion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Lab.Movil"
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   65
            Top             =   1215
            Width           =   1095
         End
         Begin VB.CheckBox chkOpcion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Height          =   240
            Index           =   4
            Left            =   135
            TabIndex        =   64
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkFMSinEspecificar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin especificar"
         Height          =   240
         Left            =   2880
         TabIndex        =   59
         Top             =   270
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1110
         Index           =   8
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2415
         Width           =   7485
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   16
         Top             =   630
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   6
         Left            =   1260
         TabIndex        =   17
         Top             =   990
         Width           =   4680
         _ExtentX        =   8255
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
      Begin MSComCtl2.DTPicker DTfechaMuestreo 
         Height          =   345
         Left            =   1260
         TabIndex        =   15
         Top             =   225
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   61669377
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaSolicitud 
         Height          =   330
         Left            =   1530
         TabIndex        =   69
         Top             =   1395
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   61669377
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker horaSolicitud 
         Height          =   330
         Left            =   3015
         TabIndex        =   70
         Top             =   1395
         Width           =   1065
         _ExtentX        =   1879
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
         Format          =   61669378
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbCalibracionId 
         Height          =   330
         Left            =   1290
         TabIndex        =   75
         Top             =   3555
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   582
      End
      Begin VB.Label lblCalibracion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calibración"
         Height          =   195
         Left            =   135
         TabIndex        =   76
         Top             =   3600
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha solicitud"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   32
         Left            =   405
         TabIndex        =   71
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizado por"
         Height          =   195
         Index           =   16
         Left            =   135
         TabIndex        =   41
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   40
         Top             =   300
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   39
         Top             =   705
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   38
         Top             =   2115
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   4680
      Left            =   7785
      TabIndex        =   29
      Top             =   630
      Width           =   5925
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   10
         Top             =   1770
         Width           =   4710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   12
         Top             =   2565
         Width           =   4710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1365
         Index           =   5
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   3210
         Width           =   5745
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmRecepcion.frx":1194
         Height          =   315
         Index           =   4
         Left            =   1140
         TabIndex        =   9
         Top             =   1380
         Width           =   4710
         _ExtentX        =   8308
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
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   5
         Left            =   1140
         TabIndex        =   11
         Top             =   2160
         Width           =   4710
         _ExtentX        =   8308
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
      Begin MSComCtl2.DTPicker DT_fechaRecepcion 
         Height          =   330
         Left            =   1140
         TabIndex        =   19
         Top             =   210
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   61669377
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmRecepcion.frx":11DA
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   630
         Width           =   4170
         _ExtentX        =   7355
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
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmRecepcion.frx":1220
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   990
         Width           =   4710
         _ExtentX        =   8308
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   135
         TabIndex        =   74
         Top             =   1050
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   36
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionada por"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   35
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   34
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precinto"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   33
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entrega"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   32
         Top             =   2190
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   31
         Top             =   2610
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   30
         Top             =   2970
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   3570
      Left            =   45
      TabIndex        =   23
      Top             =   630
      Width           =   7680
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   90
         Picture         =   "frmRecepcion.frx":1266
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   80
         Top             =   3285
         Width           =   240
      End
      Begin VB.OptionButton opUrgente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1740
         TabIndex        =   78
         Top             =   3285
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton opUrgente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1050
         TabIndex        =   77
         Top             =   3285
         Width           =   615
      End
      Begin MSDataListLib.DataCombo cmbProducto 
         Height          =   315
         Left            =   1050
         TabIndex        =   6
         Top             =   2925
         Visible         =   0   'False
         Width           =   6450
         _ExtentX        =   11377
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
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   2
         Left            =   1050
         TabIndex        =   2
         Top             =   1485
         Width           =   6450
         _ExtentX        =   11377
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
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4860
         TabIndex        =   50
         Top             =   270
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1050
         TabIndex        =   3
         Top             =   1845
         Width           =   6435
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   4080
         TabIndex        =   14
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin MSDataListLib.DataCombo cmbPedidos 
         Bindings        =   "frmRecepcion.frx":7AB8
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   2205
         Width           =   5460
         _ExtentX        =   9631
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
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         Top             =   765
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbOferta 
         Height          =   375
         Left            =   1050
         TabIndex        =   5
         Top             =   2565
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTM 
         Height          =   330
         Left            =   1050
         TabIndex        =   1
         Top             =   1125
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Urgente"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   40
         Left            =   360
         TabIndex        =   79
         Top             =   3285
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Oferta"
         Height          =   195
         Index           =   21
         Left            =   135
         TabIndex        =   62
         Top             =   2610
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   6570
         Picture         =   "frmRecepcion.frx":7AFE
         Stretch         =   -1  'True
         Top             =   2205
         Width           =   255
      End
      Begin VB.Image imgPedidos 
         Height          =   300
         Left            =   6900
         Picture         =   "frmRecepcion.frx":83C8
         Stretch         =   -1  'True
         Top             =   2205
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   60
         Top             =   2970
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   2250
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Referencia"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1890
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   810
         Width           =   495
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de muestras"
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
         Left            =   2610
         TabIndex        =   26
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T.Muestra"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. análisis"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1530
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8550
      Width           =   1050
   End
   Begin VB.Label lblplantilla 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4695
      TabIndex        =   54
      Top             =   8640
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Alta de Muestras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   45
      TabIndex        =   53
      Top             =   0
      Width           =   13650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Datos del Registro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   46
      Top             =   360
      Width           =   7680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Otros datos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   7785
      TabIndex        =   49
      Top             =   5400
      Width           =   5925
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Recepción"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   7785
      TabIndex        =   48
      Top             =   360
      Width           =   5925
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Muestreo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   47
      Top             =   4245
      Width           =   7665
   End
End
Attribute VB_Name = "frmRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim muestra As New clsMuestra
Dim esBano As Boolean


Private Sub chkFechaSolicitud_Click()
    'M1105-I
    If chkFechaSolicitud.Value = Checked Then
        fechaSolicitud = Date
        horaSolicitud = Date & " " & Time
        fechaSolicitud.Enabled = True
        horaSolicitud.Enabled = True
        chkFechaSolicitudNA.Value = Unchecked
    Else
        fechaSolicitud.Enabled = False
        horaSolicitud.Enabled = False
    End If
    'M1105-F
End Sub

Private Sub chkFechaSolicitudNA_Click()
    'M1105-I
    If chkFechaSolicitudNA.Value = Checked Then
        fechaSolicitud.Enabled = False
        horaSolicitud.Enabled = False
        chkFechaSolicitud.Enabled = False
        chkFechaSolicitud.Value = Unchecked
    Else
        chkFechaSolicitud.Enabled = True
    End If
    'M1105-F
End Sub

Private Sub chkFMSinEspecificar_Click()
    If chkFMSinEspecificar.Value = Checked Then
        DTfechaMuestreo.Enabled = False
    Else
        DTfechaMuestreo.Enabled = True
    End If
End Sub

Private Sub cmbClientes_change()
    cmbPedidos.Text = ""
    cargar_pedidos
    'M1009-I
    cmbOferta.limpiar
    If cmbClientes.getTEXTO <> "" Then
        llenar_combo cmbOferta, New clsOfertas, 0, frmOferta_Nueva2, " AND O.CLIENTE_ID = " & cmbClientes.getPK_SALIDA
        cmbOferta.activar
    Else
        cmbOferta.desactivar
    End If
    'M1009-F
End Sub

Private Sub cmbDatos_Change(Index As Integer)
  On Error GoTo fallo:
  Select Case Index
'M1105-I
'  Case 1: 'muestra los analisis para el tipo de muestra seleccionado
'       Dim tipo As New clsMuestra
'       lblCampos(20).Visible = False
'       cmbProducto.Visible = False
'       cmbProducto.Enabled = False
'       cmbProducto.Text = ""
'
'        If cmbDatos(1).Text <> "" And IsNumeric(cmbDatos(1).BoundText) Then
'        If Not tipo.esBano(cmbDatos(1).BoundText) Then  'es un baño Id_Espacial
'        Dim oAnalisis As New clsTipos_analisis
'         ' Es una determinacion de muestra
'         If oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText).RecordCount <> 0 Then 'existe un registro al menos
'            Set cmbDatos(2).RowSource = oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText) 'comboanalisis
'            cmbDatos(2).ListField = "nombre" 'lo que enseña
'            cmbDatos(2).DataField = "id_tipo_analisis" 'campo asociado
'            cmbDatos(2).BoundColumn = "id_tipo_analisis" 'lo que realmente
'            cmbDatos(2).Text = oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText).Fields("nombre").value
'            cmbDatos(2).Enabled = True
'            cmdDeterminaciones.Enabled = True
'            Dim oTA As New clsTipos_analisis
'            oTA.CARGAR cmbDatos(2).BoundText
'            txtdias = oTA.getDIAS_TRABAJO
'            carcularFechaEntrega
'         Else 'si no recupero ningun registro
'            txtdias = "0"
'            cmbDatos(2).Text = ""
'            cmbDatos(2).Enabled = False
'            cmdDeterminaciones.Enabled = False
'         End If
'            esBano = False
'        Else
'         ' Es un baño especial
'            If cmbClientes.getPK_SALIDA = 0 Then
'                MsgBox "Seleccione primero un cliente.", vbInformation, App.Title
'                cmbClientes.Limpiar
'                cmbClientes.SetFocus
'                Exit Sub
'            End If
'            cmbDatos(2).Text = ""
'            cmbDatos(2).Enabled = False
'            cmdDeterminaciones.Enabled = False
'            Dim oBANO As New clsBanos
'            Dim rsbano As New ADODB.Recordset
'            Set rsbano = oBANO.banos_cliente(cmbClientes.getPK_SALIDA, cmbDatos(1).BoundText)
'            If rsbano.RecordCount = 0 Then
'                MsgBox "No hay baños para el cliente y tipo de muestra seleccionado.", vbInformation, App.Title
'                cmbDatos(1).Text = ""
'                Exit Sub
''            ElseIf rsbano.RecordCount < val(Text1(0)) Then
''                MsgBox "El número de baños para el cliente y tipo de Muestra " & _
''                       "seleccionados es menor que el número de muestras indicado.", vbInformation, App.Title
''                cmbDatos(1).Text = ""
''                Exit Sub
'            End If
'            esBano = True
'            Set oBANO = Nothing
'        End If 'fin del esBanno
'        calcular_precio_analisis
'        Set tipo = Nothing
'        Set oAnalisis = Nothing
'       End If
'       ' Pedidos
''       If Index = 0 Then
''            cmbPedidos.Text = ""
''            If cmbClientes.getPK_SALIDA <> 0 Then
''                pedidos (cmbClientes.getPK_SALIDA)
''            End If
''       End If
'       ' Descripción del producto
'       If cmbDatos(1).Text <> "" Then
'        Dim otm As New clsTipos_muestra
'        If otm.CARGAR(cmbDatos(1).BoundText) = True Then
'         If otm.getREQUIERE_PRODUCTO = 1 Then
'             lblCampos(20).Visible = True
'             cmbProducto.Visible = True
'             cmbProducto.Enabled = True
'             cargar_producto
''             Text1(2).Visible = True
''             Text1(2).Enabled = True
''         Else
''             lblCampos(20).Visible = False
''             cmbProducto.Visible = False
''             cmbProducto.Enabled = False
''             Text1(2).Visible = False
''             Text1(2).Enabled = False
'         End If
'        End If
'       End If
'M1105-F
  Case 2: 'tipo de analisis
    calcular_precio_analisis
  Case 4:
    Dim formato As New clsformatos
    If cmbDatos(4).Text <> "" Then
        If formato.EsPrecintado(cmbDatos(4).BoundText) Then
            Text1(3).Enabled = True
            Text1(3).SetFocus
        Else
            Text1(3).Enabled = False
        End If
    End If
    Set formato = Nothing
  End Select
  cmbDatos(2).ToolTipText = cmbDatos(2).Text
  Exit Sub
fallo:
    MsgBox "Error al decodificar los campos", vbCritical, Err.Description
End Sub

Private Sub cmbTM_Change()
    Dim tipo As New clsMuestra
    lblCampos(20).visible = False
    cmbProducto.visible = False
    cmbProducto.Enabled = False
    cmbProducto.Text = ""

    If cmbTM.getTEXTO <> "" Then
        If Not tipo.esBano(cmbTM.getPK_SALIDA) Then  'es un baño Id_Espacial
            Dim oAnalisis As New clsTipos_analisis
            ' Es una determinacion de muestra
            If oAnalisis.AnalisisAsociadosMuestra(cmbTM.getPK_SALIDA).RecordCount <> 0 Then 'existe un registro al menos
               Set cmbDatos(2).RowSource = oAnalisis.AnalisisAsociadosMuestra(cmbTM.getPK_SALIDA) 'comboanalisis
               cmbDatos(2).ListField = "nombre" 'lo que enseña
               cmbDatos(2).DataField = "id_tipo_analisis" 'campo asociado
               cmbDatos(2).BoundColumn = "id_tipo_analisis" 'lo que realmente
               cmbDatos(2).Text = oAnalisis.AnalisisAsociadosMuestra(cmbTM.getPK_SALIDA).Fields("nombre").Value
               cmbDatos(2).Enabled = True
               cmdDeterminaciones.Enabled = True
               Dim oTA As New clsTipos_analisis
               oTA.CARGAR cmbDatos(2).BoundText
               txtdias = oTA.getDIAS_TRABAJO
               carcularFechaEntrega
            Else 'si no recupero ningun registro
               txtdias = "0"
               cmbDatos(2).Text = ""
               cmbDatos(2).Enabled = False
               cmdDeterminaciones.Enabled = False
            End If
            esBano = False
        Else
            ' Es un baño especial
            If cmbClientes.getPK_SALIDA = 0 Then
                MsgBox "Seleccione primero un cliente.", vbInformation, App.Title
                cmbClientes.limpiar
                cmbClientes.SetFocus
                Exit Sub
            End If
            cmbDatos(2).Text = ""
            cmbDatos(2).Enabled = False
            cmdDeterminaciones.Enabled = False
            Dim oBANO As New clsBanos
            Dim rsbano As New ADODB.Recordset
            Set rsbano = oBANO.banos_cliente(cmbClientes.getPK_SALIDA, cmbTM.getPK_SALIDA)
            If rsbano.RecordCount = 0 Then
                MsgBox "No hay baños para el cliente y tipo de muestra seleccionado.", vbInformation, App.Title
                cmbTM.limpiar
                Exit Sub
            End If
            esBano = True
            Set oBANO = Nothing
        End If 'fin del esBanno
        calcular_precio_analisis
        Set tipo = Nothing
        Set oAnalisis = Nothing
        Dim otm As New clsTipos_muestra
        If otm.CARGAR(cmbTM.getPK_SALIDA) = True Then
         If otm.getREQUIERE_PRODUCTO = 1 Then
             lblCampos(20).visible = True
             cmbProducto.visible = True
             cmbProducto.Enabled = True
             cargar_producto
         End If
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    ReDim plantilla_bano(1)
    plantilla_bano(1) = 0
    Me.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdMas_Click()
    Text1(0) = val(Text1(0)) + 1
End Sub

Private Sub cmdNuevoCliente_Click()
    frmClientes.PK = 0
    frmClientes.Show 1
    cargar_clientes
End Sub

Private Sub cmdok_Click()
    If validar_datos = False Then
        Exit Sub
    End If
    Me.MousePointer = 11
    If insertar_muestra Then
        Me.MousePointer = 0
        If esBano Then
            frmRecepcion_Multiple.CLIENTE_BANO = cmbClientes.getPK_SALIDA
            'M1105-I
            'frmRecepcion_Multiple.TIPO_ANALISIS_BANO = cmbDatos(1).BoundText
            frmRecepcion_Multiple.TIPO_ANALISIS_BANO = cmbTM.getPK_SALIDA
            'M1105-F
        Else
            frmRecepcion_Multiple.CLIENTE_BANO = 0
            frmRecepcion_Multiple.TIPO_ANALISIS_BANO = 0
        End If
'        frmRecepcion_Multiple.Show 1
'        cmdcancel_Click
        Me.MousePointer = 0
        frmRecepcion_Multiple.Show
        Unload Me
    Else
        Me.MousePointer = 0
    End If
End Sub

Private Sub DT_fechaRecepcion_Change()
    buscar_ultimo_codigo
    carcularFechaEntrega
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 25
    Me.Left = 25
    cargar_clientes
    cargar_muestras
    cargar_centros
    cargar_realizadas
    cargar_entregada
    cargar_envases
    cargar_pedidos
    permisos
    cargar_combo cmbDatos(0), New clsUsuarios
    cmbDatos(0).BoundText = USUARIO.getID_EMPLEADO
    cmbDatos(2).Enabled = False
    DTfechaMuestreo.Value = Now
    DT_fechaRecepcion = Now
    DTfechaPrevistaEntrega = Now
    lblplantilla.visible = False
    If PLANTILLA <> 0 Then
        cargar_plantilla
        PLANTILLA = 0
    Else
        ReDim plantilla_bano(1)
        plantilla_bano(1) = 0
        'M1009-I
        cmbOferta.desactivar
        'M1009-F
    End If
    buscar_ultimo_codigo
End Sub

Private Sub Image1_Click()
    cmbPedidos.Text = ""
    cmbPedidos.BoundText = ""
End Sub

Private Sub imgPedidos_Click()
    If cmbClientes.getTEXTO <> "" Then
        cmbPedidos.Text = ""
        frmClientes_Pedidos.PK = cmbClientes.getPK_SALIDA
        frmClientes_Pedidos.Show 1
        cargar_pedidos
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = &H80C0FF
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
    Select Case Index
        Case 11: 'incidencias
          ' Poner hora y dia antes de empezarescribir
          If Text1(11).Text = "" Then
            Text1(Index) = "Dia: " & Date & " Hora: " & Time & vbCrLf
            Text1(Index).SelStart = Len(Text1(Index)) + 1
          End If
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    Dim caracter As String
'    If Index <> 6 Then
'        caracter = Chr(KeyAscii) 'devuelve el numero del caracter
'        caracter = UCase(caracter) 'conviertemayusculas todo el texto
'        KeyAscii = Asc(caracter) 'da el valor de la tecla pulsadakeyascii
'    End If
    ' Escribir ',' al pulsar '.'
    If Index = 10 And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &HFFFFFF
    If Index = 10 And Text1(Index) <> "" And Not IsNumeric(Text1(Index)) Then
        MsgBox "El precio debe ser numérico.", vbCritical, "Formato"
        Text1(10).SetFocus
    Else
        Text1(10) = moneda(Text1(10))
    End If
    If Index = 1 Then
        Text1(1) = Replace(Text1(1).Text, """", " ")
    End If
End Sub
Private Sub cargar_clientes()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
End Sub

Private Sub cargar_muestras()
    llenar_combo cmbTM, New clsTipos_muestra, 0, frmTM_Detalle, " ANULADO = 0 "
End Sub
Private Sub cargar_centros()
    cargar_combo cmbCentro, New clsCentros
End Sub

Public Sub cargar_realizadas()
    Dim oenti As New clsEntidades_muestreo
    Set cmbDatos(6).RowSource = oenti.Listado
    cmbDatos(6).ListField = "descripcion"
    cmbDatos(6).DataField = "id_entidad_muestreo" 'campo asociado
    cmbDatos(6).BoundColumn = "id_entidad_muestreo" 'lo que realmente
    Set oenti = Nothing
End Sub

Private Sub cargar_entregada()
    Dim oenti As New clsEntidades_Entrega
    Set cmbDatos(5).RowSource = oenti.Listado
    cmbDatos(5).ListField = "descripcion"
    cmbDatos(5).DataField = "id_entidad_entrega" 'campo asociado
    cmbDatos(5).BoundColumn = "id_entidad_entrega" 'lo que realmente
    Set oenti = Nothing
End Sub

Private Sub cargar_envases()
    Dim oFormato As New clsformatos
    Set cmbDatos(4).RowSource = oFormato.Listado
    cmbDatos(4).ListField = "descripcion"
    cmbDatos(4).DataField = "id_formato" 'campo asociado
    cmbDatos(4).BoundColumn = "id_formato" 'lo que realmente
    Set oFormato = Nothing
End Sub
Private Sub cargar_pedidos()
    Dim oPedido As New clsClientes_pedidos
    If cmbClientes.getTEXTO <> "" Then
    '    Dim anterior As Integer
    '    If cmbPedidos.Text <> "" Then
    '        anterior = cmbPedidos.BoundText
    '    End If
    '    If ID = 0 Then
    '        Set cmbPedidos.RowSource = oPedido.Listado_completo
    '    Else
            Set cmbPedidos.RowSource = oPedido.Listado_en_fecha(cmbClientes.getPK_SALIDA, DT_fechaRecepcion.Value)
    '    End If
        cmbPedidos.ListField = "CODIGO_LARGO"
        cmbPedidos.DataField = "id_pedido"
        cmbPedidos.BoundColumn = "id_pedido"
    '    cmbPedidos.BoundText = anterior
    End If
End Sub

Private Sub calcular_precio_analisis()
    Dim aux As New clsTipos_analisis
    If cmbDatos(2).BoundText = "" Then
        Text1(10).Text = ""
    Else
        Text1(10) = moneda(aux.PrecioDelAnalisis(cmbDatos(2).BoundText))
    End If
    Set aux = Nothing
End Sub

Private Function insertar_muestra() As Boolean
    On Error GoTo fallo
    Dim i As Integer
    Dim nmuestra As Long
    ReDim muestras(val(Text1(0)))
    If mover_datos_muestra = True Then
        For i = 1 To val(Text1(0))
            nmuestra = muestra.guardarMuestra()
            If nmuestra = 0 Then
                MsgBox "Error al insertar la muestra " & CStr(i) & vbCrLf & Err.Description, vbCritical + vbOKOnly
                Exit Function
            Else
                muestras(i) = nmuestra
                ' Calibracion - Muestra
                If cmbCalibracionId.visible = True Then
                    Dim oC As New clsEquipoCalibracion
                    oC.informarMuestraId cmbCalibracionId.getPK_SALIDA, nmuestra
                    Set oC = Nothing
                End If
            End If
        Next
        If val(Text1(0)) = 1 Then
            MsgBox "La Muestra ha sido registrada con el Nº: " & muestra.getID_GENERAL & " y código: " & muestra.CodigoParticular(muestra.getID_MUESTRA), vbInformation, App.Title
        Else
            MsgBox "Las Muestras han sido registradas correctamente.", vbInformation, App.Title
        End If
        insertar_muestra = True
    Else
        MsgBox "Error al informar los datos de la muestra. No es posible la recepción. Intentelo de nuevo.", vbCritical, App.Title
        insertar_muestra = False
    End If
    Exit Function
fallo:
    MsgBox "Error al insertar la muestra." & vbCrLf & Err.Description, vbCritical + vbOKOnly, App.Title
End Function

Public Function validar_datos() As Boolean
    validar_datos = True
    If IsNumeric(Text1(0)) = False Then
        MsgBox "El número de muestras debe ser numérico.", vbExclamation, "Validación"
        Text1(0).SetFocus
        validar_datos = False
        Exit Function
    End If
    If cmbClientes.getPK_SALIDA = 0 Then
        MsgBox "El Cliente no puede estar en blanco.", vbExclamation, "Validación"
        cmbClientes.SetFocus
        validar_datos = False
        Exit Function
    End If
    'M1105 : cmbTM
    'If cmbDatos(1).Text = "" Then
    If cmbTM.getTEXTO = "" Then
        MsgBox "El tipo de muestra no puede estar en blanco.", vbExclamation, "Validación"
        'cmbDatos(1).SetFocus
        cmbTM.SetFocus
        validar_datos = False
        Exit Function
    End If
'    If IsNumeric(cmbDatos(1).BoundText) = False Then
'        MsgBox "El tipo de muestra no esta seleccionado correctamente.", vbExclamation, "Validación"
'        cmbDatos(1).SetFocus
'        validar_datos = False
'        Exit Function
'    End If
    If IsNumeric(cmbDatos(0).BoundText) = False Then
        MsgBox "El usuario de recepcion no esta seleccionado correctamente.", vbExclamation, "Validación"
        cmbDatos(0).SetFocus
        validar_datos = False
        Exit Function
    End If
    Dim tipo As New clsMuestra
'M1105    If tipo.esBano(cmbDatos(1).BoundText) = False Then 'es un baño Id_Espacial
    If tipo.esBano(cmbTM.getPK_SALIDA) = False Then  'es un baño Id_Espacial
        If Text1(1) = "" Then
            MsgBox "La referencia de la muestra no puede estar en blanco.", vbExclamation, "Validación"
            Text1(1).SetFocus
            validar_datos = False
            Exit Function
        End If
        ' Descripción del producto
        If cmbProducto.visible = True Then
            If cmbProducto.Text = "" Then
                MsgBox "El producto no puede estar en blanco.", vbExclamation, "Validación"
                cmbProducto.SetFocus
                validar_datos = False
                Exit Function
            End If
        End If
    End If
    If cmbCentro.Text = "" Then
        MsgBox "El CENTRO no puede estar en blanco.", vbExclamation, "Validación"
        cmbCentro.SetFocus
        validar_datos = False
        Exit Function
    End If
    If cmbDatos(4).Text = "" Then
        MsgBox "El envase de la muestra no puede estar en blanco.", vbExclamation, "Validación"
        cmbDatos(4).SetFocus
        validar_datos = False
        Exit Function
    End If
    'M1105-I
    If chkOpcion(0).Value = Unchecked And chkOpcion(1).Value = Unchecked And chkOpcion(2).Value = Unchecked And chkOpcion(4).Value = Unchecked Then
        MsgBox "Debe indicar uno de los Tipos en el Muestreo.", vbExclamation, "Validación"
        validar_datos = False
        Exit Function
    End If
    If chkFechaSolicitud.Value = Unchecked And chkFechaSolicitudNA.Value = Unchecked Then
        MsgBox "Debe indicar la Fecha de la Solicitud o No aplica.", vbExclamation, "Validación"
        validar_datos = False
        Exit Function
    End If
    'M1105-F
    If cmbCalibracionId.visible = True Then
        If cmbCalibracionId.getTEXTO = "" Then
            MsgBox "Debe indicar a que Calibración pertenece.", vbExclamation, "Validación"
            validar_datos = False
            Exit Function
        End If
        Dim oC As New clsEquipoCalibracion
        oC.Carga cmbCalibracionId.getPK_SALIDA
        If oC.getMUESTRA_ID <> 0 Then
            Dim oMuestra As New clsMuestra
            MsgBox "La CALIBRACIÓN ya tiene muestra asignada : " & oMuestra.CodigoParticular(oC.getMUESTRA_ID), vbExclamation, "Validación"
            validar_datos = False
            Exit Function
        End If
        Set oC = Nothing
        
    End If
    
End Function

Public Sub permisos()
    ' Proteger campo precio
    If USUARIO.getPER_FACTURACION = False Then
        Text1(10).Locked = True
        Text1(10).visible = False
        lblCampos(19).visible = False
    End If
End Sub
Private Function mover_datos_muestra() As Boolean
    mover_datos_muestra = False
    With muestra
'M1105        .setTIPO_MUESTRA_ID = cmbDatos(1).BoundText
        .setTIPO_MUESTRA_ID = cmbTM.getPK_SALIDA
        If cmbDatos(2).Text <> "" Then ' Es un baño
            .setTIPO_ANALISIS_ID = cmbDatos(2).BoundText
        Else
            .setTIPO_ANALISIS_ID = 0
        End If
        .setANALISIS_MODIFICADO = 0
        If chkFMSinEspecificar.Value = Checked Then
            .setFECHA_MUESTREO = "1900-01-01"
        Else
            .setFECHA_MUESTREO = Format(DTfechaMuestreo, "yyyy-mm-dd")
        End If
        ' 2008-04-04 Si no se marca muestreo, por defecto se pone Canagrosa
        If val(cmbDatos(6).BoundText) = 0 Then
            .setENTIDAD_MUESTREO_ID = 2
        Else
            .setENTIDAD_MUESTREO_ID = val(cmbDatos(6).BoundText)
        End If
        .setDETALLE_MUESTREO = Text1(7)
        .setOBSERVACIONES_MUESTREO = Text1(8)
        .setFECHA_RECEPCION = Format(DT_fechaRecepcion, "yyyy-mm-dd")
'        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setEMPLEADO_ID = cmbDatos(0).BoundText
        If val(cmbDatos(4).BoundText) = 0 Then
            .setFORMATO_ID = 1
        Else
            .setFORMATO_ID = val(cmbDatos(4).BoundText)
        End If
        If val(cmbDatos(5).BoundText) = 0 Then
            .setENTIDAD_ENTREGA_ID = 2
        Else
            .setENTIDAD_ENTREGA_ID = val(cmbDatos(5).BoundText)
        End If
        .setDETALLE_ENTREGA = Text1(4)
        .setOBSERVACIONES_ENTREGA = Text1(5)
        .setCLIENTE_ID = cmbClientes.getPK_SALIDA
        .setCENTRO_ID = cmbCentro.BoundText
        .setREFERENCIA_CLIENTE = Text1(1)
        .setPRECIO = moneda_bd(Text1(10))
        .setFECHA_PREV_FIN = Format(DTfechaPrevistaEntrega, "yyyy-mm-dd")
        .setOBSERVACIONES = Text1(11)
        .setANULADA = 0
        .setPRECINTO = Text1(3)
        .setBANO_ID = 0
'J51        .setFECHA_COMIENZO = 0
'J51        .setFECHA_CIERRE = 0
        .setFECHA_COMIENZO = "0000-00-00"
        .setFECHA_FINALIZACION = "0000-00-00"
        .setFECHA_CIERRE = "0000-00-00"
        .setCERRADA = 0
        .setDOCUMENTO_PAGO = 0
        .setULT_EDICION_IMP = 0
        If cmbPedidos.Text = "" Then
            .setPEDIDO_ID = 0
        Else
            .setPEDIDO_ID = cmbPedidos.BoundText
        End If
        'M1009-I
        If cmbOferta.getTEXTO = "" Then
            .setOFERTA_ID = 0
        Else
            .setOFERTA_ID = cmbOferta.getPK_SALIDA
        End If
        'M1009-F
'        If Text1(2).Visible = True Then
'            .setPRODUCTO = Text1(2)
'        End If
        If cmbProducto.visible = True Then
            .setPRODUCTO = cmbProducto.Text
        End If
        'M1105-I
        .setOP_VUELO = chkOpcion(0).Value
        .setOP_INSITU = chkOpcion(1).Value
        .setOP_LABMOVIL = chkOpcion(2).Value
        .setOP_NORUTINARIA = chkOpcion(3).Value
        .setOP_REPETICION = chkOpcion(5).Value
        ' INDICADORES
        If chkFechaSolicitud.Value = Checked Then
            .setFECHA_RECOGIDA = Format(fechaSolicitud, "yyyy-mm-dd") & " " & Format(horaSolicitud, "hh:mm:ss")
        Else
            .setFECHA_RECOGIDA = ""
        End If
        'M1105-F
        .setREPLACEMENT_ID = 0
        If opUrgente(0).Value = True Then
            .setURGENTE = 0
        Else
            .setURGENTE = 1
        End If
    End With
    mover_datos_muestra = True
End Function

Public Sub cargar_plantilla()
    Dim oplantilla As New clsPlantillas_muestras
    With oplantilla
         .CargaPlantilla (PLANTILLA)
         If UBound(plantilla_bano, 1) > 0 Then
            Text1(0) = UBound(plantilla_bano, 1)
         Else
            Text1(0) = .getCANTIDAD_MUESTRAS
         End If
         If .getCLIENTE_ID > 0 Then
            cmbClientes.MostrarElemento .getCLIENTE_ID
         End If
         If .getTIPO_MUESTRA_ID > 0 Then
'M1105            cmbDatos(1).BoundText = .getTIPO_MUESTRA_ID
            cmbTM.MostrarElemento .getTIPO_MUESTRA_ID
         End If
         If .getTIPO_ANALISIS_ID > 0 Then
            cmbDatos(2).BoundText = .getTIPO_ANALISIS_ID
         End If
         Text1(1) = .getREFERENCIA_CLIENTE
         Text1(7) = .getDETALLE_MUESTREO
         If .getENTIDAD_MUESTREO_ID > 0 Then
            cmbDatos(6).BoundText = .getENTIDAD_MUESTREO_ID
         End If
         Text1(8) = .getOBSERVACIONES_MUESTREO
         If .getFORMATO_ID > 0 Then
            cmbDatos(4).BoundText = .getFORMATO_ID
         End If
         If .getENTIDAD_ENTREGA_ID > 0 Then
            cmbDatos(6).BoundText = .getENTIDAD_ENTREGA_ID
         End If
         Text1(4) = .getDETALLE_ENTREGA
         Text1(5) = .getOBSERVACIONES_ENTREGA
         Text1(10) = moneda(.getPRECIO)
         Text1(11) = .getOBSERVACIONES
         lblplantilla.visible = True
         lblplantilla.Caption = "Plantilla : " & .getNOMBRE
    End With
End Sub

Private Sub buscar_ultimo_codigo()
    Dim oMuestra As New clsMuestra
    Label1(3).Caption = "Alta de Muestra (Próximo código general : " & Trim(str(oMuestra.buscar_ultimo_codigo_general(Year(DT_fechaRecepcion.Value)))) & ")"
End Sub

Private Sub cargar_producto()
'    If cmbDatos(1).Text <> "" And cmbProducto.Visible = True Then
'    If cmbTM.getTEXTO <> "" And cmbProducto.Visible = True Then
    If cmbTM.getTEXTO <> "" Then
        ' Si es del tipo indicado, tomar la descripcion de la solución del baño, insertarla en la decodificadora
'        If cmbTM.getPK_SALIDA = TIPOS_MUESTRAS.TM_AGUA Or _
'           cmbTM.getPK_SALIDA = TIPOS_MUESTRAS.TM_BANO Or _
'           cmbTM.getPK_SALIDA = TIPOS_MUESTRAS.TM_FLUIDO Or _
'           cmbTM.getPK_SALIDA = TIPOS_MUESTRAS.COMBUSTIBLE Or _
'           cmbTM.getPK_SALIDA = TIPOS_MUESTRAS.TM_COMBUSTIBLE_AGRUPADO Or _
'           cmbTM.getPK_SALIDA = TIPOS_MUESTRAS.TM_ACEITE_MOTOR Then
'            If cmbDatos(2).BoundText <> "" Then
'                ' Cargar la solucion del baño
'                Dim oSolucion As New clsSoluciones
'                Dim oBano As New clsBanos
'                If oBano.cargar_bano(cmbDatos(2).BoundText) = True Then
'                    If oSolucion.CARGAR(oBano.getID_SOLUCION) Then
'                        Dim oDeco As New clsDecodificadora
'                        With oDeco
'                            .setCODIGO = DECODIFICADORA.DESCRIPCION_PRODUCTO
'                            .setDESCRIPCION = oSolucion.getNOMBRE
'                            .setPARAMETROS = CStr(cmbTM.getPK_SALIDA)
'                            .Insertar
'                        End With
'                        Set oDeco = Nothing
'                    End If
'                End If
'            End If
'        End If
        ' Recuperar la descripcion del producto
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT VALOR, DESCRIPCION " & _
                   "  FROM decodificadora " & _
                   " WHERE CODIGO = " & DECODIFICADORA.DESCRIPCION_PRODUCTO & _
                   "   AND PARAMETROS = '" & CInt(cmbTM.getPK_SALIDA) & "'"
        Set rs = datos_bd(consulta)
        Set cmbProducto.RowSource = rs
        cmbProducto.ListField = "DESCRIPCION" 'lo que enseña
        cmbProducto.DataField = "VALOR" 'campo asociado
        cmbProducto.BoundColumn = "VALOR" 'lo que realmente
    End If
End Sub
Public Sub cargarCalibraciones(ID_EQUIPO As Long)
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT a.ID_CALIBRACION, concat(date_format(a.FECHA_ACTUAL,'%d-%m-%Y'),' -> ',tipo_cal.descripcion,' -> ',COALESCE(period.descripcion, ''),' -> ',COALESCE(CONCAT(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '')) AS DESCRIPCION " & _
                   "  FROM eq_calibracion_equipos a " & _
                   " INNER JOIN decodificadora tipo_cal ON a.TIPO_ID = tipo_cal.VALOR AND tipo_cal.CODIGO = 104 " & _
                   "  LEFT OUTER JOIN eq_periodicidad period ON a.PERIODICIDAD_ID = period.ID_PERIODICIDAD  " & _
                   "  LEFT OUTER JOIN proveedores calibrador_ext ON a.CALIBRADOR_EXTERNO_ID = calibrador_ext.ID_PROVEEDOR  " & _
                   "  LEFT OUTER JOIN usuarios ON a.CALIBRADOR_interno_id = usuarios.ID_EMPLEADO " & _
                   " WHERE a.EQUIPO_ID = " & ID_EQUIPO & " And a.ESTADO = 0 "

        With cmbCalibracionId
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "eq_calibracion_equipos"
            .setDESCRIPCION = "Calibraciones"
            .setPK = "ID_CALIBRACION"
            .setCAMPO = "DESCRIPCION"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = False
            Set .FORMULARIO = Me
        End With
    End If
End Sub
Private Sub carcularFechaEntrega()
'    Dim FechaEntrega As Date
'    FechaEntrega = DateAdd("d", CInt(txtdias), CDate(DT_fechaRecepcion.value))
'    DTfechaPrevistaEntrega = Format(FechaEntrega, "yyyy-mm-dd")
   On Error GoTo carcularFechaEntrega_Error

    DTfechaPrevistaEntrega = calcularFechaFinalizacion(DT_fechaRecepcion.Value, CInt(txtdias))

   On Error GoTo 0
   Exit Sub

carcularFechaEntrega_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure carcularFechaEntrega of Formulario frmRecepcion"
End Sub
