VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmEquipoEdicionVerificacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11070
   ClientLeft      =   2955
   ClientTop       =   2490
   ClientWidth     =   12720
   ClipControls    =   0   'False
   Icon            =   "frmEquipoEdicionVerificacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRevisiones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Revisiones"
      Height          =   900
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   10125
      Width           =   1530
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   900
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   10125
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3375
      Top             =   10395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10155
      Width           =   1050
   End
   Begin VB.Frame Frame1 
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
      Height          =   9570
      Left            =   90
      TabIndex        =   36
      Top             =   540
      Width           =   12615
      Begin VB.TextBox txtCalibradoEn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   7110
         MaxLength       =   255
         TabIndex        =   92
         Top             =   3555
         Width           =   5175
      End
      Begin VB.Frame fraEstadoIntervencion 
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
         Height          =   1515
         Left            =   8955
         TabIndex        =   85
         Top             =   1530
         Width           =   1500
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Revisada"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   89
            Top             =   840
            Width           =   1065
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Realizada"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   88
            Top             =   570
            Width           =   1020
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prevista"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   87
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Anulada"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   86
            Top             =   1125
            Width           =   1140
         End
      End
      Begin VB.Frame frmResultado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   10530
         TabIndex        =   81
         Top             =   1530
         Width           =   1950
         Begin VB.OptionButton optResultado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CONFORME"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   84
            Top             =   315
            Value           =   -1  'True
            Width           =   1410
         End
         Begin VB.OptionButton optResultado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO CONFORME"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   83
            Top             =   705
            Width           =   1650
         End
         Begin VB.OptionButton optResultado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "REQ. AJUSTE"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   82
            Top             =   1065
            Width           =   1560
         End
      End
      Begin VB.TextBox txtEquipoCopia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   855
         TabIndex        =   77
         Top             =   6930
         Width           =   1140
      End
      Begin VB.Frame frmReactivos 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   5985
         TabIndex        =   66
         Top             =   7335
         Width           =   6495
         Begin VB.CommandButton cmdAnadirReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   750
            Left            =   5625
            Picture         =   "frmEquipoEdicionVerificacion.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   68
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1080
            Width           =   780
         End
         Begin VB.CommandButton cmdEliminarReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   750
            Left            =   5625
            Picture         =   "frmEquipoEdicionVerificacion.frx":08D6
            Style           =   1  'Graphical
            TabIndex        =   67
            Tag             =   "Elimina el campo seleccionado"
            Top             =   270
            Width           =   780
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   1035
            Left            =   135
            TabIndex        =   69
            Top             =   270
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   1826
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
         Begin pryCombo.miCombo cmbreactivos 
            Height          =   330
            Left            =   810
            TabIndex        =   70
            Top             =   1395
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   810
            TabIndex        =   71
            Top             =   1755
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externo"
            Height          =   195
            Index           =   16
            Left            =   135
            TabIndex        =   73
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Interno"
            Height          =   195
            Index           =   15
            Left            =   135
            TabIndex        =   72
            Top             =   1800
            Width           =   495
         End
      End
      Begin VB.Frame frmEquipos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Equipos utilizados"
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
         Height          =   2175
         Left            =   45
         TabIndex        =   61
         Top             =   7335
         Width           =   5910
         Begin VB.CommandButton cmdAnadirEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            Height          =   720
            Left            =   5040
            Picture         =   "frmEquipoEdicionVerificacion.frx":11A0
            Style           =   1  'Graphical
            TabIndex        =   63
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1080
            Width           =   780
         End
         Begin VB.CommandButton cmdEliminarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   720
            Left            =   5040
            Picture         =   "frmEquipoEdicionVerificacion.frx":1A6A
            Style           =   1  'Graphical
            TabIndex        =   62
            Tag             =   "Elimina el campo seleccionado"
            Top             =   315
            Width           =   780
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   1290
            Left            =   45
            TabIndex        =   64
            Top             =   315
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   2275
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
         Begin pryCombo.miCombo cmbequipos 
            Height          =   330
            Left            =   45
            TabIndex        =   65
            Top             =   1710
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   582
         End
      End
      Begin VB.CommandButton cmdAnadirParametro 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6885
         Picture         =   "frmEquipoEdicionVerificacion.frx":2334
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Añadir accesorio"
         Top             =   6930
         Width           =   285
      End
      Begin VB.CommandButton cmdEliminarParametro 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   7215
         Picture         =   "frmEquipoEdicionVerificacion.frx":2559
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Eliminar accesorio"
         Top             =   6930
         Width           =   285
      End
      Begin VB.TextBox txtFechaProxima 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10710
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "01/01/1900"
         Top             =   630
         Width           =   1785
      End
      Begin VB.TextBox txtAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   15
         Top             =   2775
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   7215
         Picture         =   "frmEquipoEdicionVerificacion.frx":26ED
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ver norma"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarP 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion.frx":2942
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Buscar documento"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion.frx":2BB3
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Eliminar documento"
         Top             =   2760
         Width           =   360
      End
      Begin VB.CommandButton cmdEscanearEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7650
         Picture         =   "frmEquipoEdicionVerificacion.frx":2D47
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Escanear documento"
         Top             =   2760
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdEscanearHoja 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7650
         Picture         =   "frmEquipoEdicionVerificacion.frx":3101
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Escanear documento"
         Top             =   2040
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdEscanearCert 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7650
         Picture         =   "frmEquipoEdicionVerificacion.frx":34BB
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Escanear documento"
         Top             =   2400
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5580
         Picture         =   "frmEquipoEdicionVerificacion.frx":3875
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Eliminar accesorio"
         Top             =   3195
         Width           =   285
      End
      Begin VB.CommandButton cmdAnadirLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5250
         Picture         =   "frmEquipoEdicionVerificacion.frx":3A09
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Añadir accesorio"
         Top             =   3195
         Width           =   285
      End
      Begin VB.TextBox txtLimitacionesUso 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1650
         MaxLength       =   100
         TabIndex        =   20
         Top             =   3165
         Width           =   3555
      End
      Begin VB.ListBox lstLimitacionesUso 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "frmEquipoEdicionVerificacion.frx":3C2E
         Left            =   1650
         List            =   "frmEquipoEdicionVerificacion.frx":3C35
         TabIndex        =   47
         Top             =   3495
         Width           =   4215
      End
      Begin VB.TextBox txtAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         Top             =   2055
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   7215
         Picture         =   "frmEquipoEdicionVerificacion.frx":3C4D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ver norma"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarP 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion.frx":3EA2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar documento"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion.frx":4113
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Eliminar documento"
         Top             =   2400
         Width           =   360
      End
      Begin VB.CommandButton cmdAdjuntarP 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion.frx":42A7
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Buscar documento"
         Top             =   2400
         Width           =   405
      End
      Begin VB.CommandButton cmdMostrarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   7215
         Picture         =   "frmEquipoEdicionVerificacion.frx":4518
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ver norma"
         Top             =   2400
         Width           =   405
      End
      Begin VB.TextBox txtAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2415
         Width           =   4770
      End
      Begin MSComCtl2.DTPicker txtFechaActual 
         Height          =   405
         Left            =   10710
         TabIndex        =   31
         Top             =   180
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51314689
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFechaProxima_b 
         Height          =   405
         Left            =   10710
         TabIndex        =   35
         Top             =   1035
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51314689
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoVerificacion 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   270
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPeriVerificacion 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   630
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbVerificador 
         Height          =   330
         Left            =   1650
         TabIndex        =   2
         Top             =   990
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbProcedimiento 
         Height          =   330
         Left            =   1650
         TabIndex        =   4
         Top             =   1710
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVerificadorExterno 
         Height          =   330
         Left            =   1650
         TabIndex        =   3
         Top             =   1350
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion.frx":476D
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar documento"
         Top             =   2040
         Width           =   360
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2490
         Left            =   45
         TabIndex        =   60
         Top             =   4410
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton cmdTendencia 
         Height          =   300
         Left            =   4005
         TabIndex        =   75
         Top             =   6930
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Tendencia"
         Appearance      =   5
         Picture         =   "frmEquipoEdicionVerificacion.frx":4901
      End
      Begin XtremeSuiteControls.PushButton cmdCopiarPuntos 
         Height          =   300
         Left            =   2070
         TabIndex        =   79
         Top             =   6930
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Copiar Puntos"
         Appearance      =   5
         Picture         =   "frmEquipoEdicionVerificacion.frx":B163
      End
      Begin VB.Frame fraTipoParametro 
         BackColor       =   &H00C0C0C0&
         Height          =   2460
         Index           =   1
         Left            =   7560
         TabIndex        =   52
         Top             =   4500
         Visible         =   0   'False
         Width           =   4965
         Begin MSComctlLib.ListView lista_medidas 
            Height          =   2040
            Index           =   1
            Left            =   2925
            TabIndex        =   30
            Top             =   225
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   3598
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   12632319
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Resultados Medida"
               Object.Width           =   3704
            EndProperty
         End
         Begin VB.TextBox txtValor 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   29
            Top             =   1275
            Width           =   1635
         End
         Begin VB.TextBox txtNMedidas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   28
            Text            =   "1"
            Top             =   945
            Width           =   480
         End
         Begin MSDataListLib.DataCombo cmbTipoParametro 
            Height          =   315
            Index           =   1
            Left            =   495
            TabIndex        =   27
            Top             =   270
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line1 
            X1              =   45
            X2              =   2880
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label lblValor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   56
            Top             =   1335
            Width           =   450
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   54
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nº Medidas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   53
            Top             =   975
            Width           =   960
         End
      End
      Begin VB.Frame fraTipoParametro 
         BackColor       =   &H00C0C0C0&
         Height          =   2460
         Index           =   0
         Left            =   7560
         TabIndex        =   50
         Top             =   4860
         Visible         =   0   'False
         Width           =   4965
         Begin VB.TextBox txtDescripcion_Cualidad 
            Appearance      =   0  'Flat
            Height          =   885
            Left            =   60
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   540
            Width           =   4845
         End
         Begin VB.OptionButton optResultadoCualidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO CONFORME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   1800
            Width           =   3585
         End
         Begin MSDataListLib.DataCombo cmbTipoParametro 
            Height          =   315
            Index           =   0
            Left            =   2820
            TabIndex        =   23
            Top             =   180
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.OptionButton optResultadoCualidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CONFORME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   25
            Top             =   1470
            Value           =   -1  'True
            Width           =   3585
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   57
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   2370
            TabIndex        =   51
            Top             =   240
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin pryCombo.miCombo cmbUBICACION_ID 
         Height          =   330
         Left            =   7110
         TabIndex        =   90
         Top             =   3195
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incidencia"
         Height          =   195
         Index           =   18
         Left            =   6030
         TabIndex        =   93
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calibrado En"
         Height          =   195
         Index           =   17
         Left            =   6030
         TabIndex        =   91
         Top             =   3255
         Width           =   900
      End
      Begin VB.Label lblParametro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Parámetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   45
         TabIndex        =   55
         Top             =   4140
         Visible         =   0   'False
         Width           =   12495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "NºEquipo"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   78
         Top             =   7380
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Eval. Resultado"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   49
         Top             =   2835
         Width           =   1125
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Verificador Externo"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   48
         Top             =   1395
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Limitaciones uso"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   46
         Top             =   3195
         Width           =   1200
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hoja de Verificación"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   45
         Top             =   2130
         Width           =   1380
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cert. de verificación"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   44
         Top             =   2475
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo Verificación"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   43
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próx. Verificación"
         Height          =   195
         Index           =   0
         Left            =   9150
         TabIndex        =   42
         Top             =   780
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Actual Verificación"
         Height          =   195
         Index           =   10
         Left            =   9150
         TabIndex        =   41
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Ver. Interna"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   40
         Top             =   1050
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   38
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   37
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10155
      Width           =   1050
   End
   Begin XtremeSuiteControls.PushButton cmdReabrir 
      Height          =   885
      Left            =   8865
      TabIndex        =   80
      Top             =   10125
      Visible         =   0   'False
      Width           =   1635
      _Version        =   851970
      _ExtentX        =   2884
      _ExtentY        =   1561
      _StockProps     =   79
      Caption         =   "Reabrir"
      Appearance      =   5
      Picture         =   "frmEquipoEdicionVerificacion.frx":119C5
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Verificación de Equipo"
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
      TabIndex        =   39
      Top             =   120
      Width           =   2325
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   12795
   End
End
Attribute VB_Name = "frmEquipoEdicionVerificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarlngPK As Long
Public booSilencioso As Boolean
Private mvarobjEquipo As clsEquipos
Private mvarenuTipoEdicion As enumTipoEdicion
Private mvarstrId As String
Private bln_cambiando_tipo As Boolean

'Private WithEvents TecladoNumerico As frmTecladoNumerico
'Private blnEsTablet As Boolean
'Private blnPrimeraVez As Boolean

Private bln_fecha_real_editable As Boolean

Private mvarobjVerificacion As New clsEquipoVerificacion
Private mvarblnResultado As Boolean
Private mvarlngID_VERIFICACION As Long
Private mvardtmFechaProximaInicial As Date
Private mvarlngidVerificadorInternoInicial As Long
Private mvarlngIdPeriodoInicial As Long
Private mvarlngIdTipoVerificacionIncial As Long
Private mvarblnVieneDeCuaderno As Boolean

Private mvarlngidEquipo As Long
Private mvardtmFechaPrevista As Date

Private mvarlngIdEvento As Long

' Informar con el periodo de verificacion para cargar la ultima
Public copiarUltimaVerificacionPeriodo As Long

Private xR As New XArrayDB
Private xM(1 To 2) As New XArrayDB
Private xUnidades As New XArrayDB

Const filasR As Integer = 50
Const ColR As Integer = 11
Const filasM As Integer = 50
Const ColM As Integer = 1

'Private Enum ColsR
'    DESCRIPCION = 0
'    RANGO_MIN = 1
'    RANGO_MAX = 2
'    unidad = 3
'    RESULTADO_CAL = 4
'    TOLERANCIA = 5
'    INCERTIDUMBRE = 6
'    CORRECCION = 7
'    Id_resultado = 8
'    ID_UNIDAD = 9
'End Enum

Private Enum ColsR
    DESCRIPCION = 0
    Unidad = 1
    RANGO_MIN = 2
    RANGO_MAX = 3
    RESULTADO_MEDIA = 4
    ID_TIPO = 5
    Id_resultado = 6
    RESULTADO_CUALIDAD = 7
    RESULTADOS_PATRON = 8
    id_unidad = 9
    id_patron = 10
    n_medidas = 11
    LEQUIPOS = 12
    lReactivos = 13
    lreactivospropios = 14
End Enum

Private mvarlngNumParametrosResultados As Long
Private mvarlngidProcedmientoInicial As Long

Private Sub cmdAdjuntarP_Click(Index As Integer)
   On Error GoTo cmdAdjuntarP_Click_Error

    If mvarenuTipoEdicion = Alta Then
        MsgBox "Guarde primero la verificación para poder asignar documentos.", vbCritical, App.Title
        Exit Sub
    End If

    cd.ShowOpen
    If Trim(cd.FileName) = "" Then Exit Sub
    Dim oD As New clsDocumentacion
    Dim salida As String
    salida = oD.SubirEquipo(mvarlngidEquipo, 1, CLng(mvarstrId), Index, cd.FileName, cd.FileTitle)
    If salida <> "" Then
        MsgBox "Se ha producido un error al subir el documento : " & salida, vbCritical, App.Title
    Else
        txtAdjunto(Index) = cd.FileTitle
        Dim c As String
        c = "UPDATE eq_verificacion_equipos " & _
           "   set ruta_plantilla = '" & txtAdjunto(1) & "'" & _
           "      ,ruta_certificado = '" & txtAdjunto(2) & "'" & _
           "      ,ruta_evaluacion = '" & txtAdjunto(3) & "'" & _
           " where id_verificacion = " & CLng(mvarstrId)
        execute_bd c
    End If

   On Error GoTo 0
   Exit Sub

cmdAdjuntarP_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarP_Click of Formulario frmEquipoEdicionVerificacion"
End Sub

Private Sub cmdCopiarPuntos_Click()
    If txtEquipoCopia = "" Then
        MsgBox "Indique el numero de equipo del que desea copiar los parametros.", vbExclamation, App.Title
        txtEquipoCopia.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtEquipoCopia) Then
        MsgBox "El número de equipo debe ser numérico.", vbExclamation, App.Title
        txtEquipoCopia.SetFocus
        Exit Sub
    End If
    Dim oEV As New clsEquipoVerificacion
    Dim lngVerificacion As Long
    lngVerificacion = oEV.buscar_verificacion_anterior_misma_periodicidad(txtEquipoCopia, "")
    If lngVerificacion = 0 Then
        MsgBox "El equipo origen no tiene verificaciones.", vbExclamation, App.Title
        txtEquipoCopia.SetFocus
        Exit Sub
    End If
    
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs_medidas As ADODB.Recordset
    i = 0
    mvarlngNumParametrosResultados = i
    lblParametro.Caption = ""
    lblParametro.visible = False
'    If mvarenuTipoEdicion <> Alta Then
        ' Carga los Parametros de la verificacion
        
        Set rs = mvarobjVerificacion.DevolverParametrosResultados(lngVerificacion)
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
            
                With lista.ListItems.Add(, , rs!DESCRIPCION)
                    .Checked = (CInt(rs!REALIZADO) = 1)
                    .SubItems(ColsR.Unidad) = CStr(rs!Unidad)
                    If CInt(rs("tipo_id")) = 0 Then
'                        .SubItems(ColsR.RESULTADO_MEDIA) = IIf(CInt(rs("resultado_cualidad")) = 0, "NO CONFORME", "CONFORME")
                        .SubItems(ColsR.RESULTADO_MEDIA) = ""
                        .SubItems(ColsR.RANGO_MIN) = "N/A"
                        .SubItems(ColsR.RANGO_MAX) = "N/A"
                        .SubItems(ColsR.id_unidad) = "N/A"
                    Else
'                        .SubItems(ColsR.RESULTADO_MEDIA) = CStr(rs("resultado"))
                        .SubItems(ColsR.RESULTADO_MEDIA) = ""
                        .SubItems(ColsR.RANGO_MIN) = CStr(rs("rango_min"))
                        .SubItems(ColsR.RANGO_MAX) = CStr(rs("rango_max"))
                        .SubItems(ColsR.id_unidad) = CStr(rs("unidad_ID"))
                    End If
                    .SubItems(ColsR.ID_TIPO) = CStr(rs("tipo_id"))
                    .SubItems(ColsR.id_patron) = CStr(rs("patron_id"))
                    .SubItems(ColsR.Id_resultado) = CStr("0")
                    .SubItems(ColsR.n_medidas) = CStr(rs("n_medidas"))
                    .SubItems(ColsR.RESULTADO_CUALIDAD) = CStr("1")
                    .SubItems(ColsR.RESULTADOS_PATRON) = CStr("")
                    .SubItems(ColsR.LEQUIPOS) = CStr(rs("LEQUIPOS"))
                    .SubItems(ColsR.lReactivos) = CStr(rs("LREACTIVOS"))
                    .SubItems(ColsR.lreactivospropios) = CStr(rs("LREACTIVOS_PROPIOS"))
                End With
                rs.MoveNext
            Wend
            ' se va al primero
            lista.selectedItem = lista.ListItems(1)
            lista_Click
        End If
        
'    End If

End Sub

Private Sub cmdEliminarAdjunto_Click(Index As Integer)
    Dim oD As New clsDocumentacion
   On Error GoTo cmdEliminarAdjunto_Click_Error

    If oD.EliminarEquipo(mvarlngidEquipo, 1, CLng(mvarstrId), Index) = "" Then
        Dim c As String
        Dim s As String
        Select Case Index
        Case 1
            s = " ruta_plantilla = '' "
        Case 2
            s = " ruta_certificado = '' "
        Case 3
            s = " ruta_evaluacion = '' "
        End Select
        c = "UPDATE eq_verificacion_equipos set " & _
           s & _
           " where id_verificacion = " & CLng(mvarstrId)
        execute_bd c
        txtAdjunto(Index) = ""
    End If
    Set oD = Nothing

   On Error GoTo 0
   Exit Sub

cmdEliminarAdjunto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarAdjunto_Click of Formulario frmEquipoEdicionVerificacion"
End Sub

Private Sub cmdMostrarAdjunto_Click(Index As Integer)
    Dim oD As New clsDocumentacion
   On Error GoTo cmdMostrarAdjunto_Click_Error

    oD.CargarEquipo mvarlngidEquipo, 1, CLng(mvarstrId), Index, True
    Set oD = Nothing

   On Error GoTo 0
   Exit Sub

cmdMostrarAdjunto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrarAdjunto_Click of Formulario frmEquipoEdicionVerificacion"
End Sub

Private Sub cmdReabrir_Click()
    Dim verificacion As New clsEquipoVerificacion
   On Error GoTo cmdReabrir_Click_Error

    verificacion.Reabrir CLng(mvarstrId)
    MsgBox "La verificación ha sido cambiada a 'Prevista'", vbInformation, App.Title
    
    cmdReabrir.visible = False
    mvarenuTipoEdicion = EDICION
    mvarobjVerificacion.Carga CLng(mvarstrId)
    Call PresentarDatos
    Call OpcionesEdicion

   On Error GoTo 0
   Exit Sub

cmdReabrir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdReabrir_Click of Formulario frmEquipoEdicionVerificacion"
End Sub

Private Sub cmdRevisiones_Click()
    With frmRevisiones
        .TOBJETO = TOBJETO_REV_EQ_VERIFICACION
        .COBJETO = mvarlngidEquipo
        .Show 1
    End With
End Sub


' botón que permite imprimir la etiqueta de calibración
Private Sub cmdetiqueta_Click()

    If optEstado(VER_ESTADOS.VER_ESTADO_REALIZADA).Value = True Or optEstado(VER_ESTADOS.VER_ESTADO_REVISADA).Value = True Then
        Dim oEV As New clsEquipoVerificacion
        oEV.imprimir_etiqueta CLng(mvarstrId)
        Set oEV = Nothing
    Else
        MsgBox "La calibración tiene que estar cerrada para poder generar la etiqueta.", vbExclamation, App.Title
    End If

End Sub
Private Sub cabecera()
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 800, lvwColumnLeft
        .Add , , "Nombre", 2700, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
  With lista.ColumnHeaders
        
        .Item(1).Text = "Pto. Verificacion"
        .Item(1).Width = lista.Width * 0.3
        .Item(1).Alignment = lvwColumnLeft
        .Add , , "Unidad", lista.Width * 0.19, lvwColumnLeft
        .Add , , "Rango Min", lista.Width * 0.15, lvwColumnCenter
        .Add , , "Rango Max", lista.Width * 0.15, lvwColumnCenter
        .Add , , "Resultado", lista.Width * 0.15, lvwColumnCenter
        
        ' ocultas
        
        .Add , , "id_tipo", 0
        .Add , , "id_resultado", 0
        .Add , , "RESULTADO_CUALIDAD", 0
        .Add , , "RESULTADOS_PATRON", 0
        .Add , , "id_unidad", 0
        .Add , , "ID_PATRON", 0
        .Add , , "n_medidas", 0
        .Add , , "lequipos", 0
        .Add , , "lreactivos", 0
        .Add , , "lreactivospropios", 0
        
    End With
End Sub

Private Function comprobar_datos_parametros(ByRef resultado_conformidad As Boolean) As String
    Dim x As Long
    Dim tipo As Integer
    Dim cad As String
    Dim strTipo As String, rmin As Double, rmax As Double, res As Double
    
    If lista.ListItems.Count = 0 Then Exit Function
    
    resultado_conformidad = True
    
    For x = 1 To lista.ListItems.Count
        With lista.ListItems(x)
            If .Checked Then
                tipo = CInt(.SubItems(ColsR.ID_TIPO))
'                strTipo = IIf(tipo = 1, "Equipo", "Reactivo")
                    
                If tipo <> 0 Then
                    If Not optEstado(VER_ESTADOS.VER_ESTADO_PREVISTA).Value Then
                        If .SubItems(ColsR.LEQUIPOS) = "" And .SubItems(ColsR.lReactivos) = "" And .SubItems(ColsR.lreactivospropios) = "" Then
                                cad = cad & vbCrLf & " - El parámetro " & .Text & ", del tipo Patrón no tiene ningún Equipo/Reactivo."
                        End If
                        If IsNumeric(.SubItems(ColsR.RANGO_MIN)) And IsNumeric(.SubItems(ColsR.RANGO_MAX)) And IsNumeric(.SubItems(ColsR.RESULTADO_MEDIA)) Then
                            rmin = CDbl(.SubItems(ColsR.RANGO_MIN))
                            rmax = CDbl(.SubItems(ColsR.RANGO_MAX))
                            res = CDbl(.SubItems(ColsR.RESULTADO_MEDIA))
                            If rmin > rmax Then
                                cad = cad & vbCrLf & " - En el parámetro " & .Text & ", del tipo Patrón-" & strTipo & ", el Rango Mínimo es mayor que el Rango Máximo" & strTipo
                            End If
                            
                            If res < rmin Then
                                ' inferior al minimo
                                resultado_conformidad = False
                            End If
                            
                            If res > rmax Then
                                ' superior al máximo
                                resultado_conformidad = False
                            End If
                        Else
                            resultado_conformidad = True
                        End If
                    End If
                Else
                    If optResultadoCualidad(0).Value Then resultado_conformidad = False
                End If
            End If
        End With
        
    Next x
    comprobar_datos_parametros = cad
    
    
End Function

'Private Sub ConfigurarTablet()
'    Set TecladoNumerico = New frmTecladoNumerico
'
'
'    TecladoNumerico.OcultarConformidad = False
'
''    blnEsTablet = pc_es_tablet
'    blnEsTablet = False
'
'    If blnEsTablet Then
'
'        blnPrimeraVez = True
'        'On Error Resume Next
'        'grdResultados.Columns(ColsR.RESULTADO_MEDIA).Locked = True
'        'On Error GoTo 0
'        Me.top = 0
'
'
'    End If
'End Sub
Private Function devolver_medidas_resultado(ByVal prm_id_resultado As Long, ByRef rs As ADODB.Recordset) As String

    If rs.RecordCount = 0 Then Exit Function
    Dim cad As String
    
    cad = ""
    
    rs.Filter = "resultado_id = " & CStr(prm_id_resultado)
    
    If rs.RecordCount = 0 Then
        rs.MoveFirst
        While Not rs.EOF
            cad = cad & ";" & CStr(rs("resultado"))
            rs.MoveNext
        Wend
        cad = Mid(cad, 2)
    End If
    
    rs.Filter = ""
    
    devolver_medidas_resultado = cad

End Function

Private Sub modificar_parametro()

    Dim objfrm As New frmEquipoVerificacionAnadirParametro
    Dim TIPO_ID As Integer, tipo_actual As Integer

    tipo_actual = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
    
    With objfrm
        .DESCRIPCION = lista.selectedItem.Text
        .id_unidad = 0
        .rmin = 0
        .rmax = 0
        If IsNumeric(lista.selectedItem.SubItems(ColsR.id_unidad)) Then
            .id_unidad = lista.selectedItem.SubItems(ColsR.id_unidad)
            .rmax = lista.selectedItem.SubItems(ColsR.RANGO_MAX)
            .rmin = lista.selectedItem.SubItems(ColsR.RANGO_MIN)
        End If
        .tipo = lista.selectedItem.SubItems(ColsR.ID_TIPO)
        
        .medidas = lista.selectedItem.SubItems(ColsR.n_medidas)
    End With
    
    objfrm.Show vbModal
    
    If Not objfrm.resultado Then Exit Sub
    
    TIPO_ID = objfrm.tipo
                
    With lista.ListItems(lista.selectedItem.Index)
        .Text = objfrm.DESCRIPCION
        .SubItems(ColsR.Unidad) = objfrm.Unidad
        If TIPO_ID = 0 Then
            .SubItems(ColsR.RESULTADO_MEDIA) = "CONFORME"
            .SubItems(ColsR.RANGO_MIN) = "N/A"
            .SubItems(ColsR.RANGO_MAX) = "N/A"
            .SubItems(ColsR.id_unidad) = "N/A"
            .SubItems(ColsR.n_medidas) = "1"
        Else
            .SubItems(ColsR.RANGO_MIN) = objfrm.rmin
            .SubItems(ColsR.RANGO_MAX) = objfrm.rmax
            .SubItems(ColsR.id_unidad) = objfrm.id_unidad
            .SubItems(ColsR.Unidad) = objfrm.Unidad
            If .SubItems(ColsR.n_medidas) <> objfrm.medidas Then
                Dim r As Double, str_total As String
                If IsNumeric(.SubItems(ColsR.RESULTADO_MEDIA)) Then
                    r = CDbl(.SubItems(ColsR.RESULTADO_MEDIA))
                Else
                    r = 0
                End If
                .SubItems(ColsR.RESULTADOS_PATRON) = recalcular_resultados_patron(.SubItems(ColsR.RESULTADOS_PATRON), objfrm.medidas, r)
                If InStr(1, CStr(r), ",") Then
                    ' mide los decimales. Si son más de 6, los redondea a 6 decimales
                    str_total = Split(CStr(CDbl(r)), ",")(0) & "," & Left(Split(CStr(CDbl(r)), ",")(1), 6)
                Else
                    str_total = CStr(r)
                End If

                .SubItems(ColsR.RESULTADO_MEDIA) = str_total
            End If
            .SubItems(ColsR.n_medidas) = objfrm.medidas
        End If
        
        .SubItems(ColsR.ID_TIPO) = TIPO_ID
'        If tipo_actual <> TIPO_ID Then
            ' solo si cambia el tipo se reinicializan estos parámetros
'            .SubItems(ColsR.RESULTADO_MEDIA) = "0"
'            .SubItems(ColsR.id_patron) = "0"
'            .SubItems(ColsR.n_medidas) = "1"
'            .SubItems(ColsR.RESULTADO_CUALIDAD) = "1"
'            .SubItems(ColsR.RESULTADOS_PATRON) = IIf(TIPO_ID = 0, "", "0")
'        End If
    End With
    Unload objfrm
    Set objfrm = Nothing
    ' se va al añadido
    'lista.SelectedItem = lista.ListItems(lista.ListItems.Count)
    lista_Click
    
End Sub



Private Sub mostrar_datos_medidas()
    Dim TIPO_ID As Integer, i As Integer, patron_id As Long
    Dim cad As String
    Dim arrRes() As String
    Dim n_medidas As Integer, res_cual As Integer
    Dim Inicializar As Boolean
    'Muestra los resultados segun tipo
    
    If lista.ListItems.Count = 0 Then Exit Sub
    
    patron_id = 0
    TIPO_ID = -1
    n_medidas = 1
    res_cual = 1
    Inicializar = True
    
    n_medidas = CInt(lista.selectedItem.SubItems(ColsR.n_medidas))
    TIPO_ID = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
    Select Case CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
    Case 0
        TIPO_ID = 0
    Case 1, 2
        TIPO_ID = 1
        cmbTipoParametro(1).BoundText = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
    End Select
    lblParametro = lista.selectedItem
    cad = lista.selectedItem.SubItems(ColsR.RESULTADOS_PATRON)
    res_cual = lista.selectedItem.SubItems(ColsR.RESULTADO_CUALIDAD)
'    patron_id = lista.SelectedItem.SubItems(ColsR.id_patron)
    ' EQUIPOS
    listaEquipos.ListItems.Clear
    Dim Equipos As String
    Equipos = lista.selectedItem.SubItems(ColsR.LEQUIPOS)
    Dim v() As String
    v = Split(Equipos, ";")
    For i = LBound(v) To UBound(v)
       If v(i) <> "" Then
           cargar_equipo CLng(v(i))
       End If
    Next
    
    ' REACTIVOS
    listaReactivos.ListItems.Clear
    cargar_reactivos lista.selectedItem.SubItems(ColsR.lReactivos), 0
    cargar_reactivos lista.selectedItem.SubItems(ColsR.lreactivospropios), 1
    If TIPO_ID < 0 Then Exit Sub
    
    lblParametro.visible = True
    
    fraTipoParametro(0).visible = False
    fraTipoParametro(1).visible = False
    fraTipoParametro(TIPO_ID).visible = True
    ' limpia los resultados
    If TIPO_ID <> 0 Then
        
        lista_medidas(TIPO_ID).ListItems.Clear
        For i = 1 To n_medidas
            lista_medidas(TIPO_ID).ListItems.Add , , "0"
        Next i
        
        If cad <> "" Then
            arrRes = Split(cad, ";")
            For i = 0 To UBound(arrRes)
                lista_medidas(TIPO_ID).ListItems(i + 1).Text = arrRes(i)
            Next i
        End If
        txtNMedidas(TIPO_ID).Text = n_medidas
        lista_medidas_Click TIPO_ID
    Else
        txtDescripcion_Cualidad.Text = cad
        optResultadoCualidad(res_cual).Value = True
    End If

End Sub

Private Sub OpcionesEdicion()

    If mvarenuTipoEdicion = Alta Then
        txtFechaActual.Enabled = True
    ElseIf mvarenuTipoEdicion = EDICION Then
        txtFechaActual.Enabled = bln_fecha_real_editable
        cmbTipoVerificacion.Locked = False
        cmbPeriVerificacion.Locked = False
        cmbVerificador.activar
        cmbVerificadorExterno.activar
        cmbProcedimiento.activar
        cmbUBICACION_ID.activar
        txtCalibradoEn.Locked = False
        txtLimitacionesUso.Locked = False
        cmdAnadirLimitacion.Enabled = True
        cmdEliminarLimitacion.Enabled = True
        lstLimitacionesUso.Enabled = True
        txtFechaProxima_b.Enabled = True
        fraEstadoIntervencion.Enabled = True
'        txtCertificado.Locked = False
'        txtHojaVerificacion.Locked = False
'        txtEvaluacionResultado.Locked = False
        cmdAnadirParametro.Enabled = True
        cmdEliminarParametro.Enabled = True
        fraTipoParametro(0).Enabled = True
        fraTipoParametro(1).Enabled = True
        cmdok.visible = True
    
    ElseIf mvarenuTipoEdicion = visualizar Then
        cmbTipoVerificacion.Locked = True
        cmbPeriVerificacion.Locked = True
        cmbVerificador.desactivar
        cmbVerificadorExterno.desactivar
        cmbProcedimiento.desactivar
        cmbUBICACION_ID.desactivar
        txtCalibradoEn.Locked = True
'        txtHojaCalibracion.Locked = False
            cmdMostrarAdjunto(1).Left = cmdAdjuntarP(1).Left
            cmdAdjuntarP(1).visible = False
'            cmdEscanearAdjunto(1).Visible = False
            cmdEliminarAdjunto(1).visible = False
'        txtCertificado.Locked = False
            cmdMostrarAdjunto(2).Left = cmdAdjuntarP(2).Left
            cmdAdjuntarP(2).visible = False
'            cmdEscanearAdjunto(2).Visible = False
            cmdEliminarAdjunto(2).visible = False
'        txtEvaluacionResultado.Locked = False
            cmdMostrarAdjunto(3).Left = cmdAdjuntarP(3).Left
            cmdAdjuntarP(3).visible = False
'            cmdEscanearAdjunto(3).Visible = False
            cmdEliminarAdjunto(3).visible = False
        txtLimitacionesUso.Locked = True
        cmdAnadirLimitacion.Enabled = False
        cmdEliminarLimitacion.Enabled = False
        lstLimitacionesUso.Enabled = False
        txtFechaProxima_b.Enabled = False
        fraEstadoIntervencion.Enabled = False
'        txtCertificado.Locked = True
'        txtHojaVerificacion.Locked = True
'        txtEvaluacionResultado.Locked = True
        cmdAnadirParametro.Enabled = False
        cmdEliminarParametro.Enabled = False
        fraTipoParametro(0).Enabled = False
        fraTipoParametro(1).Enabled = False
        cmdok.visible = False
    End If
    If cmbPeriVerificacion.BoundText <> "" Then
        If cmbPeriVerificacion.BoundText = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO Or _
           cmbPeriVerificacion.BoundText = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
            txtFechaProxima.visible = False
            lblCampos(0).visible = False
        End If
    End If
    ' Permiso para modificar la vida
    If optEstado(VER_ESTADOS.VER_ESTADO_PREVISTA).Value <> True Then
        Dim op As New clsParametros
        Dim s() As String
        Dim i As Integer
        op.Carga parametros.PARAM_USUARIOS_MODIFICAN_EQUIPOS_MUESTRA_CERRADA, ""
        If op.getVALOR <> "" Then
            s = Split(op.getVALOR, ",")
            For i = LBound(s) To UBound(s)
                If USUARIO.getID_EMPLEADO = CInt(s(i)) Then
                    cmdReabrir.visible = True
                    Exit For
                End If
            Next
        End If
        Set op = Nothing
    End If
End Sub

Private Function recalcular_resultados_patron(res_patron As String, n_medidas As Integer, resultado As Double) As String
Dim arrPatron() As String
Dim i As Integer
Dim res As String

    
    If Trim(res_patron) <> "" And n_medidas <> 0 Then
        arrPatron = Split(res_patron, ";")
        If (UBound(arrPatron) + 1) < n_medidas Then ' se aumentan las medidas
            res = res_patron
            For i = 0 To UBound(arrPatron)
                resultado = resultado + CDbl(arrPatron(i))
            Next i
            
            For i = (UBound(arrPatron) + 1) To n_medidas
                res = res = ";0"
            Next i
            res = Mid(res, 2)
        Else ' disminuyen las medidas
            For i = 0 To (n_medidas - 1)
                resultado = resultado + CDbl(arrPatron(i))
                res = res & ";" & arrPatron(i)
            Next i
            res = Mid(res, 2)
        End If
        
        'recalcula el resultado
        resultado = resultado / n_medidas
        
    End If


End Function
' JGM
Private Sub cmbEquipos_change()
'    lista.SelectedItem.SubItems(ColsR.id_patron) = cmbequipos.getPK_SALIDA
End Sub

Private Sub cmbPeriVerificacion_Click(area As Integer)

    Call txtFechaActual_Change

End Sub
' JGM
Private Sub cmbReactivos_change()
lista.selectedItem.SubItems(ColsR.id_patron) = cmbReactivos.getPK_SALIDA
End Sub

Private Sub cmbTipoVerificacion_Change()


If cmbTipoVerificacion.BoundText = "1" Then ' Intera
    ' Es interna
    cmbVerificadorExterno.desactivar
Else
    ' Es externa
    cmbVerificadorExterno.activar
End If

End Sub

' botón que abre un cuadro de diálogo para seleccionar la plantilla excel de la verificación
'Private Sub cmdAdjuntarCertificado_Click()
'
'On Error GoTo cmdAdjuntarCertificado_Click_Error
'
'    cd.ShowOpen
'
'    If Trim(cd.FileName) = "" Then Exit Sub
'
'    mvarobjVerificacion.CERTIFICADO.setRUTA_TEMPORAL = cd.FileName
'    mvarobjVerificacion.CERTIFICADO.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtCertificado.Text = cd.FileTitle
'
'On Error GoTo 0
'    Exit Sub
'cmdAdjuntarCertificado_Click_Error:
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarCertificado_Click of Formulario frmEquipoEdicionVerificacion"
'End Sub
'
'Private Sub cmdAdjuntarEvaluacion_Click()
'
'On Error GoTo cmdAdjuntarEvaluacion_Click_Error
'
'    cd.ShowOpen
'
'    If Trim(cd.FileName) = "" Then Exit Sub
'
'    mvarobjVerificacion.Evaluacion.setRUTA_TEMPORAL = cd.FileName
'    mvarobjVerificacion.Evaluacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtEvaluacionResultado.Text = cd.FileTitle
'
'
'On Error GoTo 0
'    Exit Sub
'cmdAdjuntarEvaluacion_Click_Error:
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarEvaluacion_Click of Formulario frmEquipoEdicionVerificacion"
'End Sub
'Private Sub cmdAdjuntarHojaCal_Click()
'On Error GoTo cmdAdjuntarHojaCal_Click_Error
'    cd.ShowOpen
'    If Trim(cd.FileName) = "" Then Exit Sub
'    mvarobjVerificacion.HojaVerificacion.setRUTA_TEMPORAL = cd.FileName
'    mvarobjVerificacion.HojaVerificacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtHojaVerificacion.Text = cd.FileTitle
'On Error GoTo 0
'    Exit Sub
'cmdAdjuntarHojaCal_Click_Error:
'End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        cargar_equipo cmbEquipos.getPK_SALIDA
        informar_equipos_lista
        cmbEquipos.limpiar
    End If

End Sub

Private Sub cmdAnadirLimitacion_Click()

    mvarobjEquipo.Anadir_limitacionuso_equipo txtLimitacionesUso.Text
           
    Call PresentarDatos_LimitacionesUso
End Sub

Private Sub cmdAnadirReactivo_Click()
    ' Interno (I)
    If cmbReactivos.getTEXTO <> "" Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
        Dim oTR As New clsTipos_reactivo_ex
        oBote.CARGAR cmbReactivos.getPK_SALIDA
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "E"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Externo (E)
    If cmbReactivosInternos.getTEXTO <> "" Then
        Dim oRPR As New clsRpr_botes
        Dim oTRPR As New clsRPR_Tipos
        oRPR.Carga cmbReactivosInternos.getPK_SALIDA
        oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
        With listaReactivos.ListItems.Add(, , oRPR.getID_BOTE_PR)
            .SubItems(1) = oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    informar_reactivos_lista
    cmbReactivos.limpiar
    cmbReactivosInternos.limpiar

End Sub

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    mvarlngID_VERIFICACION = 0
    Me.Hide
End Sub

' botón que borra el documento de verificación
'Private Sub cmdEliminarCertificado_Click()
'
'txtCertificado.Text = ""
'mvarobjVerificacion.CERTIFICADO.setID_AUX = enumIdAux.ID_AUX_ELIMINADO
'
'End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
    End If
    informar_equipos_lista
End Sub

'Private Sub cmdEliminarEvaluacion_Click()
'
'txtEvaluacionResultado.Text = ""
'mvarobjVerificacion.Evaluacion.setID_AUX = enumIdAux.ID_AUX_ELIMINADO
'
'End Sub


'Private Sub cmdEliminarHojaCal_Click()
'
'txtHojaVerificacion.Text = ""
'
'mvarobjVerificacion.HojaVerificacion.setID_AUX = enumIdAux.ID_AUX_ELIMINADO
'
'End Sub

Private Sub cmdEliminarLimitacion_Click()
Dim lngid As Long

    If lstLimitacionesUso.ListIndex < 0 Then Exit Sub

    lngid = lstLimitacionesUso.ItemData(lstLimitacionesUso.ListIndex)

    mvarobjEquipo.Eliminar_LimitacionUso_equipo lngid
    
    Call PresentarDatos_LimitacionesUso
End Sub

Private Sub cmdEliminarParametro_Click()
If lista.ListItems.Count = 0 Then Exit Sub

If lista.selectedItem.Index <= 0 Then
    MsgBox "Debe señalar el parámetro a eliminar", vbInformation, "Eliminar Parámetro"
    Exit Sub
End If

lista.ListItems.Remove lista.selectedItem.Index


If lista.ListItems.Count = 0 Then
    fraTipoParametro(0).visible = False
    fraTipoParametro(1).visible = False
'    fraTipoParametro(2).Visible = False
    lblParametro.visible = False
Else
    lista.selectedItem = lista.ListItems(1)
End If

End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If

End Sub

'Private Sub cmdEscanearCert_Click()
'Dim strArchivo As String
'
'    strArchivo = EscanearATemp
'
'    If Trim(strArchivo) = "" Then Exit Sub
'
'    mvarobjVerificacion.CERTIFICADO.setRUTA_TEMPORAL = strArchivo
'    mvarobjVerificacion.CERTIFICADO.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtCertificado.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
'End Sub

'Private Sub cmdEscanearEvaluacion_Click()
'Dim strArchivo As String
'
'    strArchivo = EscanearATemp
'
'    If Trim(strArchivo) = "" Then Exit Sub
'
'    mvarobjVerificacion.Evaluacion.setRUTA_TEMPORAL = strArchivo
'    mvarobjVerificacion.Evaluacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtEvaluacionResultado.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
'
'End Sub


'Private Sub cmdEscanearHoja_Click()
'
'    Dim strArchivo As String
'
'    strArchivo = EscanearATemp
'
'    If Trim(strArchivo) = "" Then Exit Sub
'
'    mvarobjVerificacion.HojaVerificacion.setRUTA_TEMPORAL = strArchivo
'    mvarobjVerificacion.HojaVerificacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
'    txtHojaVerificacion.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
'
'End Sub

' botón que permite visualizar el archivo seleccionado
'Private Sub cmdMostrarCertificado_Click()
'
'    Dim objAI As New clsArchivoAdjunto
'    Dim destino As String, r As Double
'
'    Set objAI = mvarobjVerificacion.CERTIFICADO
'
'    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
'        destino = objAI.getRUTA_TEMPORAL
'    ElseIf (objAI.getRUTA <> "") Then
'        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & mvarobjVerificacion.getID_VERIFICACION & "\CERT\" & objAI.getNOMBRE_ARCHIVO
'    End If
'
'    On Error GoTo fallo
'    If destino = "" Then
'        MsgBox "La evaluación no se localiza." & objAI.getRUTA, vbCritical, App.Title
'        Exit Sub
'    End If
'
'    ' verificar si es hoja excel
'    If UCase(Right(destino, 3)) = "XLS" Then
'        Dim XLA As excel.Application
'        Dim XLW As excel.Workbook
'        Dim XLS As excel.Worksheet
'        Set XLA = New excel.Application
'        Set XLW = XLA.Workbooks.Open(destino, , True)
'        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
'    ElseIf Dir(destino, vbArchive) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If
'
'fallo:
'End Sub
'
'Private Sub cmdMostrarEvaluacion_Click()
'
'    Dim objAI As New clsArchivoAdjunto
'    Dim destino As String, r As Double
'
'    Set objAI = mvarobjVerificacion.Evaluacion
'
'    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
'        destino = objAI.getRUTA_TEMPORAL
'    ElseIf (objAI.getRUTA <> "") Then
'        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & mvarobjVerificacion.getID_VERIFICACION & "\EVAL\" & objAI.getNOMBRE_ARCHIVO
'    End If
'
'    On Error GoTo fallo
'    If destino = "" Then
'        MsgBox "La evaluación no se localiza." & objAI.getRUTA, vbCritical, App.Title
'        Exit Sub
'    End If
'
'    ' verificar si es hoja excel
'    If UCase(Right(destino, 3)) = "XLS" Then
'        Dim XLA As excel.Application
'        Dim XLW As excel.Workbook
'        Dim XLS As excel.Worksheet
'        Set XLA = New excel.Application
'        Set XLW = XLA.Workbooks.Open(destino, , True)
'        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
'    ElseIf Dir(destino, vbArchive) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If
'
'fallo:
'End Sub


'Private Sub cmdMostrarHojaCal_Click()
'
'
'    Dim objAI As New clsArchivoAdjunto
'    Dim destino As String, r As Double
'    Set objAI = mvarobjVerificacion.HojaVerificacion
'
'    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
'        destino = objAI.getRUTA_TEMPORAL
'    ElseIf (objAI.getRUTA <> "") Then
'        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & mvarobjVerificacion.getID_VERIFICACION & "\HOJA\" & objAI.getNOMBRE_ARCHIVO
'    End If
'
'On Error GoTo fallo
'    If destino = "" Then
'        MsgBox "La evaluación no se localiza." & objAI.getRUTA, vbCritical, App.Title
'        Exit Sub
'    End If
'
'    ' verificar si es hoja excel
'    If UCase(Right(destino, 3)) = "XLS" Then
'        Dim XLA As excel.Application
'        Dim XLW As excel.Workbook
'        Dim XLS As excel.Worksheet
'        Set XLA = New excel.Application
'        Set XLW = XLA.Workbooks.Open(destino, , True)
'        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
'    ElseIf Dir(destino, vbArchive) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If
'fallo:
'End Sub
Private Sub cmdok_Click()
    ' Recoge los datos
    Dim lngId_Verificacion As Long
    Dim bln_conformidad As Boolean
   On Error GoTo cmdok_Click_Error

    If Not ComprobarDatos() Then Exit Sub
    
    RecogerDatos
    
    If mvarenuTipoEdicion = Alta Then
        mvarobjVerificacion.setEQUIPO_ID = mvarlngidEquipo
        lngId_Verificacion = mvarobjVerificacion.Insertar(True, lista)
    Else
        lngId_Verificacion = CLng(mvarstrId)
        Call mvarobjVerificacion.Modificar(lngId_Verificacion, True, , lista)
    End If
        
    'Call mvarobjVerificacion.GuardarParametrosVerificacion(mvarlngidEquipo, lngId_Verificacion, xR, filasR)
    
    'If Not mvarblnVieneDeCuaderno Then
        ' Si no viene del cuaderno de avisos, es decir, que viene de la gestion normal y corriente, recarga las calibraciones
    '    mvarobjEquipo.Carga_Verificaciones
    'End If
    MsgBox "La verificación se ha almacenado correctamente.", vbInformation, App.Title
    mvarblnResultado = True
    mvarlngID_VERIFICACION = mvarobjVerificacion.getID_VERIFICACION
    Me.Hide

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk_Click of Formulario frmEquipoEdicionVerificacion"

End Sub


Private Sub comprobar_fecha_real_modificable()

    Dim op As New clsParametros
    
    bln_fecha_real_editable = False
    
    If op.Carga(parametros.MODIFICACION_FECHAS_CALIBRACION_VERIFICACION, "") Then
        If op.getVALOR = "1" Then
            bln_fecha_real_editable = True
        End If
    End If
    
    ' M1130-I
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
        bln_fecha_real_editable = False
    End If
    ' M1130-F

End Sub


Private Function ComprobarDatos() As Boolean
Dim strMs As String
Dim bln_conformidad As Boolean
On Error GoTo ComprobarDatos_Error

    ComprobarDatos = False

    strMs = ""

    If Not optEstado(VER_ESTADOS.VER_ESTADO_PREVISTA).Value Then
        comprobarDatosResultados strMs
    End If
    
    If Trim(cmbTipoVerificacion.BoundText) = "" Or Trim(cmbTipoVerificacion.BoundText) = "0" Then
        strMs = strMs & vbCrLf & " - Debe indicar el Tipo de Verificación"
    End If

    If cmbVerificador.getPK_SALIDA = 0 Then
        strMs = strMs & vbCrLf & " - Debe indicar el Responsable Interno de Verificación"
    End If


    If Trim(cmbPeriVerificacion.BoundText) = "" Then
        strMs = strMs & vbCrLf & " - Debe indicar el Periodo para las Verificaciones"
    End If
    
    If getDataComboSel(cmbTipoVerificacion) = 1 Then
        If Trim(cmbProcedimiento.getPK_SALIDA) < 1 Then
            strMs = strMs & vbCrLf & " - Debe indicar el el Procedimiento"
        End If
    ElseIf getDataComboSel(cmbTipoVerificacion) = 2 Then
        If Trim(cmbVerificadorExterno.getPK_SALIDA) < 1 Then
            strMs = strMs & vbCrLf & " - Debe indicar el Verificador Externo"
        End If
    End If
    
    If CDate("01/01/1900") = txtFechaActual.Value Then
        strMs = strMs & vbCrLf & " - Debe indicar una Fecha Actual de Verificación adecuada"
    End If
    
    If CDate(Format(txtFechaActual.Value, "dd/mm/yyyy")) > CDate(txtFechaProxima_b.Value) Then
        strMs = strMs & vbCrLf & " - La fecha de la próxima verificación no puede ser anterior a la de la Verificación actual"
    End If

    
    strMs = strMs & comprobar_datos_parametros(bln_conformidad)
    
    If Trim(strMs) <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strMs
        Exit Function
    End If

    ' comprobar si se cierra conforme y no lo es
    If optEstado(1).Value = True Then
        If optResultado(0).Value And Not bln_conformidad Then
            MsgBox "ATENCION: No se puede cerrar esta verificación como CONFORME, dado que uno de los resultados de los parámetros está fuera de rango", vbInformation, "Verificación NO CONFORME"
            Exit Function
        ElseIf optResultado(1).Value And bln_conformidad Then
            MsgBox "ATENCION: No se puede cerrar esta verificación como NO CONFORME, dado que todos los resultados de los parámetros está dentro de rango", vbInformation, "Verificación CONFORME"
            Exit Function
        End If
    End If
    
    ' Si es una verificación antes de uso, no se puede cerrar como prevista
    If cmbPeriVerificacion.BoundText = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO Or _
       cmbPeriVerificacion.BoundText = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
        
        If optEstado(0).Value = True Then
            MsgBox "ATENCION: Las verificaciones '" & cmbPeriVerificacion.Text & "' no se puede cerrar como pendientes.", vbInformation, App.Title
            Exit Function
        End If
    End If


    ComprobarDatos = True

On Error GoTo 0
    Exit Function
ComprobarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmEquipoEdicionVerificacion"
End Function


Private Sub comprobarDatosResultados(ByRef strMs As String)

    Dim i As Long
    i = 0
    Dim cad As String
    
    cad = ""
    'grdResultados.Refresh
    
    'For i = 0 To filasR
    '    cad = cad & xR(i, 0) & ", " & xR(i, 1) & ", " & xR(i, 2) & ", " & xR(i, 3) & ", " & xR(i, 4) & ", " & xR(i, 5) & ", " & xR(i, 6) & ", " & xR(i, 7) & ". " & vbCrLf
    'Next i
    
End Sub

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Public Property Get FechaPrevista() As Date

    FechaPrevista = mvardtmFechaPrevista

End Property

Public Property Let FechaPrevista(ByVal dtmFechaPrevista As Date)

    mvardtmFechaPrevista = dtmFechaPrevista

End Property

Public Property Get FechaProximaInicial() As Date

    FechaProximaInicial = mvardtmFechaProximaInicial

End Property

Public Property Let FechaProximaInicial(ByVal dtmFechaProximaInicial As Date)

    mvardtmFechaProximaInicial = dtmFechaProximaInicial

End Property

Private Sub cmdAnadirParametro_Click()
Dim objfrm As New frmEquipoVerificacionAnadirParametro
Dim TIPO_ID As Integer


    objfrm.Show vbModal
    
    If Not objfrm.resultado Then Exit Sub
    
    TIPO_ID = objfrm.tipo
    
    With lista.ListItems.Add(, , objfrm.DESCRIPCION)
        .SubItems(ColsR.Unidad) = objfrm.Unidad
        .Checked = True
        If TIPO_ID = 0 Then
            .SubItems(ColsR.RESULTADO_MEDIA) = "CONFORME"
            .SubItems(ColsR.RANGO_MIN) = "N/A"
            .SubItems(ColsR.RANGO_MAX) = "N/A"
            .SubItems(ColsR.id_unidad) = "N/A"
            .SubItems(ColsR.n_medidas) = "1"
            .SubItems(ColsR.RESULTADOS_PATRON) = ""
        Else
            .SubItems(ColsR.RESULTADO_MEDIA) = "0"
            .SubItems(ColsR.RANGO_MIN) = objfrm.rmin
            .SubItems(ColsR.RANGO_MAX) = objfrm.rmax
            .SubItems(ColsR.id_unidad) = objfrm.id_unidad
            .SubItems(ColsR.n_medidas) = objfrm.medidas
            .SubItems(ColsR.RESULTADOS_PATRON) = "0"
        End If
        .SubItems(ColsR.ID_TIPO) = TIPO_ID
        .SubItems(ColsR.id_patron) = "0"
        .SubItems(ColsR.Id_resultado) = "0"
        
        .SubItems(ColsR.RESULTADO_CUALIDAD) = "1"
        
    End With

    Unload objfrm
    Set objfrm = Nothing

    ' se va al añadido
    lista.selectedItem = lista.ListItems(lista.ListItems.Count)
    lista_Click
    
    
End Sub

Private Sub cmdTendencia_Click()
   On Error GoTo cmdTendencia_Click_Error

    If lista.ListItems.Count > 0 Then
        frmEquiposTendencias.PK_EQUIPO_ID = mvarlngidEquipo
        frmEquiposTendencias.PK_PERIODICIDAD = cmbPeriVerificacion.BoundText
        frmEquiposTendencias.PK_PARAMETRO = lista.ListItems(lista.selectedItem.Index).Text
        If IsNumeric(lista.ListItems(lista.selectedItem.Index).SubItems(2)) Then
            frmEquiposTendencias.PK_RANGO_MIN = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        Else
            frmEquiposTendencias.PK_RANGO_MIN = "0"
        End If
        If IsNumeric(lista.ListItems(lista.selectedItem.Index).SubItems(3)) Then
            frmEquiposTendencias.PK_RANGO_MAX = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        Else
            frmEquiposTendencias.PK_RANGO_MAX = "0"
        End If
        frmEquiposTendencias.PK_TIPO = 2 ' Verificacion
        frmEquiposTendencias.Show 1
    End If

   On Error GoTo 0
   Exit Sub

cmdTendencia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdTendencia_Click of Formulario frmEquipoEdicionVerificacion"
End Sub

'Private Sub Form_Activate()
'
'    If blnPrimeraVez Then
'        grdResultados_BeforeColEdit ColsR.RESULTADO_MEDIA, 0, 0
'        blnPrimeraVez = False
'    End If
'
'End Sub

Private Sub Form_Load()

    log Me.Name
    cargar_botones Me
    Call cabecera
    
    
    comprobar_fecha_real_modificable
    
    If mvarblnVieneDeCuaderno Then
        Set mvarobjEquipo = New clsEquipos
        Call mvarobjEquipo.Carga_Datos_Basicos(mvarlngidEquipo)
        mvarenuTipoEdicion = EDICION
        mvarstrId = CStr(mvarlngIdEvento)
    End If
    
    Call LlenarCombos
    mvarlngidEquipo = mvarobjEquipo.getID_EQUIPO
    
    Call PresentarDatos_LimitacionesUso
    
'    blnPrimeraVez = False
        
'    Call ConfigurarTablet
    
    lbltitulo.Caption = "Verificación del Equipo " & CStr(mvarobjEquipo.getID_EQUIPO) & ": " & mvarobjEquipo.getNOMBRE
    
    If mvarenuTipoEdicion = Alta Then
    
        'txtFechaActual.value = mvardtmFechaProximaInicial
        txtFechaActual.Value = Now
        
        txtFechaActual.Enabled = bln_fecha_real_editable Or True
        'txtFechaProxima_b.value = calcularFechaProxima(mvardtmFechaProximaInicial, mvarlngIdPeriodoInicial)
        Set mvarobjVerificacion = New clsEquipoVerificacion
        cmbTipoVerificacion.BoundText = mvarlngIdTipoVerificacionIncial
        cmbPeriVerificacion.BoundText = mvarlngIdPeriodoInicial
        txtFechaActual_Change
        cmbVerificador.MostrarElemento mvarlngidVerificadorInternoInicial
        cmbProcedimiento.MostrarElemento mvarlngidProcedmientoInicial
        ' Si el tipo de periodicidad para carga viene informado, buscamos la ultima verificacion de ese tipo
        ' y la cargamos para generar una igual
        If copiarUltimaVerificacionPeriodo <> 0 Then
            Dim ULTIMA As Long
            'MANTIS-810-I
            Dim TiposVerificacion As String
            TiposVerificacion = str(ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO) & "," & str(ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO)
            TiposVerificacion = Trim(TiposVerificacion)
            'ULTIMA = mvarobjVerificacion.buscar_verificacion_anterior_misma_periodicidad(mvarobjEquipo.getID_EQUIPO, copiarUltimaVerificacionPeriodo)
            ULTIMA = mvarobjVerificacion.buscar_verificacion_anterior_misma_periodicidad(mvarobjEquipo.getID_EQUIPO, TiposVerificacion)
            'MANTIS-810-F
            
            If ULTIMA <> 0 Then
                mvarenuTipoEdicion = EDICION
                mvarstrId = CStr(ULTIMA)
            Else
                OpcionesEdicion
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    'Set mvarobjVerificacion = mvarobjEquipo.Verificaciones.Item(mvarstrId)
    mvarobjVerificacion.Carga CLng(mvarstrId)
    
    Call PresentarDatos
    Call PresentarDatos_ParametrosResultados
    
    If copiarUltimaVerificacionPeriodo <> 0 Then
        mvarenuTipoEdicion = Alta
        cmbVerificador.MostrarElemento USUARIO.getID_EMPLEADO
        optEstado(0).Value = True
    End If
    
    Call OpcionesEdicion
    
    ' Si no esta pendiente, ocultamos icono Hoja de certificado
    If optEstado(0).Value = False Then
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
            cmdMostrarAdjunto(1).visible = False
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub




'Private Sub grdResultados_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'If blnEsTablet And ColIndex = ColsR.RESULTADO_MEDIA Then
'    grdResultados.Col = ColIndex
'    TecladoNumerico.TextoInicial = grdResultados.Text
'    TecladoNumerico.cabecera = xR(grdResultados.Row, 0)
'    TecladoNumerico.Subcabecera = "Resultado" 'xP(gridP.Row, 1)
'
'    TecladoNumerico.Show 1
'    grdResultados.EditActive = False
'
'End If
'
'grdResultados_RowColChange 0, 0
'
'End Sub





'Private Sub grdResultados_KeyPress(KeyAscii As Integer)
'
'    With grdResultados
'        If .Col = 1 Or .Col = 2 Or .Col = 5 Or .Col = 1 Or .Col = 6 Or .Col = 7 Then
'            KeyAscii = KeyAscii_SoloDecimal_tbgrid(.Text, KeyAscii, True)
'        End If
'        If .Col = 1 Then
'            lblParametro.Caption = .Text
'        End If
'    End With
'
'
'
'End Sub

Public Property Get ID() As String

    ID = mvarstrId

End Property

Public Property Let ID(ByVal strId As String)

    mvarstrId = strId

End Property

Public Property Get idEquipo() As Long

    idEquipo = mvarlngidEquipo

End Property

Public Property Let idEquipo(ByVal lngidEquipo As Long)

    mvarlngidEquipo = lngidEquipo

End Property

Public Property Get IdEvento() As Long

    IdEvento = mvarlngIdEvento

End Property

Public Property Let IdEvento(ByVal lngIdEvento As Long)

    mvarlngIdEvento = lngIdEvento

End Property

Public Property Get IdPeriodoInicial() As Long

    IdPeriodoInicial = mvarlngIdPeriodoInicial

End Property

Public Property Let IdPeriodoInicial(ByVal lngIdPeriodoInicial As Long)

    mvarlngIdPeriodoInicial = lngIdPeriodoInicial

End Property

Public Property Get IdTipoVerificacionIncial() As Long

    IdTipoVerificacionIncial = mvarlngIdTipoVerificacionIncial

End Property

Public Property Let IdTipoVerificacionIncial(ByVal lngIdTipoVerificacionIncial As Long)

    mvarlngIdTipoVerificacionIncial = lngIdTipoVerificacionIncial

End Property

Public Property Get idVerificadorInternoInicial() As Long

    idVerificadorInternoInicial = mvarlngidVerificadorInternoInicial

End Property

Public Property Let idVerificadorInternoInicial(ByVal lngidVerificadorInternoInicial As Long)

    mvarlngidVerificadorInternoInicial = lngidVerificadorInternoInicial

End Property
Private Sub imprimir_etiqueta(strFecha_Verificacion As String, lngOperador_ID As Long)
On Error GoTo trataError
   
    With frmReport
        .iniciar
        .informe = "Equipos\rptEquipos_ETIQUETA_Verificacion"
        .criterio = "{eq_verificacion_equipos.ID_VERIFICACION} = " & CLng(PK)
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    log ("Final impresion de etiqueta de verificación de equipo")
    
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir la etiqueta de verificación.", vbCritical, Err.Description
End Sub

Private Sub lista_Click()
    Dim x As Integer
    
    If lista.ListItems.Count = 0 Then Exit Sub
    
    x = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
        
        mostrar_datos_medidas
    
'    If blnEsTablet Then
'            TecladoNumerico.cabecera = lista.ListItems(lista.selectedItem.Index).Text
'
'            If x = 0 Then
'                TecladoNumerico.Subcabecera = "Cualidad"
'                TecladoNumerico.TextoInicial = txtValor(x).Text
'                If optResultadoCualidad(0).value Then
'                    TecladoNumerico.chkConforme.value = vbUnchecked
'                    TecladoNumerico.chkNoConforme.value = vbChecked
'                Else
'                    TecladoNumerico.chkConforme.value = vbChecked
'                    TecladoNumerico.chkNoConforme.value = vbUnchecked
'                End If
'            Else
'                TecladoNumerico.Subcabecera = "Rango: " & lista.selectedItem.SubItems(ColsR.RANGO_MIN) & " - " & lista.selectedItem.SubItems(ColsR.RANGO_MAX) & " " & lista.selectedItem.SubItems(ColsR.Unidad)
'                TecladoNumerico.TextoInicial = txtValor(x).Text
'            End If
'        If Not TecladoNumerico.Visible Then
'            TecladoNumerico.Show 1
'        End If
'    End If
End Sub


Private Sub lista_DblClick()
If cmdAnadirParametro.Enabled Then
    ' solo deja modificar si los no está cerrada.
    modificar_parametro
End If
End Sub

Private Sub lista_medidas_Click(Index As Integer)

    txtValor(Index).Text = lista_medidas(Index).selectedItem.Text
    On Error Resume Next ' necesario porque lo ejecuta con el doble clic
    txtValor(Index).SetFocus
    On Error GoTo 0
    txtvalor_GotFocus (Index)
End Sub


Private Sub lstLimitacionesUso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdEliminarLimitacion_Click
End Sub

Private Sub LlenarCombos()
Dim oDeco As New clsDecodificadora

    oDeco.cargar_combo cmbPeriVerificacion, DECODIFICADORA.EQ_periodicidad
    oDeco.cargar_combo cmbTipoVerificacion, DECODIFICADORA.EQ_TIPO_CALIBRACION
    oDeco.cargar_mi_combo cmbUBICACION_ID, DECODIFICADORA.EQ_UBICACION_ID
    
    llenar_combo cmbVerificador, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmbProcedimiento, New clsCa_documentos, 0, frmCA_Documento, ""
    llenar_combo cmbVerificadorExterno, New clsProveedor, 0, frmProveedores_Detalle, ""
    
    
    oDeco.cargar_combo cmbTipoParametro(0), DECODIFICADORA.EQ_TIPOS_PARAMETROS_RESULTADO
    oDeco.cargar_combo cmbTipoParametro(1), DECODIFICADORA.EQ_TIPOS_PARAMETROS_RESULTADO
'    oDeco.cargar_combo cmbTipoParametro(2), decodificadora.EQ_TIPOS_PARAMETROS_RESULTADO
    
    llenar_combo cmbEquipos, New clsEquipos, 1, frmEquipoEdicion, " AND ESTADO_ID NOT IN ('B','F/S','R','I','E') "
'    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, "AND ABIERTO = 1"
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, " AND ABIERTO = 1 AND FINALIZADO = 0 "
'    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, ""
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
    
    If mvarobjEquipo.getTIPO_VERIFICACION_ID = 2 Then ' es Externa
        cmbVerificadorExterno.activar
    Else
        cmbVerificadorExterno.desactivar
    End If

    
End Sub

' ----------------- Funciones auxiliares del formulario ----------------

Public Property Get PK() As Long

    PK = mvarlngPK

End Property

Public Property Let PK(ByVal lngPK As Long)

    mvarlngPK = lngPK

End Property

Private Sub PresentarDatos()

On Error GoTo PresentarDatos_Error
    
    With mvarobjVerificacion
        cmbTipoVerificacion.BoundText = .getTIPO_ID
        cmbPeriVerificacion.BoundText = .getPERIODICIDAD_ID
        cmbVerificador.MostrarElemento .getVERIFICADOR_INTERNO_ID
        If .getVERIFICADOR_EXTERNO_ID > 0 Then
            cmbVerificadorExterno.MostrarElemento .getVERIFICADOR_EXTERNO_ID
        End If
        cmbProcedimiento.MostrarElemento .getPROCEDIMIENTO_ID
        cmbUBICACION_ID.MostrarElemento .getUBICACION_ID
        txtCalibradoEn = .getINCIDENCIAS
'        txtHojaVerificacion.Text = .HojaVerificacion.getNOMBRE_ARCHIVO
'        txtCertificado.Text = .CERTIFICADO.getNOMBRE_ARCHIVO
'        txtEvaluacionResultado.Text = .Evaluacion.getNOMBRE_ARCHIVO
        
        optEstado(.getESTADO).Value = True
        optResultado(.getRESULTADO).Value = True
        
        If mvarenuTipoEdicion = Alta Then
            'txtFechaActual.value = CDate(mvardtmFechaProximaInicial)
            txtFechaActual.Value = Now
            txtFechaActual_Change
            cmbPeriVerificacion.BoundText = CStr(mvarlngIdPeriodoInicial)
            cmbTipoVerificacion.BoundText = CStr(mvarlngIdTipoVerificacionIncial)
            'txtFechaProxima_b.value = calcularFechaProxima(mvardtmFechaProximaInicial, mvarlngIdPeriodoInicial)
        Else
            'If .getESTADO = 0 Then
            '    txtFechaActual.value = Now
            '    txtFechaActual_Change
            'Else
                If IsDate(.getFECHA_ACTUAL) Then
                    txtFechaActual.Value = CDate(.getFECHA_ACTUAL)
                Else
                    txtFechaActual.Value = CDate(Date)
                End If
                txtFechaActual_Change
                'txtFechaProxima_b.value = CDate(.getFECHA_PROXIMA)
                'txtFechaProxima.Text = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
            'End If
            txtAdjunto(1) = .getRUTA_PLANTILLA
            txtAdjunto(2) = .getRUTA_CERTIFICADO
            txtAdjunto(3) = .getRUTA_EVALUACION
        End If
        
        
    End With


'    Call PresentarDatos_Adjuntos

On Error GoTo 0
    Exit Sub
PresentarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmEquipoEdicionVerificacion_nuevo"

End Sub
'Private Sub PresentarDatos_Adjuntos()
'    Dim oD As New clsDocumentacion
'    Dim i As Integer
'    For i = 1 To 3
'        txtAdjunto(i) = oD.FicheroEquipo(mvarlngidEquipo, 1, CLng(mvarstrId), i)
'    Next
'    Set oD = Nothing
    
'    Dim obja As clsArchivoAdjunto
'
'    Set obja = mvarobjVerificacion.HojaVerificacion
'    If Not obja Is Nothing Then
'        txtHojaVerificacion.Text = IIf(obja.getNOMBRE_ARCHIVO_TEMP <> "", obja.getNOMBRE_ARCHIVO_TEMP, obja.getNOMBRE_ARCHIVO)
'    End If
'
'    Set obja = mvarobjVerificacion.CERTIFICADO
'    If Not obja Is Nothing Then
'        txtCertificado.Text = IIf(obja.getNOMBRE_ARCHIVO_TEMP <> "", obja.getNOMBRE_ARCHIVO_TEMP, obja.getNOMBRE_ARCHIVO)
'    End If
    
    
'End Sub

Private Sub PresentarDatos_LimitacionesUso()
    Dim objItem As clsGenericClass

    lstLimitacionesUso.Clear
    txtLimitacionesUso.Text = ""
        
    For Each objItem In mvarobjEquipo.getLIMITACIONES_USO_COL.Iterator
        If objItem.getID_AUX <> enumIdAux.ID_AUX_ELIMINADO Then
            lstLimitacionesUso.AddItem objItem.getNOMBRE
            lstLimitacionesUso.ItemData(lstLimitacionesUso.ListCount - 1) = objItem.getID
        End If
    Next objItem

End Sub

Private Sub PresentarDatos_ParametrosResultados()

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs_medidas As ADODB.Recordset
    On Error GoTo PresentarDatos_ParametrosResultados_Error

    i = 0
    mvarlngNumParametrosResultados = i
    lblParametro.Caption = ""
    lblParametro.visible = False
    If mvarenuTipoEdicion <> Alta Then
        ' Carga los Parametros de la verificacion
        
        Set rs = mvarobjVerificacion.DevolverParametrosResultados(mvarstrId)
        
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
            
                With lista.ListItems.Add(, , rs!DESCRIPCION)
                    .Checked = (CInt(rs!REALIZADO) = 1)
                    .SubItems(ColsR.Unidad) = CStr(rs!Unidad)
                    If CInt(rs("tipo_id")) = 0 Then
                        .SubItems(ColsR.RESULTADO_MEDIA) = IIf(CInt(rs("resultado_cualidad")) = 0, "NO CONFORME", "CONFORME")
                        .SubItems(ColsR.RANGO_MIN) = "N/A"
                        .SubItems(ColsR.RANGO_MAX) = "N/A"
                        .SubItems(ColsR.id_unidad) = "N/A"
                    Else
                        .SubItems(ColsR.RESULTADO_MEDIA) = CStr(rs("resultado"))
                        .SubItems(ColsR.RANGO_MIN) = CStr(rs("rango_min"))
                        .SubItems(ColsR.RANGO_MAX) = CStr(rs("rango_max"))
                        .SubItems(ColsR.id_unidad) = CStr(rs("unidad_ID"))
                    End If
                    .SubItems(ColsR.ID_TIPO) = CStr(rs("tipo_id"))
                    .SubItems(ColsR.id_patron) = CStr(rs("patron_id"))
                    .SubItems(ColsR.Id_resultado) = CStr(rs("id_resultado"))
                    .SubItems(ColsR.n_medidas) = CStr(rs("n_medidas"))
                    .SubItems(ColsR.RESULTADO_CUALIDAD) = CStr(rs("resultado_cualidad"))
                    .SubItems(ColsR.RESULTADOS_PATRON) = CStr(rs("DATOS_PARAMETRO"))
                    .SubItems(ColsR.LEQUIPOS) = CStr(rs("LEQUIPOS"))
                    .SubItems(ColsR.lReactivos) = CStr(rs("LREACTIVOS"))
                    .SubItems(ColsR.lreactivospropios) = CStr(rs("LREACTIVOS_PROPIOS"))
                    
                    ' si estamos realizando verificacion antes de ensayo y hemos copiado la verificacion anterior
                    ' ponemos el id_resultado a cero para que se inserte el nuevo.
                    If copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO Or _
                       copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
                    
                        .SubItems(ColsR.Id_resultado) = CStr(0)
                        .SubItems(ColsR.RESULTADO_MEDIA) = CStr("")
                    End If
                End With
                rs.MoveNext
            Wend
            ' se va al primero
            lista.selectedItem = lista.ListItems(1)
            lista_Click
        End If
        
    End If
  

On Error GoTo 0
    Exit Sub
PresentarDatos_ParametrosResultados_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_ParametrosResultados of Formulario frmEquipoEdicionVerificacion_nuevo"

End Sub

Private Sub RecogerDatos()

   On Error GoTo RecogerDatos_Error

    With mvarobjVerificacion
        ' A patir del 02.09.2010, PROPUESTA
        ' Ahora que hay verificaciones previstas, la fecha se modifica siempre que sea prevista, nunca en el caso de cerrada.
        ' cuando se cierra, siempre es el momento en que se cierra.
        ' de no ser así, el usuario (no es el caso de automaticamente al cerrar una calibracion, que se crea la siguiente prevista)
        ' no se podrían crear previstas más allá del presente
        
        ' La fecha la establece solo si se cierra ahora
        If .getESTADO = 0 Then
            .setFECHA_ACTUAL = Format(txtFechaActual.Value, "dd/mm/yyyy")
            .setFECHA_PROXIMA = Format(txtFechaProxima_b.Value, "dd/mm/yyyy")
        Else
            .setFECHA_ACTUAL = Format(Now, "dd/mm/yyyy")
            .setFECHA_PROXIMA = Format(txtFechaProxima_b.Value, "dd/mm/yyyy")
        End If
    
'        If .getESTADO = 0 Then
'            .setFECHA_ACTUAL = Format(txtFechaActual.value, "dd/mm/yyyy")
'            .setFECHA_PROXIMA = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
'        End If
'
        .setTIPO_ID = CLng(cmbTipoVerificacion.BoundText)
        .setPERIODICIDAD_ID = CLng(cmbPeriVerificacion.BoundText)
        .setVERIFICADOR_INTERNO_ID = cmbVerificador.getPK_SALIDA
        .setRESPONSABLE = cmbVerificador.getTEXTO
        If .getTIPO_ID = 2 Then
            .setVERIFICADOR_EXTERNO_ID = cmbVerificadorExterno.getPK_SALIDA
        Else
            .setVERIFICADOR_EXTERNO_ID = -1
        End If
        
        .setPROCEDIMIENTO_ID = cmbProcedimiento.getPK_SALIDA
        .setPROCEDIMIENTO = cmbProcedimiento.getTEXTO
        
        .setRUTA_PLANTILLA = txtAdjunto(1)
        .setRUTA_CERTIFICADO = txtAdjunto(2)
        .setRUTA_EVALUACION = txtAdjunto(3)
        
        .setUNIDADES_ID = 0 'cmbUnidad.getPK_SALIDA
        .setRANGO_MIN = 0
        .setRANGO_MAX = 0
        
        ' Estado
        If optEstado(0).Value = True Then
            .setESTADO = 0
        ElseIf optEstado(1).Value = True Then
            .setESTADO = 1
        ElseIf optEstado(2).Value = True Then
            .setESTADO = 2
        Else
            .setESTADO = 3
        End If
        ' Resultado
        If optResultado(0).Value = True Then
            .setRESULTADO = 0
        ElseIf optResultado(1).Value = True Then
            .setRESULTADO = 1
        Else
            .setRESULTADO = 2
        End If
        
        If .getID_AUX = enumIdAux.ID_AUX_EXISTE Then
            .setID_AUX = enumIdAux.ID_AUX_MODIFICADO
        End If
        .setUBICACION_ID = cmbUBICACION_ID.getPK_SALIDA
        .setINCIDENCIAS = txtCalibradoEn
        
        
    End With
    
    
    If mvarenuTipoEdicion = Alta Then
        mvarobjVerificacion.setFECHA_PREVISTA = mvarobjVerificacion.getFECHA_ACTUAL
'M1050        Call mvarobjEquipo.Verificaciones.Add(mvarobjVerificacion)
'M1050    ElseIf mvarenuTipoEdicion = edicion Then
'M1050        Call mvarobjEquipo.Verificaciones.Replace(mvarobjVerificacion.getID_VERIFICACION, mvarobjVerificacion)
    End If
    

   On Error GoTo 0
   Exit Sub

RecogerDatos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RecogerDatos of Formulario frmEquipoEdicionVerificacion"
    
End Sub

Public Property Get resultado() As Boolean

    resultado = mvarblnResultado

End Property

Public Property Let resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Get ID_VERIFICACION() As Long

    ID_VERIFICACION = mvarlngID_VERIFICACION

End Property

Public Property Let ID_VERIFICACION(ByVal verificacion As Long)

    mvarlngID_VERIFICACION = verificacion

End Property

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

Private Sub optEstado_Click(Index As Integer)

    If fraEstadoIntervencion.Enabled = False Then Exit Sub
    If mvarenuTipoEdicion = visualizar Then Exit Sub
    
    If Index = 0 Then
'M1130        txtFechaActual.Enabled = True
        If mvarobjVerificacion.getFECHA_ACTUAL <> "" Then
            txtFechaActual.Value = mvarobjVerificacion.getFECHA_ACTUAL
            txtFechaActual_Change
        End If
    Else
'M1130        If Not bln_fecha_real_editable Then
'M1130            txtFechaActual.Enabled = False
'M1130        End If
        txtFechaActual.Value = Now
        txtFechaActual_Change
        If cmbPeriVerificacion.BoundText <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO And _
           cmbPeriVerificacion.BoundText <> ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO Then
        
            MsgBox "La Fecha de Verificación al Cerrar se Establecerá a la de Hoy (" & Format(Now, "dd/mm/yyyy") & ")." & vbCrLf & "La fecha de Próxima Verificación se recalcula a " & txtFechaProxima.Text, vbInformation, "Verificación"
        End If
    End If
End Sub

Private Sub optResultadoCualidad_Click(Index As Integer)
    lista.selectedItem.SubItems(ColsR.RESULTADO_CUALIDAD) = Index
    lista.selectedItem.SubItems(ColsR.RESULTADO_MEDIA) = IIf(Index = 0, "NO CONFORME", "CONFORME")
    lista.selectedItem.SubItems(ColsR.n_medidas) = "1"
End Sub
'Private Sub TecladoNumerico_AnteriorElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'    Dim x As Integer, Pos As Long
'
'    x = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
'
'
'    If x = 0 Then
'        txtDescripcion_Cualidad_KeyPress 0
'    Else
'        txtvalor_KeyPress x, 0
'    End If
'End Sub
'
'Private Sub TecladoNumerico_Change(ByVal res As String)
'    'grdResultados.Text = res
'    Dim x As Integer
'    x = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
'    If x = 0 Then
'        txtDescripcion_Cualidad.Text = res
'    Else
'        txtValor(x).Text = res
'    End If
'End Sub
'
'Private Sub TecladoNumerico_EstablecerConformidad(ByVal VALOR As Integer)
'
'    Dim x As Integer
'    x = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
'    If x = 0 Then
'        If VALOR = 1 Then
'            optResultadoCualidad(1).value = True
'            optResultadoCualidad_Click 1
'        Else
'            optResultadoCualidad(0).value = True
'            optResultadoCualidad_Click 0
'        End If
'
'    End If
'
'
'End Sub
'
'Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'
'    Dim x As Integer, Pos As Long
'
'    x = CInt(lista.selectedItem.SubItems(ColsR.ID_TIPO))
'
'
'    If x = 0 Then
'        txtDescripcion_Cualidad_KeyPress 13
'    Else
'        txtvalor_KeyPress x, 13
'    End If
'
'
'
'
'End Sub
'
'
'Private Sub TecladoNumerico_SiguienteElemento_old(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
''If grdResultados.Row + 1 > filasR Then
''    TecladoNumerico.Hide
''    grdResultados.EditActive = False
''    Exit Sub
''End If
''
''' si existe siguiente Fila, edita la siguiente fila
''
''If (grdResultados.Row + 1) <= xR.UpperBound(1) Then
''    If Not IsEmpty(xR(grdResultados.Row + 1, 0)) Then
''        If Trim(xR(grdResultados.Row + 1, 0)) <> "" Then
''            grdResultados.EditActive = False
''            grdResultados.Row = grdResultados.Row + 1
''            resultado = grdResultados.Text
''            cabecera = xR(grdResultados.Row, 0)
''            fecha = xR(grdResultados.Row, 1)
''            grdResultados.EditActive = True
''        End If
''    ElseIf mvarlngNumParametrosResultados = 1 Then
''        grdResultados.Row = 1
''        Cerrar = True
''        grdResultados.EditActive = False
''    ElseIf grdResultados.Row = mvarlngNumParametrosResultados - 1 Or mvarlngNumParametrosResultados = 0 Then
''        'grdResultados.EditActive = False
''        'Resultado = grdResultados.Text
''        'cabecera = xP(grdResultados.Row, 0)
''        'grdResultados.EditActive = True
''        Cerrar = True
''    End If
''Else
''    If mvarlngNumParametrosResultados = 1 Then
''        grdResultados.Row = 1
''    Else
''        grdResultados.Row = 0
''    End If
''
''    Cerrar = True
''    grdResultados.EditActive = False
''End If
'End Sub


'Private Sub tUnidades_DropDownClose()
'    On Error Resume Next
'    grdResultados.Columns(ColsR.id_unidad) = tUnidades.Columns(1)
'    On Error GoTo 0
'
'    xR(grdResultados.Row, ColsR.id_unidad) = tUnidades.Columns(1)
'    grdResultados.Col = 3
'
'End Sub
'
'
Private Sub txtDescripcion_Cualidad_Change()
'If grdResultados.Row < 0 Then Exit Sub
If lista.ListItems.Count = 0 Then Exit Sub

lista.selectedItem.SubItems(ColsR.RESULTADOS_PATRON) = Trim(txtDescripcion_Cualidad.Text)

'xR(grdResultados.Row, ColsR.RESULTADOS_PATRON) = Trim(txtDescripcion_Cualidad.Text)
'On Error Resume Next
'grdResultados.Columns(ColsR.RESULTADOS_PATRON).RefetchCell grdResultados.Row
'On Error GoTo 0
End Sub

Private Sub txtDescripcion_Cualidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ' salta al siguiente campo en la lista
    saltar_siguiente_parametro
ElseIf KeyAscii = 0 Then
    saltar_anterior_parametro
End If
End Sub

Private Sub saltar_siguiente_parametro()
    If lista.selectedItem.Index = lista.ListItems.Count Then
        lista.selectedItem = lista.ListItems(1)
    Else
        lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
    End If
    lista_Click
End Sub

Private Sub saltar_anterior_parametro()
    If lista.selectedItem.Index = 1 Then
        lista.selectedItem = lista.ListItems(lista.ListItems.Count)
    Else
        lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
    End If
    lista_Click
End Sub



Private Sub txtFechaActual_Change()

If IsDate(txtFechaActual.Value) Then
    txtFechaProxima_b.Value = calcularFechaProxima(txtFechaActual.Value, getDataComboSel(cmbPeriVerificacion))
    txtFechaProxima.Text = Format(txtFechaProxima_b.Value, "dd/mm/yyyy")
End If

End Sub

Private Sub txtLimitacionesUso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAnadirLimitacion_Click
End Sub

Public Property Get VieneDeCuaderno() As Boolean

    VieneDeCuaderno = mvarblnVieneDeCuaderno

End Property

Public Property Let VieneDeCuaderno(ByVal blnVieneDeCuaderno As Boolean)

    mvarblnVieneDeCuaderno = blnVieneDeCuaderno

End Property


Public Property Get idProcedmientoInicial() As Long

    idProcedmientoInicial = mvarlngidProcedmientoInicial

End Property

Public Property Let idProcedmientoInicial(ByVal lngidProcedmientoInicial As Long)

    mvarlngidProcedmientoInicial = lngidProcedmientoInicial

End Property

Private Sub txtNMedidas_Change(Index As Integer)

txtNMedidas(Index).Locked = True

Dim n_medidas_act As Integer ' n_medidas_actuales
Dim n_medidas As Integer ' n medidas propuestas
Dim x As Integer

If Not IsNumeric(txtNMedidas(Index).Text) Then
    txtNMedidas(Index).Text = "1"
    txtNMedidas(Index).SelStart = 0
    txtNMedidas(Index).SelLength = 1
End If
If Trim(txtNMedidas(Index).Text) = "0" Then
    txtNMedidas(Index).Text = "1"
    txtNMedidas(Index).SelStart = 0
    txtNMedidas(Index).SelLength = 1
End If

n_medidas_act = lista_medidas(Index).ListItems.Count

n_medidas = CInt(txtNMedidas(Index).Text)

If n_medidas = n_medidas_act Then
    txtNMedidas(Index).Locked = False
    Exit Sub ' no se modifican
End If

If n_medidas > n_medidas_act Then ' se añaden filas
    For x = (n_medidas_act + 1) To n_medidas
        lista_medidas(Index).ListItems.Add , , "0"
    Next x
End If

If n_medidas < n_medidas_act Then ' se quitan filas
    For x = n_medidas_act To (n_medidas + 1) Step -1
        lista_medidas(Index).ListItems.Remove lista_medidas(Index).ListItems.Count
    Next x
End If

' guarda las medidas en su fila correspondiente
lista.selectedItem.SubItems(ColsR.n_medidas) = n_medidas
txtNMedidas(Index).Locked = False

End Sub



Private Sub txtNMedidas_GotFocus(Index As Integer)
    txtNMedidas(Index).SelStart = 0
    txtNMedidas(Index).SelLength = Len(txtNMedidas(Index).Text)
End Sub

Private Sub txtNMedidas_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtValor(Index).SetFocus
    Else
        KeyAscii = KeyAscii_SoloNumerico(txtNMedidas(Index), KeyAscii)
    End If
End Sub

Private Sub txtValor_Change(Index As Integer)
    lista_medidas(Index).ListItems(lista_medidas(Index).selectedItem.Index).Text = txtValor(Index).Text
    
    txtValor(Index).Locked = True
        preguardar_medidas Index
    txtValor(Index).Locked = False
End Sub


Private Sub txtvalor_GotFocus(Index As Integer)


If Trim(txtValor(Index).Text) = "" Then
    If lista_medidas(Index).ListItems.Count > 0 Then
        txtValor(Index).Text = lista_medidas(Index).selectedItem.Text
    End If
End If

txtValor(Index).SelStart = 0
txtValor(Index).SelLength = Len(txtValor(Index).Text)
End Sub


Private Sub txtvalor_KeyPress(Index As Integer, KeyAscii As Integer)
Dim NUMERO As String

    If KeyAscii = 13 Then ' Tecla Intro
        KeyAscii = 0
        
        ' Comprueba que sea un valor correcto numérico
        If Trim(txtValor(Index).Text) = "" Then
            NUMERO = "0"
'        ElseIf Not IsNumeric(Trim(txtvalor(Index).Text)) Then
'            numero = "0"
        Else
            NUMERO = Trim(txtValor(Index).Text)
        End If

        lista_medidas(Index).ListItems(lista_medidas(Index).selectedItem.Index).Text = NUMERO
        preguardar_medidas Index
        
        If lista_medidas(Index).selectedItem.Index + 1 > lista_medidas(Index).ListItems.Count Then
            ' estamos al final de la lista
            ' Vuelve a la primera
            'Set lista_medidas(Index).SelectedItem = lista_medidas(Index).ListItems(1)
            'lista_medidas_Click Index
            
            ' salta al siguiente parámetro
            saltar_siguiente_parametro
        Else
            ' avanza una linea
            Set lista_medidas(Index).selectedItem = lista_medidas(Index).ListItems(lista_medidas(Index).selectedItem.Index + 1)
            lista_medidas_Click Index
        End If
    ElseIf KeyAscii = 0 Then
        If Trim(txtValor(Index).Text) = "" Then
            NUMERO = "0"
'        ElseIf Not IsNumeric(Trim(txtValor(Index).Text)) Then
'            numero = "0"
        Else
            NUMERO = Trim(txtValor(Index).Text)
        End If

        lista_medidas(Index).ListItems(lista_medidas(Index).selectedItem.Index).Text = NUMERO
        preguardar_medidas Index
    
        If lista_medidas(Index).selectedItem.Index = lista_medidas(Index).ListItems.Count Then
            saltar_anterior_parametro
        Else
            ' atrasa una linea
            Set lista_medidas(Index).selectedItem = lista_medidas(Index).ListItems(lista_medidas(Index).selectedItem.Index - 1)
            lista_medidas_Click Index
        End If
        
        
'    Else
'        KeyAscii = KeyAscii_SoloDecimal(txtValor(Index), KeyAscii, True)
    End If
End Sub


Private Sub preguardar_medidas(Index As Integer)
    
    ' recoge los datos modificados y los guarda en su celda de resultados
    Dim cad As String, str_total As String
    Dim fila As Long
    Dim total As Single, total_fila As Single
    Dim n_medidas As Integer
    
    
    Dim no_numero As Boolean
    Dim no_numero_valor As String
    no_numero = False
    total = 0
    
    For fila = 1 To lista_medidas(Index).ListItems.Count
        total_fila = 0
        If Trim(lista_medidas(Index).ListItems(fila)) <> "" Then
            If IsNumeric(lista_medidas(Index).ListItems(fila)) Then
                total_fila = CSng(Replace(lista_medidas(Index).ListItems(fila), ".", ","))
            Else
                no_numero = True
                no_numero_valor = Trim(lista_medidas(Index).ListItems(fila))
            End If
        End If
        total = total + total_fila
        cad = cad & ";" & lista_medidas(Index).ListItems(fila)
    Next fila
    
    If cad <> "" Then cad = Mid(cad, 2)
    n_medidas = 1
    If Trim(txtNMedidas(Index).Text) <> "" Then
        If IsNumeric(txtNMedidas(Index).Text) Then
            n_medidas = txtNMedidas(Index).Text
        Else
            no_numero = True
        End If
    End If
    
    If no_numero Then
        str_total = no_numero_valor
    Else
        ' Evaluar TIPO 1 : Media / TIPO 2 : Mediana
        If cmbTipoParametro(Index).BoundText = 1 Then
            Dim t As Single
            t = Format(total / n_medidas, "##0.000000")
            str_total = CStr(t)
        ElseIf cmbTipoParametro(Index).BoundText = 2 Then
            Dim arr() As Single
            Dim p As Integer
            p = 1
            For fila = 1 To lista_medidas(Index).ListItems.Count
                If Trim(lista_medidas(Index).ListItems(fila)) <> "" Then
                    If IsNumeric(lista_medidas(Index).ListItems(fila)) Then
                        ReDim Preserve arr(p)
                        arr(p) = CSng(Replace(lista_medidas(Index).ListItems(fila), ".", ","))
                        p = p + 1
                    End If
                End If
            Next
            str_total = Mediana(arr)
        End If
    End If
    lista.selectedItem.SubItems(ColsR.RESULTADOS_PATRON) = cad
    lista.selectedItem.SubItems(ColsR.RESULTADO_MEDIA) = str_total
    
End Sub

Private Sub informar_equipos_lista()
    Dim i As Integer
    Dim s As String
    For i = 1 To listaEquipos.ListItems.Count
        s = s & listaEquipos.ListItems(i).Text & ";"
    Next
    lista.selectedItem.SubItems(ColsR.LEQUIPOS) = s
End Sub
Private Sub informar_reactivos_lista()
    Dim i As Integer
    Dim externos As String
    Dim internos As String
    For i = 1 To listaReactivos.ListItems.Count
        If listaReactivos.ListItems(i).SubItems(3) = "E" Then
            externos = externos & listaReactivos.ListItems(i).Text & ";"
        Else
            internos = internos & listaReactivos.ListItems(i).Text & ";"
        End If
    Next
    lista.selectedItem.SubItems(ColsR.lReactivos) = externos
    lista.selectedItem.SubItems(ColsR.lreactivospropios) = internos
End Sub

Private Sub cargar_equipo(EQUIPO As Long)
    Dim oEquipo As New clsEquipos
    oEquipo.Carga_Datos_Basicos EQUIPO
    With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
        .SubItems(1) = oEquipo.getNOMBRE
        .SubItems(2) = oEquipo.getSERIE
    End With
    listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
    listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
End Sub
Private Sub cargar_reactivos(lREACTIVO As String, tipo As Integer)
    ' REACTIVOS EXTERNOS
    Dim i As Integer
    If tipo = 0 Then
        Dim REACTIVOS() As String
        Dim oReactivo As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
        Dim oTR As New clsTipos_reactivo_ex
        REACTIVOS = Split(lREACTIVO, ";")
        For i = LBound(REACTIVOS) To UBound(REACTIVOS) - 1
            oReactivo.CARGAR CLng(REACTIVOS(i))
            oTb.CARGAR oReactivo.getTIPO_BOTE_EX_ID
            oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
            With listaReactivos.ListItems.Add(, , REACTIVOS(i))
                .SubItems(1) = oTR.getNOMBRE
                .SubItems(2) = Format(oReactivo.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                .SubItems(3) = "E"
            End With
        Next
    End If
    ' REACTIVOS PROPIOS
    If tipo = 1 Then
        Dim REACTIVOS_PROPIOS() As String
        Dim oRPR As New clsRpr_botes
        Dim oTRPR As New clsRPR_Tipos
        REACTIVOS_PROPIOS = Split(lREACTIVO, ";")
        For i = LBound(REACTIVOS_PROPIOS) To UBound(REACTIVOS_PROPIOS) - 1
            oRPR.Carga CLng(REACTIVOS_PROPIOS(i))
            oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
            With listaReactivos.ListItems.Add(, , REACTIVOS_PROPIOS(i))
                .SubItems(1) = oTRPR.getNOMBRE
                .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                .SubItems(3) = "I"
            End With
        Next
    End If
End Sub
Private Function Mediana(ByRef arr() As Single) As Single
    Dim lngElementos As Single, lngMedio As Single
    lngElementos = UBound(arr) - LBound(arr)
    lngMedio = LBound(arr) + (lngElementos \ 2)
    If lngElementos And 1 Then
        Mediana = arr(lngMedio)
    Else
        Mediana = (arr(lngMedio) + arr(lngMedio - 1)) / 2
    End If
End Function
