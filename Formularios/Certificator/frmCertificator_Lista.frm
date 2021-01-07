VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.2#0"; "Codejock.CommandBars.v13.2.1.ocx"
Begin VB.Form frmCertificator_Lista 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestor de Informes"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCertificator_Lista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   17655
   WindowState     =   1  'Minimized
   Begin XtremeSuiteControls.ListView lista 
      Height          =   7170
      Left            =   45
      TabIndex        =   1
      Top             =   1170
      Width           =   9150
      _Version        =   851970
      _ExtentX        =   16140
      _ExtentY        =   12647
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HideSelection   =   0   'False
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      FlatScrollBar   =   -1  'True
      AllowColumnReorder=   -1  'True
      Appearance      =   2
      IconSize        =   16
   End
   Begin VB.Timer Timer1 
      Left            =   13950
      Top             =   8595
   End
   Begin VB.Frame frmOpciones 
      BackColor       =   &H00FFFFFF&
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
      Height          =   810
      Left            =   90
      TabIndex        =   0
      Top             =   8415
      Width           =   11865
      Begin XtremeSuiteControls.PushButton cmdAbrir 
         Height          =   510
         Left            =   90
         TabIndex        =   37
         Top             =   180
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Abrir Certificado"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdDesbloquear 
         Height          =   510
         Left            =   3960
         TabIndex        =   38
         Top             =   180
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Desbloquear"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":0CFC
      End
      Begin XtremeSuiteControls.PushButton cmdFinalizar 
         Height          =   510
         Left            =   5895
         TabIndex        =   39
         Top             =   180
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Finalizar"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":0F71
      End
      Begin XtremeSuiteControls.PushButton cmdBloquear 
         Height          =   510
         Left            =   2025
         TabIndex        =   40
         Top             =   180
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Bloquear"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":11A6
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   510
         Left            =   7830
         TabIndex        =   41
         Top             =   180
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Visualizar PDF"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":1422
      End
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   510
         Left            =   9765
         TabIndex        =   35
         Top             =   180
         Width           =   1905
         _Version        =   851970
         _ExtentX        =   3360
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":16BE
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   9150
      _Version        =   851970
      _ExtentX        =   16140
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Filtro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   7110
         Picture         =   "frmCertificator_Lista.frx":193B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   270
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   7110
         Picture         =   "frmCertificator_Lista.frx":818D
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   495
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7110
         Picture         =   "frmCertificator_Lista.frx":E9DF
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   720
         Width           =   240
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   240
         Index           =   0
         Left            =   5985
         TabIndex        =   6
         Top             =   270
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Pendiente"
         Appearance      =   2
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cmbTipoEquipo 
         Height          =   360
         Left            =   1125
         TabIndex        =   7
         Top             =   225
         Width           =   4560
         _Version        =   851970
         _ExtentX        =   8043
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   2
         UseVisualStyle  =   0   'False
         AutoComplete    =   -1  'True
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   240
         Index           =   1
         Left            =   5985
         TabIndex        =   8
         Top             =   495
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Bloqueado"
         Appearance      =   2
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   240
         Index           =   2
         Left            =   5985
         TabIndex        =   9
         Top             =   720
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Finalizado"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbEmpleado 
         Bindings        =   "frmCertificator_Lista.frx":15231
         Height          =   360
         Left            =   1125
         TabIndex        =   10
         Top             =   630
         Width           =   4560
         _Version        =   851970
         _ExtentX        =   8043
         _ExtentY        =   635
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         AutoComplete    =   -1  'True
         DropDownItemCount=   20
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   825
         Left            =   7740
         TabIndex        =   11
         Top             =   180
         Width           =   1320
         _Version        =   851970
         _ExtentX        =   2328
         _ExtentY        =   1455
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   14737632
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":1523C
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   330
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Tipo Equipo"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   330
         Index           =   6
         Left            =   90
         TabIndex        =   12
         Top             =   630
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Empleado"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8205
      Left            =   9225
      TabIndex        =   14
      Top             =   135
      Width           =   8385
      _Version        =   851970
      _ExtentX        =   14790
      _ExtentY        =   14473
      _StockProps     =   68
      Appearance      =   2
      Color           =   64
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   100
      ItemCount       =   2
      Item(0).Caption =   "Detalle"
      Item(0).ControlCount=   21
      Item(0).Control(0)=   "Label2(1)"
      Item(0).Control(1)=   "Label2(0)"
      Item(0).Control(2)=   "Label2(2)"
      Item(0).Control(3)=   "Label2(3)"
      Item(0).Control(4)=   "Label2(4)"
      Item(0).Control(5)=   "txtdatos(0)"
      Item(0).Control(6)=   "txtdatos(1)"
      Item(0).Control(7)=   "txtdatos(2)"
      Item(0).Control(8)=   "txtdatos(3)"
      Item(0).Control(9)=   "txtdatos(4)"
      Item(0).Control(10)=   "txtdatos(5)"
      Item(0).Control(11)=   "Label2(7)"
      Item(0).Control(12)=   "txtdatos(6)"
      Item(0).Control(13)=   "Label2(8)"
      Item(0).Control(14)=   "txtdatos(7)"
      Item(0).Control(15)=   "Label2(9)"
      Item(0).Control(16)=   "txtdatos(8)"
      Item(0).Control(17)=   "Label2(10)"
      Item(0).Control(18)=   "icoEstado"
      Item(0).Control(19)=   "PushButton2"
      Item(0).Control(20)=   "PushButton3"
      Item(1).Caption =   "Vida"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "listaVida"
      Begin XtremeSuiteControls.ListView listaVida 
         Height          =   7755
         Left            =   -69955
         TabIndex        =   15
         Top             =   405
         Visible         =   0   'False
         Width           =   8295
         _Version        =   851970
         _ExtentX        =   14631
         _ExtentY        =   13679
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HideSelection   =   0   'False
         View            =   3
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   2
      End
      Begin VB.PictureBox icoEstado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7965
         Picture         =   "frmCertificator_Lista.frx":154A1
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   16
         Top             =   4455
         Width           =   240
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   0
         Left            =   2115
         TabIndex        =   17
         Top             =   585
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   1
         Left            =   2115
         TabIndex        =   18
         Top             =   990
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   2
         Left            =   2115
         TabIndex        =   19
         Top             =   1395
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   3
         Left            =   2115
         TabIndex        =   20
         Top             =   1800
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   4
         Left            =   2115
         TabIndex        =   21
         Top             =   2205
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   5
         Left            =   2115
         TabIndex        =   22
         Top             =   2610
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   6
         Left            =   2115
         TabIndex        =   23
         Top             =   3015
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   7
         Left            =   2115
         TabIndex        =   24
         Top             =   4005
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   8
         Left            =   2115
         TabIndex        =   25
         Top             =   4410
         Width           =   6180
         _Version        =   851970
         _ExtentX        =   10901
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   510
         Left            =   180
         TabIndex        =   42
         Top             =   5130
         Width           =   1995
         _Version        =   851970
         _ExtentX        =   3519
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Consulta Equipo"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":1BCF3
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   510
         Left            =   180
         TabIndex        =   43
         Top             =   5715
         Width           =   1995
         _Version        =   851970
         _ExtentX        =   3519
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Consulta Calibración"
         Appearance      =   2
         Picture         =   "frmCertificator_Lista.frx":22555
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   1
         Left            =   225
         TabIndex        =   34
         Top             =   540
         Width           =   1995
         _Version        =   851970
         _ExtentX        =   3519
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Instrumento"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   33
         Top             =   990
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Número Equipo"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   32
         Top             =   1395
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Fabricante"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   3
         Left            =   225
         TabIndex        =   31
         Top             =   1845
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Modelo"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   4
         Left            =   225
         TabIndex        =   30
         Top             =   2250
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Número Serie"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   7
         Left            =   225
         TabIndex        =   29
         Top             =   2655
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Tipo de Equipo"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   8
         Left            =   225
         TabIndex        =   28
         Top             =   3060
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Plantilla"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   9
         Left            =   225
         TabIndex        =   27
         Top             =   4050
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "BLOQUEADO POR"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   10
         Left            =   225
         TabIndex        =   26
         Top             =   4455
         Width           =   1860
         _Version        =   851970
         _ExtentX        =   3281
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "ESTADO"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton cmdminimizar 
      Height          =   510
      Left            =   15795
      TabIndex        =   36
      Top             =   8595
      Width           =   1815
      _Version        =   851970
      _ExtentX        =   3201
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Minimizar"
      Appearance      =   2
      Picture         =   "frmCertificator_Lista.frx":28DB7
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   13500
      Top             =   8595
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   12510
      Top             =   8685
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCertificator_Lista.frx":29007
   End
End
Attribute VB_Name = "frmCertificator_Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub cmbEmpleado_Change()
    cargar_lista
End Sub
Private Sub cmbTipoEquipo_Change()
    cargar_lista
End Sub

Private Sub cmdAbrir_Click()
   On Error GoTo cmdAbrir_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.abrirCertificado CLng(lista.ListItems(lista.selectedItem.Index).Text), Winsock1.LocalHostName
        refrescar_lista
    End If
   On Error GoTo 0
   Exit Sub

cmdAbrir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAbrir_Click of Formulario frmMain"

End Sub

Private Sub cmdBloquear_Click()
   On Error GoTo cmdBloquear_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.bloquear CLng(lista.ListItems(lista.selectedItem.Index).Text), Winsock1.LocalHostName
        refrescar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdBloquear_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdBloquear_Click of Formulario frmCertificator_Lista"

End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdDesbloquear_Click()
   On Error GoTo cmdDesbloquear_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.desbloquear CLng(lista.ListItems(lista.selectedItem.Index).Text)
        refrescar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdDesbloquear_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDesbloquear_Click of Formulario frmCertificator_Lista"

End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.Eliminar CLng(lista.ListItems(lista.selectedItem.Index).Text)
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmCertificator_Lista"

End Sub

Private Sub cmdFinalizar_Click()
   On Error GoTo cmdFinalizar_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.finalizar CLng(lista.ListItems(lista.selectedItem.Index).Text)
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdFinalizar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFinalizar_Click of Formulario frmCertificator_Lista"

End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub
Private Sub Form_Load()
    log Me.Name
    Me.Left = (frmMenu.ScaleWidth - Me.Width) / 2
    Me.top = (frmMenu.ScaleHeight - Me.Height) / 2
    cargar_botones Me
    cabecera
    cargar_combos
    cargar_lista
    
    ' Crear carpeta si no existe
    On Error Resume Next
    Dim fichero As String
    Dim RUTA As String
    RUTA = App.Path & "\" & RUTA_CERTIFICADOS & "\"
    MkDir RUTA
End Sub
Private Sub cargar_combos()
    Dim consulta As String
    Dim rs As ADODB.Recordset
    
    consulta = "SELECT ID_EMPLEADO,CONCAT(NOMBRE,' ',APELLIDOS) AS NOMBRE FROM geslab_canagrosa.usuarios where anulado = 0 order by NOMBRE"
    Set rs = datos_bd_metrologia(consulta)
    Do Until rs.EOF
        cmbEmpleado.AddItem rs(1)
        cmbEmpleado.ItemData(cmbEmpleado.ListCount - 1) = rs(0)
        rs.MoveNext
    Loop
    
    consulta = "SELECT VALOR, DESCRIPCION FROM geslab_canagrosa.decodificadora WHERE CODIGO = " & C_TIPOS_EQUIPO & " ORDER BY DESCRIPCION "
        
    Set rs = datos_bd(consulta)
    Do Until rs.EOF
        cmbTipoEquipo.AddItem rs(1)
        cmbTipoEquipo.ItemData(cmbTipoEquipo.ListCount - 1) = rs(0)
        rs.MoveNext
    Loop
    
    Set rs = Nothing

End Sub

Private Sub cabecera()
    lista.Icons.AddIcons ImageManager1.Icons
'    lista.Icons.LoadBitmap App.Path & "\icons\0.ico", 0, xtpImageNormal
'    lista.Icons.LoadBitmap App.Path & "\icons\1.ico", 1, xtpImageNormal
'    lista.Icons.LoadBitmap App.Path & "\icons\2.ico", 2, xtpImageNormal
'    lista.Icons.LoadBitmap App.Path & "\icons\3.ico", 3, xtpImageNormal
    With lista.ColumnHeaders
        .Add , , "", 300, lvwColumnCenter, 1
        .Add , , "Tipo", 1, lvwColumnCenter
        .Add , , "Calibración", 1000, lvwColumnCenter
        .Add , , "Id.Equipo", 1000, lvwColumnCenter
        .Add , , "Descripción", 3000, lvwColumnCenter
        .Add , , "Num.Equipo", 1200, lvwColumnCenter
        .Add , , "Empleado", 1900, lvwColumnCenter
        .Add , , "Plantilla", 1, lvwColumnCenter
        .Add , , "Plantilla", 1, lvwColumnCenter
        .Add , , "Fecha Registro", 1, lvwColumnCenter
        .Add , , "Bloqueado", 1, lvwColumnCenter
        .Add , , "Estado", 1, lvwColumnCenter
    End With
    With listaVida.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 2100, lvwColumnCenter
        .Add , , "Empleado", 2800, lvwColumnCenter
        .Add , , "Estado", 2000, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista_vida(TOBJETO As Long, COBJETO As Long)
    listaVida.ListItems.Clear
    Dim oC As New clsCertificator_vida
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_lista_vida_Error

    Set rs = oC.Listado(TOBJETO, COBJETO)
    listaVida.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With listaVida.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_lista_vida_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista_vida of Formulario frmMain"
End Sub
Private Sub cargar_lista()
    On Error GoTo fallo
    Dim oC As New clsCertificator
    Dim rs As ADODB.Recordset
    ' Filtro
    Dim TIPO_EQUIPO As Integer
    Dim EMPLEADO As Integer
    Dim ESTADO As String
    TIPO_EQUIPO = 0
    EMPLEADO = 0
    ESTADO = ""
    If cmbTipoEquipo.ListIndex >= 0 Then
        TIPO_EQUIPO = cmbTipoEquipo.ItemData(cmbTipoEquipo.ListIndex)
    End If
    If cmbEmpleado.ListIndex >= 0 Then
        EMPLEADO = cmbEmpleado.ItemData(cmbEmpleado.ListIndex)
    End If
    Dim i As Integer
    For i = 0 To 2
        If CheckBox1(i).Value = xtpChecked Then
            If ESTADO <> "" Then
                ESTADO = ESTADO & ","
            End If
            ESTADO = ESTADO & i
        End If
    Next

    Set rs = oC.Listado(0, TIPO_EQUIPO, EMPLEADO, ESTADO)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0), rs("ESTADO_ID") + 1)
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(7)
             .SubItems(8) = rs(8)
             .SubItems(9) = rs(9)
             .SubItems(10) = rs(10)
             .SubItems(11) = rs(11)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Exit Sub
fallo:
    MsgBox "Error al recuperar los datos de la lista.", vbCritical, App.Title
End Sub
Private Sub refrescar_lista()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oC As New clsCertificator
    Dim rs As ADODB.Recordset
    Set rs = oC.Listado(lista.ListItems(lista.selectedItem.Index), 0, 0, 0)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems(lista.selectedItem.Index)
            .Icon = rs("ESTADO_ID") + 1
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(7)
             .SubItems(8) = rs(8)
             .SubItems(9) = rs(9)
             .SubItems(10) = rs(10)
             .SubItems(11) = rs(11)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    cargarDetalle
    Exit Sub
fallo:
    MsgBox "Error al recuperar los datos de la lista.", vbCritical, App.Title
End Sub

Private Sub lista_DblClick()
    If Not lista.selectedItem Is Nothing Then
        cargarDetalle
    End If
End Sub
Private Sub cargarDetalle()
    If Not lista.selectedItem Is Nothing Then
        ' Equipo
        Dim oC As New clsCertificator
        If oC.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
            Dim oEquipo As New clsEquipos
            If oEquipo.Carga(oC.getEQUIPO_ID) Then
                txtdatos(0) = oEquipo.getNOMBRE
                txtdatos(1) = oEquipo.getID_EQUIPO
                If oEquipo.getNUMERO_EQUIPO_CLIENTE <> "" Then
                    txtdatos(1) = txtdatos(1) & " (" & oEquipo.getNUMERO_EQUIPO_CLIENTE & ")"
                End If
                txtdatos(2) = oEquipo.getFABRICANTE
                txtdatos(3) = oEquipo.getMODELO
                txtdatos(4) = oEquipo.getSERIE
                
                txtdatos(5) = oC.getEQUIPO_TIPO_EQUIPO
                txtdatos(6) = oC.getPLANTILLA
            End If
            txtdatos(7) = oC.getPC
            txtdatos(8) = decoEstado(oC.getESTADO_ID)
'            Set icoEstado.Picture = LoadPicture(App.Path & "\icons\" & oC.getESTADO_ID & ".ico")
            Set icoEstado.Picture = ImageManager1.Icons.GetImage(oC.getESTADO_ID + 1, 16).CreatePicture(xtpImageNormal)
        End If
        ' Cargar Vida
        cargar_lista_vida lista.ListItems(lista.selectedItem.Index).SubItems(1), lista.ListItems(lista.selectedItem.Index).SubItems(2)
    End If
End Sub
Private Function decoEstado(ESTADO) As String
    Select Case ESTADO
    Case 0
        decoEstado = "PENDIENTE"
    Case 1
        decoEstado = "BLOQUEADO"
    Case 2
        decoEstado = "FINALIZADO"
    End Select
End Function

Private Sub lista_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    lista_DblClick
End Sub

Private Sub lista_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    End If
End Sub
Private Sub Text1_LostFocus()
    Text3 = Text1
End Sub

Private Sub PushButton1_Click()
   On Error GoTo PushButton1_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.generarCertificadoPDF (CLng(lista.ListItems(lista.selectedItem.Index).Text)), True
    End If

   On Error GoTo 0
   Exit Sub

PushButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton1_Click of Formulario frmCertificator_Lista"
End Sub

Private Sub PushButton2_Click()
   On Error GoTo PushButton2_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oC As New clsCertificator
        If oC.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
    
            Dim objfrm As New frmEquipoEdicion
            Dim objEquipo As New clsEquipos
            Call objEquipo.Carga(oC.getEQUIPO_ID)
            Set objfrm.EQUIPO = objEquipo
            If objEquipo.getALTA_BAJA = 1 Then
                objfrm.TipoEdicion = visualizar
            Else
                objfrm.TipoEdicion = EDICION
            End If
            objfrm.Show vbModal
            Unload objfrm
            Set objfrm = Nothing
        End If
    End If

   On Error GoTo 0
   Exit Sub

PushButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton2_Click of Formulario frmCertificator_Lista"
End Sub

Private Sub PushButton3_Click()
   On Error GoTo PushButton3_Click_Error

    If Not lista.selectedItem Is Nothing Then
        Dim oC As New clsCertificator
        If oC.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
    
            Dim objfrm  As New frmEquipoEdicionCalibracion
            Dim objEquipo As New clsEquipos
            Call objEquipo.Carga(oC.getEQUIPO_ID)
            
            With objfrm
                Set .EQUIPO = objEquipo
                .ID = oC.getCOBJETO
                .TipoEdicion = EDICION
                .Show vbModal
            End With
            Unload objfrm
            Set objfrm = Nothing

        End If
    End If

   On Error GoTo 0
   Exit Sub

PushButton3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton3_Click of Formulario frmCertificator_Lista"
End Sub
