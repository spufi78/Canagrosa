VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Certificator"
   ClientHeight    =   9660
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   17685
   DrawWidth       =   10
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   17685
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lista 
      Height          =   7170
      Left            =   45
      TabIndex        =   8
      Top             =   1890
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   45
      TabIndex        =   9
      Top             =   765
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
         Index           =   2
         Left            =   7110
         Picture         =   "frmMain.frx":6852
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         Top             =   720
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
         Picture         =   "frmMain.frx":D0A4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
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
         Index           =   0
         Left            =   7110
         Picture         =   "frmMain.frx":138F6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   27
         Top             =   270
         Width           =   240
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   240
         Index           =   0
         Left            =   5985
         TabIndex        =   26
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
         TabIndex        =   24
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
         TabIndex        =   28
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
         TabIndex        =   30
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
         Bindings        =   "frmMain.frx":1A148
         Height          =   360
         Left            =   1125
         TabIndex        =   32
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
         TabIndex        =   34
         Top             =   180
         Width           =   1320
         _Version        =   851970
         _ExtentX        =   2328
         _ExtentY        =   1455
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   14737632
         Appearance      =   2
         Picture         =   "frmMain.frx":1A153
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   330
         Index           =   6
         Left            =   90
         TabIndex        =   33
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   330
         Index           =   5
         Left            =   90
         TabIndex        =   25
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
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8205
      Left            =   9225
      TabIndex        =   7
      Top             =   855
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
      Item(0).ControlCount=   19
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
      Item(1).Caption =   "Vida"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "listaVida"
      Begin XtremeSuiteControls.ListView listaVida 
         Height          =   7755
         Left            =   -69955
         TabIndex        =   10
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
         Picture         =   "frmMain.frx":1A3B8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   43
         Top             =   4455
         Width           =   240
      End
      Begin XtremeSuiteControls.FlatEdit txtdatos 
         Height          =   375
         Index           =   0
         Left            =   2115
         TabIndex        =   12
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   35
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
         TabIndex        =   37
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
         TabIndex        =   39
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
         TabIndex        =   41
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   10
         Left            =   225
         TabIndex        =   42
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   285
         Index           =   9
         Left            =   225
         TabIndex        =   40
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
         Index           =   8
         Left            =   225
         TabIndex        =   38
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
         Index           =   7
         Left            =   225
         TabIndex        =   36
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
         Index           =   4
         Left            =   225
         TabIndex        =   16
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
         Index           =   3
         Left            =   225
         TabIndex        =   15
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
         Index           =   2
         Left            =   225
         TabIndex        =   14
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
         Index           =   0
         Left            =   225
         TabIndex        =   13
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
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   585
         Width           =   1995
         _Version        =   851970
         _ExtentX        =   3519
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Instrumento"
         BackColor       =   -2147483643
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAbrir 
      Height          =   510
      Left            =   45
      TabIndex        =   2
      Top             =   9090
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Abrir Certificado"
      Appearance      =   2
      Picture         =   "frmMain.frx":20C0A
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   13005
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstActive 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2595
      Left            =   15840
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Timer Timer1 
      Left            =   13455
      Top             =   135
   End
   Begin XtremeSuiteControls.PushButton cmdDesbloquear 
      Height          =   510
      Left            =   3915
      TabIndex        =   3
      Top             =   9090
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Desbloquear"
      Appearance      =   2
      Picture         =   "frmMain.frx":2103C
   End
   Begin XtremeSuiteControls.PushButton cmdFinalizar 
      Height          =   510
      Left            =   5850
      TabIndex        =   4
      Top             =   9090
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Finalizar"
      Appearance      =   2
      Picture         =   "frmMain.frx":212B1
   End
   Begin XtremeSuiteControls.PushButton cmdSalir 
      Height          =   510
      Left            =   16245
      TabIndex        =   5
      Top             =   9090
      Width           =   1365
      _Version        =   851970
      _ExtentX        =   2408
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Salir"
      Appearance      =   2
      Picture         =   "frmMain.frx":214E6
   End
   Begin XtremeSuiteControls.PushButton cmdBloquear 
      Height          =   510
      Left            =   1980
      TabIndex        =   6
      Top             =   9090
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Bloquear"
      Appearance      =   2
      Picture         =   "frmMain.frx":216AF
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   510
      Left            =   7785
      TabIndex        =   21
      Top             =   9090
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Visualizar PDF"
      Appearance      =   2
      Picture         =   "frmMain.frx":2192B
   End
   Begin XtremeSuiteControls.PushButton cmdminimizar 
      Height          =   510
      Left            =   14400
      TabIndex        =   22
      Top             =   9090
      Width           =   1815
      _Version        =   851970
      _ExtentX        =   3201
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Minimizar"
      Appearance      =   2
      Picture         =   "frmMain.frx":21BC7
   End
   Begin XtremeSuiteControls.PushButton cmdEliminar 
      Height          =   510
      Left            =   9720
      TabIndex        =   23
      Top             =   9090
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Eliminar"
      Appearance      =   2
      Picture         =   "frmMain.frx":21E17
   End
   Begin VB.Image imagen 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   15210
      Picture         =   "frmMain.frx":22094
      Stretch         =   -1  'True
      Top             =   135
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generador de Certificados v.1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   16
      Left            =   5760
      TabIndex        =   1
      Top             =   180
      Width           =   5775
   End
   Begin XtremeSuiteControls.TrayIcon TrayIcon1 
      Left            =   12690
      Top             =   180
      _Version        =   851970
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   16
      Text            =   "GESLAB : Generador de Informes v2.0"
      Picture         =   "frmMain.frx":23B1A
   End
   Begin VB.Menu opMenu 
      Caption         =   "Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu opRestaurar 
         Caption         =   "Restaurar"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Minimized As Boolean

Private Sub CheckBox1_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub cmbEmpleado_Change()
    cargar_lista
End Sub

Private Sub cmbEmpleado_Click()
    cargar_lista
End Sub

Private Sub cmbTipoEquipo_Change()
    cargar_lista
End Sub

Private Sub cmbTipoEquipo_Click()
    cargar_lista
End Sub

Private Sub cmdAbrir_Click()
   On Error GoTo cmdAbrir_Click_Error

    If Not lista.SelectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.abrirCertificado CLng(lista.ListItems(lista.SelectedItem.Index).Text)
        refrescar_lista
    End If
   On Error GoTo 0
   Exit Sub

cmdAbrir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAbrir_Click of Formulario frmMain"
End Sub

Private Sub cmdBloquear_Click()
   On Error GoTo cmdBloquear_Click_Error

    If Not lista.SelectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.bloquear CLng(lista.ListItems(lista.SelectedItem.Index).Text)
        refrescar_lista
    End If
   On Error GoTo 0
   Exit Sub

cmdBloquear_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdBloquear_Click of Formulario frmMain"

End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdDesbloquear_Click()
   On Error GoTo cmdDesbloquear_Click_Error

    If Not lista.SelectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.desbloquear CLng(lista.ListItems(lista.SelectedItem.Index).Text)
        refrescar_lista
    End If
   On Error GoTo 0
   Exit Sub

cmdDesbloquear_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDesbloquear_Click of Formulario frmMain"
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If Not lista.SelectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.Eliminar CLng(lista.ListItems(lista.SelectedItem.Index).Text)
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmMain"

End Sub

Private Sub cmdFinalizar_Click()
    If Not lista.SelectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.finalizar CLng(lista.ListItems(lista.SelectedItem.Index).Text)
        cargar_lista
    End If
End Sub

Private Sub cmdminimizar_Click()
    If Not Minimized Then
        TrayIcon1.MinimizeToTray Me.Hwnd
        Minimized = True
    Else
        TrayIcon1.MaximizeFromTray Me.Hwnd
        Minimized = False
    End If
End Sub

Private Sub cmdSalir_Click()
    End
End Sub
Private Sub Form_Load()
    If App.PrevInstance = True Then
        MsgBox "CERTIFICATOR ya se encuentra en ejecución. Verifique la ejecución anterior.", vbInformation, App.Title
        End
    End If
'    If CrearConexionGlobal_metrologia = False Then
'        MsgBox "Error al crear la conexión global. Contacte con mantenimiento.", vbCritical, App.Title
'        End
'    End If
    Me.Caption = Me.Caption & " (Host: " & ReadINI(App.Path + "\config.ini", "server_metrologia", "ip") & " -> BD: " & ReadINI(App.Path + "\config.ini", "server_metrologia", "bd") & ")"
'    ib.mensaje = Me.Caption
    cabecera
    cargar_combos
    cargar_lista


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
        
    Set rs = datos_bd_metrologia(consulta)
    Do Until rs.EOF
        cmbTipoEquipo.AddItem rs(1)
        cmbTipoEquipo.ItemData(cmbTipoEquipo.ListCount - 1) = rs(0)
        rs.MoveNext
    Loop
    
    Set rs = Nothing

End Sub

Private Sub cabecera()
    lista.Icons.LoadBitmap App.Path & "\icons\0.ico", 0, xtpImageNormal
    lista.Icons.LoadBitmap App.Path & "\icons\1.ico", 1, xtpImageNormal
    lista.Icons.LoadBitmap App.Path & "\icons\2.ico", 2, xtpImageNormal
    lista.Icons.LoadBitmap App.Path & "\icons\3.ico", 3, xtpImageNormal
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
            With lista.ListItems.Add(, , rs(0), rs("ESTADO_ID"))
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
    Set rs = oC.Listado(lista.ListItems(lista.SelectedItem.Index), 0, 0, 0)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems(lista.SelectedItem.Index)
            .Icon = rs("ESTADO_ID")
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
Private Sub ib_Menu()
    On Error Resume Next
    PopupMenu opMenu(0)
End Sub

Private Sub lista_DblClick()
    If Not lista.SelectedItem Is Nothing Then
        cargarDetalle
    End If
End Sub
Private Sub cargarDetalle()
    If Not lista.SelectedItem Is Nothing Then
        ' Equipo
        Dim oC As New clsCertificator
        If oC.Carga(lista.ListItems(lista.SelectedItem.Index).Text) Then
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
            Set icoEstado.Picture = LoadPicture(App.Path & "\icons\" & oC.getESTADO_ID & ".ico")
        End If
        ' Cargar Vida
        cargar_lista_vida lista.ListItems(lista.SelectedItem.Index).SubItems(1), lista.ListItems(lista.SelectedItem.Index).SubItems(2)
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

Private Sub lista_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    End If
End Sub

Private Sub opMenu_Click(Index As Integer)
    Me.Visible = True
End Sub
Private Sub Text1_LostFocus()
    Text3 = Text1
End Sub

Private Sub PushButton1_Click()
    If Not lista.SelectedItem Is Nothing Then
        Dim oCertificator As New clsCertificator
        oCertificator.generarCertificadoPDF (CLng(lista.ListItems(lista.SelectedItem.Index).Text)), True
    End If
End Sub

Private Sub Timer1_Timer()
'    DoEvents
'    KillApp "PDFGen.exe" ' Cargar lista de procesos activos (KillApp)
'    DoEvents
'        Enviar_Mail_CDO "julio.gonzalez@ixitec.net", "HAY MAS DE 10 PROCESOS PETADOS", ""
'        MsgBox "HAY MAS DE 10 PROCESOS PETADOS. PARADA TECNICA....", vbCritical, App.Title
'    End If
    ' Verificamos si hay informes pendientes y si es asi no recargamos la lista
    cargar_lista
'    If lstActive.ListCount < 15 Then
'        imprimir
'    End If
End Sub

Private Sub TrayIcon1_DblClick()
    If (Minimized) Then cmdminimizar_Click
End Sub
