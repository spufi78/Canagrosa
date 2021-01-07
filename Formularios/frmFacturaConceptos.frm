VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmFacturaConceptos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura de Conceptos"
   ClientHeight    =   12375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14625
   Icon            =   "frmFacturaConceptos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12375
   ScaleWidth      =   14625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPlasmaAnno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   10170
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.ComboBox cmbPlasmaMes 
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
      Height          =   360
      Left            =   12195
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   9900
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir Datos Plasma"
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   7515
      TabIndex        =   55
      Top             =   11610
      Width           =   5190
      Begin VB.CommandButton cmdPlasmaInsertar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asignar"
         Height          =   420
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   180
         Width           =   975
      End
      Begin MSComCtl2.DTPicker fechaDesdePlasma 
         Height          =   330
         Left            =   855
         TabIndex        =   60
         Top             =   225
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fechaHastaPlasma 
         Height          =   330
         Left            =   2475
         TabIndex        =   62
         Top             =   225
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   12
         Left            =   2295
         TabIndex        =   63
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   61
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descuento por línea"
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   90
      TabIndex        =   50
      Top             =   11610
      Width           =   7395
      Begin VB.TextBox txtDtoLinea 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5310
         TabIndex        =   52
         Text            =   "0"
         Top             =   225
         Width           =   915
      End
      Begin VB.CommandButton cmdAplicarDto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aplicar"
         Height          =   420
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indique el (%) de descuento a aplicar en todas las líneas del documento"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   53
         Top             =   315
         Width           =   5085
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conceptos desde Excel"
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   90
      TabIndex        =   46
      Top             =   9990
      Width           =   11715
      Begin VB.CommandButton cmdLimpiareXCEL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   420
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   270
         Width           =   975
      End
      Begin VB.TextBox txtExcel 
         Appearance      =   0  'Flat
         Height          =   1275
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   48
         Top             =   225
         Width           =   10350
      End
      Begin VB.CommandButton cmdAnadirConceptos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cargar"
         Height          =   420
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   945
         Width           =   975
      End
   End
   Begin VB.TextBox txtdoc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12285
      TabIndex        =   34
      Top             =   9405
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   10665
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   90
      TabIndex        =   28
      Top             =   8010
      Width           =   11715
      Begin VB.CheckBox chkDesglose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desglose (NO SUMA AL TOTAL)"
         Height          =   240
         Left            =   3285
         TabIndex        =   11
         Top             =   225
         Width           =   3075
      End
      Begin VB.TextBox txtDto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4860
         TabIndex        =   15
         Text            =   "0"
         Top             =   1530
         Width           =   915
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Text            =   "1"
         Top             =   1530
         Width           =   825
      End
      Begin VB.CommandButton cmdmodificar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   555
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   765
         Width           =   975
      End
      Begin VB.CommandButton cmdanadir2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   555
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   555
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1350
         Width           =   975
      End
      Begin VB.TextBox txtprecio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         TabIndex        =   13
         Top             =   1530
         Width           =   1635
      End
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   720
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   540
         Width           =   9720
      End
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   330
         Left            =   720
         TabIndex        =   10
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbCC 
         Height          =   345
         Left            =   6525
         TabIndex        =   16
         Top             =   1530
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "DTO (%)"
         Height          =   195
         Index           =   9
         Left            =   4185
         TabIndex        =   39
         Top             =   1575
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
         Height          =   195
         Index           =   8
         Left            =   2475
         TabIndex        =   38
         Top             =   1575
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   7
         Left            =   5940
         TabIndex        =   35
         Top             =   1575
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   31
         Top             =   1575
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   30
         Top             =   765
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   29
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   10665
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la factura"
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
      Height          =   1785
      Left            =   45
      TabIndex        =   22
      Top             =   315
      Width           =   14550
      Begin VB.CommandButton cmdProforma 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proforma"
         Enabled         =   0   'False
         Height          =   915
         Left            =   13455
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.TextBox txtiva 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   330
         Left            =   4905
         TabIndex        =   5
         Top             =   1350
         Width           =   735
      End
      Begin VB.CommandButton cmdFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Factura"
         Enabled         =   0   'False
         Height          =   915
         Left            =   11340
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton cmdAlbaran 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Albaran"
         Enabled         =   0   'False
         Height          =   915
         Left            =   12375
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   1050
      End
      Begin VB.TextBox txtdescuento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3555
         TabIndex        =   4
         Top             =   1350
         Width           =   735
      End
      Begin MSComCtl2.DTPicker ffactura 
         Height          =   330
         Left            =   1215
         TabIndex        =   3
         Top             =   1350
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   6780
         TabIndex        =   6
         Top             =   1350
         Width           =   3345
         _ExtentX        =   5900
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
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   1215
         TabIndex        =   0
         Top             =   270
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbclientesfactura 
         Height          =   345
         Left            =   1215
         TabIndex        =   1
         Top             =   630
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   345
         Left            =   1215
         TabIndex        =   2
         Top             =   990
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente Factura"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   37
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lbliva 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
         Height          =   195
         Left            =   4410
         TabIndex        =   36
         Top             =   1395
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   16
         Left            =   5805
         TabIndex        =   32
         Top             =   1395
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuento"
         Height          =   195
         Index           =   2
         Left            =   2655
         TabIndex        =   26
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha "
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5595
      Left            =   90
      TabIndex        =   9
      Top             =   2385
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   9869
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin MSComCtl2.UpDown cambiar 
      Height          =   360
      Left            =   12960
      TabIndex        =   59
      Top             =   10170
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   635
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtPlasmaAnno"
      BuddyDispid     =   196609
      OrigLeft        =   1590
      OrigTop         =   6570
      OrigRight       =   1830
      OrigBottom      =   6975
      Max             =   2015
      Min             =   2004
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   3
      Left            =   14085
      Picture         =   "frmFacturaConceptos.frx":030A
      ToolTipText     =   "Mover al Ulitmo"
      Top             =   6435
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   2
      Left            =   14085
      Picture         =   "frmFacturaConceptos.frx":044E
      ToolTipText     =   "Mover al Primero"
      Top             =   3420
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   1005
      Left            =   11835
      Top             =   8100
      Width           =   2220
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   12420
      TabIndex        =   45
      Top             =   8775
      Width           =   1575
   End
   Begin VB.Label lblrestan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
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
      Height          =   300
      Index           =   1
      Left            =   11880
      TabIndex        =   44
      Top             =   8775
      Width           =   510
   End
   Begin VB.Label lblBaseTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   12420
      TabIndex        =   43
      Top             =   8145
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Base"
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
      Height          =   300
      Index           =   0
      Left            =   11880
      TabIndex        =   42
      Top             =   8145
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IVA"
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
      Height          =   300
      Index           =   2
      Left            =   11880
      TabIndex        =   41
      Top             =   8460
      Width           =   510
   End
   Begin VB.Label lblIvaTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   12420
      TabIndex        =   40
      Top             =   8460
      Width           =   1575
   End
   Begin VB.Image flecha 
      Height          =   270
      Index           =   1
      Left            =   14085
      Picture         =   "frmFacturaConceptos.frx":0590
      ToolTipText     =   "Mover Abajo"
      Top             =   5490
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   270
      Index           =   0
      Left            =   14085
      Picture         =   "frmFacturaConceptos.frx":067E
      ToolTipText     =   "Mover Arriba"
      Top             =   4455
      Width           =   480
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la Factura"
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
      Height          =   285
      Left            =   90
      TabIndex        =   27
      Top             =   2115
      Width           =   14475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Factura de Conceptos"
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
      Height          =   360
      Index           =   4
      Left            =   0
      TabIndex        =   25
      Top             =   -15
      Width           =   14595
   End
End
Attribute VB_Name = "frmFacturaConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum COLS
    FAMILIA_ID = 0
    ALBARAN_ID = 1
    apartado = 2
    fecha = 3
    DESCRIPCION = 4
    familia = 5
    PRECIO = 6
    CANTIDAD = 7
    SUBTOTAL = 8
    dto = 9
    total = 10
End Enum
Private Enum excel
    apartado = 0
    concepto1 = 1
    concepto2 = 2
    CANTIDAD = 3
    PRECIO = 4
    SUBTOTAL = 5
    dto = 6
    fecha = 7
    total = 8
End Enum
    

Private Sub cmbclientesfactura_change()
    If cmbclientesfactura.getTEXTO <> "" Then
        cargar_pedidos CLng(cmbclientesfactura.getPK_SALIDA), ffactura.Value
        cmbPedidos.limpiar
    End If
End Sub
Private Sub cmdanadir2_Click()
   On Error GoTo cmdanadir2_Click_Error

    If valida_datos = False Then
        Exit Sub
    End If
    ' Añadimos el concepto
    With lista.ListItems.Add(, , cmbCC.getPK_SALIDA)
        .SubItems(COLS.ALBARAN_ID) = 0
        .SubItems(COLS.apartado) = chkDesglose.Value
        .SubItems(COLS.fecha) = Format(txtFecha, "dd/mm/yyyy")
        If chkDesglose.Value = Checked Then
            .SubItems(COLS.DESCRIPCION) = "     " & txtdes
        Else
            .SubItems(COLS.DESCRIPCION) = txtdes
        End If
        .SubItems(COLS.familia) = cmbCC.getTEXTO
'        .SubItems(COLS.PRECIO) = moneda(txtprecio)
        .SubItems(COLS.PRECIO) = moneda4(Replace(Replace(txtPrecio, "€", ""), ".", ""))
        .SubItems(COLS.CANTIDAD) = txtcantidad
        Dim total As Single
        total = CSng(txtcantidad) * CSng(txtPrecio)
        .SubItems(COLS.SUBTOTAL) = moneda(CStr(total))
        Dim DESCUENTO As Single
        If txtDto = "" Then
            DESCUENTO = 0
        Else
            DESCUENTO = txtDto
        End If
        .SubItems(COLS.dto) = DESCUENTO
        If DESCUENTO = 0 Then
            .SubItems(COLS.total) = moneda(CStr(total))
        Else
            .SubItems(COLS.total) = moneda(total - ((total * DESCUENTO) / 100))
        End If
    End With
    ' Limpiamos los campos
    lista.Enabled = True
    cmdEliminar2.Enabled = False
    If CLng(txtdoc) = 0 Then
        cmdAlbaran.Enabled = True
        cmdFactura.Enabled = True
        cmdProforma.Enabled = True
    End If
    calcular_totales
    borrar_campos

   On Error GoTo 0
   Exit Sub

cmdanadir2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir2_Click of Formulario frmFacturaConceptos"

End Sub
Private Sub calcular_totales()
    Dim i As Integer
    Dim BASE As Currency
    Dim IVA As Currency
    Dim total As Currency
   On Error GoTo calcular_totales_Error

    BASE = 0
    IVA = 0
    total = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).SubItems(COLS.apartado) = 0 Then
            BASE = BASE + Format((lista.ListItems(i).SubItems(COLS.total)), "0.00")
        End If
    Next
    IVA = ((BASE * CInt(txtiva)) / 100)
    total = BASE + IVA
    lblBaseTotal = moneda(CStr(BASE))
    lblIvaTotal = moneda(CStr(IVA))
    lbltotal = moneda(CStr(total))

   On Error GoTo 0
   Exit Sub

calcular_totales_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcular_totales of Formulario frmFacturaConceptos"
    
End Sub

Private Sub cmdAnadirConceptos_Click()
   On Error GoTo cmdAnadirConceptos_Click_Error

    If Trim(txtExcel.Text) <> "" Then
        If cmbCC.getTEXTO = "" Then
            MsgBox "Seleccione la familia a la que asignar los conceptos.", vbCritical, App.Title
            cmbCC.SetFocus
            Exit Sub
        End If
        
        Dim conceptos() As String
        Dim linea() As String
        Dim c As String
        conceptos = Split(txtExcel, vbNewLine)
        Dim i As Integer
        Dim total As Single
        Dim PRECIO As Single
        Dim DESCUENTO As Single
        For i = LBound(conceptos) To UBound(conceptos)
            If Trim(conceptos(i)) <> "" Then
                linea = Split(conceptos(i), vbTab)
                If UBound(linea) > 0 Then
                    If IsNumeric(linea(excel.CANTIDAD)) Then
                        With lista.ListItems.Add(, , cmbCC.getPK_SALIDA)
                            .SubItems(COLS.ALBARAN_ID) = 0
                            .SubItems(COLS.apartado) = linea(excel.apartado)
                            .SubItems(COLS.fecha) = Format(linea(excel.fecha), "dd/mm/yyyy")
                            If linea(excel.apartado) = 1 Then
                                .SubItems(COLS.DESCRIPCION) = "     " & Trim(Trim(linea(excel.concepto1)) & " " & Trim(linea(excel.concepto2)))
                            Else
                                .SubItems(COLS.DESCRIPCION) = Trim(Trim(linea(excel.concepto1)) & " " & Trim(linea(excel.concepto2)))
                            End If
                            .SubItems(COLS.familia) = cmbCC.getTEXTO
                            If linea(excel.PRECIO) = "" Then
                                PRECIO = 0
                            Else
                                PRECIO = Replace(linea(excel.PRECIO), "€", "")
                            End If
                                .SubItems(COLS.PRECIO) = moneda(CStr(PRECIO))
                            If linea(excel.CANTIDAD) = "" Then
                                .SubItems(COLS.CANTIDAD) = 0
                            Else
                                .SubItems(COLS.CANTIDAD) = linea(excel.CANTIDAD)
                            End If
                            If linea(excel.SUBTOTAL) = "" Then
                                total = 0
                            Else
                                total = Replace(linea(excel.SUBTOTAL), "€", "")
                            End If
                            .SubItems(COLS.SUBTOTAL) = moneda(CStr(total))
                            If linea(excel.dto) = "" Then
                                DESCUENTO = 0
                            Else
                                DESCUENTO = Replace(linea(excel.dto), "%", "")
                            End If
                            .SubItems(COLS.dto) = DESCUENTO
                            If DESCUENTO = 0 Then
                                .SubItems(COLS.total) = moneda(CStr(total))
                            Else
                                .SubItems(COLS.total) = moneda(total - ((total * DESCUENTO) / 100))
                            End If
                        End With
                    Else
                        If linea(excel.CANTIDAD) = "" And linea(excel.apartado) = "0" Then
                        With lista.ListItems.Add(, , cmbCC.getPK_SALIDA)
                            .SubItems(COLS.ALBARAN_ID) = 0
                            .SubItems(COLS.apartado) = linea(excel.apartado)
                            .SubItems(COLS.fecha) = Format(linea(excel.fecha), "dd/mm/yyyy")
                            .SubItems(COLS.DESCRIPCION) = Trim(Trim(linea(excel.concepto1)) & " " & Trim(linea(excel.concepto2)))
                            .SubItems(COLS.familia) = cmbCC.getTEXTO
                            
                            If linea(excel.PRECIO) = "" Then
                                PRECIO = 0
                            Else
                                PRECIO = Replace(linea(excel.PRECIO), "€", "")
                            End If
                                .SubItems(COLS.PRECIO) = moneda(CStr(PRECIO))
                            If linea(excel.CANTIDAD) = "" Then
                                .SubItems(COLS.CANTIDAD) = 0
                            Else
                                .SubItems(COLS.CANTIDAD) = linea(excel.CANTIDAD)
                            End If
                            If linea(excel.SUBTOTAL) = "" Then
                                total = 0
                            Else
                                total = Replace(linea(excel.SUBTOTAL), "€", "")
                            End If
                            .SubItems(COLS.SUBTOTAL) = moneda(CStr(total))
                            If linea(excel.dto) = "" Then
                                DESCUENTO = 0
                            Else
                                DESCUENTO = Replace(linea(excel.dto), "%", "")
                            End If
                            .SubItems(COLS.dto) = DESCUENTO
                            If DESCUENTO = 0 Then
                                .SubItems(COLS.total) = moneda(CStr(total))
                            Else
                                .SubItems(COLS.total) = moneda(total - ((total * DESCUENTO) / 100))
                            End If
                        End With
                        Else
                            ' Si es un codigo de equipo, buscar la linea donde este para añadir el concepto
                            For j = 1 To lista.ListItems.Count
                                If InStr(1, lista.ListItems(j).SubItems(4), linea(excel.concepto1)) > 0 Then
        '                                If UBound(linea) > 5 Then
        '                                    c = Trim(linea(6))
        '                                Else
        '                                    c = ""
        '                                End If
                                    lista.ListItems(j).SubItems(4) = lista.ListItems(j).SubItems(4) & "," & Trim(Trim(linea(excel.concepto1)) & " " & Trim(linea(excel.concepto2)))
                                End If
                            Next
                        End If
                    End If
'                    If IsNumeric(linea(3)) Then
'                        With lista.ListItems.Add(, , cmbCC.getPK_SALIDA)
'                            .SubItems(cols.ALBARAN_ID) = 0
'                            .SubItems(cols.apartado) = chkDesglose.value
'                            .SubItems(cols.fecha) = Format(linea(4), "dd/mm/yyyy")
'                            If UBound(linea) > 5 Then
'                                c = Trim(linea(6))
'                            Else
'                                c = ""
'                            End If
'                            If chkDesglose.value = Checked Then
'                                .SubItems(cols.descripcion) = "     " & Trim(Trim(linea(5)) & " " & c)
'                            Else
'                                .SubItems(cols.descripcion) = Trim(Trim(linea(5)) & " " & c)
'                            End If
'                            Dim total As Single
'                            total = Replace(linea(3), "€", "")
'                            .SubItems(cols.familia) = cmbCC.getTEXTO
'                            .SubItems(cols.PRECIO) = moneda(CStr(total))
'                            .SubItems(cols.cantidad) = 1
'                            .SubItems(cols.SUBTOTAL) = moneda(CStr(total))
'                            Dim DESCUENTO As Single
'
'                            If txtDto = "" Then
'                                DESCUENTO = 0
'                            Else
'                                DESCUENTO = txtDto
'                            End If
'                            .SubItems(cols.dto) = DESCUENTO
'                            If DESCUENTO = 0 Then
'                                .SubItems(cols.total) = moneda(CStr(total))
'                            Else
'                                .SubItems(cols.total) = moneda(total - ((total * DESCUENTO) / 100))
'                            End If
'                        End With
'                    Else
'                        ' Si es un codigo de equipo, buscar la linea donde este para añadir el concepto
'                        For j = 1 To lista.ListItems.Count
'                            If InStr(1, lista.ListItems(j).SubItems(4), linea(3)) > 0 Then
'                                If UBound(linea) > 5 Then
'                                    c = Trim(linea(6))
'                                Else
'                                    c = ""
'                                End If
'                                lista.ListItems(j).SubItems(4) = lista.ListItems(j).SubItems(4) & "," & Trim(Trim(linea(5)) & " " & c)
'                            End If
'                        Next
'                    End If
                End If
            End If
        Next
        lista.Enabled = True
        cmdEliminar2.Enabled = False
        If CLng(txtdoc) = 0 Then
            cmdAlbaran.Enabled = True
            cmdFactura.Enabled = True
            cmdProforma.Enabled = True
        End If
        calcular_totales
        txtExcel = ""
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadirConceptos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirConceptos_Click of Formulario frmFacturaConceptos"
End Sub

Private Sub cmdAplicarDto_Click()
    Dim total As Single
    Dim DESCUENTO As Single
    If txtDtoLinea <> "" Then
        If IsNumeric(txtDtoLinea) Then
            If MsgBox("¿Desea aplicar el descuento a todas las lineas del documento?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                Dim i As Integer
                DESCUENTO = txtDtoLinea
                For i = 1 To lista.ListItems.Count
                    total = CSng(lista.ListItems(i).SubItems(COLS.SUBTOTAL))
                    lista.ListItems(i).SubItems(COLS.dto) = DESCUENTO
                    If DESCUENTO = 0 Then
                        lista.ListItems(i).SubItems(COLS.total) = moneda(CStr(total))
                    Else
                        lista.ListItems(i).SubItems(COLS.total) = moneda(total - ((total * DESCUENTO) / 100))
                    End If
                Next
                calcular_totales
            End If
        End If
    End If
End Sub

Private Sub cmdEliminar2_Click()
    If lista.selectedItem.Index > 0 Then
     lista.ListItems.Remove (lista.selectedItem.Index)
     cmdEliminar2.Enabled = False
     calcular_totales
    End If
End Sub

Private Sub cmbClientes_change()
    If cmbclientes.getTEXTO <> "" Then
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbclientes.getPK_SALIDA
        cmbFP.BoundText = oCliente.getFP_ID
        cmbclientesfactura.MostrarElemento cmbclientes.getPK_SALIDA
    End If
End Sub
Private Sub cmdAceptar_Click()
   On Error GoTo cmdaceptar_Click_Error

    If cmbclientes.getTEXTO = "" Then
        MsgBox "Seleccione algún cliente.", vbInformation, App.Title
        cmbclientes.SetFocus
        Exit Sub
    End If
    If cmbclientesfactura.getTEXTO = "" Then
        MsgBox "Seleccione algún cliente en Cliente_Factura.", vbInformation, App.Title
        cmbclientesfactura.SetFocus
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        cmdaceptar.Enabled = False
        Me.MousePointer = 11
        ' Borramos los conceptos anteriores
        Dim oConcepto As New clsDocs_pago_conceptos
        oConcepto.EliminarConceptos (CLng(txtdoc))
        ' Insertamos los conceptos
        For i = 1 To lista.ListItems.Count
            With oConcepto
                .setDOC_ID = CLng(txtdoc)
                .setALBARAN_ID = lista.ListItems(i).SubItems(COLS.ALBARAN_ID)
                .setDESCRIPCION = lista.ListItems(i).SubItems(COLS.DESCRIPCION)
                .setFECHA = Format(lista.ListItems(i).SubItems(COLS.fecha), "yyyy-mm-dd")
                .setCANTIDAD = moneda_bd(lista.ListItems(i).SubItems(COLS.CANTIDAD))
                .setPRECIO = moneda_bd4(lista.ListItems(i).SubItems(COLS.PRECIO))
                .setSUBTOTAL = moneda_bd(lista.ListItems(i).SubItems(COLS.SUBTOTAL))
                .setDTO = moneda_bd(lista.ListItems(i).SubItems(COLS.dto))
                .setTOTAL = moneda_bd(lista.ListItems(i).SubItems(COLS.total))
                .setFAMILIA_ID = lista.ListItems(i).Text
                .setAPARTADO = lista.ListItems(i).SubItems(COLS.apartado)
                If .Insertar = False Then
                    Exit Sub
                End If
            End With
        Next
        ' Modificamos la factura
        Dim oDocPago As New clsDocs_pago
        oDocPago.setFECHA_FACTURA = Format(ffactura, "yyyy-mm-dd")
        oDocPago.setCLIENTE_ID = cmbclientes.getPK_SALIDA
        oDocPago.setCLIENTE_ID_FACTURA = cmbclientesfactura.getPK_SALIDA
        oDocPago.setDESCUENTO = Replace(Format(txtdescuento, "0.00"), ",", ".")
        oDocPago.setFP_ID = cmbFP.BoundText
        If cmbPedidos.getTEXTO = "" Then
            oDocPago.setPEDIDO_ID = 0
        Else
            oDocPago.setPEDIDO_ID = cmbPedidos.getPK_SALIDA
        End If
        oDocPago.setIVA = txtiva
        ' Insertamos el documento de pago
        If oDocPago.Modificar_Factura_Conceptos(CLng(txtdoc)) = False Then
             Exit Sub
        End If
        oDocPago.informar_factura_conceptos (CLng(txtdoc))
        oDocPago.Informar_total_factura txtdoc
        Me.MousePointer = 0
 
        MsgBox "Documento modificado correctamente.", vbOKOnly + vbInformation, App.Title
        Unload Me
    Else
        Me.MousePointer = 0
        cmdaceptar.Enabled = True
        MsgBox "Necesita algún concepto para la factura.", vbInformation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdaceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmFacturaConceptos"
End Sub

Private Sub cmdAlbaran_Click()
    If cmbclientes.getTEXTO = "" Then
        MsgBox "Seleccione algún cliente.", vbInformation, App.Title
        cmbclientes.SetFocus
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        generar_documento (C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN)
    End If
End Sub
Private Sub cmdFactura_Click()
    If cmbclientes.getTEXTO = "" Then
        MsgBox "Seleccione algún cliente.", vbInformation, App.Title
        cmbclientes.SetFocus
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        generar_documento (C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA)
    End If
End Sub

Private Sub cmdLimpiareXCEL_Click()
 txtExcel.Text = ""
End Sub

Private Sub cmdmodificar2_Click()
   On Error GoTo cmdmodificar2_Click_Error

    If lista.ListItems.Count > 0 Then
        If valida_datos Then
            With lista.ListItems(lista.selectedItem.Index)
                .Text = cmbCC.getPK_SALIDA
                .SubItems(COLS.apartado) = chkDesglose.Value
                .SubItems(COLS.fecha) = Format(txtFecha, "dd/mm/yyyy")
                If chkDesglose.Value = Checked Then
                    .SubItems(COLS.DESCRIPCION) = "     " & txtdes
                Else
                    .SubItems(COLS.DESCRIPCION) = txtdes
                End If
                .SubItems(COLS.familia) = cmbCC.getTEXTO
                .SubItems(COLS.PRECIO) = moneda4(Replace(Replace(txtPrecio, "€", ""), ".", ""))
                .SubItems(COLS.CANTIDAD) = txtcantidad
                Dim total As Single
                total = CSng(txtcantidad) * CSng(txtPrecio)
                .SubItems(COLS.SUBTOTAL) = moneda(CStr(total))
                Dim DESCUENTO As Single
                If txtDto = "" Then
                    DESCUENTO = 0
                Else
                    DESCUENTO = txtDto
                End If
                .SubItems(COLS.dto) = DESCUENTO
                If DESCUENTO = 0 Then
                    .SubItems(COLS.total) = moneda(CStr(total))
                Else
                    .SubItems(COLS.total) = moneda(total - ((total * DESCUENTO) / 100))
                End If
            End With
            borrar_campos
            calcular_totales
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdmodificar2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificar2_Click of Formulario frmFacturaConceptos"
End Sub

Private Sub cmdPlasmaInsertar_Click()
   On Error GoTo cmdPlasmaInsertar_Click_Error

'    If cmbPlasmaMes.Text <> "" And txtPlasmaAnno.Text <> "" Then
'        If MsgBox("¿Desea insertar los plasmas no facturados de " & cmbPlasmaMes.List(cmbPlasmaMes.ListIndex) & " del " & txtPlasmaAnno & "?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        If MsgBox("¿Desea insertar los plasmas no facturados y recepcionados desde " & fechaDesdePlasma & " hasta " & fechaHastaPlasma & "?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Me.MousePointer = 11
            Dim s As String
            s = ""
            s = s & "select d.NOMBRE,'MICROESTRUCTURA BOND', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and (c.MICROESTRUCTURA1 <> 2 or c.MICROESTRUCTURA2 <> 2 or c.MICROESTRUCTURA3 <> 2 or c.MICROESTRUCTURA4 <> 2 or c.MICROESTRUCTURA5 <> 2 or c.MICROESTRUCTURA6 <> 2) " & _
                    "  and c.TIPO = 1 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'MICROESTRUCTURA TOP', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and (c.MICROESTRUCTURA1 <> 2 or c.MICROESTRUCTURA2 <> 2 or c.MICROESTRUCTURA3 <> 2 or c.MICROESTRUCTURA4 <> 2 or c.MICROESTRUCTURA5 <> 2 or c.MICROESTRUCTURA6 <> 2) " & _
                    "  and c.TIPO = 2 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'TRACCION BOND', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.TRACCION_RES  <> '' " & _
                    "  and c.TIPO = 1 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'TRACCION TOP', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.TRACCION_RES  <> '' " & _
                    "  and c.TIPO = 2 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'MACRO DUREZA BOND', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.MACRO_DUREZA_RES  <> '' " & _
                    "  and c.TIPO = 1 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'MACRO DUREZA TOP', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.MACRO_DUREZA_RES  <> '' " & _
                    "  and c.TIPO = 2 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'MICRO DUREZA BOND', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.MICRO_DUREZA_RES  <> '' " & _
                    "  and c.TIPO = 1 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'MICRO DUREZA TOP', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.MICRO_DUREZA_RES  <> '' " & _
                    "  and c.TIPO = 2 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'ESPESOR BOND', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.ESPESOR_RES  <> '' " & _
                    "  and c.TIPO = 1 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            s = s & " UNION "
            s = s & "select d.NOMBRE,'ESPESOR TOP', count(distinct a.ID_MUESTRA) " & _
                    " from muestras a, plasma_recepcion b, plasma_resultados c, clientes d " & _
                    " where a.ANALISIS_MODIFICADO = " & tipo_especial.PLASMA & _
                    "  and a.ANULADA = 0 and a.REFERENCIA_CLIENTE not like '%IRR%' " & _
                    "  and a.TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.PLASMA & _
                    "  and a.DOCUMENTO_PAGO = 0 and d.IBERIA = 1 " & _
                    "  and a.ID_MUESTRA = b.MUESTRA_ID and a.ID_MUESTRA = c.MUESTRA_ID and a.CLIENTE_ID = d.ID_CLIENTE " & _
                    "  and c.BATCH <> 'N/A' " & _
                    "  and c.ESPESOR_RES  <> '' " & _
                    "  and c.TIPO = 2 " & _
                    "  and a.FECHA_RECEPCION >= '" & Format(fechaDesdePlasma, "yyyy-mm-dd") & "' and a.FECHA_RECEPCION <= '" & Format(fechaHastaPlasma, "yyyy-mm-dd") & "'"
'                    "  and month(a.FECHA_RECEPCION) = " & cmbPlasmaMes.ListIndex + 1 & " And a.ANNO = " & txtPlasmaAnno
            s = s & " group by 1 "
            Dim rs As ADODB.Recordset
            Set rs = datos_bd(s)
            Dim famId As String
            Dim famNombre As String
            Dim familias As New clsFamilias
            Dim PRECIO As Currency
            Dim total As Currency
            If rs.RecordCount > 0 Then
                Dim oParametros As New clsParametros
                
                oParametros.Carga parametros.PLASMA_FACTURACION_FAMILIA, ""
                famId = oParametros.getVALOR
                familias.CARGAR famId
                famNombre = familias.getNOMBRE
                Do
                    With lista.ListItems.Add(, , famId)
                        .SubItems(COLS.ALBARAN_ID) = 0
                        .SubItems(COLS.apartado) = 0
                        .SubItems(COLS.fecha) = Format(ffactura, "dd/mm/yyyy")
                        .SubItems(COLS.DESCRIPCION) = "Plasma: " & rs(1)
                        .SubItems(COLS.familia) = famNombre
                        If oParametros.Carga(parametros.PLASMA_FACTURACION_PRECIOS, rs(1)) = True Then
                            PRECIO = oParametros.getVALOR
                        Else
                            PRECIO = 0
                        End If
                        .SubItems(COLS.PRECIO) = moneda(CStr(PRECIO))
                        .SubItems(COLS.CANTIDAD) = rs(2)
                        total = PRECIO * rs(2)
                        .SubItems(COLS.SUBTOTAL) = moneda(CStr(total))
                        .SubItems(COLS.dto) = txtdescuento
                        
                        If txtdescuento = "0" Or txtdescuento = "" Then
                            .SubItems(COLS.total) = moneda(CStr(total))
                        Else
                            .SubItems(COLS.total) = moneda(CStr(total - ((total * txtdescuento) / 100)))
                        End If
                        
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set rs = Nothing
            ' Limpiamos los campos
            lista.Enabled = True
            cmdEliminar2.Enabled = False
            If CLng(txtdoc) = 0 Then
                cmdAlbaran.Enabled = True
                cmdFactura.Enabled = True
                cmdProforma.Enabled = True
            End If
            calcular_totales
        End If
'    End If
   Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdPlasmaInsertar_Click_Error:
   Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPlasmaInsertar_Click of Formulario frmFacturaConceptos"
End Sub

Private Sub cmdProforma_Click()
    If cmbclientes.getTEXTO = "" Then
        MsgBox "Seleccione algún cliente.", vbInformation, App.Title
        cmbclientes.SetFocus
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        generar_documento (C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA)
    End If
End Sub

Private Sub cmdSalir_Click()
    log ("Cerrando modificación factura de conceptos")
    If lista.ListItems.Count > 0 Then
       If MsgBox("Existen conceptos. ¿Esta seguro de querer salir?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Unload Me
       End If
    Else
       Unload Me
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer, j As Integer, sel As Integer
   On Error GoTo flecha_Click_Error

    If lista.ListItems.Count > 0 Then
        Select Case Index
        Case 0 ' SUBIR
           If lista.selectedItem.Index > 1 Then
              aux = lista.ListItems(lista.selectedItem.Index - 1).Text
              lista.ListItems(lista.selectedItem.Index - 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To lista.ColumnHeaders.Count - 1
                  aux = lista.ListItems(lista.selectedItem.Index - 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index - 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
           End If
        Case 1 ' BAJAR
           If lista.selectedItem.Index < lista.ListItems.Count Then
              aux = lista.ListItems(lista.selectedItem.Index + 1).Text
              lista.ListItems(lista.selectedItem.Index + 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To lista.ColumnHeaders.Count - 1
                  aux = lista.ListItems(lista.selectedItem.Index + 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index + 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
           End If
        Case 2 ' PRIMERO
           If lista.selectedItem.Index > 1 Then
                sel = lista.selectedItem.Index
                For j = sel To 2 Step -1
                    aux = lista.ListItems(j - 1).Text
                    lista.ListItems(j - 1).Text = lista.ListItems(j).Text
                    lista.ListItems(j).Text = aux
                    For i = 1 To lista.ColumnHeaders.Count - 1
                        aux = lista.ListItems(j - 1).SubItems(i)
                        lista.ListItems(j - 1).SubItems(i) = lista.ListItems(j).SubItems(i)
                        lista.ListItems(j).SubItems(i) = aux
                    Next
                    Set lista.selectedItem = lista.ListItems(j - 1)
                Next
           End If
        Case 3 ' ULTIMO
           If lista.selectedItem.Index < lista.ListItems.Count Then
                sel = lista.selectedItem.Index
                For j = sel To lista.ListItems.Count - 1
                    aux = lista.ListItems(j + 1).Text
                    lista.ListItems(j + 1).Text = lista.ListItems(j).Text
                    lista.ListItems(j).Text = aux
                    For i = 1 To lista.ColumnHeaders.Count - 1
                        aux = lista.ListItems(j + 1).SubItems(i)
                        lista.ListItems(j + 1).SubItems(i) = lista.ListItems(j).SubItems(i)
                        lista.ListItems(j).SubItems(i) = aux
                    Next
                    Set lista.selectedItem = lista.ListItems(j + 1)
                Next
           End If
        End Select
    End If

   On Error GoTo 0
   Exit Sub

flecha_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure flecha_Click of Formulario frmFacturaConceptos"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' Esc
            cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    txtdoc = gdoc
    Me.Left = 100
    Me.top = 100
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbclientesfactura, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbFP, New clsFP
    llenar_combo cmbCC, New clsFamilias, 0, Me, ""
    Dim oParametro As New clsParametros
    oParametro.Carga parametros.IVA, ""
    txtiva = oParametro.getVALOR
    cabecera_grid
    ffactura = Now
    txtFecha = Now
    ' Cargar Combo Años
    fechaDesdePlasma = Date
    fechaHastaPlasma = Date
    txtPlasmaAnno = Year(Date)
    cambiar.Max = Year(Date)
    For i = 1 To 12
        cmbPlasmaMes.AddItem UCase(CStr(MonthName(i)))
    Next
    If CLng(txtdoc) <> 0 Then
        Me.Left = 200
        Me.top = 500
        cargar_documento
'        lbliva.Visible = True
'        txtiva.Visible = True
        Dim oDoc As New clsDocs_pago
        If oDoc.esta_contabilidado(gdoc) Then
'            cmdaceptar.Enabled = False
        End If
    End If
End Sub
Private Sub cabecera_grid()
    With lista.ColumnHeaders
        .Add , , "FAMILIA_ID", 1, lvwColumnLeft
        .Add , , "ALBARAN_ID", 1, lvwColumnCenter
        .Add , , "APARTADO", 1, lvwColumnCenter
        .Add , , "Fecha", 1050, lvwColumnLeft
        .Add , , "Descripción", 6200, lvwColumnLeft
        .Add , , "Familia", 1500, lvwColumnCenter
        .Add , , "Precio", 1100, lvwColumnRight
        .Add , , "Cantidad", 700, lvwColumnCenter
        .Add , , "Subtotal", 1100, lvwColumnRight
        .Add , , "Dto", 700, lvwColumnCenter
        .Add , , "Total", 1100, lvwColumnRight
    End With
End Sub

Private Sub borrar_campos()
    txtdes = ""
    txtMuestra = ""
    txtPrecio = ""
    txtcantidad = "1"
    txtDto = "0"
'    cmbCC.Limpiar
    txtdes.SetFocus
End Sub

Private Function valida_datos() As Boolean
    valida_datos = True
    If txtdes = "" Then
        MsgBox "El concepto esta vacio.", vbInformation, App.Title
        txtdes.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtPrecio = "" Then
        MsgBox "El campo precio esta vacio.", vbInformation, App.Title
        txtPrecio.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtcantidad = "" Then
        MsgBox "El campo CANTIDAD esta vacio.", vbInformation, App.Title
        txtcantidad.SetFocus
        valida_datos = False
        Exit Function
    End If
    If Not IsNumeric(txtcantidad) Then
        MsgBox "El campo CANTIDAD no es numérico.", vbInformation, App.Title
        txtcantidad.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDto <> "" Then
        If Not IsNumeric(txtDto) Then
            MsgBox "El campo DTO no es numérico.", vbInformation, App.Title
            txtDto.SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If cmbCC.getTEXTO = "" Then
        MsgBox "Seleccione la familia a la que pertenece el concepto.", vbInformation, App.Title
        cmbCC.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmdEliminar2.Enabled = True
        chkDesglose.Value = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.apartado)
        txtFecha = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.fecha)
        txtdes = Trim(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.DESCRIPCION))
        
        txtPrecio = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.PRECIO)
        
        txtcantidad = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.CANTIDAD)
        txtDto = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.dto)
        cmbCC.MostrarElemento lista.ListItems(lista.selectedItem.Index).Text
    End If
End Sub

Private Sub txtcantidad_GotFocus()
    txtcantidad.SelStart = 0
    txtcantidad.SelLength = Len(txtcantidad)
End Sub

Private Sub txtdes_GotFocus()
    txtdes.BackColor = &H80C0FF
    txtdes.SelStart = 0
    txtdes.SelLength = Len(txtdes)
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdes_LostFocus()
    txtdes.BackColor = &HFFFFFF
End Sub

Private Sub txtDto_GotFocus()
    txtDto.SelStart = 0
    txtDto.SelLength = Len(txtDto)
End Sub

Private Sub txtDto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub
Private Sub txtDtoLinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub txtiva_LostFocus()
    If Not IsNumeric(txtiva) Then
        MsgBox "El IVA debe ser numérico.", vbCritical, App.Title
        txtiva.SetFocus
    End If
End Sub

Private Sub txtprecio_LostFocus()
    txtPrecio.BackColor = &HFFFFFF
    If txtPrecio <> "" Then
        If Not IsNumeric(txtPrecio) Then
            MsgBox "El precio debe ser numérico.", vbInformation, App.Title
            txtPrecio = ""
            txtPrecio.SetFocus
        End If
    End If
End Sub
Private Sub txtprecio_GotFocus()
    txtPrecio.BackColor = &H80C0FF
    txtPrecio.SelStart = 0
    txtPrecio.SelLength = Len(txtPrecio)
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub
Private Sub generar_documento(TIPO_DOCUMENTO As Integer)
    Dim i As Integer
    Dim num_doc As Long
    Dim oCliente As New clsCliente
    Dim oDocPago As New clsDocs_pago
'    Dim IVA As Integer
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
   
   On Error GoTo generar_documento_Error

    If oCliente.CargaCliente(cmbclientes.getPK_SALIDA) = False Then
        Exit Sub
    End If
    Me.MousePointer = 11
    oDocPago.setTIPO = TIPO_DOCUMENTO
    oDocPago.setFECHA_FACTURA = Format(ffactura, "yyyy-mm-dd")
    oDocPago.setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
    oDocPago.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
    oDocPago.setCLIENTE_ID = cmbclientes.getPK_SALIDA
    oDocPago.setCLIENTE_ID_FACTURA = cmbclientesfactura.getPK_SALIDA
    oDocPago.setTOTAL = "0.00"
    If txtdescuento <> "" Then
        oDocPago.setDESCUENTO = Replace(Format(txtdescuento, "0.00"), ",", ".")
    Else
        oDocPago.setDESCUENTO = 0
    End If
'    If TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA Or TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA Then
'         oDocPago.setIVA = IVA
'    Else
'         oDocPago.setIVA = 0
'    End If
    oDocPago.setPAGADO = 0
    oDocPago.setANULADO = 0
    oDocPago.setFP_ID = cmbFP.BoundText
    If cmbPedidos.getTEXTO = "" Then
        oDocPago.setPEDIDO_ID = 0
    Else
        oDocPago.setPEDIDO_ID = cmbPedidos.getPK_SALIDA
    End If
    oDocPago.setFACTURA_CONCEPTOS = 1
    ' Insertamos el documento de pago
    num_doc = oDocPago.InsertarDocPago
    If num_doc = 0 Then
         Exit Sub
    End If
    ' Insertamos los conceptos
    Dim oConcepto As New clsDocs_pago_conceptos
    Dim oMuestra As New clsMuestra
    For i = 1 To lista.ListItems.Count
        With oConcepto
            .setDOC_ID = num_doc
            .setDESCRIPCION = lista.ListItems(i).SubItems(COLS.DESCRIPCION)
            .setFECHA = Format(lista.ListItems(i).SubItems(COLS.fecha), "yyyy-mm-dd")
            .setCANTIDAD = moneda_bd(lista.ListItems(i).SubItems(COLS.CANTIDAD))
            .setPRECIO = moneda_bd(lista.ListItems(i).SubItems(COLS.PRECIO))
            .setSUBTOTAL = moneda_bd(lista.ListItems(i).SubItems(COLS.SUBTOTAL))
            .setDTO = moneda_bd(lista.ListItems(i).SubItems(COLS.dto))
            .setTOTAL = moneda_bd(lista.ListItems(i).SubItems(COLS.total))
            .setFAMILIA_ID = lista.ListItems(i).Text
            .setAPARTADO = lista.ListItems(i).SubItems(COLS.apartado)
            If .Insertar = False Then
                Exit Sub
            End If
        End With
    Next
    ' Modificamos el total de la factura
    oDocPago.Informar_total_factura (num_doc)
    Dim sTIPO As String
    Me.MousePointer = 0
 
    Select Case TIPO_DOCUMENTO
    Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
        sTIPO = "Albaran"
    Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
        sTIPO = "Factura"
    Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA
        sTIPO = "Proforma"
    End Select
    MsgBox sTIPO & " registrado correctamente.", vbOKOnly + vbInformation, App.Title
    restaurar_formulario

   On Error GoTo 0
   Exit Sub

generar_documento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_documento of Formulario frmFacturaConceptos"
End Sub
Private Sub restaurar_formulario()
   cmbclientes.limpiar
   cmbclientesfactura.limpiar
   ffactura = Date
   txtdescuento = ""
   cmdFactura.Enabled = False
   cmdAlbaran.Enabled = False
   cmdProforma.Enabled = False
   lista.ListItems.Clear
   cmdEliminar2.Enabled = False
   txtFecha = Date
   cmbCC.limpiar
   cmbclientes.SetFocus
End Sub

Private Sub cargar_documento()
   On Error GoTo cargar_documento_Error

    Label1(4).BackColor = &H80C0FF
    Label1(4) = "Modificación de Factura de Conceptos"
    cmdaceptar.visible = True
    cmdFactura.Enabled = False
    cmdAlbaran.Enabled = False
    cmdProforma.Enabled = False
    lista.Enabled = True
    ' Documento
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.CargarDocumento (CLng(txtdoc))
    ffactura = oDoc_pago.getFECHA_FACTURA
    cmbclientes.MostrarElemento oDoc_pago.getCLIENTE_ID
    cmbclientesfactura.MostrarElemento oDoc_pago.getCLIENTE_ID_FACTURA
    
    txtdescuento = oDoc_pago.getDESCUENTO
    cmbFP.BoundText = oDoc_pago.getFP_ID
    
    cargar_pedidos CLng(oDoc_pago.getCLIENTE_ID_FACTURA), oDoc_pago.getFECHA_FACTURA
    cmbPedidos.MostrarElemento oDoc_pago.getPEDIDO_ID
    txtiva = oDoc_pago.getIVA
    ' Conceptos
    Dim oDoc_pago_conceptos As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Set rs = oDoc_pago_conceptos.ConceptosListado(CLng(txtdoc))
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(COLS.ALBARAN_ID) = rs(1)
                .SubItems(COLS.apartado) = rs(2)
                .SubItems(COLS.fecha) = Format(rs(3), "dd/mm/yyyy")
                If rs(2) = 0 Then
                    .SubItems(COLS.DESCRIPCION) = Trim(rs(4))
                Else
                    .SubItems(COLS.DESCRIPCION) = "     " & Trim(rs(4))
                End If
                .SubItems(COLS.familia) = rs(5)
                .SubItems(COLS.PRECIO) = moneda4(rs(6))
                .SubItems(COLS.CANTIDAD) = rs(7)
                .SubItems(COLS.SUBTOTAL) = moneda(rs(8))
                .SubItems(COLS.dto) = rs(9)
                .SubItems(COLS.total) = moneda(rs(10))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    calcular_totales
   On Error GoTo 0
   Exit Sub

cargar_documento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_documento of Formulario frmFacturaConceptos"
End Sub

Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim consulta As String
    consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
               "  FROM CLIENTES_PEDIDOS " & _
               " WHERE ID_PEDIDO <> 0 " & _
               "   AND CLIENTE_ID = " & cliente & _
               "   AND FECHA_PEDIDO <= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
               "   AND FECHA_BAJA >= '" & Format(fecha, "yyyy-mm-dd") & "' "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbPedidos
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = 0
        .setTABLA = "CLIENTES_PEDIDOS"
        .setDESCRIPCION = "Pedidos"
        .setPK = "ID_PEDIDO"
        .setCAMPO = "CONCAT(CODIGO,' (',DESCRIPCION,')')"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = Me
        End With
    End If
End Sub

