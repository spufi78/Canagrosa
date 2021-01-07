VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmModificarFactura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación de factura"
   ClientHeight    =   10890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   24795
   Icon            =   "frmModificarFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   24795
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7155
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   10260
      Visible         =   0   'False
      Width           =   2265
   End
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
      Left            =   9585
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   10260
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir Datos Plasma"
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   10800
      TabIndex        =   60
      Top             =   9855
      Width           =   5235
      Begin VB.CommandButton cmdPlasmaInsertar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asignar"
         Height          =   420
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   180
         Width           =   975
      End
      Begin MSComCtl2.DTPicker fechaDesdePlasma 
         Height          =   330
         Left            =   855
         TabIndex        =   65
         Top             =   270
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
         Format          =   52101121
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fechaHastaPlasma 
         Height          =   330
         Left            =   2475
         TabIndex        =   66
         Top             =   270
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
         Format          =   52101121
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   68
         Top             =   360
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   12
         Left            =   2295
         TabIndex        =   67
         Top             =   315
         Width           =   90
      End
   End
   Begin VB.Frame frmDeter 
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
      Height          =   7635
      Left            =   6210
      TabIndex        =   26
      Top             =   1935
      Visible         =   0   'False
      Width           =   10725
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   330
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   6660
         Width           =   1410
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6660
         Width           =   1410
      End
      Begin VB.CommandButton cmdAceptarDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   930
         Left            =   8190
         Picture         =   "frmModificarFactura.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6645
         Width           =   1155
      End
      Begin VB.CommandButton cmdSalirDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   930
         Left            =   9390
         Picture         =   "frmModificarFactura.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6645
         Width           =   1155
      End
      Begin MSComctlLib.ListView lista2 
         Height          =   6180
         Left            =   90
         TabIndex        =   29
         Top             =   450
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   10901
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
         BackColor       =   14737632
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Muestras pendientes de facturación del cliente"
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
         TabIndex        =   30
         Top             =   135
         Width           =   10515
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   10800
      TabIndex        =   43
      Top             =   8100
      Width           =   13425
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   720
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   540
         Width           =   11520
      End
      Begin VB.TextBox txtprecioConcepto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         TabIndex        =   47
         Top             =   1305
         Width           =   1635
      End
      Begin VB.CommandButton cmdmodificar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   465
         Left            =   12330
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   675
         Width           =   975
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         TabIndex        =   48
         Text            =   "1"
         Top             =   1305
         Width           =   825
      End
      Begin VB.TextBox txtDto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4860
         TabIndex        =   50
         Text            =   "0"
         Top             =   1305
         Width           =   915
      End
      Begin VB.CheckBox chkDesglose 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desglose (NO SUMA AL TOTAL)"
         Height          =   240
         Left            =   3285
         TabIndex        =   44
         Top             =   225
         Width           =   3075
      End
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   330
         Left            =   720
         TabIndex        =   45
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
         Format          =   52101121
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbCC 
         Height          =   345
         Left            =   6525
         TabIndex        =   52
         Top             =   1305
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdanadir2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   465
         Left            =   12330
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   12330
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   59
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   57
         Top             =   765
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   55
         Top             =   1350
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   7
         Left            =   5940
         TabIndex        =   53
         Top             =   1350
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
         Height          =   195
         Index           =   8
         Left            =   2475
         TabIndex        =   51
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "DTO (%)"
         Height          =   195
         Index           =   9
         Left            =   4185
         TabIndex        =   49
         Top             =   1350
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView listaConceptos 
      Height          =   6315
      Left            =   10800
      TabIndex        =   41
      Top             =   1755
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   11139
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
      NumItems        =   0
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar Muestras"
      Height          =   930
      Left            =   1755
      Picture         =   "frmModificarFactura.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9900
      Width           =   1605
   End
   Begin VB.CommandButton cmdinsertar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Insertar Muestras"
      Height          =   930
      Left            =   90
      Picture         =   "frmModificarFactura.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9900
      Width           =   1605
   End
   Begin VB.CommandButton cmdbajar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bajar datos"
      Height          =   780
      Left            =   2880
      Picture         =   "frmModificarFactura.frx":2632
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdmodificarPrecio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   330
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9900
      Width           =   1095
   End
   Begin VB.TextBox txtprecio 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   7875
      TabIndex        =   10
      Top             =   9900
      Width           =   1635
   End
   Begin VB.CommandButton cmdSubir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subir datos"
      Height          =   780
      Left            =   3915
      Picture         =   "frmModificarFactura.frx":2A74
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   22005
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9945
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   23130
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9945
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
      Height          =   1035
      Left            =   90
      TabIndex        =   14
      Top             =   360
      Width           =   24120
      Begin VB.TextBox txtiva 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   330
         Left            =   15570
         TabIndex        =   5
         Top             =   585
         Width           =   735
      End
      Begin VB.TextBox txtdescuento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   13995
         TabIndex        =   4
         Top             =   585
         Width           =   690
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   11610
         TabIndex        =   3
         Top             =   585
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
         Format          =   52101121
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   17325
         TabIndex        =   6
         Top             =   585
         Width           =   2940
         _ExtentX        =   5186
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
         Left            =   1170
         TabIndex        =   0
         Top             =   225
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbclientesfactura 
         Height          =   345
         Left            =   11610
         TabIndex        =   1
         Top             =   225
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   345
         Left            =   1170
         TabIndex        =   2
         Top             =   585
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente Factura"
         Height          =   195
         Index           =   5
         Left            =   10485
         TabIndex        =   25
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
         Height          =   195
         Index           =   0
         Left            =   15075
         TabIndex        =   24
         Top             =   675
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   23
         Top             =   675
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   16
         Left            =   16425
         TabIndex        =   22
         Top             =   675
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   14760
         TabIndex        =   21
         Top             =   630
         Width           =   195
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuento"
         Height          =   195
         Index           =   3
         Left            =   13095
         TabIndex        =   20
         Top             =   675
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   16
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   10485
         TabIndex        =   15
         Top             =   675
         Width           =   450
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   8070
      Left            =   90
      TabIndex        =   7
      Top             =   1755
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   14235
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
   Begin MSComCtl2.UpDown cambiar 
      Height          =   360
      Left            =   10335
      TabIndex        =   64
      Top             =   10260
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   635
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtPlasmaAnno"
      BuddyDispid     =   196611
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle de Conceptos de la Muestra"
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
      Left            =   10845
      TabIndex        =   42
      Top             =   1440
      Width           =   13395
   End
   Begin VB.Image flecha 
      Height          =   270
      Index           =   0
      Left            =   24255
      Picture         =   "frmModificarFactura.frx":2EB6
      ToolTipText     =   "Mover Arriba"
      Top             =   3915
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   270
      Index           =   1
      Left            =   24255
      Picture         =   "frmModificarFactura.frx":2FA1
      ToolTipText     =   "Mover Abajo"
      Top             =   4950
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   2
      Left            =   24255
      Picture         =   "frmModificarFactura.frx":308F
      ToolTipText     =   "Mover al Primero"
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   3
      Left            =   24255
      Picture         =   "frmModificarFactura.frx":31D1
      ToolTipText     =   "Mover al Ulitmo"
      Top             =   5895
      Width           =   480
   End
   Begin VB.Label lblbase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   19665
      TabIndex        =   38
      Top             =   9855
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Base"
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
      Height          =   345
      Index           =   3
      Left            =   18180
      TabIndex        =   37
      Top             =   9855
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
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
      Height          =   315
      Index           =   2
      Left            =   18180
      TabIndex        =   36
      Top             =   10515
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   19665
      TabIndex        =   35
      Top             =   10515
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dto."
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
      Height          =   345
      Index           =   1
      Left            =   18180
      TabIndex        =   34
      Top             =   10185
      Width           =   1455
   End
   Begin VB.Label lbliva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   19665
      TabIndex        =   33
      Top             =   10185
      Width           =   2160
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar Precio del análisis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   4995
      TabIndex        =   19
      Top             =   9945
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Modificación de factura"
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
      Height          =   330
      Index           =   4
      Left            =   105
      TabIndex        =   18
      Top             =   0
      Width           =   24630
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Muestras de la Muestra"
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
      TabIndex        =   17
      Top             =   1440
      Width           =   10620
   End
End
Attribute VB_Name = "frmModificarFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public muestras_modificadas As Boolean
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

Private Sub cmbClientes_change()
    If cmbClientes.getTEXTO <> "" Then
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbClientes.getPK_SALIDA
        cmbFP.BoundText = oCliente.getFP_ID
        cargar_muestras_pendientes cmbClientes.getPK_SALIDA
    End If
End Sub

Private Sub cmbclientesfactura_change()
    If cmbclientesfactura.getTEXTO <> "" Then
        ' Cargamos los pedido del cliente factura
        cargar_pedidos CLng(cmbclientesfactura.getPK_SALIDA), fdesde.Value
        cmbPedidos.limpiar
    End If
End Sub

Private Sub cmdAceptarDeter_Click()
    
   On Error GoTo cmdAceptarDeter_Click_Error
    Dim odocm As New clsDocs_pago_muestras
    Dim oMuestra As New clsMuestra
    Me.MousePointer = 11
    For i = 1 To lista2.ListItems.Count
        If lista2.ListItems(i).Checked = True Then
            With odocm
             .setMUESTRA_ID = CLng(lista2.ListItems(i).SubItems(6))
             .setDOC_ID = gdoc
             .setFECHA = Format(lista2.ListItems(i).SubItems(4), "yyyy-mm-dd")
             .setTIPO_ANALISIS = lista2.ListItems(i).SubItems(2)
             .setREFERENCIA_CLIENTE = lista2.ListItems(i).SubItems(3)
             .setPRECIO = Replace(Format(lista2.ListItems(i).SubItems(5), "0.00"), ",", ".")
             
             .setORDEN = odocm.CalcularOrden(gdoc)
             
             ORDEN = .Insertar_doc_pago_muestra(lista2.ListItems(i).SubItems(7))
             If ORDEN = -1 Then
                 MsgBox "Error al insertar las muestras en el documento. Contacte con mantenimiento.", vbCritical, App.Title
                 Exit Sub
             End If
            End With
            ' Modificar el documento de pago de la muestra
            If oMuestra.Informar_Documento_Pago(CLng(lista2.ListItems(i).SubItems(6)), 2) = False Then
                Exit Sub
            End If
            If cmbPedidos.getTEXTO = "" Then
                oMuestra.informar_pedido CLng(lista2.ListItems(i).SubItems(6)), 0
            Else
                oMuestra.informar_pedido CLng(lista2.ListItems(i).SubItems(6)), cmbPedidos.getPK_SALIDA
            End If
            oMuestra.actualizar_precio CLng(lista2.ListItems(i).SubItems(6)), Replace(Format(lista2.ListItems(i).SubItems(5), "0.00"), ",", ".")
        End If
    Next
    Me.MousePointer = 0
    Dim oDoc As New clsDocs_pago
    oDoc.Informar_total_factura gdoc
    Set oDoc = Nothing
    cargar_muestras
    frmDeter.visible = False

   On Error GoTo 0
   Exit Sub

cmdAceptarDeter_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptarDeter_Click of Formulario frmModificarFactura"
    
End Sub

Private Sub cmdanadir2_Click()
   On Error GoTo cmdanadir2_Click_Error

    If valida_datos_conceptos = False Then
        Exit Sub
    End If
    ' Añadimos el concepto
    With listaConceptos.ListItems.Add(, , cmbCC.getPK_SALIDA)
        .SubItems(COLS.ALBARAN_ID) = 0
        .SubItems(COLS.apartado) = chkDesglose.Value
        .SubItems(COLS.fecha) = Format(txtFecha, "dd/mm/yyyy")
        If chkDesglose.Value = Checked Then
            .SubItems(COLS.DESCRIPCION) = "     " & txtdes
        Else
            .SubItems(COLS.DESCRIPCION) = txtdes
        End If
        .SubItems(COLS.familia) = cmbCC.getTEXTO
        .SubItems(COLS.PRECIO) = moneda(Replace(Replace(txtprecioConcepto, "€", ""), ".", ""))
        .SubItems(COLS.CANTIDAD) = txtcantidad
        Dim total As Single
        total = CSng(txtcantidad) * CSng(txtprecioConcepto)
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
    cmdEliminar2.Enabled = False
    calcular_total
    borrar_campos_conceptos

   On Error GoTo 0
   Exit Sub

cmdanadir2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir2_Click of Formulario frmModificarFactura"
End Sub

Private Sub cmdbajar_Click()
    If lista.ListItems.Count > 0 Then
        Dim omue As New clsMuestra
        With lista2.ListItems.Add(, , lista.ListItems(lista.selectedItem.Index).Text)
                .SubItems(1) = cmbClientes.getTEXTO
                .SubItems(2) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
                .SubItems(3) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
                .SubItems(4) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
                .SubItems(5) = lista.ListItems(lista.selectedItem.Index).SubItems(4)
                .SubItems(6) = lista.ListItems(lista.selectedItem.Index).SubItems(5)
                .SubItems(7) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
                ' Recalcular Precio de la muestra
                If .SubItems(7) = 1 Then
                    .SubItems(5) = Format(omue.ImporteMuestraPorDeterminaciones(lista.ListItems(lista.selectedItem.Index).SubItems(5), cmbClientes.getPK_SALIDA), "currency")
                End If
        End With
        lista.ListItems.Remove (lista.selectedItem.Index)
        muestras_modificadas = True
    End If
End Sub

Private Sub cmdBorrar_Click()
   On Error GoTo cmdborrar_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim CANTIDAD As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            CANTIDAD = CANTIDAD + 1
        End If
    Next
    If CANTIDAD = 0 Then
        MsgBox "No hay muestras marcadas.", vbCritical, App.Title
        Exit Sub
    End If
    If MsgBox("Va a eliminar " & CANTIDAD & " muestras. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oMuestra As New clsMuestra
        Dim odpm As New clsDocs_pago_muestras
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                ' Marcamos la muestra como no facturada
                oMuestra.Informar_Documento_Pago CLng(lista.ListItems(i).SubItems(5)), 0
                ' Eliminamos la muestra de la factura
                odpm.EliminarIdMuestra gdoc, CLng(lista.ListItems(i).SubItems(5))
            End If
        Next
        Set oMuestra = Nothing
        Set odpm = Nothing
        Dim oDoc As New clsDocs_pago
        oDoc.Informar_total_factura gdoc
        Set oDoc = Nothing
        cargar_muestras
    End If

   On Error GoTo 0
   Exit Sub

cmdborrar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdborrar_Click of Formulario frmModificarFactura"
    
End Sub

Private Sub cmdcancel_Click()
    log ("Cerrando frmModificarFactura")
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    For i = 1 To lista2.ListItems.Count
        lista2.ListItems(i).Checked = False
    Next

End Sub

Private Sub cmdEliminar2_Click()
    If listaConceptos.selectedItem.Index > 0 Then
     listaConceptos.ListItems.Remove (listaConceptos.selectedItem.Index)
     cmdEliminar2.Enabled = False
     calcular_total
     listaConceptos.SetFocus
    End If
End Sub

Private Sub cmdinsertar_Click()
    frmDeter.visible = True
End Sub

Private Sub cmdMarcar_Click()
    For i = 1 To lista2.ListItems.Count
        lista2.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdmodificar2_Click()
   On Error GoTo cmdmodificar2_Click_Error

    If listaConceptos.ListItems.Count > 0 Then
        If valida_datos_conceptos Then
            With listaConceptos.ListItems(listaConceptos.selectedItem.Index)
                .Text = cmbCC.getPK_SALIDA
                .SubItems(COLS.apartado) = chkDesglose.Value
                .SubItems(COLS.fecha) = Format(txtFecha, "dd/mm/yyyy")
                If chkDesglose.Value = Checked Then
                    .SubItems(COLS.DESCRIPCION) = "     " & txtdes
                Else
                    .SubItems(COLS.DESCRIPCION) = txtdes
                End If
                .SubItems(COLS.familia) = cmbCC.getTEXTO
                .SubItems(COLS.PRECIO) = moneda(Replace(Replace(txtprecioConcepto, "€", ""), ".", ""))
                .SubItems(COLS.CANTIDAD) = txtcantidad
                Dim total As Single
                total = CSng(txtcantidad) * CSng(txtprecioConcepto)
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
            borrar_campos_conceptos
            calcular_total
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdmodificar2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificar2_Click of Formulario frmModificarFactura"

End Sub

Private Sub cmdmodificarPrecio_Click()
   On Error GoTo cmdmodificarPrecio_Click_Error

    If lista.ListItems.Count > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = Format(txtPrecio, "currency")
        Dim oMuestra As New clsMuestra
        oMuestra.actualizar_precio CLng(lista.ListItems(lista.selectedItem.Index).SubItems(5)), Replace(Format(lista.ListItems(lista.selectedItem.Index).SubItems(4), "0.00"), ",", ".")
        Set oMuestra = Nothing
        Dim odpm As New clsDocs_pago_muestras
        odpm.modificarPrecio gdoc, CLng(lista.ListItems(lista.selectedItem.Index).SubItems(5)), Replace(Format(lista.ListItems(lista.selectedItem.Index).SubItems(4), "0.00"), ",", ".")
        Set odpm = Nothing
        cargar_muestras
    End If

   On Error GoTo 0
   Exit Sub

cmdmodificarPrecio_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificarPrecio_Click of Formulario frmModificarFactura"
End Sub

Private Sub cmdok_Click()
    If MsgBox("Va a modificar el documento de pago. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim i As Integer
        Dim ORDEN As Integer
        Dim oBANO As New clsBanos
        Dim oTA As New clsTipos_analisis
        Dim oCodigo As New clsTarifas_codigos
        Dim sgrupo As String
        Dim odocm As New clsDocs_pago_muestras
        Dim oMuestra As New clsMuestra
        Dim rs As ADODB.Recordset
        
        On Error GoTo fallo
        Me.MousePointer = 11
        ' Conceptos del documento
        Dim oConcepto As New clsDocs_pago_conceptos
        oConcepto.EliminarConceptos (CLng(gdoc))
        ' Insertamos los conceptos
        For i = 1 To listaConceptos.ListItems.Count
            With oConcepto
                .setDOC_ID = CLng(gdoc)
                .setALBARAN_ID = listaConceptos.ListItems(i).SubItems(COLS.ALBARAN_ID)
                .setDESCRIPCION = listaConceptos.ListItems(i).SubItems(COLS.DESCRIPCION)
                .setFECHA = Format(listaConceptos.ListItems(i).SubItems(COLS.fecha), "yyyy-mm-dd")
                .setCANTIDAD = moneda_bd(listaConceptos.ListItems(i).SubItems(COLS.CANTIDAD))
                .setPRECIO = moneda_bd(listaConceptos.ListItems(i).SubItems(COLS.PRECIO))
                .setSUBTOTAL = moneda_bd(listaConceptos.ListItems(i).SubItems(COLS.SUBTOTAL))
                .setDTO = moneda_bd(listaConceptos.ListItems(i).SubItems(COLS.dto))
                .setTOTAL = moneda_bd(listaConceptos.ListItems(i).SubItems(COLS.total))
                .setFAMILIA_ID = listaConceptos.ListItems(i).Text
                .setAPARTADO = listaConceptos.ListItems(i).SubItems(COLS.apartado)
                If .Insertar = False Then
                    Exit Sub
                End If
            End With
        Next
        ' Informamos los datos del documento
        Dim oDoc As New clsDocs_pago
        With oDoc
            .setCLIENTE_ID = cmbClientes.getPK_SALIDA
            .setCLIENTE_ID_FACTURA = cmbclientesfactura.getPK_SALIDA
            .setDESCUENTO = numerico_bd(txtdescuento)
            .setFECHA_FACTURA = Format(fdesde.Value, "yyyy-mm-dd")
            .setFP_ID = cmbFP.BoundText
            If cmbPedidos.getTEXTO <> "" Then
                .setPEDIDO_ID = cmbPedidos.getPK_SALIDA
            Else
                .setPEDIDO_ID = 0
            End If
            .setIVA = txtiva
            .Modificar_Factura_Muestras gdoc
        End With
        oDoc.informar_factura_conceptos gdoc
        oDoc.Informar_total_factura gdoc
        
        MsgBox "La factura se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
        Set oMuestra = Nothing
        Me.MousePointer = 0
        Unload Me
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al modificar la factura número " & gdoc & "." & Err.Description
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
                    With listaConceptos.ListItems.Add(, , famId)
                        .SubItems(COLS.ALBARAN_ID) = 0
                        .SubItems(COLS.apartado) = 0
                        .SubItems(COLS.fecha) = Format(fdesde, "dd/mm/yyyy")
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
            calcular_total
        End If
'    End If
   Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdPlasmaInsertar_Click_Error:
   Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPlasmaInsertar_Click of Formulario frmModificarFactura"

End Sub

Private Sub cmdSalirDeter_Click()
    frmDeter.visible = False
End Sub

Private Sub cmdSubir_Click()
    If lista2.ListItems.Count > 0 Then
        With lista.ListItems.Add(, , lista2.ListItems(lista2.selectedItem.Index).Text)
                .SubItems(1) = lista2.ListItems(lista2.selectedItem.Index).SubItems(4)
                .SubItems(2) = lista2.ListItems(lista2.selectedItem.Index).SubItems(2)
                .SubItems(3) = lista2.ListItems(lista2.selectedItem.Index).SubItems(3)
                .SubItems(4) = lista2.ListItems(lista2.selectedItem.Index).SubItems(5)
                .SubItems(5) = lista2.ListItems(lista2.selectedItem.Index).SubItems(6)
                .SubItems(6) = lista2.ListItems(lista2.selectedItem.Index).SubItems(7)
        End With
        lista2.ListItems.Remove lista2.selectedItem.Index
        muestras_modificadas = True
    End If
End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer, j As Integer, sel As Integer
   On Error GoTo flecha_Click_Error

    If listaConceptos.ListItems.Count > 0 Then
        Select Case Index
        Case 0 ' SUBIR
           If listaConceptos.selectedItem.Index > 1 Then
              aux = listaConceptos.ListItems(listaConceptos.selectedItem.Index - 1).Text
              listaConceptos.ListItems(listaConceptos.selectedItem.Index - 1).Text = listaConceptos.ListItems(listaConceptos.selectedItem.Index).Text
              listaConceptos.ListItems(listaConceptos.selectedItem.Index).Text = aux
              For i = 1 To listaConceptos.ColumnHeaders.Count - 1
                  aux = listaConceptos.ListItems(listaConceptos.selectedItem.Index - 1).SubItems(i)
                  listaConceptos.ListItems(listaConceptos.selectedItem.Index - 1).SubItems(i) = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(i)
                  listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(i) = aux
              Next
              Set listaConceptos.selectedItem = listaConceptos.ListItems(listaConceptos.selectedItem.Index - 1)
           End If
        Case 1 ' BAJAR
           If listaConceptos.selectedItem.Index < listaConceptos.ListItems.Count Then
              aux = listaConceptos.ListItems(listaConceptos.selectedItem.Index + 1).Text
              listaConceptos.ListItems(listaConceptos.selectedItem.Index + 1).Text = listaConceptos.ListItems(listaConceptos.selectedItem.Index).Text
              listaConceptos.ListItems(listaConceptos.selectedItem.Index).Text = aux
              For i = 1 To listaConceptos.ColumnHeaders.Count - 1
                  aux = listaConceptos.ListItems(listaConceptos.selectedItem.Index + 1).SubItems(i)
                  listaConceptos.ListItems(listaConceptos.selectedItem.Index + 1).SubItems(i) = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(i)
                  listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(i) = aux
              Next
              Set listaConceptos.selectedItem = listaConceptos.ListItems(listaConceptos.selectedItem.Index + 1)
           End If
        Case 2 ' PRIMERO
           If listaConceptos.selectedItem.Index > 1 Then
                sel = listaConceptos.selectedItem.Index
                For j = sel To 2 Step -1
                    aux = listaConceptos.ListItems(j - 1).Text
                    listaConceptos.ListItems(j - 1).Text = listaConceptos.ListItems(j).Text
                    listaConceptos.ListItems(j).Text = aux
                    For i = 1 To listaConceptos.ColumnHeaders.Count - 1
                        aux = listaConceptos.ListItems(j - 1).SubItems(i)
                        listaConceptos.ListItems(j - 1).SubItems(i) = listaConceptos.ListItems(j).SubItems(i)
                        listaConceptos.ListItems(j).SubItems(i) = aux
                    Next
                    Set listaConceptos.selectedItem = listaConceptos.ListItems(j - 1)
                Next
           End If
        Case 3 ' ULTIMO
           If listaConceptos.selectedItem.Index < listaConceptos.ListItems.Count Then
                sel = listaConceptos.selectedItem.Index
                For j = sel To listaConceptos.ListItems.Count - 1
                    aux = listaConceptos.ListItems(j + 1).Text
                    listaConceptos.ListItems(j + 1).Text = listaConceptos.ListItems(j).Text
                    listaConceptos.ListItems(j).Text = aux
                    For i = 1 To listaConceptos.ColumnHeaders.Count - 1
                        aux = listaConceptos.ListItems(j + 1).SubItems(i)
                        listaConceptos.ListItems(j + 1).SubItems(i) = listaConceptos.ListItems(j).SubItems(i)
                        listaConceptos.ListItems(j).SubItems(i) = aux
                    Next
                    Set listaConceptos.selectedItem = listaConceptos.ListItems(j + 1)
                Next
           End If
        End Select
    End If

   On Error GoTo 0
   Exit Sub

flecha_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure flecha_Click of Formulario frmFacturaConceptos"

End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
'    cargar_clientes
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbclientesfactura, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbFP, New clsFP
    llenar_combo cmbCC, New clsFamilias, 0, Me, ""
    cabecera
    ' Cargar Combo Años
    fechaDesdePlasma = Date
    fechaHastaPlasma = Date
    txtPlasmaAnno = Year(Date)
    cambiar.Max = Year(Date)
    For i = 1 To 12
        cmbPlasmaMes.AddItem UCase(CStr(MonthName(i)))
    Next
    If gdoc <> 0 Then
        cargar_documento
        cargarConceptos gdoc
    End If
    ' Verificar si esta contabilidado
    Dim oDoc As New clsDocs_pago
    If oDoc.esta_contabilidado(gdoc) Then
'        cmdok.Enabled = False
    End If
    muestras_modificadas = False
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºEnsayo", 1000, lvwColumnLeft
        .Add , , "Fecha", 1150, lvwColumnCenter
        .Add , , "Tipo Analisis", 2950, lvwColumnLeft
        .Add , , "Ref.Cliente", 3900, lvwColumnLeft
        .Add , , "Importe", 1300, lvwColumnRight
        .Add , , "Id", 1, lvwColumnCenter
        .Add , , "FACTURA_DETERMINACIONES", 1, lvwColumnCenter
    End With
    ' Lista 2
    With lista2.ColumnHeaders
        .Add , , "NºEnsayo", 900, lvwColumnLeft
        .Add , , "Cliente", 1, lvwColumnLeft
        .Add , , "Analisis", 3000, lvwColumnLeft
        .Add , , "Ref.Cliente", 4000, lvwColumnLeft
        .Add , , "Fecha", 1150, lvwColumnCenter
        .Add , , "Precio", 1000, lvwColumnRight
        .Add , , "Id", 1, lvwColumnCenter
        .Add , , "FACTURA_DETERMINACIONES", 1, lvwColumnCenter
    End With
    
    With listaConceptos.ColumnHeaders
        .Add , , "FAMILIA_ID", 1, lvwColumnLeft
        .Add , , "ALBARAN_ID", 1, lvwColumnCenter
        .Add , , "APARTADO", 1, lvwColumnCenter
        .Add , , "Fecha", 1050, lvwColumnLeft
        .Add , , "Descripción", 5850, lvwColumnLeft
        .Add , , "Familia", 1500, lvwColumnCenter
        .Add , , "Precio", 1100, lvwColumnRight
        .Add , , "Cantidad", 700, lvwColumnCenter
        .Add , , "Subtotal", 1100, lvwColumnRight
        .Add , , "Dto", 700, lvwColumnCenter
        .Add , , "Total", 1100, lvwColumnRight
    End With

End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtPrecio = lista.ListItems(lista.selectedItem.Index).SubItems(4)
    End If
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = CLng(lista.ListItems(lista.selectedItem.Index).SubItems(5))
        frmVerMuestra.Show 1
    End If
End Sub

Private Sub lista_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        cmdcancel_Click
    End Select
End Sub
Private Sub cargar_muestras()
    Dim odoc_m As New clsDocs_pago_muestras
    Dim oCliente As New clsCliente
   On Error GoTo cargar_muestras_Error

    Set rs = odoc_m.MuestrasDocumentoEdicion(gdoc)
    lista.ListItems.Clear
'    Dim oTA As New clsTipos_analisis
'    Dim oBANO As New clsBanos
    
    Dim BANO_ID As Long
    Dim ANALISIS_MODIFICADO As Long
    Dim TA_FACTURA_DETERMINACIONES As Long
    Dim BANO_FACTURA_DETERMINACIONES As Long
    
    If rs.RecordCount > 0 Then
        oCliente.CargaCliente cmbclientesfactura.getPK_SALIDA
        Do
                With lista.ListItems.Add(, , rs.Fields(1))
                    .SubItems(1) = rs.Fields(2)
                    .SubItems(2) = rs.Fields(3)
                    .SubItems(3) = rs.Fields(4)
                    .SubItems(4) = Format(rs.Fields(5), "currency")
                    .SubItems(5) = rs.Fields(6)
'                    oMuestra.CargaMuestra (rs.Fields(6))
                    
'                    oTA.CARGAR oMuestra.getTIPO_ANALISIS_ID
'                    If oMuestra.getBANO_ID <> 0 Then
'                        oBANO.cargar_bano oMuestra.getBANO_ID
'                    End If
                    '** rs(9) El cliente factura por determinaciones
                    '** rs(13) El tipo de analisis es por determinaciones
                    '** rs(14) BANO_ID
                    '** rs(19) BANO -> FACTURA_DETERMINACIONES
                    ' Cliente factura por determinaciones O
                    ' No es baño y el tipo de analisis se factura por determinaciones O
                    ' Es baño y el baño se factura por determinaciones Y NO ES CE
                    BANO_ID = 0
                    ANALISIS_MODIFICADO = 0
                    TA_FACTURA_DETERMINACIONES = 0
                    BANO_FACTURA_DETERMINACIONES = 0
                    If Not IsNull(rs(12)) Then
                        BANO_ID = rs(12)
                    End If
                    If Not IsNull(rs(13)) Then
                        ANALISIS_MODIFICADO = rs(13)
                    End If
                    If Not IsNull(rs(14)) Then
                        TA_FACTURA_DETERMINACIONES = rs(14)
                    End If
                    If Not IsNull(rs(15)) Then
                        BANO_FACTURA_DETERMINACIONES = rs(15)
                    End If
                    
                    If oCliente.getFACTURA_DETERMINACIONES = 1 Or _
                      (BANO_ID = 0 And TA_FACTURA_DETERMINACIONES = 1) Or _
                      (BANO_ID <> 0 And BANO_FACTURA_DETERMINACIONES = 1 And ANALISIS_MODIFICADO <> 2) Then
                        .SubItems(6) = 1
                    Else
                        .SubItems(6) = 0
                    End If
                    
'                    If oCliente.getFACTURA_DETERMINACIONES = 1 Or _
'                      (oMuestra.getBANO_ID = 0 And oTA.getFACTURA_DETERMINACIONES = 1) Or _
'                      (oMuestra.getBANO_ID <> 0 And oBANO.getFACTURA_DETERMINACIONES = 1 And oMuestra.getANALISIS_MODIFICADO <> 2) Then
'                        .SubItems(6) = 1
'                    Else
'                        .SubItems(6) = 0
'                    End If
'                    If oCliente.getFACTURA_DETERMINACIONES = 1 Or _
'                       oTA.getFACTURA_DETERMINACIONES = 1 Then
'                        .SubItems(6) = 1
'                    Else
'                        .SubItems(6) = 0
'                    End If
                End With
'                lista.ListItems(lista.ListItems.Count).Checked = True
                rs.MoveNext
         Loop Until rs.EOF
    End If
    calcular_total
    cargar_muestras_pendientes cmbClientes.getPK_SALIDA

   On Error GoTo 0
   Exit Sub

cargar_muestras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestras of Formulario frmModificarFactura"

End Sub
Private Sub cargar_documento()
    On Error GoTo fallo
    Dim oDoc As New clsDocs_pago
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    oDoc.CargarDocumento (gdoc)
    fdesde.Value = Format(oDoc.getFECHA_FACTURA, "dd-mm-yyyy")
    txtdescuento = oDoc.getDESCUENTO
    txtiva = oDoc.getIVA
    cmbClientes.MostrarElemento oDoc.getCLIENTE_ID
    cmbclientesfactura.MostrarElemento oDoc.getCLIENTE_ID_FACTURA
    
    cmbFP.BoundText = oDoc.getFP_ID
    cargar_pedidos CLng(oDoc.getCLIENTE_ID_FACTURA), fdesde.Value
    cmbPedidos.MostrarElemento oDoc.getPEDIDO_ID
    
    cargar_muestras
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    error_grave ("Cargar_Documento. Error al cargar los datos para modificar el documento." & Err.Description)
End Sub

Private Sub lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista2.ListItems.Count > 0 Then
     lista2.SortKey = ColumnHeader.Index - 1
     If lista2.SortOrder = 0 Then
        lista2.SortOrder = 1
     Else
        lista2.SortOrder = 0
     End If
     lista2.Sorted = True
   End If
End Sub

Private Sub lista2_DblClick()
    If lista2.ListItems.Count > 0 Then
        gmuestra = CLng(lista2.ListItems(lista2.selectedItem.Index).SubItems(6))
        frmVerMuestra.Show 1
    End If
End Sub

Public Sub cargar_muestras_pendientes(cliente As Long)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT cl.id_cliente, " & _
                    "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
                    "cl.nombre, " & _
                    "mu.tipo_analisis_id, " & _
                    "mu.referencia_cliente, " & _
                    "mu.fecha_recepcion, " & _
                    "mu.id_general, " & _
                    "mu.precio, " & _
                    "mu.id_muestra, " & _
                    "cl.factura_determinaciones, " & _
                    "ta.nombre,tm.id_tipo_muestra,ta.factura_determinaciones " & _
                   "FROM clientes as cl,tipos_muestra as tm, tipos_analisis as ta, " & _
                        "muestras as mu " & _
                   "WHERE  mu.cliente_id=cl.id_cliente AND " & _
                         " mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                         " mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                         " mu.cliente_id = " & cliente & _
                         " AND mu.documento_pago=0 And (mu.anulada Is Null " & _
                         " or mu.anulada = 0) order by cl.id_cliente, mu.id_muestra"
    Set rs = datos_bd(consulta)
    lista2.ListItems.Clear
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista2.ListItems.Add(, , rs.Fields(6))
            .SubItems(1) = rs.Fields(2)
            .SubItems(2) = rs(10)
            .SubItems(3) = rs.Fields(4)
            If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
            End If
            ' Si el cliente factura por determinaciones o el ta es por determinaciones
            If rs(9) = 1 Or rs(12) = 1 Then
                Dim oMuestra As New clsMuestra
                .SubItems(5) = Format(oMuestra.ImporteMuestraPorDeterminaciones(rs(8), rs(0)), "currency")
                .SubItems(7) = 1
            Else
                If Not IsNull(rs.Fields(7)) Then
                .SubItems(5) = Format(rs.Fields(7), "currency")
                End If
                .SubItems(7) = 0
            End If
            .SubItems(6) = rs.Fields(8)
            End With
            lista2.ListItems(lista2.ListItems.Count).Checked = True
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
End Sub
Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim consulta As String
    consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
               "  FROM CLIENTES_PEDIDOS " & _
               " WHERE ID_PEDIDO <> 0 " & _
               "   AND CLIENTE_ID = " & cliente & _
               "   AND FECHA_PEDIDO <= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
               "   AND FECHA_BAJA >= '" & Format(fecha, "yyyy-mm-dd") & "' "
'               " ORDER BY FECHA_PEDIDO DESC"
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

Private Sub listaConceptos_Click()
    If listaConceptos.ListItems.Count > 0 Then
        cmdEliminar2.Enabled = True
        chkDesglose.Value = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(COLS.apartado)
        txtFecha = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(COLS.fecha)
        txtdes = Trim(listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(COLS.DESCRIPCION))
        txtprecioConcepto = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(COLS.PRECIO)
        txtcantidad = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(COLS.CANTIDAD)
        txtDto = listaConceptos.ListItems(listaConceptos.selectedItem.Index).SubItems(COLS.dto)
        cmbCC.MostrarElemento listaConceptos.ListItems(listaConceptos.selectedItem.Index).Text
    End If
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtiva_LostFocus()
    If Not IsNumeric(txtiva) Then
        MsgBox "El IVA debe ser numérico.", vbCritical, App.Title
        txtiva.SetFocus
    End If
End Sub
Private Sub calcular_total()
    Dim i As Integer
    On Error Resume Next
    Dim total As Currency
    total = 0
    ' muestras
    For i = 1 To lista.ListItems.Count
        total = total + CSng(Format(lista.ListItems(i).SubItems(4), "0.00"))
    Next
    ' Conceptos
    For i = 1 To listaConceptos.ListItems.Count
        If listaConceptos.ListItems(i).SubItems(COLS.apartado) = 0 Then
            total = total + Format((listaConceptos.ListItems(i).SubItems(COLS.total)), "0.00")
        End If
    Next
    
    lblBase = Format(total, "#,##0.00")
    Dim dto As Currency
    If txtdescuento.Text <> "" Then
        dto = Format((CCur(lblBase) * CInt(txtdescuento.Text) / 100), "#,##0.00")
    Else
        dto = 0
    End If
'    dto = dto + Format(((CCur(lblbase) - dto) * CInt(txtdescuento.Text) / 100), "#,##0.00")
'    dto = Format((CCur(lblbase) - dto), "#,##0.00")
    lblIVA = Format(dto, "#,##0.00")
    lbltotal = Format(CCur(lblBase) - CCur(lblIVA), "#,##0.00")
    
    
    
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub cargarConceptos(ID_DOC As Long)
'    Dim oDoc_pago As New clsDocs_pago
'    oDoc_pago.CargarDocumento (CLng(txtdoc))
'    ffactura = oDoc_pago.getFECHA_FACTURA
'    cmbclientes.MostrarElemento oDoc_pago.getCLIENTE_ID
'    cmbclientesfactura.MostrarElemento oDoc_pago.getCLIENTE_ID_FACTURA
'    txtdescuento = oDoc_pago.getDESCUENTO
'    cmbFP.BoundText = oDoc_pago.getFP_ID
'    cargar_pedidos CLng(oDoc_pago.getCLIENTE_ID_FACTURA), oDoc_pago.getFECHA_FACTURA
'    cmbPedidos.MostrarElemento oDoc_pago.getPEDIDO_ID
'    txtiva = oDoc_pago.getIVA
    ' Conceptos
    Dim oDoc_pago_conceptos As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Set rs = oDoc_pago_conceptos.ConceptosListado(ID_DOC)
    If rs.RecordCount > 0 Then
        Do
            With listaConceptos.ListItems.Add(, , rs(0))
                .SubItems(COLS.ALBARAN_ID) = rs(1)
                .SubItems(COLS.apartado) = rs(2)
                .SubItems(COLS.fecha) = Format(rs(3), "dd/mm/yyyy")
                If rs(2) = 0 Then
                    .SubItems(COLS.DESCRIPCION) = Trim(rs(4))
                Else
                    .SubItems(COLS.DESCRIPCION) = "     " & Trim(rs(4))
                End If
                .SubItems(COLS.familia) = rs(5)
                .SubItems(COLS.PRECIO) = moneda(rs(6))
                .SubItems(COLS.CANTIDAD) = rs(7)
                .SubItems(COLS.SUBTOTAL) = moneda(rs(8))
                .SubItems(COLS.dto) = rs(9)
                .SubItems(COLS.total) = moneda(rs(10))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    calcular_total
End Sub

Private Function valida_datos_conceptos() As Boolean
    valida_datos_conceptos = True
    If txtdes = "" Then
        MsgBox "El concepto esta vacio.", vbInformation, App.Title
        txtdes.SetFocus
        valida_datos_conceptos = False
        Exit Function
    End If
    If txtprecioConcepto = "" Then
        MsgBox "El campo precio esta vacio.", vbInformation, App.Title
        txtprecioConcepto.SetFocus
        valida_datos_conceptos = False
        Exit Function
    End If
    If txtcantidad = "" Then
        MsgBox "El campo CANTIDAD esta vacio.", vbInformation, App.Title
        txtcantidad.SetFocus
        valida_datos_conceptos = False
        Exit Function
    End If
    If Not IsNumeric(txtcantidad) Then
        MsgBox "El campo CANTIDAD no es numérico.", vbInformation, App.Title
        txtcantidad.SetFocus
        valida_datos_conceptos = False
        Exit Function
    End If
    If txtDto <> "" Then
        If Not IsNumeric(txtDto) Then
            MsgBox "El campo DTO no es numérico.", vbInformation, App.Title
            txtDto.SetFocus
            valida_datos_conceptos = False
            Exit Function
        End If
    End If
    If cmbCC.getTEXTO = "" Then
        MsgBox "Seleccione la familia a la que pertenece el concepto.", vbInformation, App.Title
        cmbCC.SetFocus
        valida_datos_conceptos = False
        Exit Function
    End If
End Function

Private Sub borrar_campos_conceptos()
    txtdes = ""
    txtMuestra = ""
    txtprecioConcepto = ""
    txtcantidad = "1"
    txtDto = "0"
    txtdes.SetFocus
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
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub
Private Sub txtprecioConcepto_LostFocus()
    txtprecioConcepto.BackColor = &HFFFFFF
    If txtprecioConcepto <> "" Then
        If Not IsNumeric(txtprecioConcepto) Then
            MsgBox "El precio debe ser numérico.", vbInformation, App.Title
            txtprecioConcepto = ""
            txtprecioConcepto.SetFocus
        End If
    End If
End Sub
Private Sub txtprecioConcepto_GotFocus()
    txtprecioConcepto.BackColor = &H80C0FF
    txtprecioConcepto.SelStart = 0
    txtprecioConcepto.SelLength = Len(txtPrecio)
End Sub

Private Sub txtprecioConcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

