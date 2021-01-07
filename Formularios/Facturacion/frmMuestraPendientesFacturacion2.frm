VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmMuestraPendientesFacturacion2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muestras pendientes de facturacion"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   120
   ClientWidth     =   17475
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMuestraPendientesFacturacion2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   17475
   Begin VB.CheckBox chkRevision 
      Caption         =   "Check1"
      Height          =   195
      Left            =   7065
      TabIndex        =   47
      Top             =   8280
      Value           =   1  'Checked
      Width           =   240
   End
   Begin VB.Frame frmDatosEspeciales 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Pedido a la muestra seleccionada"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   4860
      TabIndex        =   39
      Top             =   3780
      Visible         =   0   'False
      Width           =   9135
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   330
         Left            =   945
         TabIndex        =   40
         Top             =   765
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   795
         Left            =   7605
         TabIndex        =   42
         Top             =   315
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "Informar Pedido"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":08CA
      End
      Begin pryCombo.miCombo cmbClienteFactura 
         Height          =   330
         Left            =   945
         TabIndex        =   43
         Top             =   360
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   44
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   135
         TabIndex        =   41
         Top             =   855
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdlog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7965
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   45
      TabIndex        =   30
      Top             =   8280
      Width           =   3795
      Begin XtremeSuiteControls.PushButton cmdbano 
         Height          =   435
         Left            =   135
         TabIndex        =   31
         Top             =   225
         Width           =   1725
         _Version        =   851970
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Baño/Análisis"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":712C
      End
      Begin XtremeSuiteControls.PushButton cmdrec 
         Height          =   435
         Left            =   1890
         TabIndex        =   32
         Top             =   225
         Width           =   1725
         _Version        =   851970
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Recalcular Seleccionada"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":D98E
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   435
         Left            =   135
         TabIndex        =   34
         Top             =   1125
         Width           =   1725
         _Version        =   851970
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Ofertas Cliente"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":141F0
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   435
         Left            =   1890
         TabIndex        =   35
         Top             =   675
         Width           =   1725
         _Version        =   851970
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "F5 - Informar Pedido"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":1AA52
      End
      Begin XtremeSuiteControls.PushButton cmdNoFacturable 
         Height          =   435
         Left            =   1890
         TabIndex        =   36
         Top             =   1125
         Width           =   1725
         _Version        =   851970
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Marcar como NO facturable"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":212B4
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   435
         Left            =   135
         TabIndex        =   33
         Top             =   675
         Width           =   1725
         _Version        =   851970
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Datos Cliente"
         Appearance      =   5
         Picture         =   "frmMuestraPendientesFacturacion2.frx":27B16
      End
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   13365
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0,00 €"
      Top             =   7920
      Width           =   1905
   End
   Begin VB.CommandButton cmdRecalculo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recalcular todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Recalcula el precio de todas las muestras pendientes de facturar"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Leyenda Revisión Facturas"
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   3915
      TabIndex        =   18
      Top             =   8280
      Width           =   3585
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Todo"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   315
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Verde: precio no coincide tarifa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1065
         Width           =   3075
      End
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Azul: muestras con precio 0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   825
         Width           =   3105
      End
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rojo: Revisar Facturación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   585
         Width           =   2805
      End
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Morado: Alguna Determinacion con --"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1305
         Width           =   3075
      End
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7920
      Width           =   1410
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   1410
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6405
      Left            =   45
      TabIndex        =   12
      Top             =   1485
      Width           =   17355
      _ExtentX        =   30612
      _ExtentY        =   11298
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Facturar"
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   7560
      TabIndex        =   10
      Top             =   8280
      Width           =   2895
      Begin VB.CommandButton cmdAlbaran 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear &Albaran"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   1485
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear &Factura"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   45
      TabIndex        =   6
      Top             =   315
      Width           =   17310
      Begin VB.CheckBox chkIberia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iberia"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11115
         TabIndex        =   46
         Top             =   495
         Width           =   870
      End
      Begin VB.CheckBox chkAirbus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10080
         TabIndex        =   45
         Top             =   495
         Width           =   870
      End
      Begin VB.CheckBox chkNF 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ver muestras marcadas como no facturables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6930
         TabIndex        =   23
         Top             =   765
         Width           =   3975
      End
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   19
         Top             =   765
         Width           =   1365
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   375
         Left            =   765
         TabIndex        =   17
         Top             =   270
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdiniciar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   16020
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   1035
      End
      Begin VB.CheckBox chkCerradas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestras abiertas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   765
         Width           =   2310
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10080
         TabIndex        =   0
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14895
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1530
         TabIndex        =   1
         Top             =   720
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3555
         TabIndex        =   2
         Top             =   720
         Width           =   1305
         _ExtentX        =   2302
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
         CalendarTitleBackColor=   12632256
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3015
         TabIndex        =   8
         Top             =   765
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   480
      End
   End
   Begin XtremeSuiteControls.PushButton cmdListado 
      Height          =   1245
      Left            =   14310
      TabIndex        =   29
      Top             =   8595
      Width           =   1500
      _Version        =   851970
      _ExtentX        =   2646
      _ExtentY        =   2196
      _StockProps     =   79
      Caption         =   "Listado Impresora / Excel"
      Appearance      =   5
      Picture         =   "frmMuestraPendientesFacturacion2.frx":2E378
   End
   Begin XtremeSuiteControls.PushButton cmdcancel 
      Height          =   1245
      Left            =   15840
      TabIndex        =   38
      Top             =   8595
      Width           =   1500
      _Version        =   851970
      _ExtentX        =   2646
      _ExtentY        =   2196
      _StockProps     =   79
      Caption         =   "Salir"
      Appearance      =   5
      Picture         =   "frmMuestraPendientesFacturacion2.frx":34BDA
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10530
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":3B43C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":3BD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":3C5F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":3CECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":3D7A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":3E07E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMuestraPendientesFacturacion2.frx":448E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   12060
      TabIndex        =   22
      Top             =   7965
      Width           =   1230
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Botón izquierdo detalle muestra / Botón Derecho Vista Previa Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5310
      TabIndex        =   13
      Top             =   7920
      Width           =   5235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muestras pendientes de facturación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   17835
   End
End
Attribute VB_Name = "frmMuestraPendientesFacturacion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_cTT As New cTooltip
Private Enum COLS
    CODIGO = 0
    CLIENTE_NOMBRE = 1
    TIPO_ANALISIS = 2
    REFERENCIA_CLIENTE = 3
    fecha = 4
    pedido = 5
    CODIGO_TARIFA = 6
    PRECIO = 7
    ID_GENERAL = 8
    ID_MUESTRA = 9
    CLIENTE_ID = 10
    FP_ID = 11
    FACTURA_DETERMINACIONES = 12
    BANO_ID = 13
    ANALISIS_MODIFICADO = 14
    TIPO_ANALISIS_ID = 15
    TARIFA_ID = 16
    PEDIDO_ID = 17
    familia = 18
    fecha_cierre = 19
    urgente = 20
    ajuste = 21
End Enum
Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub chkRevision_Click()
    If chkRevision.Value = Checked Then
        Frame3.Enabled = True
    Else
        Frame3.Enabled = False
    End If
End Sub

Private Sub chkTodos_Click()
    chkAirbus.Enabled = chkTodos.Value
    chkIberia.Enabled = chkTodos.Value
End Sub

Private Sub cmbClienteFactura_change()
    cmbPedidos.limpiar
    If cmbClienteFactura.getTEXTO <> "" Then
        pedidos cmbClienteFactura.getPK_SALIDA
    End If
End Sub

Public Sub cmdbano_Click()
   On Error GoTo cmdbano_Click_Error

        If lista.ListItems.Count = 0 Then
            Exit Sub
        End If
        Dim oMuestra As New clsMuestra
        If oMuestra.CargaMuestra(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)) Then
            If oMuestra.getBANO_ID = 0 Or oMuestra.getANALISIS_MODIFICADO = 2 Then
                frmTA_Detalle.PK = oMuestra.getTIPO_ANALISIS_ID
                frmTA_Detalle.Show 1
            Else
                frmBANO_Detalle.PK = oMuestra.getBANO_ID
                frmBANO_Detalle.Show 1
            End If
        End If
        lista_Click
   On Error GoTo 0
   Exit Sub

cmdbano_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdbano_Click of Formulario frmMuestraPendientesFacturacion2"
End Sub

Private Sub cmdlog_Click()
        On Error GoTo fallo
        If lista.ListItems.Count = 0 Then
            Exit Sub
        End If
        Dim men As String
        Dim total As Currency
        Dim odd As New clsDeterminaciones_analisis
        Dim oCliente As New clsCliente
        Dim oTarifa As New clsTarifas_precios
        Dim oMuestra As New clsMuestra
        If oMuestra.CargaMuestra(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)) Then
            oCliente.CargaCliente (oMuestra.getCLIENTE_ID)
            Dim otar As New clsTarifas
            otar.Carga oCliente.getTARIFA_ID
            men = "TARIFA DEL CLIENTE : " & otar.getNOMBRE & " (" & oCliente.getTARIFA_ID & ")" & vbNewLine
            ' Factura por determinaciones
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.FACTURA_DETERMINACIONES) = 1 Then
                men = men & "El cliente o el análisis se factura por DETERMINACIONES" & vbNewLine
                Dim consulta As String
                consulta = " SELECT td.nombre,tp.PRECIO " & _
                           "  FROM determinaciones d,tipos_determinacion td " & _
                           "  LEFT JOIN tarifas_precios tp on tp.tipo_determinacion_id = td.id_tipo_determinacion " & _
                           " WHERE d.tipo_determinacion_id = td.id_tipo_determinacion" & _
                           "   AND tp.tarifa_id = " & oCliente.getTARIFA_ID & _
                           "   AND d.muestra_id=" & lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)
                Dim rs As ADODB.Recordset
                Set rs = datos_bd(consulta)
                men = men & "--------------------------------------------------------" & vbNewLine
                If rs.RecordCount > 0 Then
                    Do
                        men = men & rs(0) & " (Precio : " & moneda(rs(1)) & ")" & vbNewLine
                        rs.MoveNext
                    Loop Until rs.EOF
                Else
                    men = men & "Las determinaciones no tienen el precio introducido." & vbNewLine
                End If
                men = men & "--------------------------------------------------------" & vbNewLine
                men = men & "Precio TOTAL : " & Format(oMuestra.ImporteMuestraPorDeterminaciones(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA), lista.ListItems(lista.selectedItem.Index).SubItems(COLS.CLIENTE_ID)), "currency")
            Else
                ' No factura por determinaciones
                ' Recuperamos el precio del analisis o bano por tarifa
                ' Miramos si se factura por tipo analisis o control de eficacia
                If oMuestra.getBANO_ID = 0 Or oMuestra.getANALISIS_MODIFICADO = 2 Then
                    Dim oTA As New clsTipos_analisis
                    oTA.CARGAR oMuestra.getTIPO_ANALISIS_ID
                    If oMuestra.getANALISIS_MODIFICADO = 2 Then
                        men = men & "CONTROL DEL EFICACIA : " & oTA.getNOMBRE & " (" & oMuestra.getTIPO_ANALISIS_ID & ")" & vbNewLine
                    Else
                        men = men & "TIPO DE ANÁLISIS : " & oTA.getNOMBRE & " (" & oMuestra.getTIPO_ANALISIS_ID & ")" & vbNewLine
                    End If
                    If oTarifa.Carga_por_TA(oMuestra.getTIPO_ANALISIS_ID, oCliente.getTARIFA_ID) Then
                        total = CCur(Replace(oTarifa.getPRECIO, ".", ","))
                        men = men & "PRECIO DEL ANÁLISIS PARA LA TARIFA : " & Format(total, "CURRENCY") & vbNewLine
                    End If
                Else
                    Dim oBANO As New clsBanos
                    oBANO.cargar_bano (oMuestra.getBANO_ID)
                    men = men & "BAÑO : " & oBANO.getNOMBRE & " (" & oMuestra.getBANO_ID & ")" & vbNewLine
                    If oTarifa.Carga_por_BANO(oMuestra.getBANO_ID, oCliente.getTARIFA_ID) Then
                        total = CCur(Replace(oTarifa.getPRECIO, ".", ","))
                        men = men & "PRECIO DEL BAÑO PARA LA TARIFA : " & Format(total, "CURRENCY") & vbNewLine
                    End If
                End If
                ' Recuperamos los datos por defecto del analisis o bano
                If oMuestra.getBANO_ID = 0 Then
                    men = men & "PRECIO POR DETERMINACIONES : " & Format(odd.Precio_determinaciones_por_tipo_analisis(oMuestra.getTIPO_ANALISIS_ID, lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA), oCliente.getTARIFA_ID), "CURRENCY") & vbNewLine
                    total = total + odd.Precio_determinaciones_por_tipo_analisis(oMuestra.getTIPO_ANALISIS_ID, lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA), oCliente.getTARIFA_ID)
                Else
                    men = men & "PRECIO POR DETERMINACIONES : " & Format(odd.Precio_determinaciones_por_bano(oMuestra.getBANO_ID, lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA), oCliente.getTARIFA_ID), "CURRENCY") & vbNewLine
                    total = total + odd.Precio_determinaciones_por_bano(oMuestra.getBANO_ID, lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA), oCliente.getTARIFA_ID)
                End If
                men = men & "PRECIO TOTAL MUESTRA : " & Format(total, "CURRENCY") & vbNewLine
            End If
        ' Actualizamos el precio de la muestra
'        omuestra.actualizar_precio MUESTRA, Replace(total, ",", ".")
        m_cTT.ToolText(lista) = men
'        MsgBox men
        End If
    Set odd = Nothing
'    Set oDeter = Nothing
    Set oMuestra = Nothing
    Exit Sub
fallo:
    MsgBox "Error al obtener el tipo de documento de la muestra.", vbCritical, Err.Description

End Sub

Private Sub cmdNoFacturable_Click()
    Dim strcadena As String
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar alguna muestra", vbInformation, App.Title
        Exit Sub
    End If
    If contar_marcados = 1 Then
        strcadena = "Va a marcar como no facturable 1 muestra. ¿Desea continuar?"
    Else
        strcadena = "Va a marcar como no facturable " & contar_marcados & " muestras. ¿Desea continuar?"
    End If
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        Dim i As Integer
        Dim oMuestra As New clsMuestra
        For i = lista.ListItems.Count To 1 Step -1
            If lista.ListItems(i).Checked = True Then
                oMuestra.Informar_Documento_Pago lista.ListItems(i).SubItems(COLS.ID_MUESTRA), 99
                lista.ListItems.Remove i
            End If
        Next
        Me.MousePointer = 0
    End If

End Sub

Public Sub cmdrec_Click()
    If lista.ListItems.Count > 0 Then
        Dim oMuestra As New clsMuestra
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Selected = True Then
                oMuestra.informar_precio_muestra (lista.ListItems(i).SubItems(COLS.ID_MUESTRA))
'                buscar CLng(lista.ListItems(i).SubItems(COLS.ID_MUESTRA)), lista.selectedItem.Index
                buscar CLng(lista.ListItems(i).SubItems(COLS.ID_MUESTRA)), lista.ListItems(i).Index
            End If
        Next
    End If
End Sub

Private Sub cmdRecalculo_Click()
    If MsgBox("¿Esta seguro de recalcular el precio de las muestras sin facturar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        Dim oMuestra As New clsMuestra
        If oMuestra.recalcular_precios_muestras_sin_facturar Then
            Me.MousePointer = 0
            MsgBox "Se han recalculado los precios correctamente.", vbInformation, App.Title
            cmdBuscar_Click
        End If
        Me.MousePointer = 0
    End If
End Sub
Private Sub cmdAlbaran_Click()
    Dim strcadena As String
   On Error GoTo cmdAlbaran_Click_Error

    If contar_marcados = 0 Then
         MsgBox "Debe seleccionar alguna muestra", vbInformation, App.Title
         Exit Sub
    End If
    If contar_marcados = 1 Then
        strcadena = "Va a generar un albaran a 1 muestra. ¿Desea continuar?"
    Else
        strcadena = "Va a generar albaranes a " & contar_marcados & " muestras. ¿Desea continuar?"
    End If
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
         generar_documentos (1) ' Factura
    End If
    ' Modificar para que no recalcule, si no que elimine de la lista lo
    ' que se acaba de facturar
    Dim i As Integer
    For i = lista.ListItems.Count To 1 Step -1
        If lista.ListItems(i).Checked = True Then
            lista.ListItems.Remove i
        End If
    Next
'    cmdBuscar_Click

   On Error GoTo 0
   Exit Sub

cmdAlbaran_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAlbaran_Click of Formulario frmMuestraPendientesFacturacion2"
End Sub

Private Sub cmdBuscar_Click()
   buscar 0, 0
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdFactura_Click()
    Dim strcadena As String
    Dim i As Integer
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar alguna muestra", vbInformation, App.Title
        Exit Sub
    End If
    ' Validar que todas las muestras tengan el mismo pedido
    Dim pedido As Long
    pedido = 1 * (-1)
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If pedido = -1 Then
                pedido = lista.ListItems(i).SubItems(COLS.PEDIDO_ID)
            Else
                If pedido <> lista.ListItems(i).SubItems(COLS.PEDIDO_ID) Then
                    MsgBox "No se pueden facturar muestras de distintos pedidos.", vbExclamation, App.Title
                    Exit Sub
                End If
            End If
        End If
    Next
    If contar_marcados = 1 Then
        strcadena = "Va a facturar 1 muestra. ¿Desea continuar?"
    Else
        strcadena = "Va a facturar " & contar_marcados & " muestras. ¿Desea continuar?"
    End If
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        generar_documentos (2) ' Factura
        Me.MousePointer = 0
    End If
    ' Modificar para que no recalcule, si no que elimine de la lista lo
    ' que se acaba de facturar
    For i = lista.ListItems.Count To 1 Step -1
        If lista.ListItems(i).Checked = True Then
            lista.ListItems.Remove i
        End If
    Next
'    Call cmdBuscar_Click
End Sub

Private Sub cmdiniciar_Click()
    cmbclientes.limpiar
    chkTodos.Value = Unchecked
    chkCerradas.Value = Unchecked
    fdesde.Value = Date
    fhasta.Value = Date
    lista.ListItems.Clear
End Sub

Private Sub cmdListado_Click()
    If MsgBox("¿Desea exportar a excel?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        generar_excel_listado
        
    Else
        Dim total As Currency
        Dim i As Integer
        On Error GoTo fallo
        If lista.ListItems.Count = 0 Then
            MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
            Exit Sub
        End If
        Dim rs As New ADODB.Recordset
        rs.Fields.Append "c1", adChar, 5, adFldUpdatable
        rs.Fields.Append "c2", adChar, 50, adFldUpdatable
        rs.Fields.Append "c3", adChar, 50, adFldUpdatable
        rs.Fields.Append "c4", adChar, 12, adFldUpdatable
        rs.Open
        total = 0
        For i = 1 To lista.ListItems.Count
            rs.AddNew
            rs("c1") = lista.ListItems(i).SubItems(COLS.ID_GENERAL)
            rs("c2") = Left(lista.ListItems(i).SubItems(COLS.CLIENTE_NOMBRE), 50)
            rs("c3") = Left(lista.ListItems(i) & " " & lista.ListItems(i).SubItems(COLS.TIPO_ANALISIS), 50)
            rs("c4") = lista.ListItems(i).SubItems(COLS.PRECIO)
            If Trim(lista.ListItems(i).SubItems(COLS.PRECIO)) <> "" Then
                total = total + Format(lista.ListItems(i).SubItems(COLS.PRECIO), "currency")
            End If
            rs.Update
        Next
        ' Generar Listado
        Dim Listado As New dataListadoMuestrasPendientes
        ' Cabecera
        With Listado.Sections("cabecera")
            .Controls("lbltitulo").Caption = "Análisis pendientes de facturar del " & Format(fdesde, "dd/mm/yyyy") & " al " & Format(fhasta, "dd/mm/yyyy")
            If chkTodos.Value = Checked Then
                .Controls("lblcliente").Caption = "Cliente : *** TODOS ***"
            Else
                Dim oCliente As New clsCliente
                oCliente.CargaCliente cmbclientes.getPK_SALIDA
                .Controls("lblcliente").Caption = "Cliente : " & oCliente.getNOMBRE
        
            End If
        End With
        Set Listado.Sections("cabecera").Controls("logo").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
        'Detalle
        With Listado.Sections("detalle")
            .Controls("c1").DataField = rs.Fields("c1").Name
            .Controls("c2").DataField = rs.Fields("c2").Name
            .Controls("c3").DataField = rs.Fields("c3").Name
            .Controls("c4").DataField = rs.Fields("c4").Name
        End With
        ' Pie de Pagina
        With Listado.Sections("pie")
            .Controls("lbltotal").Caption = Format(total, "currency")
        End With
        Set Listado.DataSource = rs
        Listado.Caption = "Listado de Análisis Pendientes"
        Listado.WindowState = vbNormal
        Listado.Show
        Set rs = Nothing
        '    Me.Height = 7890
        '    Me.Width = 12780
    End If
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de Analisis pendientes.", vbCritical, Err.Description
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        PushButton4_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cargar_combo_clientes
'    cargar_combo_clientes_pedidos
    cabecera_grid
    fhasta = Now
    fdesde = Now - 30
    tool
End Sub
Private Sub buscar(MUESTRA_ID As Long, linea As Long)  'obtengo el listado de las muestras pendientes de facturacion para el cliente seleccionado y lo vuelco en el listbox
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim oMuestra As New clsMuestra
    Dim oCodigo As New clsTarifas_codigos
    Dim cliente As Long
    Dim color As Long
    If chkTodos.Value = 0 And cmbclientes.getTEXTO = "" Then
        MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
        Exit Sub
    End If
    Me.MousePointer = 11
    If MUESTRA_ID = 0 Then
        lista.ListItems.Clear
    End If
    If cmbclientes.getTEXTO <> "" Then
        cliente = cmbclientes.getPK_SALIDA
    End If
    Dim pedido As Long
    Set rs = oMuestra.Muestras_pendientes_facturar(MUESTRA_ID, fdesde.Value, fhasta.Value, cliente, chkCerradas.Value, pedido, chkFecha.Value, chkNF.Value, chkAirbus.Value, chkIberia.Value)
    Dim cont As Integer
    cont = 0
    Dim i As Integer
    If rs.RecordCount > 0 Then
        If MUESTRA_ID = 0 Then
            Label1(4).Caption = "Muestras pendientes de facturación. Total: " & rs.RecordCount
            While Not rs.EOF
                i = i + 1
                With lista.ListItems.Add(, , rs.Fields(1))
                .SubItems(COLS.CLIENTE_NOMBRE) = rs.Fields(2)
                .SubItems(COLS.TIPO_ANALISIS) = rs.Fields(10)
                .SubItems(COLS.REFERENCIA_CLIENTE) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                    .SubItems(COLS.fecha) = rs.Fields(5)
                End If
                '** rs(9) El cliente factura por determinaciones
                '** rs(13) El tipo de analisis es por determinaciones
                '** rs(14) BANO_ID
                '** rs(19) BANO -> FACTURA_DETERMINACIONES
                ' Cliente factura por determinaciones O
                ' No es baño y el tipo de analisis se factura por determinaciones O
                ' Es baño y el baño se factura por determinaciones Y NO ES CE
                If rs(9) = 1 Or (rs(14) = 0 And rs(13) = 1) Or (rs(14) <> 0 And rs(19) = 1 And rs(15) <> 2) Then
                    .SubItems(COLS.PRECIO) = Format(oMuestra.ImporteMuestraPorDeterminaciones(rs(8), rs(0)), "currency")
                    .SubItems(COLS.FACTURA_DETERMINACIONES) = 1
'                    lista.ListItems(lista.ListItems.Count).SmallIcon = 7
                Else
                    If Not IsNull(rs.Fields(7)) Then
                        .SubItems(COLS.PRECIO) = Format(rs.Fields(7), "currency")
                    End If
                    .SubItems(COLS.FACTURA_DETERMINACIONES) = 0
'                    lista.ListItems(lista.ListItems.Count).SmallIcon = 6
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(COLS.ID_GENERAL) = Format(rs.Fields(6), "00000")
                End If
                If Not IsNull(rs.Fields(8)) Then
                    .SubItems(COLS.ID_MUESTRA) = rs.Fields(8)  ' ID_MUESTRA
                End If
                .SubItems(COLS.CLIENTE_ID) = rs.Fields(0)  ' CLIENTE_ID
                .SubItems(COLS.FP_ID) = rs.Fields(12)  'FP
                .SubItems(COLS.BANO_ID) = rs(14)  ' BANO_ID
                .SubItems(COLS.ANALISIS_MODIFICADO) = rs(15)  ' ANALISIS_MODIFICADO
                .SubItems(COLS.TIPO_ANALISIS_ID) = rs(3)    ' TIPO_ANALISIS_ID
                .SubItems(COLS.TARIFA_ID) = rs(16)  ' TARIFA
                ' PEDIDO
                If rs(17) = 0 Then
                    .SubItems(COLS.pedido) = " "
                    .SubItems(COLS.PEDIDO_ID) = "0"
                Else
                    .SubItems(COLS.pedido) = rs(18)
                    .SubItems(COLS.PEDIDO_ID) = rs(17)
                End If
                'CÓDIGO TARIFA
                 If lista.ListItems(lista.ListItems.Count).SubItems(COLS.FACTURA_DETERMINACIONES) = 0 Then
                   If rs(14) <> 0 And rs(15) <> 2 Then
                      .SubItems(COLS.CODIGO_TARIFA) = oCodigo.Codigo_Bano(rs(14))
                   Else
                      .SubItems(COLS.CODIGO_TARIFA) = oCodigo.Codigo_TipoAnalisis(rs(3))
                   End If
                 End If
                 If Not IsNull(rs(24)) Then
                     .SubItems(COLS.familia) = rs(24)
                 End If
                 If Not IsNull(rs(25)) Then
                     .SubItems(COLS.fecha_cierre) = Format(rs(25), "dd/mm/yyyy")
                 Else
                     .SubItems(COLS.fecha_cierre) = ""
                 End If
                 If rs(26) = 1 Then
                     .SubItems(COLS.urgente) = "X"
                 Else
                     .SubItems(COLS.urgente) = ""
                 End If
                 If rs(27) = 1 Then
                     .SubItems(COLS.ajuste) = "X"
                 Else
                     .SubItems(COLS.ajuste) = ""
                 End If
                 
                End With
            
                'ANALISIS MUESTRA FACTURADA
                If chkRevision.Value = Checked Then
                    color = analizar_muestra(rs(8), rs(14), rs(15), rs(3), lista.ListItems(lista.ListItems.Count).SubItems(COLS.PRECIO), lista.ListItems(lista.ListItems.Count).SubItems(COLS.FACTURA_DETERMINACIONES), rs(0), rs(16))
                    If color <> 0 Then
                        colorear lista.ListItems.Count, color
                    End If
                    If color <> 0 Or opLeyenda(4).Value = False Then
                        ' ROJO : 255
                        ' AZUL : 16711680
                        ' MORADO : 12583104
                        If opLeyenda(0).Value = True And color <> 255 Then
                            lista.ListItems.Remove (lista.ListItems.Count)
                        End If
                        If opLeyenda(1).Value = True And color <> 16711680 Then
                            lista.ListItems.Remove (lista.ListItems.Count)
                        End If
                        If opLeyenda(2).Value = True And color <> &H8000& Then
                            lista.ListItems.Remove (lista.ListItems.Count)
                        End If
                        If opLeyenda(3).Value = True And color <> 12583104 Then
                            lista.ListItems.Remove (lista.ListItems.Count)
                        End If
                    End If
                End If
                icono_lista lista.ListItems.Count, rs(20), rs(21), rs(22), rs(23)
                rs.MoveNext
            Wend
            cmdFactura.Enabled = True
            cmdAlbaran.Enabled = True
            If lista.ListItems.Count > 0 Then
                lista.ListItems(1).EnsureVisible
            End If
        Else
            lista.ListItems(linea).Text = rs.Fields(1)
            lista.ListItems(linea).SubItems(COLS.CLIENTE_NOMBRE) = rs.Fields(2)
            lista.ListItems(linea).SubItems(COLS.TIPO_ANALISIS) = rs.Fields(10)
            lista.ListItems(linea).SubItems(COLS.REFERENCIA_CLIENTE) = rs.Fields(4)
            If Not IsNull(rs.Fields(5)) Then
               lista.ListItems(linea).SubItems(COLS.fecha) = rs.Fields(5)
            End If
            If rs(9) = 1 Or (rs(14) = 0 And rs(13) = 1) Or (rs(14) <> 0 And rs(19) = 1 And rs(15) <> 2) Then
                lista.ListItems(linea).SubItems(COLS.PRECIO) = Format(oMuestra.ImporteMuestraPorDeterminaciones(rs(8), rs(0)), "currency")
                lista.ListItems(linea).SubItems(COLS.FACTURA_DETERMINACIONES) = 1
'                lista.ListItems(linea).SmallIcon = 7
            Else
                If Not IsNull(rs.Fields(7)) Then
                    lista.ListItems(linea).SubItems(COLS.PRECIO) = Format(rs.Fields(7), "currency")
                End If
                lista.ListItems(linea).SubItems(COLS.FACTURA_DETERMINACIONES) = 0
'                lista.ListItems(linea).SmallIcon = 6
            End If
            If Not IsNull(rs.Fields(6)) Then
                   lista.ListItems(linea).SubItems(COLS.ID_GENERAL) = Format(rs.Fields(6), "00000")
            End If
            If Not IsNull(rs.Fields(8)) Then
                   lista.ListItems(linea).SubItems(COLS.ID_MUESTRA) = rs.Fields(8)
            End If
            lista.ListItems(linea).SubItems(COLS.CLIENTE_ID) = rs.Fields(0)
            lista.ListItems(linea).SubItems(COLS.FP_ID) = rs.Fields(12)
            lista.ListItems(linea).SubItems(COLS.BANO_ID) = rs(14) ' BANO_ID
            lista.ListItems(linea).SubItems(COLS.ANALISIS_MODIFICADO) = rs(15) ' ANALISIS_MODIFICADO
            lista.ListItems(linea).SubItems(COLS.TIPO_ANALISIS_ID) = rs(3)  ' TIPO_ANALISIS_ID
            lista.ListItems(linea).SubItems(COLS.TARIFA_ID) = rs(16)  ' TIPO_ANALISIS_ID
            ' PEDIDO
            If rs(17) = 0 Then
                lista.ListItems(linea).SubItems(COLS.pedido) = " "
                lista.ListItems(linea).SubItems(COLS.PEDIDO_ID) = "0"
            Else
                lista.ListItems(linea).SubItems(COLS.pedido) = rs(18)
                lista.ListItems(linea).SubItems(COLS.PEDIDO_ID) = rs(17)
            End If
            'CÓDIGO TARIFA
            If lista.ListItems(linea).SubItems(COLS.FACTURA_DETERMINACIONES) = 0 Then
               If rs(14) <> 0 And rs(15) <> 2 Then
                  lista.ListItems(linea).SubItems(COLS.CODIGO_TARIFA) = oCodigo.Codigo_Bano(rs(14))
               Else
                  lista.ListItems(linea).SubItems(COLS.CODIGO_TARIFA) = oCodigo.Codigo_TipoAnalisis(rs(3))
               End If
            End If
                 If Not IsNull(rs(25)) Then
                     lista.ListItems(linea).SubItems(COLS.fecha_cierre) = Format(rs(25), "dd/mm/yyyy")
                 Else
                     lista.ListItems(linea).SubItems(COLS.fecha_cierre) = ""
                 End If
                 If rs(26) = 1 Then
                     lista.ListItems(linea).SubItems(COLS.urgente) = "X"
                 Else
                     lista.ListItems(linea).SubItems(COLS.urgente) = ""
                 End If
                 If rs(27) = 1 Then
                     lista.ListItems(linea).SubItems(COLS.ajuste) = "X"
                 Else
                     lista.ListItems(linea).SubItems(COLS.ajuste) = ""
                 End If
            
            'ANALISIS MUESTRA FACTURADA
            If chkRevision.Value = Checked Then
                color = analizar_muestra(rs(8), rs(14), rs(15), rs(3), lista.ListItems(linea).SubItems(COLS.PRECIO), lista.ListItems(linea).SubItems(COLS.FACTURA_DETERMINACIONES), rs(0), rs(16))
                icono_lista linea, rs(20), rs(21), rs(22), rs(23)
                colorear CInt(linea), color
            End If
            cont = cont + 1
            If cont = 100 Then
                cont = 0
                DoEvents
            End If
        End If
        lista_Click
    Else
        cmdFactura.Enabled = False
        cmdAlbaran.Enabled = False
'        lblmsg.Caption = "No existe ninguna muestra por facturar con esos criterios."
    End If
'    Dim i As Integer
'    If MUESTRA_ID = 0 Then
        
'        For i = 1 To lista.ListItems.Count
'            lista.ListItems(i).Checked = True
'        Next
'    End If
    calcular_total
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras (frmMuestrasPendientesFacturar). Indice : " & i, vbCritical, Err.Description
End Sub

Private Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To lista.ListItems.Count
       If lista.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function

Private Sub generar_documentos(TIPO_DOCUMENTO As Integer)
    Dim i As Integer
    Dim num_doc As Long
    Dim cliente_ant As Long
    Dim total_doc As Integer
   On Error GoTo generar_documentos_Error

    total_doc = 0
    cliente_ant = 0
    ReDim documentos_pago(lista.ListItems.Count)
    Dim oDocPago As New clsDocs_pago
    Dim odoc_muestra As New clsDocs_pago_muestras
    Dim oMuestra As New clsMuestra
    'cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    Dim ORDEN As Integer
    ORDEN = 1
    Dim oCliente As New clsCliente
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If cliente_ant <> lista.ListItems(i).SubItems(COLS.CLIENTE_ID) Then
                oDocPago.setTIPO = TIPO_DOCUMENTO
                oDocPago.setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                oDocPago.setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                oDocPago.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                oDocPago.setCLIENTE_ID = lista.ListItems(i).SubItems(COLS.CLIENTE_ID)
                oDocPago.setCLIENTE_ID_FACTURA = lista.ListItems(i).SubItems(COLS.CLIENTE_ID)
                oDocPago.setFP_ID = lista.ListItems(i).SubItems(COLS.FP_ID)
'REVISAR
                oDocPago.setPEDIDO_ID = lista.ListItems(i).SubItems(COLS.PEDIDO_ID)
'REVISAR
                oDocPago.setTOTAL = "0.00"
                oDocPago.setDESCUENTO = "0.00"
'                If TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA Then
'                    oCliente.CargaCliente lista.ListItems(i).SubItems(COLS.CLIENTE_ID)
'                    If oCliente.getINTRA = 1 Or oCliente.getEXTRANJERO = 0 Then
'                        oDocPago.setIVA = 0
'                    Else
'                        oDocPago.setIVA = IVA
'                    End If
'                Else
'                    oDocPago.setIVA = 0
'                End If
                oDocPago.setPAGADO = 0
                oDocPago.setANULADO = 0
                oDocPago.setFACTURA_CONCEPTOS = 0
                ' Insertamos el documento de pago
                num_doc = oDocPago.InsertarDocPago
                If num_doc = 0 Then
                    MsgBox "Error al generar las facturas, contacte con mantenimiento.", vbCritical, App.Title
                    Exit Sub
                End If
                total_doc = total_doc + 1
                documentos_pago(total_doc) = num_doc
            End If
            ' Insertar el Documento de Pago de Muestras
            log "num_doc : " & num_doc
            odoc_muestra.setDOC_ID = num_doc
            odoc_muestra.setMUESTRA_ID = lista.ListItems(i).SubItems(COLS.ID_MUESTRA)
            odoc_muestra.setTIPO_ANALISIS = lista.ListItems(i).SubItems(COLS.TIPO_ANALISIS)
            odoc_muestra.setFECHA = Format(lista.ListItems(i).SubItems(COLS.fecha), "yyyy-mm-dd")
            odoc_muestra.setREFERENCIA_CLIENTE = lista.ListItems(i).SubItems(COLS.REFERENCIA_CLIENTE)
            log "precio :  " & Replace(Format(lista.ListItems(i).SubItems(COLS.PRECIO), "0.00"), ",", ".")
            odoc_muestra.setPRECIO = Replace(Format(lista.ListItems(i).SubItems(COLS.PRECIO), "0.00"), ",", ".")
            log "deter : " & lista.ListItems(i).SubItems(COLS.FACTURA_DETERMINACIONES)
            odoc_muestra.setORDEN = ORDEN
            ORDEN = odoc_muestra.Insertar_doc_pago_muestra(lista.ListItems(i).SubItems(COLS.FACTURA_DETERMINACIONES))
            If ORDEN = -1 Then
                MsgBox "Error al generar las facturas (2), contacte con mantenimiento.", vbCritical, App.Title
                Exit Sub
            Else
                ORDEN = ORDEN + 1
            End If
            ' Modificar el documento de pago de la muestra
            If oMuestra.Informar_Documento_Pago(lista.ListItems(i).SubItems(COLS.ID_MUESTRA), TIPO_DOCUMENTO) = False Then
                MsgBox "Error al informar el documento de pago, contacte con mantenimiento.", vbCritical, App.Title
                Exit Sub
            End If
            cliente_ant = lista.ListItems(i).SubItems(COLS.CLIENTE_ID)
        End If
    Next
    Set oMuestra = Nothing
    Set oDocPago = Nothing
    Set odoc_muestra = Nothing
    Dim sTIPO As String
    If TIPO_DOCUMENTO = 1 Then
        sTIPO = "Albaran"
    Else
        sTIPO = "Factura"
    End If
    If total_doc = 1 Then
        MsgBox "Se ha registrado 1 " & sTIPO & ".", vbOKOnly + vbInformation, App.Title
    Else
        MsgBox "Se han registrado " & total_doc & " " & sTIPO & "s.", vbOKOnly + vbInformation, App.Title
    End If
    ' LlamarMas Datos de la factura
    numero_documentos_pago = total_doc
    frmMasDatosFactura.Show 1

   On Error GoTo 0
   Exit Sub

generar_documentos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_documentos of Formulario frmMuestraPendientesFacturacion2"
End Sub

Private Sub cabecera_grid()
    With lista.ColumnHeaders
        .Add , , "Código", 1300, lvwColumnLeft
        .Add , , "Cliente", 2250, lvwColumnLeft
        .Add , , "Analisis", 2250, lvwColumnLeft
        .Add , , "Ref.Cliente", 2250, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Pedido", 1100, lvwColumnCenter
        .Add , , "C.Tarifa", 1000, lvwColumnCenter
        .Add , , "Precio", 1000, lvwColumnRight
        .Add , , "General", 700, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
        .Add , , "CLIENTE_ID", 1, lvwColumnCenter
        .Add , , "FP_ID", 1, lvwColumnCenter
        .Add , , "FACTURA_DETERMINACIONES", 1, lvwColumnCenter
        .Add , , "BANO_ID", 1, lvwColumnCenter
        .Add , , "ANALISIS_MODIFICADO", 1, lvwColumnCenter
        .Add , , "TIPO_ANALISIS_ID", 1, lvwColumnCenter
        .Add , , "TARIFA_ID", 1, lvwColumnCenter
        .Add , , "PEDIDO_ID", 1, lvwColumnCenter
        .Add , , "Familia", 1600, lvwColumnLeft
        .Add , , "F.Cierre", 1050, lvwColumnCenter
        .Add , , "URGENTE", 750, lvwColumnCenter
        .Add , , "AJUSTE", 750, lvwColumnCenter
    End With
End Sub

Public Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmdlog_Click
        cmbClienteFactura.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.CLIENTE_ID)
        If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.PEDIDO_ID) <> 0 Then
            cmbPedidos.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.PEDIDO_ID)
        Else
            pedidos cmbClienteFactura.getPK_SALIDA
        End If
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

Public Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)
        frmVerMuestra.Show 1
        buscar CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)), lista.selectedItem.Index
        gmuestra = 0
    End If
End Sub
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub tool()
   On Error GoTo tool_Error

   With m_cTT
    ' Creamos el toolTip pasandole el nombre del Formulario
    Call .Create(Me)
    'Establecemos el Ancho del ToolTip
    .MaxTipWidth = 600
    ' establece los márgenes
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    ' Establecemos el tiempo que se muestra ( 7 segundos )
    .DelayTime(ttDelayShow) = 10000
    ' Agregamos un ToolTip al FileListBox
    'Para agregar mas controles solo hay que añadir uno por uno
    'Nota: solo es valido usar controles que posean HWND
    .AddTool lista
   End With

   On Error GoTo 0
   Exit Sub

tool_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tool of Formulario frmMuestraPendientesFacturacion2"
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        PushButton4_Click
    End If
End Sub

Private Sub lista_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
'    If Button And vbRightButton Then PopupMenu frmMenu.menuopciones
    If Button And vbRightButton Then cmdprevia_Click
End Sub
Private Sub colorear(fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub
Private Sub cargar_combo_clientes()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
                   "  FROM CLIENTES AS C, MUESTRAS AS M " & _
                   " WHERE C.ID_CLIENTE = M.CLIENTE_ID " & _
                   "   AND M.DOCUMENTO_PAGO=0 AND ANULADA = 0"
        With cmbclientes
            .setCONN = conn
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "CLIENTES"
                .setDESCRIPCION = "Clientes"
                .setPK = "ID_CLIENTE"
                .setCAMPO = "NOMBRE"
                .setQUERY = consulta
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = frmClientes
        End With
    
        With cmbClienteFactura
            .setCONN = conn
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "CLIENTES"
                .setDESCRIPCION = "Clientes"
                .setPK = "ID_CLIENTE"
                .setCAMPO = "NOMBRE"
                .setQUERY = consulta
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = frmClientes
        End With
    End If
End Sub
Private Sub cargar_combo_clientes_pedidos()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
                   "  FROM CLIENTES AS C, CLIENTES_PEDIDOS CP " & _
                   " WHERE C.ID_CLIENTE = CP.CLIENTE_ID " & _
                   "   AND C.ANULADO = 0"
        With cmbclientes
            .setCONN = conn
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "CLIENTES"
                .setDESCRIPCION = "Clientes"
                .setPK = "ID_CLIENTE"
                .setCAMPO = "NOMBRE"
                .setQUERY = consulta
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = frmClientes
        End With
    End If
End Sub

Private Function analizar_muestra(muestra As Long, BANO As Long, ANALISIS_MODIFICADO As Long, TIPO_ANALISIS As Long, PRECIO As String, FACTURA_DETERMINACIONES As Integer, cliente As Long, TARIFA As Long) As Long
    Dim i As Integer
'    Dim omuestra As New clsMuestra
    Dim oBANO As New clsBanos
    Dim oTA As New clsTipos_analisis
    Dim oDeterminacion As New clsDeterminaciones
    Dim oTarifa As New clsTarifas_precios
    Dim oCodigo As New clsTarifas_codigos
    Dim CODIGO As String
'    Dim ocliente As New clsCliente
    Dim color As Long
    ' ROJO: Muestras con TA, BANO o DETERMICION con check de REVISAR_FACTURA
    ' AZUL: Muestras con precio 0
    ' VERDE: Muestras con precio distinto al de su código de tarifa
    color = 0
    ' Analizar check de REVISAR_FACTURACION (TA y BANO)
    If BANO = 0 Or ANALISIS_MODIFICADO = 2 Then
        If oTA.revisar_facturacion(TIPO_ANALISIS) Then
            color = &HFF& ' ROJO
        End If
    Else
        If oBANO.revisar_facturacion(BANO) Then
            color = &HFF& ' ROJO
        End If
    End If
    ' Analizar check de REVISAR_FACTURACION (DETERMINACIONES)
    If color = 0 Then
        If oDeterminacion.revisar_facturacion(muestra) Then
            color = &HFF& ' ROJO
        End If
    End If
    If color = 0 Then
        If oDeterminacion.revisar_no_realizado(muestra) Then
            color = &HC000C0 ' MORADO
        End If
    End If
    ' Analizar precio de la muestra
    If color = 0 Then
        If CCur(PRECIO) = 0 Then
            color = &HFF0000    ' AZUL
        End If
    End If
    ' Analizar precio de la tarifa
    If color = 0 Then
        If FACTURA_DETERMINACIONES = 0 Then ' No factura determinaciones
           ' Baño
           If BANO <> 0 And ANALISIS_MODIFICADO <> 2 Then
              oTarifa.Carga_por_BANO BANO, TARIFA
           Else
              ' Tipo analisis o CE
              oTarifa.Carga_por_TA TIPO_ANALISIS, TARIFA
           End If
           If Trim(oTarifa.getPRECIO) = "" Then
                oTarifa.setPRECIO = 0
           End If
           If CCur(PRECIO) <> CCur(oTarifa.getPRECIO) Then
              color = &H8000&     ' VERDE
           End If
        End If
    End If
    analizar_muestra = color
End Function
Private Sub cmdprevia_Click()
    If lista.ListItems.Count > 0 Then
        frmDocumento_Previo_Facturacion.PK_MUESTRA = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)
        frmDocumento_Previo_Facturacion.FACTURA_DETERMINACIONES = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.FACTURA_DETERMINACIONES)
        frmDocumento_Previo_Facturacion.Show 1
        buscar CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)), lista.selectedItem.Index
    End If
End Sub


Private Sub calcular_total()
    Dim i As Integer
    Dim total As Currency
    total = 0
    For i = 1 To lista.ListItems.Count
        total = total + lista.ListItems(i).SubItems(COLS.PRECIO)
    Next
    txttotal = moneda(CStr(total))
End Sub

Private Sub opLeyenda_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub PushButton1_Click()
    Dim oMuestra As New clsMuestra
    oMuestra.informar_pedido CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)), cmbPedidos.getPK_SALIDA
    Set oMuestra = Nothing
    buscar CLng(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.ID_MUESTRA)), lista.selectedItem.Index
    If lista.ListItems.Count > lista.selectedItem.Index Then
        If lista.ListItems(lista.selectedItem.Index).SubItems(10) <> lista.ListItems(lista.selectedItem.Index + 1).SubItems(10) Then
            pedidos lista.ListItems(lista.selectedItem.Index + 1).SubItems(10)
        End If
        Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
        lista.selectedItem.EnsureVisible
        frmDatosEspeciales.top = lista.ListItems(lista.selectedItem.Index).top + 600
        lista.SetFocus
    End If
End Sub

Private Sub PushButton2_Click()
    If lista.ListItems.Count > 0 Then
        frmClientes.PK = lista.ListItems(lista.selectedItem.Index).SubItems(10)
        frmClientes.Show 1
    End If
End Sub

Private Sub PushButton3_Click()
    If lista.ListItems.Count > 0 Then
        frmOferta_Listado.pk_CLIENTE = lista.ListItems(lista.selectedItem.Index).SubItems(10)
        frmOferta_Listado.Show
    End If
End Sub

Private Sub PushButton4_Click()
'    pedidos lista.ListItems(lista.selectedItem.Index).SubItems(10)
    frmDatosEspeciales.visible = Not frmDatosEspeciales.visible
    If lista.ListItems.Count = 0 Then
       frmDatosEspeciales.top = Me.Height / 2 - frmDatosEspeciales.Height
    Else
        frmDatosEspeciales.top = lista.ListItems(lista.selectedItem.Index).top + 600
    End If
End Sub
Private Sub pedidos(ID As Long)
'    cmbPedidos.Limpiar
    Dim filtro As String
    If ID <> 0 Then
        filtro = " AND CLIENTE_ID = " & ID & " AND FECHA_BAJA >= '" & Format(lista.ListItems(lista.selectedItem.Index).SubItems(4), "YYYY-MM-DD") & "'"
    End If
    llenar_combo cmbPedidos, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub
Private Sub icono_lista(linea As Long, enviado_correo As Integer, ANULADA As Integer, CERRADA As Integer, revision_usuario As Integer)
    If enviado_correo <> 0 Then 'ENVIADO_CORREO
        lista.ListItems(linea).SmallIcon = 1
        lista.ListItems(linea).ToolTipText = "Enviado Correo"
    Else
        If ANULADA <> 0 Then ' ANULADA
            lista.ListItems(linea).SmallIcon = 2
            lista.ListItems(linea).ToolTipText = "Anulada"
        Else
            Select Case CERRADA ' Cerrada
                Case 0 ' Abierta
                    lista.ListItems(linea).SmallIcon = 5
                    lista.ListItems(linea).ToolTipText = "Abierta"
                Case 1 ' Cerrada
                    If revision_usuario = 0 Then ' Revision Usuario
                        lista.ListItems(linea).SmallIcon = 6
                        lista.ListItems(linea).ToolTipText = "Cerrada Pendiente Revisar"
                    Else
                        lista.ListItems(linea).SmallIcon = 4
                        lista.ListItems(linea).ToolTipText = "Cerrada y Revisada por Usuario : " & revision_usuario
                    End If
                Case 2 ' Pdte. Cierre
                    lista.ListItems(linea).SmallIcon = 3
                    lista.ListItems(linea).ToolTipText = "Pdte. Cierre"
            End Select
        End If
    End If
End Sub

Private Sub generar_excel_listado()
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
   On Error GoTo generar_excel_listado_Error

    Me.MousePointer = 11
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Add
    Set XLS = XLW.Worksheets(1)
    XLW.Worksheets(3).Delete
    XLW.Worksheets(2).Delete
    XLW.Worksheets(1).Name = "Muestras pendientes de facturar"
    XLA.visible = False
    XLS.Range("1:1").HorizontalAlignment = xlCenter
    XLS.Range("1:1").VerticalAlignment = xlCenter
    XLS.Range("1:1").RowHeight = 30
    XLS.Range("1:1").WrapText = True
    'Cabeceras Excel
    Dim i As Integer
    Dim j As Integer
    Dim Col As Integer
    Col = 1
    For i = 1 To lista.ColumnHeaders.Count
        If lista.ColumnHeaders(i).Width > 10 Then
            XLS.Cells(1, Col) = lista.ColumnHeaders(i).Text
            Col = Col + 1
        End If
    Next
    'Detalle
    Dim fila As Integer
    fila = 2
    For i = 1 To lista.ListItems.Count
        Col = 1
        XLS.Range(XLS.Cells(fila, 6), XLS.Cells(fila, 6)) = "0.00"
        For j = 1 To lista.ColumnHeaders.Count
            If lista.ColumnHeaders(j).Width > 1 Then
                If j = 1 Then
                    XLS.Cells(fila, Col) = lista.ListItems(i).Text
                Else
                    If Col = 5 Then ' Fecha
                    XLS.Cells(fila, Col) = Format(lista.ListItems(i).SubItems(j - 1), "yyyy-mm-dd")
                    Else
                    XLS.Cells(fila, Col) = lista.ListItems(i).SubItems(j - 1)
                    End If
                End If
                Col = Col + 1
            End If
        Next
        fila = fila + 1
        ' Detalle determinaciones
        Dim rs As ADODB.Recordset
        Dim oDeter As New clsDeterminaciones
        Set rs = oDeter.lista_determinaciones(lista.ListItems(i).SubItems(COLS.ID_MUESTRA))
        If rs.RecordCount > 0 Then
            Do
                XLS.Cells(fila, 3) = rs("nombre") ' DETERMINACION
                XLS.Cells(fila, 4) = rs("pnt") ' PNT
                rs.MoveNext
                fila = fila + 1
            Loop Until rs.EOF
        End If
        
    Next
    Me.MousePointer = 0
    XLA.visible = True
    MsgBox "Listado generado correctamente.", vbInformation, App.Title

   On Error GoTo 0
   Exit Sub

generar_excel_listado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_excel_listado of Formulario frmMuestraPendientesFacturacion2"
End Sub


