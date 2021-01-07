VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmProveedores_Facturas_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Facturas de Proveedores"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17850
   Icon            =   "frmProveedores_Facturas_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   17850
   Begin VB.Frame frmLeyenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leyenda"
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
      Height          =   780
      Left            =   5850
      TabIndex        =   64
      Top             =   9225
      Width           =   6540
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2610
         Picture         =   "frmProveedores_Facturas_Listado.frx":08CA
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "APROBADA"
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
         Index           =   7
         Left            =   3105
         TabIndex        =   67
         Top             =   360
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   4410
         Picture         =   "frmProveedores_Facturas_Listado.frx":0CE3
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "RECHAZADA"
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
         Index           =   6
         Left            =   4905
         TabIndex        =   66
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PDTE.REVISIÓN"
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
         Index           =   4
         Left            =   675
         TabIndex        =   65
         Top             =   360
         Width           =   1545
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   180
         Picture         =   "frmProveedores_Facturas_Listado.frx":110B
         Top             =   225
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   5265
      TabIndex        =   45
      Top             =   4905
      Visible         =   0   'False
      Width           =   6450
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Generando documento EXCEL. Por favor, espere."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   1
         Left            =   585
         TabIndex        =   46
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pdtes. Pago"
      Height          =   870
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9270
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   2355
      Left            =   45
      TabIndex        =   28
      Top             =   585
      Width           =   17745
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TODAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   13680
         TabIndex        =   63
         Top             =   2070
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "APROBADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   10215
         TabIndex        =   62
         Top             =   2070
         Width           =   1725
      End
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RECHAZADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   11925
         TabIndex        =   61
         Top             =   2070
         Width           =   1860
      End
      Begin VB.OptionButton opSituacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PDTE.REVISION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   8235
         TabIndex        =   60
         Top             =   2070
         Width           =   2265
      End
      Begin VB.TextBox txtCodigoEquipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   13860
         TabIndex        =   19
         Top             =   1665
         Width           =   1770
      End
      Begin VB.CheckBox chkIntra 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intracomunitarios"
         Height          =   285
         Left            =   15660
         TabIndex        =   57
         Top             =   225
         Width           =   1635
      End
      Begin VB.TextBox txtCC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10980
         TabIndex        =   18
         Top             =   1665
         Width           =   1770
      End
      Begin VB.CheckBox chkFCobro 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   45
         TabIndex        =   55
         Top             =   1350
         Width           =   240
      End
      Begin VB.CheckBox chkFVenci 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   45
         TabIndex        =   54
         Top             =   990
         Width           =   240
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   960
         Left            =   16245
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   585
         Width           =   1140
      End
      Begin VB.TextBox txtimportehasta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3030
         TabIndex        =   8
         Top             =   1665
         Width           =   1320
      End
      Begin VB.TextBox txtImporteDesde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Top             =   1665
         Width           =   1320
      End
      Begin VB.CheckBox chkVencidas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar las vencidas"
         Height          =   240
         Left            =   4500
         TabIndex        =   12
         Top             =   1440
         Width           =   1905
      End
      Begin VB.CheckBox chkPagoPrevisto 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Pago Previsto"
         Height          =   240
         Left            =   4500
         TabIndex        =   13
         Top             =   1755
         Width           =   1995
      End
      Begin VB.CheckBox chkIncidencias 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar INCIDENCIAS"
         Height          =   240
         Left            =   4500
         TabIndex        =   11
         Top             =   1125
         Width           =   2040
      End
      Begin VB.TextBox txtconcepto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8235
         TabIndex        =   17
         Top             =   1665
         Width           =   1680
      End
      Begin VB.CheckBox chkNoEnviadas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo no enviadas"
         Height          =   240
         Left            =   4500
         TabIndex        =   10
         Top             =   855
         Width           =   2220
      End
      Begin VB.CheckBox chkPendientesPago 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo pdtes. pago"
         Height          =   285
         Left            =   4500
         TabIndex        =   9
         Top             =   540
         Width           =   2130
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   1050
         TabIndex        =   0
         Top             =   225
         Width           =   14520
         _ExtentX        =   25612
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbFamilia 
         Height          =   345
         Left            =   8235
         TabIndex        =   14
         Top             =   585
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbGasto 
         Height          =   345
         Left            =   8235
         TabIndex        =   15
         Top             =   945
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPago 
         Height          =   345
         Left            =   8235
         TabIndex        =   16
         Top             =   1305
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1050
         TabIndex        =   1
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   74907649
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3030
         TabIndex        =   2
         Top             =   585
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   74907649
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fVencimientoDesde 
         Height          =   330
         Left            =   1050
         TabIndex        =   3
         Top             =   945
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   74907649
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fVencimientoHasta 
         Height          =   330
         Left            =   3030
         TabIndex        =   4
         Top             =   945
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   74907649
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fCobroDesde 
         Height          =   330
         Left            =   1050
         TabIndex        =   5
         Top             =   1305
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   74907649
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fCobroHasta 
         Height          =   330
         Left            =   3030
         TabIndex        =   6
         Top             =   1305
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   74907649
         CurrentDate     =   38002
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisión"
         Height          =   195
         Index           =   6
         Left            =   6840
         TabIndex        =   59
         Top             =   2070
         Width           =   750
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod.Equipo"
         Height          =   195
         Index           =   5
         Left            =   12960
         TabIndex        =   58
         Top             =   1710
         Width           =   840
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.Contable"
         Height          =   195
         Index           =   4
         Left            =   10080
         TabIndex        =   56
         Top             =   1710
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   53
         Top             =   1395
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Pago"
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   52
         Top             =   1350
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   51
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Vencim."
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   50
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   3
         Left            =   2655
         TabIndex        =   49
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   48
         Top             =   1710
         Width           =   750
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   47
         Top             =   1710
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Factura"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   44
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   43
         Top             =   675
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcuenta Pago"
         Height          =   195
         Index           =   2
         Left            =   6840
         TabIndex        =   34
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcuenta Gasto"
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   33
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   32
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   31
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9270
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1130
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9270
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9270
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   16740
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9225
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6090
      Left            =   45
      TabIndex        =   27
      Top             =   2970
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   10742
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13185
      Top             =   9090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Facturas_Listado.frx":1508
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Facturas_Listado.frx":199F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Facturas_Listado.frx":1E35
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9270
      Width           =   1050
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Index           =   2
      Left            =   12780
      TabIndex        =   42
      Top             =   9945
      Width           =   825
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13995
      TabIndex        =   41
      Top             =   9945
      Width           =   2085
   End
   Begin VB.Label lblIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13995
      TabIndex        =   40
      Top             =   9405
      Width           =   2085
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Height          =   285
      Index           =   0
      Left            =   12780
      TabIndex        =   39
      Top             =   9405
      Width           =   645
   End
   Begin VB.Label lblBase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13995
      TabIndex        =   38
      Top             =   9135
      Width           =   2085
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Index           =   1
      Left            =   12780
      TabIndex        =   37
      Top             =   9135
      Width           =   870
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retención"
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
      Height          =   285
      Index           =   3
      Left            =   12780
      TabIndex        =   36
      Top             =   9675
      Width           =   1275
   End
   Begin VB.Label lblRetencion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13995
      TabIndex        =   35
      Top             =   9675
      Width           =   2085
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Proveedores"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   45
      TabIndex        =   30
      Top             =   315
      Width           =   1680
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Facturas de Proveedores"
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
      Left            =   45
      TabIndex        =   29
      Top             =   45
      Width           =   3810
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   17820
   End
End
Attribute VB_Name = "frmProveedores_Facturas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Enum COLS
    C_ID = 0
    C_PROVEEDOR = 1
    C_CCONTABLE = 2
    C_fecha = 3
    C_concepto = 4
    C_NUMERO = 5
    C_FAMILIA = 6
    C_SUBCUENTA = 7
    C_BASE = 8
    C_IVA_PORCENTAJE = 9
    C_IVA = 10
    C_RETENCION = 11
    C_total = 12
    C_FP = 13
    C_vencimiento = 14
    C_PAGO = 15
    C_TOBJETO = 16
    C_cOBJETO = 17
    C_IDPROVEEDOR = 18
'M1335-I
    C_CUENTA = 19
'M1335-F
    C_ENVIADA = 20
    C_REVISADA = 21
    C_CIF = 22
End Enum

Private Sub chkFCobro_Click()
    If chkFCobro.Value = Checked Then
        fCobroDesde.Enabled = True
        fCobroHasta.Enabled = True
    Else
        fCobroDesde.Enabled = False
        fCobroHasta.Enabled = False
    End If
End Sub

Private Sub chkFVenci_Click()
    If chkFVenci.Value = Checked Then
        fVencimientoDesde.Enabled = True
        fVencimientoHasta.Enabled = True
    Else
        fVencimientoDesde.Enabled = False
        fVencimientoHasta.Enabled = False
    End If
End Sub

Private Sub chkIncidencias_Click()
    cargar_lista
End Sub

Private Sub chkIntra_Click()
    If chkIntra.Value = Checked Then
        cmbProveedor.desactivar
    Else
        cmbProveedor.activar
    End If
    cargar_lista
End Sub

Private Sub chkNoEnviadas_Click()
    cargar_lista
End Sub

Private Sub chkPagoPrevisto_Click()
    cargar_lista
End Sub

Private Sub chkPendientesPago_Click()
    cargar_lista
End Sub

Private Sub chkVencidas_Click()
    cargar_lista
End Sub

Private Sub cmbfamilia_Change()
    cargar_lista
End Sub

Private Sub cmbGasto_change()
    cargar_lista
End Sub

Private Sub cmbPago_change()
    cargar_lista
End Sub

Private Sub cmbProveedor_change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    With frmProveedores_Facturas
        .PK = 0
        .PK_FACTURA_ID = 0
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR la factura de proveedor " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oPF As New clsProveedores_Facturas
        If oPF.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            MsgBox "Factura eliminada correctamente.", vbInformation, App.Title
            cargar_lista
        End If
        Set oPF = Nothing
    End If
End Sub

Private Sub cmdImprimir_Click()
       Me.MousePointer = vbHourglass
       Frame3.visible = True
       Dim rs As New ADODB.Recordset
       Dim fecha As String
      
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable    'NUMERO
       rs.Fields.Append "c2", adChar, 150, adFldUpdatable   'Proveedor
       rs.Fields.Append "c3", adChar, 20, adFldUpdatable   'CC
       rs.Fields.Append "c4", adChar, 20, adFldUpdatable    'Fecha
       rs.Fields.Append "c5", adChar, 50, adFldUpdatable    'Concepto
       rs.Fields.Append "c6", adChar, 100, adFldUpdatable   'Numero
       rs.Fields.Append "c7", adChar, 20, adFldUpdatable   'Base
       rs.Fields.Append "c8", adChar, 20, adFldUpdatable   'IVA
       rs.Fields.Append "c9", adChar, 20, adFldUpdatable   'Retención
       rs.Fields.Append "c10", adChar, 20, adFldUpdatable    'Total
       rs.Fields.Append "c11", adChar, 100, adFldUpdatable   'FormaPago
       rs.Fields.Append "c12", adChar, 20, adFldUpdatable    'F.Vencimiento
       rs.Fields.Append "c13", adChar, 100, adFldUpdatable   'F.Pago
       rs.Fields.Append "c14", adChar, 35, adFldUpdatable    'Cuenta Bancaria
       rs.Fields.Append "c15", adChar, 35, adFldUpdatable    'CIF
       rs.Fields.Append "c16", adChar, 250, adFldUpdatable    'FAMILIA
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
'            If lista.ListItems(i).SubItems(C_PAGO) <> "" Then
                rs.AddNew
                rs("c1") = lista.ListItems(i).Text
                rs("c2") = lista.ListItems(i).SubItems(C_PROVEEDOR)
                rs("c3") = lista.ListItems(i).SubItems(C_CCONTABLE)
                rs("c4") = lista.ListItems(i).SubItems(C_fecha)
                rs("c5") = lista.ListItems(i).SubItems(C_concepto)
                rs("c6") = lista.ListItems(i).SubItems(C_NUMERO)
                rs("c7") = lista.ListItems(i).SubItems(C_BASE)
                rs("c8") = lista.ListItems(i).SubItems(C_IVA)
                rs("c9") = lista.ListItems(i).SubItems(C_RETENCION)
                rs("c10") = lista.ListItems(i).SubItems(C_total)
                rs("c11") = lista.ListItems(i).SubItems(C_FP)
                rs("c12") = lista.ListItems(i).SubItems(C_vencimiento)
                rs("c13") = lista.ListItems(i).SubItems(C_PAGO)
                rs("c14") = lista.ListItems(i).SubItems(C_CUENTA)
                rs("c15") = lista.ListItems(i).SubItems(C_CIF)
                rs("c16") = lista.ListItems(i).SubItems(C_FAMILIA)
                rs.Update
'            End If
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Listado de Facturas"
 
        'Cabecera
        With XLS.Range("A1:P1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:P1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:P1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 12
        XLS.Range("B1:B1").ColumnWidth = 55
        XLS.Range("C1:C1").ColumnWidth = 20 ' CIF
        XLS.Range("D1:D1").ColumnWidth = 20
        XLS.Range("E1:E1").ColumnWidth = 12
        XLS.Range("F1:F1").ColumnWidth = 25
        XLS.Range("G1:G1").ColumnWidth = 25
        XLS.Range("H1:H1").ColumnWidth = 12
        XLS.Range("I1:I1").ColumnWidth = 12
        XLS.Range("J1:J1").ColumnWidth = 12
        XLS.Range("K1:K1").ColumnWidth = 12
        XLS.Range("L1:L1").ColumnWidth = 20
        XLS.Range("M1:M1").ColumnWidth = 15
        XLS.Range("N1:N1").ColumnWidth = 15
        XLS.Range("O1:O1").ColumnWidth = 30
        XLS.Range("P1:P1").ColumnWidth = 30 ' FAMILIA

        XLS.Cells(1, 1) = "NºAsiento"
        XLS.Cells(1, 2) = "Proveedor"
        XLS.Cells(1, 3) = "CIF"
        XLS.Cells(1, 4) = "C.Contable"
        XLS.Cells(1, 5) = "Fecha"
        XLS.Cells(1, 6) = "Concepto"
        XLS.Cells(1, 7) = "Número"
        XLS.Cells(1, 8) = "Base"
        XLS.Cells(1, 9) = "IVA"
        XLS.Cells(1, 10) = "Retención"
        XLS.Cells(1, 11) = "Total"
        XLS.Cells(1, 12) = "Forma Pago"
        XLS.Cells(1, 13) = "F.Vencimiento"
        XLS.Cells(1, 14) = "F.Pago"
        XLS.Cells(1, 15) = "Cuenta Bancaria"
        XLS.Cells(1, 16) = "Familia"
        
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = CLng(rs("c1")) ' Asiento
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True) ' Proveedor
            XLS.Cells(i, 3) = rs("c15") ' CIF
            XLS.Cells(i, 4) = rs("c3") ' CC
            XLS.Cells(i, 5) = CDate(Trim(rs("c4"))) ' Fecha
            XLS.Cells(i, 6) = CStr(ClrStr(rs("c5"), False, True, True)) ' Concepto
            XLS.Cells(i, 7) = CStr(ClrStr(rs("c6"), False, True, True)) ' Numero
            XLS.Cells(i, 8) = CDbl(rs("c7")) 'Base
            XLS.Cells(i, 9) = CDbl(rs("c8")) ' Iva
            XLS.Cells(i, 10) = CDbl(rs("c9")) ' Retencion
            XLS.Cells(i, 11) = CDbl(rs("c10")) ' Total
            XLS.Cells(i, 12) = rs("C11") ' Forma Pago
            XLS.Cells(i, 13) = CDate(Trim(rs("C12")))      ' F.Vencimiento
            If Trim(rs("c13")) <> "" Then
                XLS.Cells(i, 14) = CDate(Trim(rs("c13"))) ' F.Pago
            End If
            XLS.Cells(i, 15) = rs("c14") ' Cuenta bancaria
            XLS.Cells(i, 16) = rs("c16") ' Familia
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame3.visible = False
        Me.MousePointer = vbNormal
        XLA.visible = True
        Set rs = Nothing

End Sub

Private Sub cmdListado_Click()
'Listado de facturas pendientes de cobro
       Me.MousePointer = vbHourglass
       Frame3.visible = True
       Dim rs As New ADODB.Recordset
       Dim fecha As String
      
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable    'ID
       rs.Fields.Append "c2", adChar, 150, adFldUpdatable   'Proveedor
       rs.Fields.Append "c3", adChar, 20, adFldUpdatable   'CC
       rs.Fields.Append "c4", adChar, 35, adFldUpdatable    'Cuenta Bancaria
       rs.Fields.Append "c5", adChar, 20, adFldUpdatable    'Total
       rs.Fields.Append "c6", adChar, 20, adFldUpdatable    'Fecha
       rs.Fields.Append "c7", adChar, 50, adFldUpdatable   'Concepto
       rs.Fields.Append "c8", adChar, 100, adFldUpdatable    'Numero
       rs.Fields.Append "c9", adChar, 100, adFldUpdatable    'Familia
       rs.Fields.Append "c10", adChar, 150, adFldUpdatable    'Subcuenta
       rs.Fields.Append "c11", adChar, 20, adFldUpdatable   'Base
       rs.Fields.Append "c12", adChar, 10, adFldUpdatable   'IVA Porcentaje
       rs.Fields.Append "c13", adChar, 20, adFldUpdatable   'IVA
       rs.Fields.Append "c14", adChar, 20, adFldUpdatable   'CIF
       rs.Fields.Append "c15", adChar, 150, adFldUpdatable   'Familia
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).SubItems(C_PAGO) <> "" Then
                rs.AddNew
                rs("c1") = lista.ListItems(i).Text
                rs("c2") = lista.ListItems(i).SubItems(C_PROVEEDOR)
                rs("c3") = lista.ListItems(i).SubItems(C_CCONTABLE)
                rs("c4") = lista.ListItems(i).SubItems(C_CUENTA)
                rs("c5") = lista.ListItems(i).SubItems(C_total)
                rs("c6") = lista.ListItems(i).SubItems(C_fecha)
                rs("c7") = lista.ListItems(i).SubItems(C_concepto)
                rs("c8") = lista.ListItems(i).SubItems(C_NUMERO)
                rs("c9") = lista.ListItems(i).SubItems(C_FAMILIA)
                rs("c10") = lista.ListItems(i).SubItems(C_SUBCUENTA)
                rs("c11") = lista.ListItems(i).SubItems(C_BASE)
                rs("c12") = lista.ListItems(i).SubItems(C_IVA_PORCENTAJE)
                rs("c13") = lista.ListItems(i).SubItems(C_IVA)
                rs("c14") = lista.ListItems(i).SubItems(C_CIF)
                rs("c15") = lista.ListItems(i).SubItems(C_FAMILIA)
                rs.Update
            End If
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Facturas pendientes de pago"
 
        'Cabecera
        With XLS.Range("A1:N1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:N1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:N1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 12
        XLS.Range("B1:B1").ColumnWidth = 55
        XLS.Range("C1:C1").ColumnWidth = 20
        XLS.Range("D1:D1").ColumnWidth = 20
        XLS.Range("E1:E1").ColumnWidth = 55
        XLS.Range("F1:F1").ColumnWidth = 20
        XLS.Range("G1:G1").ColumnWidth = 12
        XLS.Range("H1:H1").ColumnWidth = 12
        XLS.Range("I1:I1").ColumnWidth = 12
        XLS.Range("J1:J1").ColumnWidth = 10
        XLS.Range("K1:K1").ColumnWidth = 5
        XLS.Range("L1:L1").ColumnWidth = 30
        XLS.Range("M1:M1").ColumnWidth = 30
        XLS.Range("N1:N1").ColumnWidth = 30
        XLS.Range("O1:O1").ColumnWidth = 30

        XLS.Cells(1, 1) = "ID"
        XLS.Cells(1, 2) = "Proveedor"
        XLS.Cells(1, 3) = "CIF"
        XLS.Cells(1, 4) = "C.Contable"
        XLS.Cells(1, 5) = "Cuenta Bancaria"
        XLS.Cells(1, 6) = "Total a pagar"
        XLS.Cells(1, 7) = "Fecha"
        XLS.Cells(1, 8) = "Concepto"
        XLS.Cells(1, 9) = "Número"
        XLS.Cells(1, 10) = "Familia"
        XLS.Cells(1, 11) = "Subcuenta"
        XLS.Cells(1, 12) = "Base"
        XLS.Cells(1, 13) = "IVA %"
        XLS.Cells(1, 14) = "IVA"
        XLS.Cells(1, 15) = "Familia"
        
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = CLng(rs("c1"))
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = ClrStr(rs("c14"), False, True, True) ' CIF
            XLS.Cells(i, 4) = ClrStr(rs("c3"), False, True, True)
            XLS.Cells(i, 5) = rs("c4")
            XLS.Cells(i, 6) = CDbl(rs("c5"))
            XLS.Cells(i, 7) = Trim(rs("c6"))
            XLS.Cells(i, 8) = rs("c7")
            XLS.Cells(i, 9) = rs("c8")
            XLS.Cells(i, 10) = rs("C9")
            XLS.Cells(i, 11) = rs("C10")
            XLS.Cells(i, 12) = CDbl(rs("c11"))
            XLS.Cells(i, 13) = rs("c12")
            XLS.Cells(i, 14) = CDbl(rs("c13"))
            XLS.Cells(i, 15) = rs("c15")
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame3.visible = False
        Me.MousePointer = vbNormal
        XLA.visible = True
        Set rs = Nothing
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmProveedores_Facturas
        .PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_IDPROVEEDOR)
        .PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
'    actualizar_lista
End Sub
Private Sub fCobroDesde_Change()
    cargar_lista
End Sub
Private Sub fCobroHasta_Change()
    cargar_lista
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
'    fdesde = "01/01/" & Year(Date)
'    fhasta = "31/12/" & Year(Date)
    
    fdesde = Date - 180
    fVencimientoDesde = "01/" & Month(Date) & "/" & Year(Date)
    fCobroDesde = "01/" & Month(Date) & "/" & Year(Date)
    
    fhasta = Date
    fVencimientoHasta = Date
    fCobroHasta = Date
    
    cargarCombos
    permisos
    cabecera
    cargar_lista
End Sub
Private Sub cargarCombos()
    cargarProveedores
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbGasto, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS
    oDeco.cargar_mi_combo cmbPago, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS
    llenar_combo cmbFamilia, New clsFamilias, 0, Me, ""
    Set oDeco = Nothing
End Sub
Private Sub cargarProveedores()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT P.ID_PROVEEDOR,P.NOMBRE " & _
                   "  FROM PROVEEDORES AS P, PROVEEDORES_FACTURAS AS PF " & _
                   " WHERE P.ID_PROVEEDOR = PF.PROVEEDOR_ID "
        With cmbProveedor
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "PROVEEDORES"
            .setDESCRIPCION = "Proveedores"
            .setPK = "ID_PROVEEDOR"
            .setCAMPO = "NOMBRE"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmProveedores_Detalle
        End With
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº", 1100, lvwColumnLeft
        .Add , , "Proveedor", 2000, lvwColumnLeft
        .Add , , "C.Contable", 1150, lvwColumnCenter
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Concepto", 1700, lvwColumnCenter
        .Add , , "Numero", 1000, lvwColumnCenter
        .Add , , "Familia", 1400, lvwColumnLeft
        .Add , , "Subcuenta", 1, lvwColumnLeft
        .Add , , "Base", 1050, lvwColumnRight
        .Add , , "Iva %", 1, lvwColumnCenter
        .Add , , "Iva", 1000, lvwColumnRight
        .Add , , "Retención", 1000, lvwColumnRight
        .Add , , "Total", 1050, lvwColumnRight
        .Add , , "Forma Pago", 1100, lvwColumnCenter
        .Add , , "Fecha Vencimiento", 1050, lvwColumnCenter
        .Add , , "Fecha Pago", 1050, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "ID_PROVEEDOR", 1, lvwColumnLeft
        'M1335-I
        .Add , , "CUENTA_BANCARIA", 1, lvwColumnLeft
        'M1335-F
        .Add , , "Env", 350, lvwColumnLeft
        .Add , , "REV.", 350, lvwColumnCenter
        .Add , , "CIF", 1, lvwColumnCenter
    End With
End Sub
Private Sub permisos()
    If USUARIO.getPER_TESORERIA_FP = False Then
    End If
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPF As New clsProveedores_Facturas
    Dim ID As Long
   On Error GoTo cargar_lista_Error

    ID = 0
    If cmbProveedor.getTEXTO <> "" Then
        ID = cmbProveedor.getPK_SALIDA
    End If
    Dim familiaid As Long
    Dim subcuentagasto As Long
    Dim subcuentapago As Long
    If cmbFamilia.getTEXTO <> "" Then
        familiaid = cmbFamilia.getPK_SALIDA
    End If
    If cmbGasto.getTEXTO <> "" Then
        subcuentagasto = cmbGasto.getPK_SALIDA
    End If
    If cmbPago.getTEXTO <> "" Then
        subcuentapago = cmbPago.getPK_SALIDA
    End If
    'REVISION
    Dim revision As Integer
    revision = 3
    If opSituacion(0).Value = True Then
        revision = 0
    ElseIf opSituacion(1).Value = True Then
        revision = 1
    ElseIf opSituacion(2).Value = True Then
        revision = 2
    End If
    Me.MousePointer = 11
    Set rs = oPF.ListadoCompleto(ID, chkPendientesPago.Value, fdesde, fhasta, familiaid, subcuentagasto, subcuentapago, chkNoEnviadas.Value, txtConcepto, chkVencidas.Value, chkPagoPrevisto.Value, chkIncidencias.Value, txtImporteDesde, txtimportehasta, chkFVenci.Value, fVencimientoDesde, fVencimientoHasta, chkFCobro.Value, fCobroDesde, fCobroHasta, txtCC, chkIntra.Value, txtCodigoEquipo, revision)
    Dim BASE As Currency
    Dim IVA As Currency
    Dim retencion As Currency
    Dim total As Currency
    BASE = 0
    IVA = 0
    retencion = 0
    total = 0
    lista.ListItems.Clear
    lblsubtitulo = "Se han detectado " & rs.RecordCount & " registros."
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000000")) ' ID
           .SubItems(COLS.C_PROVEEDOR) = rs(17)
           .SubItems(COLS.C_IDPROVEEDOR) = rs(18)
           .SubItems(COLS.C_CCONTABLE) = rs(21) 'CC
            .SubItems(COLS.C_fecha) = Format(rs(1), "dd/mm/yyyy")  ' Fecha
            If Not IsNull(rs(2)) Then
                .SubItems(COLS.C_concepto) = rs(2)  ' Concepto
            End If
            If Not IsNull(rs(3)) Then
                .SubItems(COLS.C_NUMERO) = rs(3)  ' Numero
            End If
            If Not IsNull(rs(4)) Then
                .SubItems(COLS.C_FAMILIA) = rs(4)  ' Familia
            End If
            If Not IsNull(rs(5)) Then
                .SubItems(COLS.C_SUBCUENTA) = rs(5)  ' Subcuenta
            End If
            .SubItems(COLS.C_BASE) = Format(rs(6), "currency")  ' BI
            .SubItems(COLS.C_IVA_PORCENTAJE) = rs(7)  ' IVA PORCENTAJE
            .SubItems(COLS.C_IVA) = Format(rs(8), "currency")  ' IVA
            .SubItems(COLS.C_total) = Format(rs(9), "currency")  ' TOTAL
            BASE = BASE + rs(6)
            IVA = IVA + rs(8)
            retencion = retencion + rs(16)
            total = total + rs(9)
            If Not IsNull(rs(10)) Then
                .SubItems(COLS.C_FP) = rs(10)  ' FP
            End If
            If Not IsNull(rs(11)) Then
                .SubItems(COLS.C_vencimiento) = rs(11)  ' F.Vencimiento
            End If
            If Not IsNull(rs(13)) Then
                .SubItems(COLS.C_TOBJETO) = rs(13)  ' Tobjeto
            End If
            If Not IsNull(rs(14)) Then
                .SubItems(COLS.C_cOBJETO) = rs(14)  ' Cobjeto
            End If
            If Not IsNull(rs(15)) Then
                .Checked = True
            Else
                .Checked = False
            End If
            If Not IsNull(rs(16)) Then
                .SubItems(COLS.C_RETENCION) = Format(rs(16), "currency") ' RETENCION
            End If
            If Not IsNull(rs(12)) Then
                .SubItems(COLS.C_PAGO) = rs(12)
            End If
            'M1335-I
            If Not IsNull(rs(19)) Then
                .SubItems(COLS.C_CUENTA) = rs(19)
            End If
            'M1335-F
            If rs(20) = 0 Then
                .SubItems(COLS.C_ENVIADA) = "N"
            Else
                .SubItems(COLS.C_ENVIADA) = "S"
            End If
            ' REVISION 21: REVISADA_POR, 22: SITUACION
            If rs(22) = 0 Then ' Si no hay revisor, no bola
                .ListSubItems.Add , , "", vbNothing
            Else
                If rs(23) = 0 Then 'ENVIADA
                    .ListSubItems.Add , , "", 2
                ElseIf rs(23) = 1 Then 'aprobada
                    .ListSubItems.Add , , "", 1
                ElseIf rs(23) = 2 Then 'Rechazada
                    .ListSubItems.Add , , "", 3
                End If
            End If
            .SubItems(COLS.C_CIF) = rs(24) ' CIF
            
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lblBase = Format(BASE, "currency")
    lblIVA = Format(IVA, "currency")
    lblRetencion = Format(retencion, "currency")
    lbltotal = Format(total, "currency")
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmProveedores_Facturas_Listado"

End Sub

Private Sub fVencimientoDesde_Change()
    cargar_lista
End Sub

Private Sub fVencimientoHasta_Change()
    cargar_lista
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
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtCC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cargar_lista
    End If

End Sub
Private Sub txtCodigoEquipo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cargar_lista
    End If
End Sub

Private Sub txtconcepto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cargar_lista
    End If
End Sub

Private Sub txtImporteDesde_GotFocus()
    txtImporteDesde.SelStart = 0
    txtImporteDesde.SelLength = Len(txtImporteDesde)
End Sub

Private Sub txtImporteDesde_LostFocus()
    If txtImporteDesde <> "" Then
        txtImporteDesde = moneda(txtImporteDesde)
        txtimportehasta = moneda_bd(txtImporteDesde)
    End If
End Sub

Private Sub txtimportehasta_GotFocus()
    txtimportehasta.SelStart = 0
    txtimportehasta.SelLength = Len(txtimportehasta)
End Sub

Private Sub txtimportehasta_LostFocus()
    If txtimportehasta <> "" Then
        txtimportehasta = moneda(txtimportehasta)
    End If
End Sub
