VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmFacturacion_henkel 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modulo de Facturación de probetas HENKEL"
   ClientHeight    =   10755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17910
   Icon            =   "frmFacturacion_henkel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10755
   ScaleWidth      =   17910
   Begin VB.CommandButton cmdFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Facturar"
      Height          =   885
      Index           =   2
      Left            =   15480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9765
      UseMaskColor    =   -1  'True
      Width           =   1170
   End
   Begin VB.Frame frmDatosFactura 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3675
      Left            =   4455
      TabIndex        =   57
      Top             =   3690
      Visible         =   0   'False
      Width           =   9135
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   885
         Index           =   1
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   2655
         Width           =   1035
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   2745
         TabIndex        =   62
         Top             =   2385
         Width           =   3840
         Begin VB.CommandButton cmdAlbaran 
            BackColor       =   &H00E0E0E0&
            Caption         =   "ALBARAN"
            Height          =   885
            Left            =   1350
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   1170
         End
         Begin VB.CommandButton cmdFactura 
            BackColor       =   &H00E0E0E0&
            Caption         =   "FACTURA"
            Height          =   885
            Index           =   0
            Left            =   135
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   1170
         End
         Begin VB.CommandButton cmdFactura 
            BackColor       =   &H00E0E0E0&
            Caption         =   "PROFORMA"
            Height          =   885
            Index           =   1
            Left            =   2565
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   225
            UseMaskColor    =   -1  'True
            Width           =   1170
         End
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   330
         Left            =   1125
         TabIndex        =   58
         Top             =   1530
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClienteFactura 
         Height          =   330
         Left            =   1125
         TabIndex        =   59
         Top             =   1170
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCliente 
         Height          =   330
         Left            =   1125
         TabIndex        =   69
         Top             =   810
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaFactura 
         Height          =   330
         Left            =   1125
         TabIndex        =   71
         Top             =   405
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   72
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   70
         Top             =   855
         Width           =   690
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "DATOS PARA LA FACTURA"
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
         Index           =   3
         Left            =   0
         TabIndex        =   67
         Top             =   0
         Width           =   9525
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   135
         TabIndex        =   61
         Top             =   1575
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cli. Factura"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   60
         Top             =   1215
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      Height          =   1860
      Left            =   8055
      TabIndex        =   31
      Top             =   7650
      Width           =   9645
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   18
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1125
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   17
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   765
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   16
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1485
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   15
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   405
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   14
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   13
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1080
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   12
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   11
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   10
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1485
         Width           =   1230
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   9
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   405
         Width           =   1230
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   8
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   765
         Width           =   1230
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Index           =   7
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1125
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº PROBETAS DTO"
         Height          =   240
         Index           =   18
         Left            =   3330
         TabIndex        =   56
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRECIO MEDIO"
         Height          =   240
         Index           =   17
         Left            =   3330
         TabIndex        =   54
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "IMPORTE DTO"
         Height          =   240
         Index           =   16
         Left            =   3330
         TabIndex        =   53
         Top             =   1530
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRIMER"
         Height          =   240
         Index           =   15
         Left            =   3330
         TabIndex        =   50
         Top             =   450
         Width           =   1500
      End
      Begin VB.Line Line3 
         X1              =   3150
         X2              =   3150
         Y1              =   315
         Y2              =   1800
      End
      Begin VB.Line Line1 
         X1              =   6210
         X2              =   6210
         Y1              =   315
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO PRIMER"
         Height          =   240
         Index           =   14
         Left            =   180
         TabIndex        =   48
         Top             =   405
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº PROBETAS DTO"
         Height          =   240
         Index           =   13
         Left            =   180
         TabIndex        =   46
         Top             =   1125
         Width           =   1500
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "IMPORTES"
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
         Index           =   2
         Left            =   45
         TabIndex        =   44
         Top             =   0
         Width           =   9525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRECIO MEDIO"
         Height          =   240
         Index           =   12
         Left            =   180
         TabIndex        =   43
         Top             =   765
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "IMPORTE DTO"
         Height          =   240
         Index           =   11
         Left            =   180
         TabIndex        =   42
         Top             =   1485
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL BASE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   6930
         TabIndex        =   41
         Top             =   1530
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "IMPORTE PROBETAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   5850
         TabIndex        =   40
         Top             =   810
         Width           =   2400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.F.M."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   6885
         TabIndex        =   39
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "DTO APLICABLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   6390
         TabIndex        =   38
         Top             =   1170
         Width           =   1860
      End
   End
   Begin VB.CommandButton cmdPrecios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tabla de Precios"
      Height          =   885
      Left            =   90
      Picture         =   "frmFacturacion_henkel.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9720
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      TabIndex        =   23
      Top             =   495
      Width           =   17820
      Begin VB.CheckBox chkCerradas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir muestras abiertas"
         Height          =   255
         Left            =   5580
         TabIndex        =   25
         Top             =   225
         Width           =   2130
      End
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Desde"
         Height          =   285
         Left            =   135
         TabIndex        =   24
         Top             =   225
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1530
         TabIndex        =   26
         Top             =   180
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3555
         TabIndex        =   27
         Top             =   180
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   390
         Left            =   15525
         TabIndex        =   29
         Top             =   180
         Width           =   2145
         _Version        =   851970
         _ExtentX        =   3784
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "BUSCAR"
         Appearance      =   5
         Picture         =   "frmFacturacion_henkel.frx":0BD4
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   3015
         TabIndex        =   28
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.Frame frmTotales 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      Height          =   1860
      Left            =   90
      TabIndex        =   5
      Top             =   7650
      Width           =   7800
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1125
         Width           =   1230
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   765
         Width           =   1230
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   405
         Width           =   1230
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1125
         Width           =   960
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   765
         Width           =   960
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   3825
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   405
         Width           =   960
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "TASA VOLUMEN"
         Height          =   240
         Index           =   6
         Left            =   4905
         TabIndex        =   22
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.F.M."
         Height          =   240
         Index           =   5
         Left            =   4905
         TabIndex        =   18
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "RANGO"
         Height          =   240
         Index           =   4
         Left            =   5040
         TabIndex        =   16
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO CONFORMES"
         Height          =   240
         Index           =   3
         Left            =   2430
         TabIndex        =   14
         Top             =   1170
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "CONFORMES"
         Height          =   240
         Index           =   2
         Left            =   2520
         TabIndex        =   12
         Top             =   810
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº PROBETAS"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOTALES"
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
         Index           =   0
         Left            =   45
         TabIndex        =   8
         Top             =   0
         Width           =   7680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº MUESTRAS"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   450
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   885
      Index           =   0
      Left            =   16695
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9765
      Width           =   1035
   End
   Begin MSComctlLib.ListView probetas 
      Height          =   6090
      Left            =   10665
      TabIndex        =   0
      Top             =   1485
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10742
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
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdmarcarmuestras 
      Height          =   300
      Index           =   1
      Left            =   2205
      TabIndex        =   19
      Top             =   1170
      Width           =   2115
      _Version        =   851970
      _ExtentX        =   3731
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Desmarcar Todas"
      Appearance      =   5
      Picture         =   "frmFacturacion_henkel.frx":7436
   End
   Begin XtremeSuiteControls.PushButton cmdmarcarmuestras 
      Height          =   300
      Index           =   0
      Left            =   45
      TabIndex        =   20
      Top             =   1170
      Width           =   2145
      _Version        =   851970
      _ExtentX        =   3784
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Marcar Todas"
      Appearance      =   5
      Picture         =   "frmFacturacion_henkel.frx":DC98
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   9090
      Top             =   9945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion_henkel.frx":144FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFacturacion_henkel.frx":1AD5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView muestras 
      Height          =   6090
      Left            =   45
      TabIndex        =   2
      Top             =   1485
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   10742
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
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1485
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modulo de Facturación de probetas HENKEL"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4650
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de muestras pendientes de Facturar"
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
      Index           =   1
      Left            =   45
      TabIndex        =   3
      Top             =   1170
      Width           =   17820
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   -135
      Top             =   0
      Width           =   18015
   End
End
Attribute VB_Name = "frmFacturacion_henkel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum COLS_MUESTRAS
    C_ID_MUESTRA = 0
    C_CODIGO = 1
    C_REFERENCIA_CLIENTE = 2
    C_fecha = 3
    C_GENERAL = 4
    C_Probetas = 5
    C_CONFORMES = 6
    C_NO_CONFORMES = 7
    C_IMPORTE = 8
    C_TASA = 9
    C_total = 10
    C_PROBETAS_E = 11
    C_PROBETAS_ABCD = 12
    C_TIPO_ANALISIS = 13
End Enum
Private Enum COLS_PROBETAS
    CP_IMPORTE = 12
    CP_TASA = 13
    CP_TOTAL = 14
    CP_DOC_ID = 15
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
Private Sub cmbClienteFactura_change()
    cmbPedidos.limpiar
    If cmbClienteFactura.getTEXTO <> "" Then
        pedidos cmbClienteFactura.getPK_SALIDA
    End If
End Sub
Private Sub pedidos(ID As Long)
    Dim filtro As String
    If ID <> 0 Then
        filtro = " AND CLIENTE_ID = " & ID & " AND FECHA_BAJA >= '" & Format(Date, "YYYY-MM-DD") & "'"
    End If
    llenar_combo cmbPedidos, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub

Private Sub cmdAlbaran_Click()
    facturar C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
End Sub

Private Sub cmdBuscar_Click()
    cargar_muestras
End Sub
Private Sub facturar(TIPO As Integer)
    Dim strcadena As String
    Dim i As Integer
    Dim marcadas As Integer
   On Error GoTo facturar_Error

    marcadas = contar_marcados
    If marcadas = 0 Then
        MsgBox "Debe seleccionar alguna muestra para Facturar.", vbInformation, App.Title
        Exit Sub
    End If
    Dim t As String
    Select Case TIPO
    Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
        t = "FACTURA"
    Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
        t = "ALBARAN"
    Case C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA
        t = "FACTURA PROFORMA"
    End Select
    If marcadas = 1 Then
        strcadena = "Va a generar " & t & " para 1 muestra. ¿Desea continuar?"
    Else
        strcadena = "Va a generar " & t & " para " & contar_marcados & " muestras. ¿Desea continuar?"
    End If
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        generar_documentos (TIPO) ' Factura
        Me.MousePointer = 0
        cargar_muestras
    End If

   On Error GoTo 0
   Exit Sub

facturar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure facturar of Formulario frmFacturacion_henkel"
End Sub

Private Sub cmdFactura_Click(Index As Integer)
    If Index = 0 Then
        facturar C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
    End If
    If Index = 1 Then
        facturar C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA
    End If
    If Index = 2 Then
        ' Cargar Datos para factura
        fechaFactura = Date
        cargar_combo_clientes
        frmDatosFactura.visible = True
    End If
End Sub

Private Sub cmdmarcarmuestras_Click(Index As Integer)
    Dim i As Integer
    If Index = 0 Or Index = 1 Then
        If muestras.ListItems.Count = 0 Then Exit Sub
        For i = 1 To muestras.ListItems.Count
            If Index = 0 Then
                muestras.ListItems(i).Checked = True
            Else
                muestras.ListItems(i).Checked = False
            End If
        Next
    End If
    calcular_total
End Sub

Private Sub cmdPrecios_Click()
'    frmHenkel_Precios.Show 1
    frmHenkel_Price.Show 1
End Sub


Private Sub conceptos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    calcular_total
End Sub

Private Sub cmdSalir_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    End If
    If Index = 1 Then
        frmDatosFactura.visible = False
    End If
    
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.top = 50
    cargar_botones Me
    cabecera_grid
    fdesde = "01/" & Month(Date) & "/" & Year(Date)
    fhasta = Date
    cargar_parametros
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        chkFecha.Value = Checked
        fdesde = "01/08/2018"
        fhasta = Date
    End If
    cargar_muestras
End Sub
Private Sub cargar_parametros()
    Dim op As New clsParametros
    If op.Carga(parametros.PARAM_HENKEL_PROBETAS_DTO, "") Then
        txttotal(13) = op.getVALOR
        txttotal(18) = op.getVALOR
    Else
        txttotal(13) = "0"
        txttotal(18) = "0"
    End If
    Set op = Nothing
End Sub
Private Sub cabecera_grid()
    With muestras.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "Código", 950, lvwColumnCenter
        .Add , , "Ref.Cliente", 2500, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "General", 850, lvwColumnCenter
        .Add , , "NºProbetas", 700, lvwColumnCenter
        .Add , , "Conformes", 700, lvwColumnCenter
        .Add , , "No Conformes", 700, lvwColumnCenter
        .Add , , "Importe", 900, lvwColumnRight
        .Add , , "Tasa", 650, lvwColumnCenter
        .Add , , "Total", 900, lvwColumnRight
        .Add , , "Probetas E", 0, lvwColumnRight
        .Add , , "Probetas ABCD", 0, lvwColumnRight
        .Add , , "TIPO_ANALISIS", 0, lvwColumnLeft
    End With
    With probetas.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "Designacion", 0, lvwColumnCenter
        .Add , , "Probeta", 0, lvwColumnLeft
        .Add , , "Area", 0, lvwColumnRight
        .Add , , "Material", 0, lvwColumnCenter
        .Add , , "Dimension", 1700, lvwColumnCenter
        .Add , , "Identificación", 2000, lvwColumnCenter
        .Add , , "Identificación Canagrosa", 0, lvwColumnCenter
        .Add , , "Criterio", 0, lvwColumnCenter
        .Add , , "Resultado", 0, lvwColumnCenter
        .Add , , "Fecha", 0, lvwColumnCenter
        .Add , , "Conforme", 1000, lvwColumnCenter
        .Add , , "Precio", 750, lvwColumnRight
        .Add , , "Tasa", 650, lvwColumnCenter
        .Add , , "Total", 750, lvwColumnRight
        .Add , , "DOC_ID", 0, lvwColumnRight
    End With
End Sub
Private Sub cargar_muestras()
    Dim rs As ADODB.Recordset
    ' Muestras
    Dim oHenkel As New clsHenkel
   On Error GoTo cargar_muestras_Error
    muestras.ListItems.Clear
    probetas.ListItems.Clear
    Set rs = oHenkel.ListadoPendienteFacturar(IIf(chkFecha.Value = vbChecked, fdesde, ""), IIf(chkFecha.Value = vbChecked, fhasta, ""), chkCerradas.Value, 0)
    lblMsg(1) = "Listado de muestras pendientes de Facturar : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With muestras.ListItems.Add(, , rs.Fields(0))
                .SubItems(COLS_MUESTRAS.C_CODIGO) = rs.Fields(1)
                .SubItems(COLS_MUESTRAS.C_REFERENCIA_CLIENTE) = rs.Fields(2)
                .SubItems(COLS_MUESTRAS.C_fecha) = rs.Fields(3)
                .SubItems(COLS_MUESTRAS.C_GENERAL) = rs.Fields(4)
                .SubItems(COLS_MUESTRAS.C_Probetas) = rs.Fields(5) ' TOTAL
                .SubItems(COLS_MUESTRAS.C_CONFORMES) = rs.Fields(6) ' CONFORMES
                .SubItems(COLS_MUESTRAS.C_NO_CONFORMES) = rs.Fields(7) ' NO CONFORMES
                If Not IsNull(rs.Fields(8)) Then
                    .SubItems(COLS_MUESTRAS.C_IMPORTE) = moneda(rs.Fields(8))
                    If rs.Fields(8) = 0 Then
                        muestras.ListItems(muestras.ListItems.Count).ListSubItems(COLS_MUESTRAS.C_IMPORTE).ForeColor = vbRed
                    End If
                End If
                .SubItems(COLS_MUESTRAS.C_PROBETAS_E) = rs.Fields(9)
                .SubItems(COLS_MUESTRAS.C_PROBETAS_ABCD) = rs.Fields(10)
                .SubItems(COLS_MUESTRAS.C_TIPO_ANALISIS) = rs.Fields(11)
            End With
            muestras.ListItems(muestras.ListItems.Count).Checked = True
            rs.MoveNext
         Loop Until rs.EOF
    End If
    Set rs = Nothing
    calcular_total
   On Error GoTo 0
   Exit Sub

cargar_muestras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestras of Formulario frmFacturacion_henkel"
End Sub
Private Sub recargar_muestra(ID_MUESTRA As Long)
    Dim rs As ADODB.Recordset
    Dim oHenkel As New clsHenkel
    Set rs = oHenkel.ListadoPendienteFacturar(IIf(chkFecha.Value = vbChecked, fdesde, ""), IIf(chkFecha.Value = vbChecked, fhasta, ""), chkCerradas.Value, ID_MUESTRA)
    If rs.RecordCount > 0 Then
        Do
            With muestras.ListItems(muestras.selectedItem.Index)
                .SubItems(COLS_MUESTRAS.C_CODIGO) = rs.Fields(1)
                .SubItems(COLS_MUESTRAS.C_REFERENCIA_CLIENTE) = rs.Fields(2)
                .SubItems(COLS_MUESTRAS.C_fecha) = rs.Fields(3)
                .SubItems(COLS_MUESTRAS.C_GENERAL) = rs.Fields(4)
                .SubItems(COLS_MUESTRAS.C_Probetas) = rs.Fields(5) ' TOTAL
                .SubItems(COLS_MUESTRAS.C_CONFORMES) = rs.Fields(6) ' CONFORMES
                .SubItems(COLS_MUESTRAS.C_NO_CONFORMES) = rs.Fields(7) ' NO CONFORMES
                If Not IsNull(rs.Fields(8)) Then
                    .SubItems(COLS_MUESTRAS.C_IMPORTE) = moneda(rs.Fields(8))
                    If rs.Fields(8) = 0 Then
                        muestras.ListItems(muestras.ListItems.Count).ListSubItems(COLS_MUESTRAS.C_IMPORTE).ForeColor = vbRed
                    End If
                End If
                .SubItems(COLS_MUESTRAS.C_PROBETAS_E) = rs.Fields(9)
                .SubItems(COLS_MUESTRAS.C_PROBETAS_ABCD) = rs.Fields(10)
                .SubItems(COLS_MUESTRAS.C_TIPO_ANALISIS) = rs.Fields(11)
            End With
            rs.MoveNext
         Loop Until rs.EOF
    End If
    Set rs = Nothing
    calcular_total
   On Error GoTo 0
   Exit Sub

cargar_muestras_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestras of Formulario frmFacturacion_henkel"
End Sub

Private Sub calcular_total()
'    Dim total As Currency
    Dim nMuestras As Integer
    Dim nProbetas As Integer
    Dim nProbetasE As Integer
    Dim nProbetasABCD As Integer
    Dim nConformes As Integer
    Dim nNoConformes As Integer
    Dim i As Integer
   On Error GoTo calcular_total_Error

    For i = 1 To muestras.ListItems.Count
        If muestras.ListItems(i).Checked = True Then
            nMuestras = nMuestras + 1
            nProbetas = nProbetas + muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_Probetas)
            nProbetasE = nProbetasE + muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_PROBETAS_E)
            nProbetasABCD = nProbetasABCD + muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_PROBETAS_ABCD)
            nConformes = nConformes + muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_CONFORMES)
            nNoConformes = nNoConformes + muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_NO_CONFORMES)
        End If
    Next
    txttotal(0) = nMuestras
    txttotal(1) = nProbetas
    txttotal(14) = nProbetasE
    txttotal(15) = nProbetasABCD
    txttotal(2) = nConformes
    txttotal(3) = nNoConformes
    ' Calcular CFM
    Dim cfm As New clsHenkel_cfm
    Dim tasa As Single
    tasa = 0
    If cfm.CargaRango(nProbetas) Then
        txttotal(4) = cfm.getP_INICIO & " - " & cfm.getP_FIN
        txttotal(5) = moneda(cfm.getPRECIO)
        txttotal(6) = cfm.getTASA & " %"
        tasa = cfm.getTASA
    Else
        txttotal(4) = ""
        txttotal(5) = ""
        txttotal(6) = ""
    End If
    ' Calcular TASAS
    Dim IMPORTE As Currency
    Dim totalE As Currency
    Dim totalABCD As Currency
    For i = 1 To muestras.ListItems.Count
        muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_TASA) = CStr(tasa)
        IMPORTE = monedaNum(muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_IMPORTE))
        IMPORTE = IMPORTE + ((IMPORTE * tasa) / 100)
        muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_total) = moneda(CStr(IMPORTE))
        If muestras.ListItems(i).Checked = True Then
            If muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_PROBETAS_E) > 0 Then
                totalE = totalE + moneda(CStr(IMPORTE))
            Else
                totalABCD = totalABCD + moneda(CStr(IMPORTE))
            End If
        End If
    Next
    ' Calcular TASAS DE PROBETAS
    cargar_probetas_tasa
    ' Importes E
    Dim media As Currency
    If totalE <> 0 And nProbetasE <> 0 Then
        media = moneda(totalE / nProbetasE)
    Else
        media = 0
    End If
    txttotal(12) = moneda(CStr(media))
    If nProbetasE > CInt(txttotal(13)) Then
        txttotal(11) = moneda(media * txttotal(13))
    Else
        txttotal(11) = moneda("0")
    End If
    ' Importes ABCD
    If totalABCD <> 0 And nProbetasABCD <> 0 Then
        media = moneda(totalABCD / nProbetasABCD)
    Else
        media = 0
    End If
    txttotal(17) = moneda(CStr(media))
    If nProbetasABCD > CInt(txttotal(18)) Then
        txttotal(16) = moneda(media * txttotal(18))
    Else
        txttotal(16) = moneda("0")
    End If
    ' Importe
    txttotal(9) = txttotal(5) ' CFM
    txttotal(8) = moneda(totalE + totalABCD) ' IMPORTE PROBETAS
    txttotal(7) = moneda(monedaNum(txttotal(11)) + monedaNum(txttotal(16))) ' IMPORTE DTO
    txttotal(10) = moneda(monedaNum(txttotal(9)) + monedaNum(txttotal(8)) - monedaNum(txttotal(7)))   ' TOTAL : CFM + PROBETAS - DTO
    
   On Error GoTo 0
   Exit Sub

calcular_total_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcular_total of Formulario frmFacturacion_henkel"
End Sub
Private Sub muestras_Click()
    If muestras.ListItems.Count = 0 Then Exit Sub
    cargar_probetas CLng(muestras.ListItems(muestras.selectedItem.Index).Text)
End Sub

Private Sub muestras_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If muestras.ListItems.Count > 0 Then
     muestras.SortKey = ColumnHeader.Index - 1
     If muestras.SortOrder = 0 Then
        muestras.SortOrder = 1
     Else
        muestras.SortOrder = 0
     End If
     muestras.Sorted = True
   End If
End Sub

Private Sub muestras_DblClick()
    If muestras.ListItems.Count = 0 Then Exit Sub
    gmuestra = CLng(muestras.ListItems(muestras.selectedItem.Index).Text)
    frmVerMuestra.Show 1
End Sub

Private Sub muestras_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    calcular_total
End Sub
Private Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To muestras.ListItems.Count
       If muestras.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function

Private Sub cargar_probetas(MUESTRA_ID As Long)
    Dim rs As ADODB.Recordset
    Dim oListItem As ListItem, oListSubItem As ListSubItem
    Dim oHenkel As New clsHenkel
   On Error GoTo cargar_probetas_Error
    probetas.ListItems.Clear
    Set rs = oHenkel.ListadoProbetas(MUESTRA_ID)
    If rs.RecordCount > 0 Then
        Do
            Set oListItem = probetas.ListItems.Add(, , rs.Fields(0))
            With oListItem
                .SubItems(1) = rs.Fields(1)
                .SubItems(2) = rs.Fields(2)
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = rs.Fields(4)
                .SubItems(5) = rs.Fields(5)
                .SubItems(6) = rs.Fields(6)
                .SubItems(7) = rs.Fields(7)
                .SubItems(8) = rs.Fields(8)
                .SubItems(9) = rs.Fields(9)
                .SubItems(10) = rs.Fields(10)
                Set oListSubItem = .ListSubItems.Add(, , IIf(rs.Fields(11) = 0, "No Conforme", "Conforme"), rs.Fields(11) + 1)
'                .SubItems(11) = rs.Fields(11)
                If IsNull(rs.Fields(13)) Then
                    .SubItems(COLS_PROBETAS.CP_IMPORTE) = moneda("0")
                Else
                    .SubItems(COLS_PROBETAS.CP_IMPORTE) = moneda(rs.Fields(13)) ' PRECIO
                End If
                .SubItems(COLS_PROBETAS.CP_DOC_ID) = rs(12)
                If rs(12) = 0 Then
                    .Checked = True
                ElseIf rs(12) = -1 Then
                    .Checked = False
                End If
            End With
            rs.MoveNext
         Loop Until rs.EOF
    End If
    Set rs = Nothing
    cargar_probetas_tasa
   On Error GoTo 0
   Exit Sub

cargar_probetas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_probetas of Formulario frmFacturacion_henkel"
End Sub
Private Sub cargar_probetas_tasa()
    Dim i As Integer
   On Error GoTo cargar_probetas_tasa_Error
    Dim tasa As Single
    Dim IMPORTE As Currency
    If txttotal(6) = "" Then
        tasa = 0
    Else
        tasa = Replace(txttotal(6), "%", "")
    End If
    For i = 1 To probetas.ListItems.Count
'        If probetas.ListItems(i).Checked = True Then
         probetas.ListItems(i).SubItems(COLS_PROBETAS.CP_TASA) = CStr(tasa)
         IMPORTE = Replace(moneda_bd(probetas.ListItems(i).SubItems(COLS_PROBETAS.CP_IMPORTE)), ".", ",")
         IMPORTE = IMPORTE + ((IMPORTE * tasa) / 100)
         probetas.ListItems(i).SubItems(COLS_PROBETAS.CP_TOTAL) = moneda(CStr(IMPORTE))
'        End If
    Next

   On Error GoTo 0
   Exit Sub

cargar_probetas_tasa_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_probetas_tasa of Formulario frmFacturacion_henkel"
End Sub

Private Sub muestras_KeyUp(KeyCode As Integer, Shift As Integer)
    If muestras.ListItems.Count = 0 Then Exit Sub
    cargar_probetas CLng(muestras.ListItems(muestras.selectedItem.Index).Text)
End Sub
Private Sub generar_documentos(TIPO_DOCUMENTO As Integer)
    Dim i As Integer
    Dim ID_DOC As Long
    'cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
   On Error GoTo generar_documentos_Error

'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Generar cabecera factura
    Dim oDocPago As New clsDocs_pago
    With oDocPago
        .setTIPO = TIPO_DOCUMENTO
        .setFECHA_FACTURA = Format(fechaFactura, "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setCLIENTE_ID = cmbCliente.getPK_SALIDA
        .setCLIENTE_ID_FACTURA = cmbClienteFactura.getPK_SALIDA
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbClienteFactura.getPK_SALIDA
        .setFP_ID = oCliente.getFP_ID
        If cmbPedidos.getTEXTO = "" Then
            .setPEDIDO_ID = 0
        Else
            .setPEDIDO_ID = cmbPedidos.getPK_SALIDA
        End If
        .setTOTAL = "0.00"
        .setDESCUENTO = "0.00"
'        If TIPO_DOCUMENTO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN Then
'            .setIVA = 0
'        Else
'            .setIVA = IVA
'        End If
        .setPAGADO = 0
        .setANULADO = 0
        .setFACTURA_CONCEPTOS = 2 ' MUESTRAS Y CONCEPTOS
        ' Insertamos el documento de pago
        ID_DOC = oDocPago.InsertarDocPago
        If ID_DOC = 0 Then
            MsgBox "Error al generar el documento, contacte con mantenimiento.", vbCritical, App.Title
            Exit Sub
        End If
        ' MUESTRAS
        Dim odpm As New clsDocs_pago_muestras
        Dim oMuestra As New clsMuestra
        Dim ORDEN As Integer
        ORDEN = 1
    
        Dim IMPORTE As Currency
        Dim tasa As Single
        If txttotal(6) = "" Then
            tasa = 0
        Else
            tasa = Replace(txttotal(6), "%", "")
        End If
        Dim oCE As New clsCe_resultados
        For i = 1 To muestras.ListItems.Count
            If muestras.ListItems(i).Checked = True Then
                With odpm
                    .setDOC_ID = ID_DOC
                    .setMUESTRA_ID = muestras.ListItems(i).Text
                    .setTIPO_ANALISIS = muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_TIPO_ANALISIS)
                    .setFECHA = Format(muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_fecha), "yyyy-mm-dd")
                    .setREFERENCIA_CLIENTE = muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_REFERENCIA_CLIENTE)
                    .setPRECIO = Replace(Format(muestras.ListItems(i).SubItems(COLS_MUESTRAS.C_total), "0.00"), ",", ".")
                    .setORDEN = ORDEN
                    ORDEN = .Insertar_doc_pago_muestra(False)
                    If ORDEN = -1 Then
                        MsgBox "Error al generar las facturas (2), contacte con mantenimiento.", vbCritical, App.Title
                        Exit Sub
                    Else
                        ORDEN = ORDEN + 1
                    End If
                    ' Generar el detalle de las probetas
                    Dim oHenkel As New clsHenkel
                    Dim rsprobetas As ADODB.Recordset
                    Set rsprobetas = oHenkel.ListadoProbetas(muestras.ListItems(i).Text)
                    Dim muestraCompleta As Boolean
                    muestraCompleta = True
                    If rsprobetas.RecordCount > 0 Then
                        Do
                            If rsprobetas("DOC_ID") = 0 Then
                                .setMUESTRA_ID = 0
                                .setTIPO_ANALISIS = rsprobetas("DIMENSION")
                                .setREFERENCIA_CLIENTE = rsprobetas("IDENTIFICACION_CLIENTE")
                                IMPORTE = rsprobetas("PRECIO")
                                IMPORTE = IMPORTE + ((IMPORTE * tasa) / 100)
                                .setPRECIO = moneda_bd(CStr(IMPORTE))
                                .setORDEN = ORDEN
                                ORDEN = .Insertar_doc_pago_muestra(False)
                                If ORDEN = -1 Then
                                    MsgBox "Error al generar las facturas (2), contacte con mantenimiento.", vbCritical, App.Title
                                    Exit Sub
                                Else
                                    ORDEN = ORDEN + 1
                                End If
                                ' Informar DOC_ID
                                oCE.modificarDOC_ID rsprobetas("MUESTRA_ID"), rsprobetas("DESIGNACION"), rsprobetas("PROBETA"), rsprobetas("AREA"), ID_DOC
                            Else
                                muestraCompleta = False
                            End If
                            rsprobetas.MoveNext
                        Loop Until rsprobetas.EOF
                    End If
                End With
                ' Actualizar resultados -1 a 0 para que se facture la proxima vez
                If Not muestraCompleta Then
                    oCE.restaurarDOC_ID muestras.ListItems(i).Text
                End If
                ' Modificar el documento de pago de las muestras facturadas
                If TIPO_DOCUMENTO <> C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_PROFORMA Then
                    If muestraCompleta Then
                        If oMuestra.Informar_Documento_Pago(muestras.ListItems(i).Text, TIPO_DOCUMENTO) = False Then
                            MsgBox "Error al informar el documento de pago en la muestra " & muestras.ListItems(i).Text & ", contacte con mantenimiento.", vbCritical, App.Title
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next
        ' Generar los conceptos FIJOS
        insertarConcepto ID_DOC, "PRECIOS PARA EL RANGO DE SUMINISTROS DE " & txttotal(4) & " PROBETAS AL MES", 0, 0, 0, 0
        insertarConcepto ID_DOC, "IMPORTE TOTAL PROBETAS PROCESADAS (TASA VOLUMEN APLICADA " & txttotal(6) & ")", txttotal(8), 0, 0, 0
        insertarConcepto ID_DOC, "TOTAL PROBETAS PROCESADAS " & txttotal(1) & " (CONFORMES : " & txttotal(2) & ", NO CONFORMES : " & txttotal(3) & ")", "0", 0, 0, 1
        insertarConcepto ID_DOC, "COSTE FIJO MENSUAL", txttotal(9), 1, 0, 0
        insertarConcepto ID_DOC, "DTO. APLICABLE PRIMERAS " & txttotal(13) & " PROBETAS DE TIPO E : " & txttotal(11) & ")", txttotal(11) * -1, 1, 0, 0
        insertarConcepto ID_DOC, "TOTAL PROBETAS E : " & txttotal(14) & " (PRECIO MEDIO : " & txttotal(12) & ")", "0", 0, 0, 1
        insertarConcepto ID_DOC, "DTO. APLICABLE PRIMERAS " & txttotal(18) & " PROBETAS DE TIPO A,B,C,D : " & txttotal(16) & ")", txttotal(16) * -1, 1, 0, 0
        insertarConcepto ID_DOC, "TOTAL PROBETAS A,B,C,D : " & txttotal(15) & " (PRECIO MEDIO : " & txttotal(17) & ")", "0", 0, 0, 1
        
        ' Recalcular total factura con los importes
        oDocPago.Informar_total_factura ID_DOC
    End With
    MsgBox "Se ha registrado el documento correctamente con numero : " + oDocPago.getNUMERO_FORMATEADO + "/" & Year(oDocPago.getFECHA_FACTURA), vbOKOnly + vbInformation, App.Title
    frmDatosFactura.visible = False
    cargar_muestras

   On Error GoTo 0
   Exit Sub

generar_documentos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_documentos of Formulario frmFacturacion_henkel"
End Sub
Private Sub insertarConcepto(ID_DOC As Long, DESCRIPCION As String, PRECIO As String, CANTIDAD As Integer, FAMILIA_ID As Integer, apartado As Integer)
    Dim oConcepto As New clsDocs_pago_conceptos
    With oConcepto
        .setDOC_ID = ID_DOC
        .setALBARAN_ID = 0
        .setDESCRIPCION = DESCRIPCION
        .setFECHA = Format(fechaFactura, "yyyy-mm-dd")
        .setCANTIDAD = CANTIDAD
        .setPRECIO = moneda_bd4(PRECIO)
        .setSUBTOTAL = moneda_bd4(CANTIDAD * PRECIO)
        .setDTO = 0
        .setTOTAL = moneda_bd4(CANTIDAD * PRECIO)
        .setFAMILIA_ID = FAMILIA_ID ' REVISAR
        .setAPARTADO = apartado
        If .Insertar = False Then
            MsgBox "Error al informar el concepto " & .getDESCRIPCION & ", contacte con mantenimiento.", vbCritical, App.Title
            Exit Sub
        End If
    End With
End Sub
Private Sub cargar_combo_clientes()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        
        Dim oParametros As New clsParametros
        oParametros.Carga parametros.PARAM_HENKEL_CLIENTE, ""
        
        consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
                   "  FROM CLIENTES AS C " & _
                   " WHERE C.ID_CLIENTE = " & oParametros.getVALOR
        With cmbCliente
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
        
        cmbCliente.MostrarElemento oParametros.getVALOR
        cmbClienteFactura.MostrarElemento oParametros.getVALOR
    End If
End Sub
Private Sub probetas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   On Error GoTo probetas_ItemCheck_Error

    If Item.ListSubItems(COLS_PROBETAS.CP_DOC_ID) > 0 Then
        MsgBox "No se puede modificar, la probeta ya esta facturada", vbCritical, App.Title
    Else
        Dim DOC_ID As Long
        DOC_ID = 0
        If Item.Checked = False Then
            DOC_ID = -1
        End If
        Item.ListSubItems(COLS_PROBETAS.CP_DOC_ID) = DOC_ID
        Dim oCE As New clsCe_resultados
        oCE.modificarDOC_ID Item.Text, Item.ListSubItems(1), Item.ListSubItems(2), Item.ListSubItems(3), DOC_ID
        recargar_muestra CLng(Item.Text)
    End If

   On Error GoTo 0
   Exit Sub

probetas_ItemCheck_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure probetas_ItemCheck of Formulario frmFacturacion_henkel"
End Sub
