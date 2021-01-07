VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmObras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Obras"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmObras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContratos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contratos"
      Height          =   915
      Left            =   4005
      Picture         =   "frmObras.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   9420
      Width           =   1275
   End
   Begin VB.CommandButton cmdCobro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cobro"
      Height          =   915
      Left            =   2690
      Picture         =   "frmObras.frx":12B4
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   9420
      Width           =   1275
   End
   Begin VB.CommandButton cmdSobre 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sobre"
      Height          =   915
      Left            =   1375
      Picture         =   "frmObras.frx":1B7E
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   9420
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   1
      Left            =   45
      TabIndex        =   57
      Top             =   6795
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Height          =   345
         Index           =   10
         Left            =   8100
         TabIndex        =   24
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Height          =   345
         Index           =   8
         Left            =   2700
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Height          =   345
         Index           =   4
         Left            =   1365
         TabIndex        =   22
         Top             =   240
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo cmbTipoIva 
         Height          =   360
         Left            =   4410
         TabIndex        =   62
         Top             =   225
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Libro"
         Height          =   195
         Index           =   2
         Left            =   7620
         TabIndex        =   60
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Iva"
         Height          =   195
         Index           =   0
         Left            =   3420
         TabIndex        =   59
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contrapartida"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   58
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdTarifa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tarifa"
      Height          =   915
      Left            =   60
      Picture         =   "frmObras.frx":2448
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9420
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Avisos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Index           =   0
      Left            =   60
      TabIndex        =   45
      Top             =   8415
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   585
         Index           =   12
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   8985
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Condiciones Especiales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   60
      TabIndex        =   41
      Top             =   3645
      Width           =   9165
      Begin VB.TextBox txtdatos 
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
         Height          =   300
         Index           =   17
         Left            =   1350
         TabIndex        =   21
         Top             =   2700
         Width           =   7695
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   300
         Index           =   15
         Left            =   1350
         TabIndex        =   20
         Top             =   2385
         Width           =   7695
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   14
         Left            =   6090
         TabIndex        =   14
         Top             =   990
         Width           =   945
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   300
         Index           =   13
         Left            =   1350
         TabIndex        =   12
         Top             =   630
         Width           =   7650
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   315
         Index           =   11
         Left            =   1350
         TabIndex        =   10
         Top             =   270
         Width           =   3390
      End
      Begin MSMask.MaskEdBox txtcuenta 
         Height          =   315
         Left            =   6090
         TabIndex        =   11
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   23
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-####-##-##########"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Top             =   990
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cmbTarifa 
         Height          =   315
         Left            =   1350
         TabIndex        =   15
         Top             =   1335
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cmbTipoFacturacion 
         Height          =   315
         Left            =   6090
         TabIndex        =   16
         Top             =   1335
         Width           =   2985
         _ExtentX        =   5265
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
      Begin MSDataListLib.DataCombo cmbTipoObra 
         Height          =   315
         Left            =   1350
         TabIndex        =   17
         Top             =   1695
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cmbTipoImporte 
         Height          =   315
         Left            =   6090
         TabIndex        =   18
         Top             =   1695
         Width           =   2985
         _ExtentX        =   5265
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
      Begin MSDataListLib.DataCombo cmbComercial 
         Height          =   315
         Left            =   1350
         TabIndex        =   19
         Top             =   2055
         Width           =   7725
         _ExtentX        =   13626
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dest. Factura"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   64
         Top             =   2760
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref. Factura"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   63
         Top             =   2445
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuento (%)"
         Height          =   195
         Index           =   4
         Left            =   4860
         TabIndex        =   61
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comercial"
         Height          =   195
         Left            =   180
         TabIndex        =   56
         Top             =   2115
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Importe"
         Height          =   195
         Left            =   4845
         TabIndex        =   55
         Top             =   1755
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Obra"
         Height          =   195
         Left            =   180
         TabIndex        =   51
         Top             =   1755
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Facturación"
         Height          =   195
         Left            =   4845
         TabIndex        =   50
         Top             =   1365
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa Porte"
         Height          =   195
         Left            =   180
         TabIndex        =   47
         Top             =   1395
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dirección"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   46
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Cuenta"
         Height          =   195
         Left            =   4920
         TabIndex        =   42
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   915
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9405
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9405
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
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
      Index           =   13
      Left            =   45
      TabIndex        =   37
      Top             =   7560
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   495
         Index           =   9
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   240
         Width           =   8985
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos de la Obra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   60
      TabIndex        =   29
      Top             =   375
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   5
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   52
         Top             =   330
         Width           =   2430
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   16
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2745
         Width           =   7695
      End
      Begin VB.CommandButton cmdaddprovincia 
         Caption         =   "+"
         Height          =   345
         Left            =   8760
         TabIndex        =   48
         Top             =   1890
         Width           =   315
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   0
         Left            =   7335
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2415
         Width           =   1695
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   7
         Left            =   5175
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2400
         Width           =   1440
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   6
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2400
         Width           =   3195
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Height          =   300
         Index           =   3
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1740
         Width           =   960
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   2
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   2
         Top             =   1380
         Width           =   7710
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   1
         Top             =   1020
         Width           =   7710
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   3690
         TabIndex        =   4
         Top             =   1740
         Width           =   4995
         _ExtentX        =   8811
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
      Begin MSDataListLib.DataCombo cmbMunicipio 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   2070
         Width           =   7335
         _ExtentX        =   12938
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
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   375
         Left            =   1350
         TabIndex        =   54
         Top             =   675
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   53
         Top             =   375
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "e-Mail"
         Height          =   195
         Index           =   9
         Left            =   225
         TabIndex        =   49
         Top             =   2805
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         Height          =   195
         Index           =   8
         Left            =   6975
         TabIndex        =   40
         Top             =   2505
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   39
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   36
         Top             =   2130
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Movil"
         Height          =   195
         Index           =   15
         Left            =   4725
         TabIndex        =   35
         Top             =   2460
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   225
         TabIndex        =   34
         Top             =   2460
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   33
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   32
         Top             =   1770
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   31
         Top             =   1425
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   30
         Top             =   1065
         Width           =   345
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nueva Obra"
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
      Height          =   315
      Left            =   60
      TabIndex        =   38
      Top             =   30
      Width           =   9180
   End
End
Attribute VB_Name = "frmObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk As Long

Private Sub cmbCliente_change()
    If cmbCliente.getTEXTO <> "" And pk = 0 Then
        ' Calcular el ID_OBRA
        Dim oObra As New clsObras
        oObra.CrearID cmbCliente.getPK_SALIDA
        txtdatos(5) = oObra.getID_OBRA
        ' Copiar los datos del cliente
        If MsgBox("¿Desea informar los datos de la obra con los del cliente?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim ocliente As New clsCliente
            With ocliente
                .CargaCliente cmbCliente.getPK_SALIDA
                txtdatos(2) = .getDIRECCION
                txtdatos(3) = .getCP
                cmbProvincia.BoundText = .getPROVINCIA_ID
                cmbMunicipio.BoundText = .getMUNICIPIO_ID
                txtdatos(6) = .getTELEFONO
                txtdatos(7) = .getMOVIL
                txtdatos(0) = .getFAX
                txtdatos(16) = .getEMAIL
                cmbfp.BoundText = .getFORMA_PAGO
            End With
            Set ocliente = Nothing
            txtdatos(1).SetFocus
        End If
    End If
End Sub

Private Sub cmbProvincia_Change()
    If cmbProvincia.Text <> "" Then
        cargar_municipios (cmbProvincia.BoundText)
    End If
End Sub
Private Sub cmdaddprovincia_Click()
    frmProvincias.Show 1
    Dim aux As Long
    aux = 0
    If cmbProvincia.Text <> "" Then
        aux = cmbProvincia.BoundText
    End If
    Cargar_Combo cmbProvincia, New clsProvincias
    cmbProvincia.BoundText = aux
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCobro_Click()
    frmObras_Cobros.pk = pk
    frmObras_Cobros.Show 1
End Sub

Private Sub cmdContratos_Click()
    frmObras_Contratos.pk = pk
    frmObras_Contratos.Show 1
End Sub

Private Sub cmdok_Click()
    If pk > 0 Then
        Modificar
    Else
        Insertar
    End If
End Sub

Private Sub cmdSobre_Click()
    Dim FILTRO As String
    FILTRO = FILTRO & " {obras.ID_OBRA} = " & pk
    With frmReport
        .iniciar
        .consulta = ""
        .CRITERIO = FILTRO
        .informe = "rptSobre"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport

End Sub

Private Sub cmdTarifa_Click()
    frmObras_Tarifas.pk = pk
    frmObras_Tarifas.Show 1
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    If pk > 0 Then
        cmdTarifa.Enabled = True
        cmdSobre.Enabled = True
        cmdCobro.Enabled = True
        consulta
    Else
        cmdTarifa.Enabled = False
        cmdSobre.Enabled = False
        cmdCobro.Enabled = False
        cmbTipoImporte.BoundText = 1
        txtdatos(8) = "2"
        txtdatos(10) = "1"
        txtdatos(14) = "0"
        txtdatos(4) = "7000001" ' Contrapartida
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmObras = Nothing
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 12 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtdatos(12).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 9 And Index <> 12 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
    If Index = 4 Then
        If Not IsNumeric(txtdatos(4)) Then
            txtdatos(4) = ""
        End If
    End If
End Sub
Public Sub Insertar()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta la Obra. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set oObra = mover_datos
        aux = oObra.Insertar
        If aux > 0 Then
            If MsgBox("La obra se ha insertado correctamente. ¿Desea generar la tarifa?", vbInformation + vbYesNo, App.Title) = vbYes Then
                frmObras_Tarifas.pk = aux
                frmObras_Tarifas.Show 1
                Unload Me
            Else
               Unload Me
            End If
        End If
        Set oObra = Nothing
    End If
End Sub

Public Sub Modificar()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim cliente As Integer
    pregunta = "Va a modificar los datos de la Obra. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oObra = mover_datos
        If oObra.Modificar(pk) = True Then
            MsgBox "La Obra se ha modificado correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set oObra = Nothing
    End If

End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If cmbCliente.getTEXTO = "" Then
        MsgBox "Indique el cliente asociado a la obra.", vbExclamation, App.Title
        cmbCliente.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(1) = "" Then
        MsgBox "El nombre de la obra no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(3) <> "" Then
        If IsNumeric(txtdatos(3)) = False Then
            MsgBox "El CP debe ser numérico.", vbCritical, "Error"
            txtdatos(3).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If cmbProvincia.Text = "" Then
        MsgBox "Seleccione una provincia.", vbInformation, App.Title
        cmbProvincia.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipio.Text = "" Then
        MsgBox "Seleccione un municipio.", vbInformation, App.Title
        cmbMunicipio.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbfp.Text = "" Then
        MsgBox "Seleccione la forma de pago.", vbInformation, App.Title
        cmbfp.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbTarifa.Text = "" Then
        MsgBox "Indique la tarifa de porte de la obra.", vbExclamation, App.Title
        cmbTarifa.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbTipoFacturacion.Text = "" Then
        MsgBox "Indique el tipo de facturación de la obra.", vbExclamation, App.Title
        cmbTipoFacturacion.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbTipoObra.Text = "" Then
        MsgBox "Indique el tipo de la obra.", vbExclamation, App.Title
        cmbTipoObra.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbTipoImporte.Text = "" Then
        MsgBox "Indique el tipo de importe", vbExclamation, App.Title
        cmbTipoImporte.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbTipoIva.Text = "" Then
        MsgBox "Indique el tipo de IVA", vbExclamation, App.Title
        cmbTipoIva.SetFocus
        valida_datos = False
        Exit Function
    End If
'    If txtdatos(8) = "" Then
'        MsgBox "El tipo de IVA deber ser numérico de 1 a 5", vbExclamation, App.Title
'        txtdatos(8).SetFocus
'        valida_datos = False
'        Exit Function
'    End If
'    If Not IsNumeric(txtdatos(8)) Then
'        MsgBox "El tipo de IVA deber ser numérico de 1 a 5", vbExclamation, App.Title
'        txtdatos(8).SetFocus
'        valida_datos = False
'        Exit Function
'    ElseIf CInt(txtdatos(8)) < 0 Or CInt(txtdatos(8)) > 5 Then
'        MsgBox "El tipo de IVA deber ser numérico de 1 a 5", vbExclamation, App.Title
'        txtdatos(8).SetFocus
'        valida_datos = False
'        Exit Function
'    End If
    If txtdatos(10) = "" Then
        MsgBox "El libro deber ser numérico", vbExclamation, App.Title
        txtdatos(10).SetFocus
        valida_datos = False
        Exit Function
    End If
    If Not IsNumeric(txtdatos(10)) Then
        MsgBox "El libro deber ser numérico", vbExclamation, App.Title
        txtdatos(10).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(14) = "" Then
        MsgBox "El descuento deber ser numérico", vbExclamation, App.Title
        txtdatos(14).SetFocus
        valida_datos = False
        Exit Function
    End If
    If Not IsNumeric(txtdatos(14)) Then
        MsgBox "El descuento deber ser numérico", vbExclamation, App.Title
        txtdatos(14).SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta()
    On Error GoTo fallo
    Dim oObra As New clsObras
    lbltitulo.Caption = "Modificacion de Obra"
    With oObra
        If .Carga(pk) = True Then
            txtdatos(5) = .getID_OBRA
            txtdatos(1) = .getNOMBRE
            cmbCliente.MostrarElemento .getCLIENTE_ID
            txtdatos(2) = .getDIRECCION
            txtdatos(3) = .getCP
            cmbProvincia.BoundText = .getPROVINCIA_ID
            cargar_municipios (.getPROVINCIA_ID)
            cmbMunicipio.BoundText = .getMUNICIPIO_ID
            txtdatos(6) = .getTELEFONO
            cmbfp.BoundText = .getFORMA_PAGO_ID
            txtdatos(4) = .getCONTRAPARTIDA
            txtdatos(16) = .getEMAIL
            
            txtdatos(11) = .getBANCO
            txtcuenta = .getCCC
            txtdatos(13) = .getBANCO_DIRECCION
            txtdatos(9) = .getOBSERVACIONES
            txtdatos(12) = .getAVISOS
            
'            txtdatos(8) = .getTIPO_IVA
            cmbTipoIva.BoundText = .getTIPO_IVA
            
            txtdatos(10) = .getLIBRO
            txtdatos(14) = .getDESCUENTO
             
            txtdatos(15) = .getREFERENCIA_FACTURA
            txtdatos(17) = .getDESTINO_FACTURA
            
            cmbTarifa.BoundText = .getTARIFA_PORTE_ID
            cmbTipoFacturacion.BoundText = .getTIPO_FACTURACION
            cmbTipoObra.BoundText = .getTIPO_OBRA_ID
            cmbTipoImporte.BoundText = .getTIPO_IMPORTE_ID
            cmbComercial.BoundText = .getCOMERCIAL_ID

        End If
    End With
    Set oObra = Nothing
    Exit Sub
fallo:
    log ("Error al consultar los datos de la Obra : " & Err.Description)
    MsgBox "Error al consultar los datos de la Obra.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsObras
    On Error GoTo fallo
    Dim oObra As New clsObras
    With oObra
        .setNOMBRE = txtdatos(1)
        .setCLIENTE_ID = cmbCliente.getPK_SALIDA
        .setDIRECCION = txtdatos(2)
        
        If txtdatos(3) = "" Then
            .setCP = 0
        Else
            .setCP = txtdatos(3)
        End If
        If cmbProvincia.Text = "" Then
            .setPROVINCIA_ID = 0
        Else
            .setPROVINCIA_ID = cmbProvincia.BoundText
        End If
        If cmbMunicipio.Text = "" Then
            .setMUNICIPIO_ID = 0
        Else
            .setMUNICIPIO_ID = cmbMunicipio.BoundText
        End If
        .setTELEFONO = txtdatos(6)
        If cmbfp.BoundText = "" Then
            .setFORMA_PAGO_ID = 0
        Else
            .setFORMA_PAGO_ID = cmbfp.BoundText
        End If
        If txtdatos(4) = "" Then
            .setCONTRAPARTIDA = 0
        Else
            .setCONTRAPARTIDA = txtdatos(4)
        End If
        .setEMAIL = txtdatos(16)
        .setBANCO = txtdatos(11)
        .setCCC = txtcuenta
        .setBANCO_DIRECCION = txtdatos(13)
        If cmbTarifa.Text = "" Then
            .setTARIFA_PORTE_ID = 0
        Else
            .setTARIFA_PORTE_ID = cmbTarifa.BoundText
        End If
        If cmbTipoFacturacion.Text = "" Then
            .setTIPO_FACTURACION = ""
        Else
            .setTIPO_FACTURACION = cmbTipoFacturacion.BoundText
        End If
        If cmbTipoObra.Text = "" Then
            .setTIPO_OBRA_ID = 0
        Else
            .setTIPO_OBRA_ID = cmbTipoObra.BoundText
        End If
        If cmbTipoImporte.Text = "" Then
            .setTIPO_IMPORTE_ID = 0
        Else
            .setTIPO_IMPORTE_ID = cmbTipoImporte.BoundText
        End If
        .setOBSERVACIONES = txtdatos(9)
        .setREFERENCIA_FACTURA = txtdatos(15)
        .setDESTINO_FACTURA = txtdatos(17)
        
        .setAVISOS = txtdatos(12)
        If cmbComercial.Text = "" Then
            .setCOMERCIAL_ID = 0
        Else
            .setCOMERCIAL_ID = cmbComercial.BoundText
        End If
        .setTIPO_IVA = cmbTipoIva.BoundText
'        .setTIPO_IVA = txtdatos(8)
        .setLIBRO = txtdatos(10)
        .setDESCUENTO = moneda_bd(txtdatos(14))
    End With
    Set mover_datos = oObra
    Set ocliente = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos de la obra.", vbCritical, Err.Description
End Function

Public Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    Cargar_Combo cmbfp, New clsForma_pago
    Cargar_Combo cmbTarifa, New clsTarifas_portes
    Cargar_Combo cmbProvincia, New clsProvincias
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo cmbTipoFacturacion, DECODIFICADORA.D_TIPOS_FACTURACION
    oDeco.Cargar_Combo cmbTipoObra, DECODIFICADORA.D_TIPOS_OBRAS
    oDeco.Cargar_Combo cmbTipoImporte, DECODIFICADORA.D_OBRAS_TIPOS_IMPORTE
    oDeco.Cargar_Combo cmbTipoIva, DECODIFICADORA.D_TIPOS_IVA
    Cargar_Combo cmbComercial, New clsComercial

End Sub
Public Sub cargar_municipios(PROVINCIA As Long)
    cmbMunicipio.Text = ""
    cargar_combo_FK cmbMunicipio, New clsMunicipios, PROVINCIA
End Sub
