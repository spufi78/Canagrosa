VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmSE_Recepcion 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Sellantes"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   Icon            =   "frmSE_Recepcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10845
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   9900
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   870
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   9900
      Width           =   1050
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10125
      MaxLength       =   255
      TabIndex        =   53
      Text            =   "1"
      Top             =   90
      Width           =   1275
   End
   Begin VB.Frame frmAIM 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Datos ADS (Sólo clientes Airbus)"
      DragMode        =   1  'Automatic
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
      Height          =   2460
      Left            =   45
      TabIndex        =   40
      Top             =   8325
      Width           =   9090
      Begin pryCombo.miCombo cmbPrograma 
         Height          =   330
         Left            =   1080
         TabIndex        =   41
         Top             =   765
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbEnsayo 
         Height          =   330
         Left            =   1080
         TabIndex        =   42
         Top             =   360
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSection 
         Height          =   330
         Left            =   1080
         TabIndex        =   43
         Top             =   1170
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFluid 
         Height          =   330
         Left            =   1080
         TabIndex        =   44
         Top             =   1575
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFacility 
         Height          =   330
         Left            =   1080
         TabIndex        =   45
         Top             =   1980
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   480
         Left            =   7785
         TabIndex        =   51
         Top             =   270
         Width           =   1185
         _Version        =   851970
         _ExtentX        =   2090
         _ExtentY        =   847
         _StockProps     =   79
         Caption         =   "Plantas"
         Appearance      =   5
         Picture         =   "frmSE_Recepcion.frx":2AFA
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Programa"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   135
         TabIndex        =   50
         Top             =   810
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   49
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Section"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   48
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fluid"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   135
         TabIndex        =   47
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facility"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   46
         Top             =   2025
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Sellante"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   30
      TabIndex        =   19
      Top             =   585
      Width           =   11445
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   11
         Top             =   2610
         Width           =   2115
      End
      Begin VB.CheckBox chkfmezcla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         Height          =   255
         Left            =   2700
         TabIndex        =   4
         Top             =   1755
         Width           =   1395
      End
      Begin VB.CheckBox chkFLimite 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         Height          =   255
         Left            =   2715
         TabIndex        =   8
         Top             =   2190
         Width           =   1395
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1755
         Width           =   2130
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   9045
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2205
         Width           =   2265
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   9
         Top             =   2205
         Width           =   2115
      End
      Begin MSDataListLib.DataCombo cmbproducto 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   1365
         Width           =   10245
         _ExtentX        =   18071
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   1725
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbPedido 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   990
         Width           =   9465
         _ExtentX        =   16695
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
      Begin MSComCtl2.DTPicker fechaLimite 
         Height          =   330
         Left            =   1095
         TabIndex        =   7
         Top             =   2160
         Width           =   1485
         _ExtentX        =   2619
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmSE_Recepcion.frx":935C
         Height          =   315
         Left            =   9045
         TabIndex        =   6
         Top             =   1755
         Width           =   2280
         _ExtentX        =   4022
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
      Begin pryCombo.miCombo cmbTiposAnalisis 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   270
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1080
         TabIndex        =   56
         Top             =   630
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Análisis"
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   39
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ratio Mezcla"
         Height          =   195
         Index           =   14
         Left            =   4560
         TabIndex        =   38
         Top             =   2625
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   8415
         TabIndex        =   37
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Límite"
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   36
         Top             =   2205
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   34
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   10695
         Picture         =   "frmSE_Recepcion.frx":93A2
         Stretch         =   -1  'True
         Top             =   1020
         Width           =   255
      End
      Begin VB.Image imgPedidos 
         Height          =   300
         Left            =   11025
         Picture         =   "frmSE_Recepcion.frx":9C6C
         Stretch         =   -1  'True
         Top             =   1035
         Width           =   315
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hora"
         Height          =   195
         Index           =   10
         Left            =   4530
         TabIndex        =   30
         Top             =   1785
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Lote y Kit"
         Height          =   195
         Index           =   1
         Left            =   8040
         TabIndex        =   29
         Top             =   2265
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Mezcla"
         Height          =   195
         Index           =   9
         Left            =   4560
         TabIndex        =   28
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   23
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Mezcla"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   21
         Top             =   1785
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   1425
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de recepción de la muestra"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   15
      TabIndex        =   24
      Top             =   3615
      Width           =   11445
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   4320
         MaxLength       =   255
         TabIndex        =   13
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1290
         MaxLength       =   255
         TabIndex        =   12
         Top             =   270
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo cmbentregada 
         Height          =   315
         Left            =   1305
         TabIndex        =   15
         Top             =   630
         Width           =   4230
         _ExtentX        =   7461
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
      Begin MSDataListLib.DataCombo cmbenvases 
         Height          =   315
         Left            =   7065
         TabIndex        =   14
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSDataListLib.DataCombo cmbrealizada 
         Height          =   315
         Left            =   7065
         TabIndex        =   16
         Top             =   630
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSDataListLib.DataCombo cmbDesProducto 
         Height          =   315
         Left            =   1305
         TabIndex        =   17
         Top             =   1035
         Width           =   9870
         _ExtentX        =   17410
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Des. Producto"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   33
         Top             =   1065
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Higrometría"
         Height          =   195
         Index           =   8
         Left            =   3210
         TabIndex        =   32
         Top             =   300
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Temperatura"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   315
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tomada por"
         Height          =   195
         Index           =   7
         Left            =   5985
         TabIndex        =   27
         Top             =   675
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         Height          =   195
         Index           =   5
         Left            =   5985
         TabIndex        =   26
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   25
         Top             =   690
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView lista 
      DragMode        =   1  'Automatic
      Height          =   2895
      Left            =   15
      TabIndex        =   18
      Top             =   5385
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   5106
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
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nº Muestras a Recepcionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   7200
      TabIndex        =   52
      Top             =   135
      Width           =   2865
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Sellantes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   150
      TabIndex        =   35
      Top             =   45
      Width           =   3540
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Ensayos"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   22
      Top             =   5085
      Width           =   11445
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   525
      Left            =   -180
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "frmSE_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const TA_DEFECTO As Long = 284

Private Sub cmbClientes_change()
    cargar_sellantes_cliente
    cargar_pedidos
    cargar_aim
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
MsgBox "OK"
End Sub

Private Sub PushButton3_Click()
    frmAirbus_Decodificadora.Show 1
    cargar_aim
End Sub

Private Sub chkFLimite_Click()
    If chkFLimite.Value = Checked Then
        fechaLimite.Enabled = False
    Else
        fechaLimite.Enabled = True
    End If
End Sub
Private Sub chkfmezcla_Click()
    If chkfMezcla.Value = Checked Then
        fecha.Enabled = False
    Else
        fecha.Enabled = True
    End If
End Sub

Private Sub Image1_Click()
    cmbPedido.Text = ""
    cmbPedido.BoundText = ""
End Sub

Private Sub imgPedidos_Click()
    If cmbClientes.getTEXTO <> "" Then
        frmClientes_Pedidos.PK = cmbClientes.getPK_SALIDA
        frmClientes_Pedidos.Show 1
        cargar_pedidos
    End If
End Sub
Private Sub cargar_pedidos()
    If cmbClientes.getTEXTO <> "" Then
        Dim oPedido As New clsClientes_pedidos
        Set cmbPedido.RowSource = oPedido.Listado_en_fecha(cmbClientes.getPK_SALIDA, CStr(Date))
        cmbPedido.ListField = "CODIGO_LARGO"
        cmbPedido.DataField = "ID_PEDIDO"
        cmbPedido.BoundColumn = "ID_PEDIDO"
    End If
End Sub
Private Sub cmbproducto_Change()
    If cmbProducto.Text <> "" Then
        cargar_sellante (cmbProducto.BoundText)
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim nMuestras As Integer
        Dim salida As String
        Me.MousePointer = 11
        cmdok.Enabled = False
        For nMuestras = 1 To CInt(txtcantidad)
            ' Generamos el registro de las muestras
            Dim oMuestra As New clsMuestra
            Dim muestra As Long
            With oMuestra
                .setTIPO_MUESTRA_ID = TIPOS_MUESTRAS.SELLANTE
                .setTIPO_ANALISIS_ID = cmbTiposAnalisis.getPK_SALIDA
                .setANALISIS_MODIFICADO = 3 ' Para identificar que es un Sellante
                .setFECHA_MUESTREO = Format(Date, "yyyy-mm-dd")
                .setENTIDAD_MUESTREO_ID = cmbrealizada.BoundText
                .setDETALLE_MUESTREO = ""
                .setOBSERVACIONES_MUESTREO = ""
                .setFECHA_RECEPCION = Format(Date, "yyyy-mm-dd")
                .setHORA_RECEPCION = Format(Time, "hh:mm")
                .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                .setFORMATO_ID = cmbenvases.BoundText
                .setENTIDAD_ENTREGA_ID = cmbentregada.BoundText
                .setDETALLE_ENTREGA = ""
                .setOBSERVACIONES_ENTREGA = ""
                .setCLIENTE_ID = cmbClientes.getPK_SALIDA
                .setCENTRO_ID = cmbCentro.BoundText
                .setREFERENCIA_CLIENTE = txtDatos(0)
                Dim oTipo_analisis As New clsTipos_analisis
                oTipo_analisis.CARGAR (cmbTiposAnalisis.getPK_SALIDA)
    '            .setPRECIO = Replace(Trim(oTipo_analisis.getPRECIO), ",", ".")
                Dim FechaEntrega As Date
                FechaEntrega = DateAdd("d", oTipo_analisis.getDIAS_TRABAJO, Date)
                .setFECHA_PREV_FIN = Format(FechaEntrega, "yyyy-mm-dd")
    '            .setFECHA_PREV_FIN = Format(Date, "yyyy-mm-dd")
                Dim oSellante As New clsSellantes
                oSellante.Carga cmbProducto.BoundText
                .setOBSERVACIONES = oSellante.getOBSERVACIONES
                .setANULADA = 0
                .setPRECINTO = ""
                .setBANO_ID = 0
                .setFECHA_COMIENZO = "0000-00-00"
                .setFECHA_FINALIZACION = "0000-00-00"
                .setFECHA_CIERRE = "0000-00-00"
                .setCERRADA = 0
                .setDOCUMENTO_PAGO = 0
                .setULT_EDICION_IMP = 0
                .setPRECIO = moneda_bd("0")
    '            .setPRODUCTO = txtDatos(5)
                .setPRODUCTO = cmbDesProducto.Text
                If cmbPedido.Text = "" Then
                    .setPEDIDO_ID = 0
                Else
                    .setPEDIDO_ID = cmbPedido.BoundText
                End If
                .setREPLACEMENT_ID = 0
                muestra = .guardarMuestra
                .informar_precio_muestra muestra
            End With
            ' Recepcion del Sellante
            Dim oSellante_recepcion As New clsSellantes_recepcion
            With oSellante_recepcion
                .setMUESTRA_ID = muestra
                .setSELLANTE_ID = cmbProducto.BoundText
                .setN_MEZCLA = txtDatos(0)
                .setR_MEZCLA = txtDatos(5)
                .setLOTE = txtDatos(1)
                .setTEMPERATURA = txtDatos(3)
                .setHIGROMETRIA = txtDatos(4)
                If chkfMezcla.Value = Checked Then
                    .setFECHA = Format("0000-00-00", "yyyy-mm-dd")
                Else
                    .setFECHA = Format(fecha.Value, "yyyy-mm-dd")
                End If
                If chkFLimite.Value = Checked Then
                    .setFECHA_LIMITE = Format("0000-00-00", "yyyy-mm-dd")
                Else
                    .setFECHA_LIMITE = Format(fechaLimite.Value, "yyyy-mm-dd")
                End If
                .setHORA = txtDatos(2)
                .Insertar
            End With
            ' Resultados
            Dim i As Integer
            Dim oSellante_resultados As New clsSellantes_resultados
            Dim oDE As New clsTipos_determinacion_equipos
            Dim oDR As New clsTipos_determinacion_botes_ex
            Dim rs As ADODB.Recordset
            Dim oSE_Equipos As New clsSellantes_equipos
            Dim cont As Integer
            Dim Equipos As String
            Dim REACTIVOS As String
            Dim PROPIOS As String
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    With oSellante_resultados
                        .setMUESTRA_ID = muestra
                        .setSELLANTE_ID = cmbProducto.BoundText
                        .setORDEN = lista.ListItems(i).SubItems(6)
                        .setTIPO_DETERMINACION_ID = lista.ListItems(i).SubItems(7) ' TIPO_DETERMINACION_ID
                        .setFORMULA_ID = lista.ListItems(i).SubItems(8) ' FORMULA_ID
                        .setRESULTADO = ""
                        .setCONFORME = 0
                        .Insertar
                    End With
                    If lista.ListItems(i).SubItems(7) <> 0 Then
                        'M1137-I
                        'Set rs = oDE.Listado(lista.ListItems(i).SubItems(7))
                        Set rs = oDE.ListadoActivos(lista.ListItems(i).SubItems(7))
                        'M1137-F
                        cont = 1
                        If rs.RecordCount > 0 Then
                            Do
                                Equipos = Equipos & rs(0) & ";"
                                With oSE_Equipos
                                    .setMUESTRA_ID = muestra
                                    .setORDEN = cont
                                    .setEQUIPO_ID = rs(0)
                                    .setVERIFICACION_ID = 0
                                    .setEN_INFORME = 0 'Falta meterlo en la tabla
                                    .Insertar
                                End With
                                rs.MoveNext
                            Loop Until rs.EOF
                        End If
                        Set rs = oDR.Listado(lista.ListItems(i).SubItems(7))
                        If rs.RecordCount > 0 Then
                            Do
                                If rs(1) = "E" Then
                                    REACTIVOS = REACTIVOS & rs(0) & ";"
                                Else
                                    PROPIOS = PROPIOS & rs(0) & ";"
                                End If
                                rs.MoveNext
                            Loop Until rs.EOF
                        End If
                    
                    End If
                End If
            Next
            If Equipos <> "" Then
               oSellante_recepcion.ModificarEquipos muestra, Equipos
            End If
            If REACTIVOS <> "" Then
               oSellante_recepcion.ModificarReactivos muestra, REACTIVOS, PROPIOS
            End If
            ' ADS
            If frmAIM.Enabled = True Then
                Dim oM As New clsMuestras_airbus
                With oM
                    .setMUESTRA_ID = muestra
                    .setENSAYO_ID = IIf(cmbEnsayo.getTEXTO = "", 0, cmbEnsayo.getPK_SALIDA)
                    .setPROGRAMA_ID = IIf(cmbPrograma.getTEXTO = "", 0, cmbPrograma.getPK_SALIDA)
                    .setSECTION_ID = IIf(cmbSection.getTEXTO = "", 0, cmbSection.getPK_SALIDA)
                    .setFLUID_ID = IIf(cmbFluid.getTEXTO = "", 0, cmbFluid.getPK_SALIDA)
                    .setFACILITY_ID = IIf(cmbFacility.getTEXTO = "", 0, cmbFacility.getPK_SALIDA)
                    .Insertar True, True, True, True, True
                End With
            End If
            ' M3338 : Informar observacion del sellante si la tuviera
            If oSellante.getOBSERVACIONES_DE <> "" Then
                Dim oDatos_especificos As New clsDatos_valores
                With oDatos_especificos
                       .setMUESTRA_ID = muestra
                       .setBANO_ID = 0
                       .setTIPO_DATO_ID = 1
                       .setVALOR = oSellante.getOBSERVACIONES_DE
                       .setORDEN = 1
                       .Insertar
                End With
            End If
            
            salida = salida & vbNewLine & oMuestra.CodigoParticular(muestra)
        Next
        cmdok.Enabled = True
        Me.MousePointer = 0
        
        txtDatos(0) = ""
        txtDatos(2) = ""
        txtDatos(3) = ""
        txtDatos(4) = ""
        
        MsgBox "La recepción se ha realizado correctamente : " & vbNewLine & salida & vbNewLine & vbNewLine & " Puede seguir registrando más sellantes modificando los datos necesarios o cerrar la ventana.", vbInformation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    cmdok.Enabled = True
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmSE_Recepcion")
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.top = 100
    
    Call cabecera
    Call cargar_combos
    fecha = Date
    fechaLimite = Date
    
    txtDatos(2) = Format(Time, "hh:mm")
End Sub
Private Function validar() As Boolean
    validar = True
    
    If txtcantidad = "" Then
        MsgBox "Debe indicar el número de probetas a recepcionar.", vbExclamation, App.Title
        validar = False
        Exit Function
    Else
        If Not IsNumeric(txtcantidad) Then
            MsgBox "Debe indicar el número de probetas a recepcionar.", vbExclamation, App.Title
            validar = False
            Exit Function
        End If
    End If
    If cmbProducto.BoundText = "" Then
        MsgBox "Debe asignar un producto a la selección.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Debe asignar un cliente a la selección.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbCentro.Text = "" Then
        MsgBox "El CENTRO no puede estar en blanco.", vbExclamation, "Validación"
        cmbCentro.SetFocus
        validar = False
        Exit Function
    End If
    If cmbTiposAnalisis.getTEXTO = "" Then
        MsgBox "El TIPO DE ANALISIS no puede estar en blanco.", vbExclamation, "Validación"
        cmbTiposAnalisis.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "Inserte la mezcla.", vbExclamation, App.Title
        validar = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If txtDatos(1) = "" Then
        MsgBox "Inserte el lote.", vbExclamation, App.Title
        validar = False
        txtDatos(1).SetFocus
        Exit Function
    End If
    If txtDatos(3) = "" Then
        MsgBox "Inserte la temperatura de la muestra.", vbExclamation, App.Title
        validar = False
        txtDatos(3).SetFocus
        Exit Function
    End If
    If txtDatos(4) = "" Then
        MsgBox "Inserte la higrometría de la muestra.", vbExclamation, App.Title
        validar = False
        txtDatos(4).SetFocus
        Exit Function
    End If
'    If txtDatos(5) = "" Then
'        MsgBox "Inserte la descripción del producto.", vbExclamation, App.Title
'        validar = False
'        txtDatos(5).SetFocus
'        Exit Function
'    End If
'    If cmbDesProducto.Text = "" Then
'        MsgBox "Inserte la descripción del producto.", vbExclamation, App.Title
'        validar = False
'        cmbDesProducto.SetFocus
'        Exit Function
'    End If
    If cmbenvases.BoundText = "" Then
        MsgBox "Debe asignar un envase a la selección.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbentregada.BoundText = "" Then
        MsgBox "Debe asignar la entrega.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbrealizada.BoundText = "" Then
        MsgBox "Debe asignar la realización.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbPedido.Text <> "" Then
        If Not IsNumeric(cmbPedido) Then
            MsgBox "El pedido no esta correctamente informado.", vbExclamation, App.Title
            validar = False
            Exit Function
        End If
    End If
End Function

Private Sub cargar_combos()
    cargar_clientes
    cargar_combo cmbCentro, New clsCentros
    cargar_combo cmbenvases, New clsformatos
    cargar_combo cmbentregada, New clsEntidades_Entrega
    cargar_combo cmbrealizada, New clsEntidades_muestreo
    cargar_producto
    llenar_combo cmbTiposAnalisis, New clsTipos_analisis, 0, frmTA_Detalle, " TIPO_MUESTRA_ID = " & TIPOS_MUESTRAS.SELLANTE
    cmbTiposAnalisis.MostrarElemento TA_DEFECTO
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Ensayo", 3400, lvwColumnLeft
        .Add , , "Test", 3400, lvwColumnLeft
        .Add , , "R.Inferior", 1000, lvwColumnCenter
        .Add , , "R.Superior", 1000, lvwColumnCenter
        .Add , , "Unidades", 2300, lvwColumnCenter
        .Add , , "UNIDAD_ID", 1, lvwColumnCenter
        .Add , , "ORDEN", 1, lvwColumnCenter
        .Add , , "TIPO_DETERMINACION_ID", 1, lvwColumnCenter
        .Add , , "FORMULA_ID", 1, lvwColumnCenter
    End With
End Sub

Private Sub cargar_sellante(ID As Long)
    Dim oSellantes_ensayos As New clsSellantes_ensayos
    Dim rs As ADODB.Recordset
    ' Sellante
'    Dim oSellante As New clsSellantes
   On Error GoTo cargar_sellante_Error

    lista.ListItems.Clear
    Set rs = oSellantes_ensayos.Listado(ID)
    If rs.RecordCount > 0 Then
       Do
          With lista.ListItems.Add(, , rs(1))
                .SubItems(1) = rs(2)
                .SubItems(2) = rs(3)
                .SubItems(3) = rs(4)
                .SubItems(4) = rs(5)
                .SubItems(5) = rs(6)
                .SubItems(6) = rs(0)
                .SubItems(7) = rs(7)
                ' Cargamos la formula
                If rs(7) = 0 Then
                    .SubItems(8) = 0
                Else
                    Dim oTD As New clsTipos_determinacion
                    oTD.CargarTipoDeterminacion rs(7)
                    .SubItems(8) = oTD.getFORMULA_ID
                End If
                If rs(11) = 1 Then
                    .Checked = True
                Else
                    .Checked = False
                End If
          End With
'          lista.ListItems(lista.ListItems.Count).Checked = True
          rs.MoveNext
       Loop Until rs.EOF
    End If
    Set oSellantes_ensayos = Nothing

   On Error GoTo 0
   Exit Sub

cargar_sellante_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_sellante of Formulario frmSE_Recepcion"
End Sub
Private Sub cargar_clientes()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
                   "  FROM SELLANTES S, CLIENTES C " & _
                   " WHERE S.CLIENTE_ID = C.ID_CLIENTE "
        With cmbClientes
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
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80FFFF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    ' Unidades temperatura
    If Index = 3 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = Replace(txtDatos(Index), "ºC", "")
            txtDatos(Index) = txtDatos(Index) & "ºC"
        End If
    End If
    If Index = 4 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = Replace(txtDatos(Index), "%", "")
            txtDatos(Index) = txtDatos(Index) & "%"
        End If
    End If
    
End Sub

Private Sub cargar_sellantes_cliente()
    If cmbClientes.getTEXTO = "" Then
        Exit Sub
    End If
    Dim oSellante As New clsSellantes
    Set cmbProducto.RowSource = oSellante.Listado_Combo_Sellantes_de_Clientes(cmbClientes.getPK_SALIDA)
    cmbProducto.ListField = "C2"
    cmbProducto.DataField = "C1" 'campo asociado
    cmbProducto.BoundColumn = "C1" 'lo que realmente
    Set oSellante = Nothing
End Sub

Private Sub cargar_producto()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT VALOR, DESCRIPCION " & _
               "  FROM decodificadora " & _
               " WHERE CODIGO = " & DECODIFICADORA.DESCRIPCION_PRODUCTO & _
               "   AND PARAMETROS = '" & TIPOS_MUESTRAS.SELLANTE & "'" & _
               " ORDER BY DESCRIPCION "
    Set rs = datos_bd(consulta)
    Set cmbDesProducto.RowSource = rs
    cmbDesProducto.ListField = "DESCRIPCION" 'lo que enseña
    cmbDesProducto.DataField = "VALOR" 'campo asociado
    cmbDesProducto.BoundColumn = "VALOR" 'lo que realmente
End Sub
Private Sub cargar_aim()
   On Error GoTo cargar_aim_Error

    frmAIM.Enabled = False
    cmbEnsayo.limpiar
    cmbPrograma.limpiar
    cmbSection.limpiar
    cmbFluid.limpiar
    cmbFacility.limpiar
    
    If cmbClientes.getTEXTO <> "" Then
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbClientes.getPK_SALIDA
        Dim ID_PLANTA As String
        ID_PLANTA = CStr(oCliente.getPLANT_ID)
        If oCliente.getAIRBUS = 1 Then
            If ID_PLANTA = "0" Then
                MsgBox "El cliente ADS no tiene informada la planta. Es necesario informarla en la ficha de cliente.", vbCritical, App.Title
                Exit Sub
            Else
                frmAIM.Enabled = True
                Dim oDeco As New clsDecodificadora
                oDeco.cargar_mi_combo_parametro cmbEnsayo, DECODIFICADORA.AIRBUS_TIPOS_ENSAYOS, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbPrograma, DECODIFICADORA.AIRBUS_PROGRAMAS, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbSection, DECODIFICADORA.AIRBUS_SECTION, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbFluid, DECODIFICADORA.AIRBUS_FLUID, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbFacility, DECODIFICADORA.AIRBUS_FACILITY, ID_PLANTA
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cargar_aim_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_aim of Formulario frmSE_Recepcion"
End Sub

