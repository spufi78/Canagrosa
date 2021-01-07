VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmPinturasRecepcionAdministrativa 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción Administrativa de PINTURAS"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin Geslab.ControlPanelXP cpDatos 
      Height          =   1650
      Left            =   45
      TabIndex        =   57
      Top             =   585
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   2910
      Caption         =   "Datos de recepción"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   1650
      Begin MSDataListLib.DataCombo cmbPedido 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Top             =   1215
         Width           =   6855
         _ExtentX        =   12091
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
      Begin pryCombo.miCombo cmbOfertas 
         Height          =   330
         Left            =   1170
         TabIndex        =   1
         Top             =   855
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   10800
         TabIndex        =   3
         Top             =   495
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   51642369
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1170
         TabIndex        =   0
         Top             =   495
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   582
      End
      Begin VB.Image imgPedidos 
         Height          =   300
         Left            =   8430
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   315
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   8100
         Stretch         =   -1  'True
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   61
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Recepción"
         Height          =   195
         Index           =   6
         Left            =   9315
         TabIndex        =   60
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Oferta"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   59
         Top             =   900
         Width           =   435
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   58
         Top             =   540
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8475
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   11805
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8475
      Width           =   1050
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   4890
      Left            =   45
      TabIndex        =   62
      Top             =   2250
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8625
      Caption         =   "Características Técnicas"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   4890
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   4380
         Left            =   45
         TabIndex        =   69
         Top             =   450
         Width           =   12885
         _Version        =   851970
         _ExtentX        =   22728
         _ExtentY        =   7726
         _StockProps     =   68
         Color           =   2
         ItemCount       =   4
         Item(0).Caption =   "BASE"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "frmCT(0)"
         Item(1).Caption =   "ENDURECEDOR"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "frmCT(1)"
         Item(2).Caption =   "ACTIVADOR"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "frmCT(2)"
         Item(3).Caption =   "DISOLVENTE"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "frmCT(4)"
         Begin VB.Frame frmCT 
            BackColor       =   &H00C0C0C0&
            Height          =   3975
            Index           =   4
            Left            =   -70000
            TabIndex        =   126
            Top             =   315
            Visible         =   0   'False
            Width           =   12795
            Begin VB.TextBox txtNOMBRE_BASE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   2565
               TabIndex        =   42
               Top             =   270
               Width           =   10125
            End
            Begin VB.TextBox txtFABRICANTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   2565
               TabIndex        =   43
               Top             =   630
               Width           =   10125
            End
            Begin VB.TextBox txtESPECIFICACIONES 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   2565
               TabIndex        =   44
               Top             =   990
               Width           =   10125
            End
            Begin VB.TextBox txtN_LATAS 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   2565
               TabIndex        =   45
               Top             =   1350
               Width           =   10125
            End
            Begin VB.TextBox txtCANTIDAD_LATA 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   2565
               TabIndex        =   46
               Top             =   1710
               Width           =   10125
            End
            Begin VB.TextBox txtLOTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   2565
               TabIndex        =   47
               Top             =   2070
               Width           =   10125
            End
            Begin VB.TextBox txtCONDICIONES_ALMACENAMIENTO 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   3600
               TabIndex        =   52
               Top             =   3150
               Width           =   9090
            End
            Begin VB.TextBox txtESTADO_EMBALAJE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   4
               Left            =   3600
               TabIndex        =   53
               Top             =   3510
               Width           =   9090
            End
            Begin VB.CheckBox chkFechaEnvasado 
               Caption         =   "Check1"
               Height          =   195
               Index           =   4
               Left            =   2565
               TabIndex        =   48
               Top             =   2475
               Width           =   240
            End
            Begin VB.CheckBox chkFechaCaducidad 
               Caption         =   "Check1"
               Height          =   195
               Index           =   4
               Left            =   2565
               TabIndex        =   50
               Top             =   2835
               Width           =   240
            End
            Begin MSComCtl2.DTPicker fechaCaducidad 
               Height          =   330
               Index           =   4
               Left            =   2835
               TabIndex        =   51
               Top             =   2790
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fechaEnvasado 
               Height          =   330
               Index           =   4
               Left            =   2835
               TabIndex        =   49
               Top             =   2430
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESPECIFICACIONES"
               Height          =   195
               Index           =   53
               Left            =   180
               TabIndex        =   136
               Top             =   1035
               Width           =   1515
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FABRICANTE"
               Height          =   195
               Index           =   52
               Left            =   180
               TabIndex        =   135
               Top             =   675
               Width           =   1005
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "NOMBRE DE LA BASE"
               Height          =   195
               Index           =   51
               Left            =   180
               TabIndex        =   134
               Top             =   315
               Width           =   1680
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nº DE LATAS / BOTES"
               Height          =   195
               Index           =   50
               Left            =   180
               TabIndex        =   133
               Top             =   1395
               Width           =   1710
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CANTIDAD POR LATA / BOTE"
               Height          =   195
               Index           =   49
               Left            =   180
               TabIndex        =   132
               Top             =   1755
               Width           =   2265
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "LOTE"
               Height          =   195
               Index           =   48
               Left            =   180
               TabIndex        =   131
               Top             =   2115
               Width           =   420
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE ENVASADO"
               Height          =   195
               Index           =   47
               Left            =   180
               TabIndex        =   130
               Top             =   2475
               Width           =   1725
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE CADUCIDAD"
               Height          =   195
               Index           =   46
               Left            =   180
               TabIndex        =   129
               Top             =   2835
               Width           =   1785
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CONDICIONES DE ALMACENAMIENTO"
               Height          =   195
               Index           =   45
               Left            =   180
               TabIndex        =   128
               Top             =   3195
               Width           =   2925
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESTADO DE EMBALAJE Y DEL CONTENIDO"
               Height          =   195
               Index           =   44
               Left            =   180
               TabIndex        =   127
               Top             =   3555
               Width           =   3315
            End
         End
         Begin VB.Frame frmCT 
            BackColor       =   &H00C0C0C0&
            Height          =   3975
            Index           =   2
            Left            =   -70000
            TabIndex        =   92
            Top             =   315
            Visible         =   0   'False
            Width           =   12795
            Begin VB.Frame frmCT 
               BackColor       =   &H00C0C0C0&
               Height          =   3975
               Index           =   3
               Left            =   0
               TabIndex        =   115
               Top             =   0
               Width           =   12795
               Begin VB.CheckBox chkFechaCaducidad 
                  Caption         =   "Check1"
                  Height          =   195
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   38
                  Top             =   2835
                  Width           =   240
               End
               Begin VB.CheckBox chkFechaEnvasado 
                  Caption         =   "Check1"
                  Height          =   195
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   36
                  Top             =   2475
                  Width           =   240
               End
               Begin VB.TextBox txtESTADO_EMBALAJE 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   41
                  Top             =   3510
                  Width           =   9090
               End
               Begin VB.TextBox txtCONDICIONES_ALMACENAMIENTO 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   40
                  Top             =   3150
                  Width           =   9090
               End
               Begin VB.TextBox txtLOTE 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   35
                  Top             =   2070
                  Width           =   10125
               End
               Begin VB.TextBox txtCANTIDAD_LATA 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   34
                  Top             =   1710
                  Width           =   10125
               End
               Begin VB.TextBox txtN_LATAS 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   33
                  Top             =   1350
                  Width           =   10125
               End
               Begin VB.TextBox txtESPECIFICACIONES 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   32
                  Top             =   990
                  Width           =   10125
               End
               Begin VB.TextBox txtFABRICANTE 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   31
                  Top             =   630
                  Width           =   10125
               End
               Begin VB.TextBox txtNOMBRE_BASE 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Height          =   330
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   30
                  Top             =   270
                  Width           =   10125
               End
               Begin MSComCtl2.DTPicker fechaCaducidad 
                  Height          =   330
                  Index           =   3
                  Left            =   2835
                  TabIndex        =   39
                  Top             =   2790
                  Width           =   1290
                  _ExtentX        =   2275
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
                  Format          =   51642369
                  CurrentDate     =   38000
                  MinDate         =   2
               End
               Begin MSComCtl2.DTPicker fechaEnvasado 
                  Height          =   330
                  Index           =   3
                  Left            =   2835
                  TabIndex        =   37
                  Top             =   2430
                  Width           =   1290
                  _ExtentX        =   2275
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
                  Format          =   51642369
                  CurrentDate     =   38000
                  MinDate         =   2
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ESTADO DE EMBALAJE Y DEL CONTENIDO"
                  Height          =   195
                  Index           =   43
                  Left            =   180
                  TabIndex        =   125
                  Top             =   3555
                  Width           =   3315
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "CONDICIONES DE ALMACENAMIENTO"
                  Height          =   195
                  Index           =   42
                  Left            =   180
                  TabIndex        =   124
                  Top             =   3195
                  Width           =   2925
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "FECHA DE CADUCIDAD"
                  Height          =   195
                  Index           =   41
                  Left            =   180
                  TabIndex        =   123
                  Top             =   2835
                  Width           =   1785
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "FECHA DE ENVASADO"
                  Height          =   195
                  Index           =   40
                  Left            =   180
                  TabIndex        =   122
                  Top             =   2475
                  Width           =   1725
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "LOTE"
                  Height          =   195
                  Index           =   39
                  Left            =   180
                  TabIndex        =   121
                  Top             =   2115
                  Width           =   420
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "CANTIDAD POR LATA / BOTE"
                  Height          =   195
                  Index           =   38
                  Left            =   180
                  TabIndex        =   120
                  Top             =   1755
                  Width           =   2265
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Nº DE LATAS / BOTES"
                  Height          =   195
                  Index           =   37
                  Left            =   180
                  TabIndex        =   119
                  Top             =   1395
                  Width           =   1710
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "NOMBRE DE LA BASE"
                  Height          =   195
                  Index           =   36
                  Left            =   180
                  TabIndex        =   118
                  Top             =   315
                  Width           =   1680
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "FABRICANTE"
                  Height          =   195
                  Index           =   35
                  Left            =   180
                  TabIndex        =   117
                  Top             =   675
                  Width           =   1005
               End
               Begin VB.Label lblCampos 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ESPECIFICACIONES"
                  Height          =   195
                  Index           =   34
                  Left            =   180
                  TabIndex        =   116
                  Top             =   1035
                  Width           =   1515
               End
            End
            Begin VB.CheckBox chkFechaCaducidad 
               Caption         =   "Check1"
               Height          =   195
               Index           =   2
               Left            =   2565
               TabIndex        =   102
               Top             =   2835
               Width           =   240
            End
            Begin VB.CheckBox chkFechaEnvasado 
               Caption         =   "Check1"
               Height          =   195
               Index           =   2
               Left            =   2565
               TabIndex        =   101
               Top             =   2475
               Width           =   240
            End
            Begin VB.TextBox txtESTADO_EMBALAJE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   3600
               TabIndex        =   100
               Top             =   3510
               Width           =   9090
            End
            Begin VB.TextBox txtCONDICIONES_ALMACENAMIENTO 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   3600
               TabIndex        =   99
               Top             =   3150
               Width           =   9090
            End
            Begin VB.TextBox txtLOTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   2565
               TabIndex        =   98
               Top             =   2070
               Width           =   10125
            End
            Begin VB.TextBox txtCANTIDAD_LATA 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   2565
               TabIndex        =   97
               Top             =   1710
               Width           =   10125
            End
            Begin VB.TextBox txtN_LATAS 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   2565
               TabIndex        =   96
               Top             =   1350
               Width           =   10125
            End
            Begin VB.TextBox txtESPECIFICACIONES 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   2565
               TabIndex        =   95
               Top             =   990
               Width           =   10125
            End
            Begin VB.TextBox txtFABRICANTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   2565
               TabIndex        =   94
               Top             =   630
               Width           =   10125
            End
            Begin VB.TextBox txtNOMBRE_BASE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   2565
               TabIndex        =   93
               Top             =   270
               Width           =   10125
            End
            Begin MSComCtl2.DTPicker fechaCaducidad 
               Height          =   330
               Index           =   2
               Left            =   2835
               TabIndex        =   103
               Top             =   2790
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fechaEnvasado 
               Height          =   330
               Index           =   2
               Left            =   2835
               TabIndex        =   104
               Top             =   2430
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESTADO DE EMBALAJE Y DEL CONTENIDO"
               Height          =   195
               Index           =   33
               Left            =   180
               TabIndex        =   114
               Top             =   3555
               Width           =   3315
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CONDICIONES DE ALMACENAMIENTO"
               Height          =   195
               Index           =   32
               Left            =   180
               TabIndex        =   113
               Top             =   3195
               Width           =   2925
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE CADUCIDAD"
               Height          =   195
               Index           =   31
               Left            =   180
               TabIndex        =   112
               Top             =   2835
               Width           =   1785
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE ENVASADO"
               Height          =   195
               Index           =   30
               Left            =   180
               TabIndex        =   111
               Top             =   2475
               Width           =   1725
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "LOTE"
               Height          =   195
               Index           =   29
               Left            =   180
               TabIndex        =   110
               Top             =   2115
               Width           =   420
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CANTIDAD POR LATA / BOTE"
               Height          =   195
               Index           =   28
               Left            =   180
               TabIndex        =   109
               Top             =   1755
               Width           =   2265
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nº DE LATAS / BOTES"
               Height          =   195
               Index           =   27
               Left            =   180
               TabIndex        =   108
               Top             =   1395
               Width           =   1710
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "NOMBRE DE LA BASE"
               Height          =   195
               Index           =   26
               Left            =   180
               TabIndex        =   107
               Top             =   315
               Width           =   1680
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FABRICANTE"
               Height          =   195
               Index           =   25
               Left            =   180
               TabIndex        =   106
               Top             =   675
               Width           =   1005
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESPECIFICACIONES"
               Height          =   195
               Index           =   24
               Left            =   180
               TabIndex        =   105
               Top             =   1035
               Width           =   1515
            End
         End
         Begin VB.Frame frmCT 
            BackColor       =   &H00C0C0C0&
            Height          =   3975
            Index           =   1
            Left            =   -70000
            TabIndex        =   81
            Top             =   315
            Visible         =   0   'False
            Width           =   12795
            Begin VB.TextBox txtNOMBRE_BASE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   2565
               TabIndex        =   18
               Top             =   270
               Width           =   10125
            End
            Begin VB.TextBox txtFABRICANTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   2565
               TabIndex        =   19
               Top             =   630
               Width           =   10125
            End
            Begin VB.TextBox txtESPECIFICACIONES 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   2565
               TabIndex        =   20
               Top             =   990
               Width           =   10125
            End
            Begin VB.TextBox txtN_LATAS 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   2565
               TabIndex        =   21
               Top             =   1350
               Width           =   10125
            End
            Begin VB.TextBox txtCANTIDAD_LATA 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   2565
               TabIndex        =   22
               Top             =   1710
               Width           =   10125
            End
            Begin VB.TextBox txtLOTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   2565
               TabIndex        =   23
               Top             =   2070
               Width           =   10125
            End
            Begin VB.TextBox txtCONDICIONES_ALMACENAMIENTO 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   3600
               TabIndex        =   28
               Top             =   3150
               Width           =   9090
            End
            Begin VB.TextBox txtESTADO_EMBALAJE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   1
               Left            =   3600
               TabIndex        =   29
               Top             =   3510
               Width           =   9090
            End
            Begin VB.CheckBox chkFechaEnvasado 
               Caption         =   "Check1"
               Height          =   195
               Index           =   1
               Left            =   2565
               TabIndex        =   24
               Top             =   2475
               Width           =   240
            End
            Begin VB.CheckBox chkFechaCaducidad 
               Caption         =   "Check1"
               Height          =   195
               Index           =   1
               Left            =   2565
               TabIndex        =   26
               Top             =   2835
               Width           =   240
            End
            Begin MSComCtl2.DTPicker fechaCaducidad 
               Height          =   330
               Index           =   1
               Left            =   2835
               TabIndex        =   27
               Top             =   2790
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fechaEnvasado 
               Height          =   330
               Index           =   1
               Left            =   2835
               TabIndex        =   25
               Top             =   2430
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESPECIFICACIONES"
               Height          =   195
               Index           =   14
               Left            =   180
               TabIndex        =   91
               Top             =   1035
               Width           =   1515
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FABRICANTE"
               Height          =   195
               Index           =   15
               Left            =   180
               TabIndex        =   90
               Top             =   675
               Width           =   1005
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "NOMBRE DE LA BASE"
               Height          =   195
               Index           =   16
               Left            =   180
               TabIndex        =   89
               Top             =   315
               Width           =   1680
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nº DE LATAS / BOTES"
               Height          =   195
               Index           =   17
               Left            =   180
               TabIndex        =   88
               Top             =   1395
               Width           =   1710
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CANTIDAD POR LATA / BOTE"
               Height          =   195
               Index           =   18
               Left            =   180
               TabIndex        =   87
               Top             =   1755
               Width           =   2265
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "LOTE"
               Height          =   195
               Index           =   19
               Left            =   180
               TabIndex        =   86
               Top             =   2115
               Width           =   420
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE ENVASADO"
               Height          =   195
               Index           =   20
               Left            =   180
               TabIndex        =   85
               Top             =   2475
               Width           =   1725
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE CADUCIDAD"
               Height          =   195
               Index           =   21
               Left            =   180
               TabIndex        =   84
               Top             =   2835
               Width           =   1785
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CONDICIONES DE ALMACENAMIENTO"
               Height          =   195
               Index           =   22
               Left            =   180
               TabIndex        =   83
               Top             =   3195
               Width           =   2925
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESTADO DE EMBALAJE Y DEL CONTENIDO"
               Height          =   195
               Index           =   23
               Left            =   180
               TabIndex        =   82
               Top             =   3555
               Width           =   3315
            End
         End
         Begin VB.Frame frmCT 
            BackColor       =   &H00C0C0C0&
            Height          =   3975
            Index           =   0
            Left            =   0
            TabIndex        =   70
            Top             =   315
            Width           =   12795
            Begin VB.TextBox txtNOMBRE_BASE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   2565
               TabIndex        =   4
               Top             =   270
               Width           =   10125
            End
            Begin VB.TextBox txtFABRICANTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   2565
               TabIndex        =   5
               Top             =   630
               Width           =   10125
            End
            Begin VB.TextBox txtESPECIFICACIONES 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   2565
               TabIndex        =   6
               Top             =   990
               Width           =   10125
            End
            Begin VB.TextBox txtN_LATAS 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   2565
               TabIndex        =   7
               Top             =   1350
               Width           =   10125
            End
            Begin VB.TextBox txtCANTIDAD_LATA 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   2565
               TabIndex        =   8
               Top             =   1710
               Width           =   10125
            End
            Begin VB.TextBox txtLOTE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   2565
               TabIndex        =   9
               Top             =   2070
               Width           =   10125
            End
            Begin VB.TextBox txtCONDICIONES_ALMACENAMIENTO 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   3600
               TabIndex        =   14
               Top             =   3150
               Width           =   9090
            End
            Begin VB.TextBox txtESTADO_EMBALAJE 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   0
               Left            =   3600
               TabIndex        =   15
               Top             =   3510
               Width           =   9090
            End
            Begin VB.CheckBox chkFechaEnvasado 
               Caption         =   "Check1"
               Height          =   195
               Index           =   0
               Left            =   2565
               TabIndex        =   10
               Top             =   2475
               Width           =   240
            End
            Begin VB.CheckBox chkFechaCaducidad 
               Caption         =   "Check1"
               Height          =   195
               Index           =   0
               Left            =   2565
               TabIndex        =   12
               Top             =   2835
               Width           =   240
            End
            Begin MSComCtl2.DTPicker fechaCaducidad 
               Height          =   330
               Index           =   0
               Left            =   2835
               TabIndex        =   13
               Top             =   2790
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fechaEnvasado 
               Height          =   330
               Index           =   0
               Left            =   2835
               TabIndex        =   11
               Top             =   2430
               Width           =   1290
               _ExtentX        =   2275
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
               Format          =   51642369
               CurrentDate     =   38000
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESPECIFICACIONES"
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   80
               Top             =   1035
               Width           =   1515
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FABRICANTE"
               Height          =   195
               Index           =   5
               Left            =   180
               TabIndex        =   79
               Top             =   675
               Width           =   1005
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "NOMBRE DE LA BASE"
               Height          =   195
               Index           =   7
               Left            =   180
               TabIndex        =   78
               Top             =   315
               Width           =   1680
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nº DE LATAS / BOTES"
               Height          =   195
               Index           =   4
               Left            =   180
               TabIndex        =   77
               Top             =   1395
               Width           =   1710
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CANTIDAD POR LATA / BOTE"
               Height          =   195
               Index           =   8
               Left            =   180
               TabIndex        =   76
               Top             =   1755
               Width           =   2265
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "LOTE"
               Height          =   195
               Index           =   9
               Left            =   180
               TabIndex        =   75
               Top             =   2115
               Width           =   420
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE ENVASADO"
               Height          =   195
               Index           =   10
               Left            =   180
               TabIndex        =   74
               Top             =   2475
               Width           =   1725
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "FECHA DE CADUCIDAD"
               Height          =   195
               Index           =   11
               Left            =   180
               TabIndex        =   73
               Top             =   2835
               Width           =   1785
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "CONDICIONES DE ALMACENAMIENTO"
               Height          =   195
               Index           =   12
               Left            =   180
               TabIndex        =   72
               Top             =   3195
               Width           =   2925
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "ESTADO DE EMBALAJE Y DEL CONTENIDO"
               Height          =   195
               Index           =   13
               Left            =   180
               TabIndex        =   71
               Top             =   3555
               Width           =   3315
            End
         End
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   2280
      Left            =   45
      TabIndex        =   63
      Top             =   7155
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   4022
      Caption         =   "Estado de las Latas a la Entrada"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   2280
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         Index           =   10
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   495
         Width           =   4905
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   330
         Left            =   2880
         TabIndex        =   65
         Top             =   2655
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   51642369
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   330
         Left            =   2880
         TabIndex        =   64
         Top             =   3015
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   51642369
         CurrentDate     =   38000
         MinDate         =   2
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP3 
      Height          =   2280
      Left            =   5175
      TabIndex        =   66
      Top             =   7155
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   4022
      Caption         =   "Marcado de las latas"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   2280
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   330
         Left            =   2880
         TabIndex        =   68
         Top             =   3015
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   51642369
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   330
         Left            =   2880
         TabIndex        =   67
         Top             =   2655
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   51642369
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         Index           =   11
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   495
         Width           =   4995
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción Administrativa de PINTURAS"
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
      Height          =   360
      Left            =   135
      TabIndex        =   56
      Top             =   75
      Width           =   5610
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   600
      Left            =   -45
      Top             =   -45
      Width           =   14085
   End
End
Attribute VB_Name = "frmPinturasRecepcionAdministrativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private Sub chkFechaCaducidad_Click(Index As Integer)
    If chkFechaCaducidad(Index).Value = Checked Then
        fechaCaducidad(Index).Enabled = True
        fechaCaducidad(Index).Value = Date
    Else
        fechaCaducidad(Index).Enabled = False
    End If
End Sub

Private Sub chkFechaEnvasado_Click(Index As Integer)
    If chkFechaEnvasado(Index).Value = Checked Then
        fechaEnvasado(Index).Enabled = True
        fechaEnvasado(Index).Value = Date
    Else
        fechaEnvasado(Index).Enabled = False
    End If

End Sub

Private Sub cmbClientes_change()
    cmbOfertas.desactivar
    If cmbClientes.getTEXTO <> "" Then
        cmbOfertas.activar
        cargarOfertas (cmbClientes.getPK_SALIDA)
        cargar_pedidos CLng(cmbClientes.getPK_SALIDA), fecha.Value
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    If validar = True Then
        Me.MousePointer = 11
        Dim oPR As New clsPinturas_radmin
        Dim RECEPCION As Long
        Dim i As Integer
        With oPR
            .setCLIENTE_ID = cmbClientes.getPK_SALIDA
            If cmbOfertas.getTEXTO = "" Then
                .setOFERTA_ID = 0
            Else
                .setOFERTA_ID = cmbOfertas.getPK_SALIDA
            End If
            If cmbPedido.BoundText = "" Then
                .setPEDIDO_ID = 0
            Else
                .setPEDIDO_ID = cmbPedido.BoundText
            End If
            .setFECHA_RECEPCION = Format(fecha, "yyyy-mm-dd")
            .setESTADOS_LATAS = txtDatos(10)
            .setMARCADO_LATAS = txtDatos(11)
            .setESTADO_ID = PINTURAS_ESTADOS.PINTURAS_PENDIENTE
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            If PK = 0 Then
                RECEPCION = .Insertar
            Else
                .Modificar (PK)
                RECEPCION = PK
            End If
        End With
        Dim oPRC As New clsPinturas_radmin_car
        oPRC.Eliminar RECEPCION
        For i = 0 To 3
            With oPRC
                .setNOMBRE_BASE = txtNOMBRE_BASE(i)
                .setFABRICANTE = txtFABRICANTE(i)
                .setESPECIFICACIONES = txtESPECIFICACIONES(i)
                .setN_LATAS = txtN_LATAS(i)
                .setCANTIDAD_LATA = txtCANTIDAD_LATA(i)
                .setLOTE = txtLOTE(i)
                If chkFechaEnvasado(i).Value = Checked Then
                    .setFECHA_ENVASADO = Format(fechaEnvasado(i), "yyyy-mm-dd")
                Else
                    .setFECHA_ENVASADO = 0
                End If
                If chkFechaCaducidad(i).Value = Checked Then
                    .setFECHA_CADUCIDAD = Format(fechaCaducidad(i), "yyyy-mm-dd")
                Else
                    .setFECHA_CADUCIDAD = 0
                End If
                
                .setCONDICIONES_ALMACENAMIENTO = txtCONDICIONES_ALMACENAMIENTO(i)
                .setESTADO_EMBALAJE = txtESTADO_EMBALAJE(i)
                .Insertar
            End With
        Next
        MsgBox "La RECEPCION se ha almacenado correctamente.", vbInformation + vbOKOnly, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPinturasRecepcionAdministrativa")
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
'    Me.SetFocus
End Sub

Private Sub Form_Load()
    Me.Left = 50
    Me.top = 50
    log (Me.Name)
    cargar_botones Me
    cmbOfertas.desactivar
    cargar_combos
    fecha = Date
    If PK <> 0 Then
        cargarRecepcion PK
    End If
End Sub
Public Function validar() As Boolean
   On Error GoTo validar_Error

    validar = True
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Debe indicar el CLIENTE.", vbExclamation, App.Title
        cmbClientes.SetFocus
        validar = False
        Exit Function
    End If
   On Error GoTo 0
   Exit Function

validar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validar of Formulario frmPinturasRecepcionAdministrativa"
End Function
Private Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, " ANULADO = 0 "
End Sub
Private Sub Image1_Click()
    cmbPedido.Text = ""
    cmbPedido.BoundText = ""
End Sub

Private Sub imgPedidos_Click()
    If cmbClientes.getTEXTO <> "" Then
        frmClientes_Pedidos.PK = cmbClientes.getPK_SALIDA
        frmClientes_Pedidos.Show 1
        cargar_pedidos CLng(cmbClientes.getPK_SALIDA), fecha.Value
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80FFFF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim oPedido As New clsClientes_pedidos
    Set cmbPedido.RowSource = oPedido.Listado_en_fecha(CInt(cliente), CStr(fecha))
    cmbPedido.ListField = "CODIGO_LARGO"
    cmbPedido.DataField = "ID_PEDIDO"
    cmbPedido.BoundColumn = "ID_PEDIDO"
End Sub

Private Sub cargarOfertas(cliente As Long)
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT A.ID_OFERTA, CONCAT('NUMERO : ', A.NUMERO, ', FECHA : ',A.FECHA, A.DESCRIPCION) " & _
                   "  FROM OFERTAS A,CLIENTES B " & _
                   " Where A.CLIENTE_ID = B.ID_CLIENTE " & _
                   "   AND A.CLIENTE_ID = " & cliente & _
                   "   AND A.ESTADO_OFERTA <> " & OFERTAS_ESTADOS.OFERTAS_ESTADOS_ANULADA & _
                   "   AND A.ESTADO_OFERTA <> " & OFERTAS_ESTADOS.OFERTAS_ESTADOS_RECHAZADA
        With cmbOfertas
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "OFERTAS"
            .setDESCRIPCION = "Ofertas"
            .setPK = "A.ID_OFERTA"
            .setCAMPO = "CONCAT('NUMERO : ', A.NUMERO, ', FECHA : ',A.FECHA, A.DESCRIPCION)"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmOferta_Nueva2
        End With
    End If
End Sub

Private Sub cargarRecepcion(ID As Long)
    Dim oPR As New clsPinturas_radmin
   On Error GoTo cargarRecepcion_Error
    
    Dim i As Integer
    Dim oPRC As New clsPinturas_radmin_car
    With oPR
        If .Carga(ID) Then
            cmbClientes.MostrarElemento .getCLIENTE_ID
            cmbOfertas.MostrarElemento .getOFERTA_ID
            cmbPedido.BoundText = .getPEDIDO_ID
            fecha = .getFECHA_RECEPCION
            
            For i = 0 To 3
                If oPRC.Carga(ID, i) = True Then
                    txtNOMBRE_BASE(i) = oPRC.getNOMBRE_BASE
                    txtFABRICANTE(i) = oPRC.getFABRICANTE
                    txtESPECIFICACIONES(i) = oPRC.getESPECIFICACIONES
                    txtN_LATAS(i) = oPRC.getN_LATAS
                    txtCANTIDAD_LATA(i) = oPRC.getCANTIDAD_LATA
                    txtLOTE(i) = oPRC.getLOTE
                    If oPRC.getFECHA_ENVASADO <> "" Then
                        chkFechaEnvasado(i).Value = Checked
                        fechaEnvasado(i) = oPRC.getFECHA_ENVASADO
                    End If
                    If oPRC.getFECHA_CADUCIDAD <> "" Then
                        chkFechaCaducidad(i).Value = Checked
                        fechaCaducidad(i) = oPRC.getFECHA_CADUCIDAD
                    End If
                    txtDatos(8) = oPRC.getCONDICIONES_ALMACENAMIENTO
                    txtDatos(9) = oPRC.getESTADO_EMBALAJE
                End If
            Next
            txtDatos(10) = .getESTADOS_LATAS
            txtDatos(11) = .getMARCADO_LATAS
            
        End If
    End With
    Set oPR = Nothing

   On Error GoTo 0
   Exit Sub

cargarRecepcion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarRecepcion of Formulario frmPinturasRecepcionAdministrativa"
End Sub
