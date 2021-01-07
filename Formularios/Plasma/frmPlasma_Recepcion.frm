VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmPlasma_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Ensayos Físicos IBERIA"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "frmPlasma_Recepcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   12600
   Begin XtremeSuiteControls.Resizer Resizer1 
      Height          =   7665
      Left            =   45
      TabIndex        =   27
      Top             =   405
      Width           =   12480
      _Version        =   851970
      _ExtentX        =   22013
      _ExtentY        =   13520
      _StockProps     =   1
      VScrollLargeChange=   1000
      VScrollSmallChange=   200
      VScrollMaximum  =   6500
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "TESTING SPECIMEN"
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
         Height          =   660
         Left            =   810
         TabIndex        =   52
         Top             =   -315
         Visible         =   0   'False
         Width           =   12120
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   41
            Left            =   7425
            MaxLength       =   255
            TabIndex        =   23
            Top             =   270
            Width           =   4425
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   31
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   22
            Top             =   270
            Width           =   3795
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "BATCH TOP COAT:"
            Height          =   195
            Index           =   25
            Left            =   5895
            TabIndex        =   54
            Top             =   315
            Width           =   1440
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "BATCH BOND COAT:"
            Height          =   195
            Index           =   18
            Left            =   90
            TabIndex        =   53
            Top             =   315
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "SAMPLE DATA"
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
         Height          =   2865
         Left            =   45
         TabIndex        =   42
         Top             =   90
         Width           =   12120
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   8460
            Picture         =   "frmPlasma_Recepcion.frx":2AFA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   58
            Top             =   2430
            Width           =   240
         End
         Begin VB.OptionButton opUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   10290
            TabIndex        =   56
            Top             =   2430
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton opUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   9690
            TabIndex        =   55
            Top             =   2430
            Width           =   615
         End
         Begin MSComCtl2.DTPicker fecha 
            Height          =   330
            Left            =   1170
            TabIndex        =   2
            Top             =   945
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
            Left            =   1170
            TabIndex        =   1
            Top             =   600
            Width           =   9780
            _ExtentX        =   17251
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
         Begin MSComCtl2.DTPicker horaRecepcion 
            Height          =   330
            Left            =   3510
            TabIndex        =   3
            Top             =   945
            Width           =   1170
            _ExtentX        =   2064
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
            Format          =   16515074
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin pryCombo.miCombo cmbclientes 
            Height          =   345
            Left            =   1170
            TabIndex        =   0
            Top             =   225
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   609
         End
         Begin MSDataListLib.DataCombo cmbenvases 
            Height          =   315
            Left            =   1170
            TabIndex        =   5
            Top             =   1305
            Width           =   9765
            _ExtentX        =   17224
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
         Begin MSDataListLib.DataCombo cmbentregada 
            Height          =   315
            Left            =   1170
            TabIndex        =   6
            Top             =   1665
            Width           =   9765
            _ExtentX        =   17224
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
            Left            =   1170
            TabIndex        =   7
            Top             =   2025
            Width           =   9765
            _ExtentX        =   17224
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
         Begin MSComCtl2.DTPicker fechaMuestreo 
            Height          =   330
            Left            =   9395
            TabIndex        =   4
            Top             =   945
            Width           =   1560
            _ExtentX        =   2752
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
            Bindings        =   "frmPlasma_Recepcion.frx":934C
            Height          =   315
            Left            =   1170
            TabIndex        =   8
            Top             =   2385
            Width           =   3495
            _ExtentX        =   6165
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
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "URGENTE"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   40
            Left            =   8775
            TabIndex        =   57
            Top             =   2475
            Width           =   780
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora"
            Height          =   195
            Index           =   13
            Left            =   3045
            TabIndex        =   51
            Top             =   1020
            Width           =   345
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pedido"
            Height          =   195
            Index           =   12
            Left            =   75
            TabIndex        =   50
            Top             =   645
            Width           =   495
         End
         Begin VB.Image Image1 
            Height          =   300
            Left            =   11045
            Picture         =   "frmPlasma_Recepcion.frx":9392
            Stretch         =   -1  'True
            Top             =   570
            Width           =   255
         End
         Begin VB.Image imgPedidos 
            Height          =   300
            Left            =   11700
            Picture         =   "frmPlasma_Recepcion.frx":9C5C
            Stretch         =   -1  'True
            Top             =   585
            Width           =   315
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cliente"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   49
            Top             =   300
            Width           =   480
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Recepción"
            Height          =   195
            Index           =   6
            Left            =   75
            TabIndex        =   48
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Envase"
            Height          =   195
            Index           =   5
            Left            =   75
            TabIndex        =   47
            Top             =   1365
            Width           =   540
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Entregada por"
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   46
            Top             =   1725
            Width           =   1005
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tomada por"
            Height          =   195
            Index           =   7
            Left            =   75
            TabIndex        =   45
            Top             =   2070
            Width           =   855
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "F. Muestreo"
            Height          =   195
            Index           =   1
            Left            =   8325
            TabIndex        =   44
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Centro"
            Height          =   195
            Index           =   22
            Left            =   90
            TabIndex        =   43
            Top             =   2430
            Width           =   465
         End
      End
      Begin VB.Frame DECODIFICADORA_PLASMA_PRODUCT_TYPE 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "SPECIMEN ID AND DESCRIPTION"
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
         Height          =   3000
         Left            =   45
         TabIndex        =   32
         Top             =   4560
         Width           =   12120
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   345
            Left            =   11655
            TabIndex        =   60
            Top             =   1305
            Width           =   420
            _Version        =   851970
            _ExtentX        =   741
            _ExtentY        =   609
            _StockProps     =   79
            Appearance      =   5
            Picture         =   "frmPlasma_Recepcion.frx":A526
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   14
            Top             =   990
            Width           =   2895
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   17
            Top             =   2025
            Width           =   2895
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   19
            Top             =   2340
            Width           =   2895
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   7290
            MaxLength       =   255
            TabIndex        =   20
            Top             =   2385
            Width           =   3615
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1710
            MaxLength       =   255
            TabIndex        =   21
            Top             =   2655
            Width           =   2895
         End
         Begin pryCombo.miCombo cmbProcess 
            Height          =   345
            Left            =   1710
            TabIndex        =   12
            Top             =   270
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbCustomer 
            Height          =   345
            Left            =   1710
            TabIndex        =   13
            Top             =   630
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbProductType 
            Height          =   345
            Left            =   7290
            TabIndex        =   18
            Top             =   2025
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbnatype 
            Height          =   345
            Left            =   1710
            TabIndex        =   16
            Top             =   1665
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbDenominacion 
            Height          =   345
            Left            =   1710
            TabIndex        =   15
            Top             =   1305
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   609
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "DENOMINACIÓN"
            Height          =   195
            Index           =   21
            Left            =   135
            TabIndex        =   59
            Top             =   1350
            Width           =   1260
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "S/N:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   41
            Top             =   2385
            Width           =   345
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "SPECIMEN ID. (O.R.)"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   40
            Top             =   1035
            Width           =   1545
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "P/N:"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   39
            Top             =   2085
            Width           =   345
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PROCESS"
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   38
            Top             =   345
            Width           =   765
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "CUSTOMER"
            Height          =   195
            Index           =   14
            Left            =   135
            TabIndex        =   37
            Top             =   705
            Width           =   915
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PRODUCT S/N:"
            Height          =   195
            Index           =   15
            Left            =   5940
            TabIndex        =   36
            Top             =   2430
            Width           =   1185
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PRODUCT TYPE:"
            Height          =   195
            Index           =   16
            Left            =   5940
            TabIndex        =   35
            Top             =   2085
            Width           =   1305
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "MODULE S/N:"
            Height          =   195
            Index           =   17
            Left            =   135
            TabIndex        =   34
            Top             =   2700
            Width           =   1080
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nº AND TYPE"
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   33
            Top             =   1710
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST AND REPLACEMENT"
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
         Height          =   1470
         Left            =   45
         TabIndex        =   28
         Top             =   3030
         Width           =   12120
         Begin pryCombo.miCombo cmbTest 
            Height          =   345
            Left            =   1350
            TabIndex        =   9
            Top             =   270
            Width           =   10620
            _ExtentX        =   18733
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbReplacement 
            Height          =   345
            Left            =   1350
            TabIndex        =   10
            Top             =   630
            Width           =   10620
            _ExtentX        =   18733
            _ExtentY        =   609
         End
         Begin MSDataListLib.DataCombo cmbProducto 
            Height          =   315
            Left            =   1350
            TabIndex        =   11
            Top             =   990
            Width           =   9575
            _ExtentX        =   16880
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
            Caption         =   "REPLACEMENT"
            Height          =   195
            Index           =   19
            Left            =   90
            TabIndex        =   31
            Top             =   705
            Width           =   1200
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "TEST"
            Height          =   195
            Index           =   20
            Left            =   90
            TabIndex        =   30
            Top             =   345
            Width           =   420
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PRODUCT DES."
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   29
            Top             =   1035
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8085
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8085
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Ensayos Físicos IBERIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   90
      TabIndex        =   26
      Top             =   0
      Width           =   5280
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   390
      Left            =   0
      Top             =   0
      Width           =   12600
   End
End
Attribute VB_Name = "frmPlasma_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTest_change()
    cmbProcess.limpiar
    If cmbTest.getTEXTO <> "" Then
        llenar_combo cmbProcess, New clsPlasma_procesos, 0, frmPlasma_Procesos_Detalle, " TIPO = " & cmbTest.getPK_SALIDA
        cargar_producto
    End If
End Sub
Private Sub cargar_producto()
    If cmbTest.getTEXTO <> "" Then
        Dim odeco As New clsDecodificadora
        odeco.Carga_valor DECODIFICADORA.IBERIA_ENSAYOS_FISICOS, cmbTest.getPK_SALIDA
        cadena = Split(odeco.getPARAMETROS, ";")
        
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT VALOR, DESCRIPCION " & _
                   "  FROM decodificadora " & _
                   " WHERE CODIGO = " & DECODIFICADORA.DESCRIPCION_PRODUCTO & _
                   "   AND PARAMETROS = '" & CInt(cadena(0)) & "'"
        Set rs = datos_bd(consulta)
        cmbProducto.Text = ""
        Set cmbProducto.RowSource = rs
        cmbProducto.ListField = "DESCRIPCION" 'lo que enseña
        cmbProducto.DataField = "VALOR" 'campo asociado
        cmbProducto.BoundColumn = "VALOR" 'lo que realmente
    End If
End Sub

Private Sub Image1_Click()
    cmbPedido.Text = ""
    cmbPedido.BoundText = ""
End Sub

Private Sub imgPedidos_Click()
    If cmbCliente.Text <> "" Then
        frmClientes_Pedidos.PK = cmbCliente.BoundText
        frmClientes_Pedidos.Show 1
        cargar_pedidos
    End If
End Sub
Private Sub cmbCliente_Change()
    cargar_pedidos
End Sub
Private Sub cargar_pedidos()
    If cmbCliente.Text <> "" Then
        Dim oPedido As New clsClientes_pedidos
        Set cmbPedido.RowSource = oPedido.Listado_en_fecha(cmbCliente.BoundText, CStr(Date))
        cmbPedido.ListField = "CODIGO_LARGO"
        cmbPedido.DataField = "ID_PEDIDO"
        cmbPedido.BoundColumn = "ID_PEDIDO"
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        ' confirmar si es replacement
        If cmbReplacement.getPK_SALIDA <> PLASMA_REPLACEMENT.PR_ENSAYO Then
            If MsgBox("¿Esta seguro de dar de alta un " & cmbReplacement.getTEXTO & "?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
        Me.MousePointer = 11
        ' Generamos el registro de las muestras
        Dim oParametros As New clsParametros
        Dim odeco As New clsDecodificadora
        Dim oMuestra As New clsMuestra
        Dim MUESTRA As Long
        Dim cadena() As String
        With oMuestra
'            oParametros.Carga parametros.PARAM_PLASMA_TM, ""
'            oParametros.Carga parametros.PARAM_PLASMA_TA, ""
'            .setTIPO_MUESTRA_ID = oParametros.getVALOR
'            .setTIPO_ANALISIS_ID = oParametros.getVALOR
            odeco.Carga_valor DECODIFICADORA.IBERIA_ENSAYOS_FISICOS, cmbTest.getPK_SALIDA
            cadena = Split(odeco.getPARAMETROS, ";")
            .setTIPO_MUESTRA_ID = cadena(0)
            .setTIPO_ANALISIS_ID = cadena(1)
            .setANALISIS_MODIFICADO = tipo_especial.PLASMA
            .setFECHA_MUESTREO = Format(fechaMuestreo, "yyyy-mm-dd")
            .setENTIDAD_MUESTREO_ID = cmbrealizada.BoundText
            .setDETALLE_MUESTREO = ""
            .setOBSERVACIONES_MUESTREO = ""
            .setFECHA_RECEPCION = Format(fecha, "yyyy-mm-dd")
            .setHORA_RECEPCION = Format(Time, "hh:mm")
            .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
            .setFORMATO_ID = cmbenvases.BoundText
            .setENTIDAD_ENTREGA_ID = cmbentregada.BoundText
            .setDETALLE_ENTREGA = ""
            .setOBSERVACIONES_ENTREGA = ""
            .setCLIENTE_ID = cmbClientes.getPK_SALIDA
            .setCENTRO_ID = cmbCentro.BoundText
            Dim ref As String
            ref = "OR: " & txtDatos(0)
            If cmbDenominacion.getTEXTO <> "" Then
                ref = ref & " - " & cmbDenominacion.getTEXTO
            End If
            .setREFERENCIA_CLIENTE = ref 'SPECIMEN_ID
            ' Fecha Prevista de Entrega
            Dim FechaEntrega As Date
            If oParametros.Carga(parametros.PARAM_PLASMA_HORAS, "") Then
                FechaEntrega = DateAdd("h", oParametros.getVALOR, fecha)
            Else
               Dim oTA As New clsTipos_analisis
               oTA.CARGAR .getTIPO_ANALISIS_ID
               FechaEntrega = DateAdd("d", oTA.getDIAS_TRABAJO, fecha)
            End If
            .setFECHA_PREV_FIN = Format(FechaEntrega, "yyyy-mm-dd")
            ' Resto de campos
            .setOBSERVACIONES = ""
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
            .setPRODUCTO = cmbProducto.Text
            If cmbPedido.Text = "" Then
                .setPEDIDO_ID = 0
            Else
                .setPEDIDO_ID = cmbPedido.BoundText
            End If
            .setREPLACEMENT_ID = cmbReplacement.getPK_SALIDA
            If opUrgente(0).Value = True Then
                .setURGENTE = 0
            Else
                .setURGENTE = 1
            End If
            MUESTRA = .guardarMuestra
'            .informar_precio_muestra MUESTRA
        End With
        ' PLASMA_RECEPCION
        Dim oPlasma_recepcion As New clsPlasma_recepcion
        With oPlasma_recepcion
            .setMUESTRA_ID = MUESTRA
            .setPROCESO_ID = cmbProcess.getPK_SALIDA
            .setCUSTOMER_ID = cmbCustomer.getPK_SALIDA
'            .setSPECIMEN_ID = txtDatos(0) & " - " & cmbDenominacion.getTEXTO
            .setSPECIMEN_ID = ref
            .setNTYPE = cmbnatype.getPK_SALIDA
            .setPN = txtDatos(1)
            
            .setPRODUCT_TYPE = cmbProductType.getTEXTO
            
            .setSN = txtDatos(3)
            .setPRODUCT_SN = txtDatos(4)
            .setMODULE_SN = txtDatos(5)
            .setBOND_CONTROL_SPECIFICATION = ""
            .setTOP_CONTROL_SPECIFICATION = ""
            .setMP = 0
            .setMP_FECHA = "NULL"
            .setMP_USUARIO_ID = 0
            .setMP_PASS = 0
            .setMACRO_DUREZA_T1 = ""
            .setMACRO_DUREZA_T2 = ""
            .setMICRO_DUREZA_T1 = ""
            .setMICRO_DUREZA_T2 = ""
            .Insertar
            .informarControlSpecification MUESTRA
            .informarEquiposRecepcion MUESTRA, cmbProcess.getPK_SALIDA
        End With
        Me.MousePointer = 0
        oMuestra.CargaMuestra MUESTRA
        Dim otm As New clsTipos_muestra
        otm.CARGAR oMuestra.getTIPO_MUESTRA_ID
        MsgBox "La recepción del Plasma se ha realizado correctamente. Número : " & oMuestra.getID_GENERAL & "/" & oMuestra.getANNO & " (" & otm.getCODIGO & "-" & oMuestra.getID_PARTICULAR & ")", vbInformation, App.Title
        If cadena(0) = TIPOS_MUESTRAS.DUREZA_ROCKWELL Or _
           cadena(0) = TIPOS_MUESTRAS.DUREZA_BRINELL Or _
           cadena(0) = TIPOS_MUESTRAS.DUREZA_VICKERS Then
            With frmPlasma_Dureza
                .PK = MUESTRA
                .Show 1
            End With
        ElseIf cadena(0) = TIPOS_MUESTRAS.DUREZA_SHORE_PIEZAS Then
            With frmPlasma_Dureza_Shore
                .PK = MUESTRA
                .Show 1
            End With
        Else
            With frmPlasma_Resultados
                .PK = MUESTRA
                .Show 1
            End With
        End If
        Unload Me
    End If
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Recepcion")
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.top = 100
    
    Call cargar_combos
    fecha = Date
    fechaMuestreo = Date
    horaRecepcion = Date & " " & Format(Time, "hh:mm:ss")
    
    ' Cargar datos defecto
    Dim datos As String
    Dim lista() As String
    Dim parametros As New clsParametros
    parametros.Carga IBERIA_DATOS_DEFECTO_RECEPCION_PLASMA, ""
'    datos = "3136;2;2;2;3;1;1;93"
    datos = parametros.getVALOR
    lista = Split(datos, ";")
    If lista(0) <> "" Then
        cmbClientes.MostrarElemento CLng(lista(0))
    End If
    If lista(1) <> "" Then
        cmbenvases.BoundText = lista(1)
    End If
    If lista(2) <> "" Then
        cmbentregada.BoundText = lista(2)
    End If
    If lista(3) <> "" Then
        cmbrealizada.BoundText = lista(3)
    End If
    If lista(4) <> "" Then
        cmbCentro.BoundText = lista(4)
    End If
    If lista(5) <> "" Then
        cmbTest.MostrarElemento CLng(lista(5))
    End If
    If lista(6) <> "" Then
        cmbReplacement.MostrarElemento CLng(lista(6))
    End If
    If lista(7) <> "" Then
        cmbProducto.BoundText = lista(7)
    End If
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmPlasma_Recepcion"
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Debe seleccionar el cliente.", vbExclamation, App.Title
        cmbClientes.SetFocus
        validar = False
        Exit Function
    End If
    If cmbenvases.BoundText = "" Then
        MsgBox "Debe asignar un envase a la selección.", vbExclamation, App.Title
        cmbenvases.SetFocus
        validar = False
        Exit Function
    End If
    If cmbentregada.BoundText = "" Then
        MsgBox "Debe asignar la entrega.", vbExclamation, App.Title
        cmbentregada.SetFocus
        validar = False
        Exit Function
    End If
    If cmbrealizada.BoundText = "" Then
        MsgBox "Debe asignar la realización.", vbExclamation, App.Title
        cmbrealizada.SetFocus
        validar = False
        Exit Function
    End If
    If cmbProcess.getTEXTO = "" Then
        MsgBox "Debe asignar el campo PROCESS.", vbExclamation, App.Title
        cmbProcess.SetFocus
        validar = False
        Exit Function
    End If
    If cmbCustomer.getTEXTO = "" Then
        MsgBox "Debe asignar el campo PROCESS.", vbExclamation, App.Title
        cmbCustomer.SetFocus
        validar = False
        Exit Function
    End If
    If cmbCentro.Text = "" Then
        MsgBox "El CENTRO no puede estar en blanco.", vbExclamation, "Validación"
        cmbCentro.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "Debe indicar el campo SPECIMEN ID.", vbExclamation, App.Title
        validar = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If cmbnatype.getTEXTO = "" Then
        MsgBox "Debe indicar el campo N.AND TYPE.", vbExclamation, App.Title
        validar = False
        cmbnatype.SetFocus
        Exit Function
    End If
    If txtDatos(1) = "" Then
        MsgBox "Debe indicar el campo P/N.", vbExclamation, App.Title
        validar = False
        txtDatos(1).SetFocus
        Exit Function
    End If
    If cmbProductType.getTEXTO = "" Then
        MsgBox "Debe indicar el campo PRODUCT TYPE.", vbExclamation, App.Title
        validar = False
        cmbProductType.SetFocus
        Exit Function
    End If
    If txtDatos(3) = "" Then
        MsgBox "Debe indicar el campo S/N.", vbExclamation, App.Title
        validar = False
        txtDatos(3).SetFocus
        Exit Function
    End If
    If txtDatos(4) = "" Then
        MsgBox "Debe indicar el campo PRODUCT S/N.", vbExclamation, App.Title
        validar = False
        txtDatos(4).SetFocus
        Exit Function
    End If
    If txtDatos(5) = "" Then
        MsgBox "Debe indicar el campo MODULE S/N.", vbExclamation, App.Title
        validar = False
        txtDatos(5).SetFocus
        Exit Function
    End If
    If cmbTest.getTEXTO = "" Then
        MsgBox "Debe indicar el seleccionable TEST.", vbExclamation, App.Title
        validar = False
        cmbTest.SetFocus
        Exit Function
    End If
    If cmbReplacement.getTEXTO = "" Then
        MsgBox "Debe indicar el seleccionable REPLACEMENT.", vbExclamation, App.Title
        validar = False
        cmbReplacement.SetFocus
        Exit Function
    End If
    
    If cmbProducto.Text = "" Then
        MsgBox "Debe indicar el seleccionable PRODUCT DESCRIPTION.", vbExclamation, App.Title
        validar = False
        cmbProducto.SetFocus
        Exit Function
    End If
    
    If cmbDenominacion.getTEXTO = "" Then
        MsgBox "Debe indicar el seleccionable DENOMINACION.", vbExclamation, App.Title
        validar = False
        cmbDenominacion.SetFocus
        Exit Function
    End If
    
End Function

Private Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    cargar_combo cmbCentro, New clsCentros
    cargar_combo cmbenvases, New clsformatos
    cargar_combo cmbentregada, New clsEntidades_Entrega
    cargar_combo cmbrealizada, New clsEntidades_muestreo
    
    llenar_combo cmbProcess, New clsPlasma_procesos, 0, frmPlasma_Procesos_Detalle, ""
    Dim odeco As New clsDecodificadora
    odeco.cargar_mi_combo cmbCustomer, DECODIFICADORA.DECODIFICADORA_PLASMA_CLIENTES_INTERNOS
    odeco.cargar_mi_combo cmbProductType, DECODIFICADORA.DECODIFICADORA_PLASMA_PRODUCT_TYPE
    odeco.cargar_mi_combo cmbTest, DECODIFICADORA.IBERIA_ENSAYOS_FISICOS
    odeco.cargar_mi_combo cmbReplacement, DECODIFICADORA.IBERIA_REPLACEMENT
    odeco.cargar_mi_combo cmbnatype, DECODIFICADORA.DECODIFICADORA_PLASMA_NUMBER_AND_TYPE
    odeco.cargar_mi_combo cmbDenominacion, DECODIFICADORA.SPECIMEN_ID_DENOMINACION
End Sub
Public Function cargarComboDenominacion() As Boolean
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With combo
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "DECODIFICADORA"
            .setDESCRIPCION = "Decodificadora"
            .setQUERY = "SELECT VALOR,DESCRIPCION FROM decodificadora WHERE CODIGO = " & CCODTABL
            .setPK = "VALOR"
            .setCAMPO = "DESCRIPCION"
            .setMUESTRA_DETALLE = False
            .setFILTRO = ""
            Set .FORMULARIO = Nothing
        End With
    End If
    Set conn = Nothing
    
End Function

Private Sub PushButton1_Click()
    cmbDenominacion.limpiar
    Dim oform As New frmDecodificadoraModal
    oform.CODIGO = DECODIFICADORA.SPECIMEN_ID_DENOMINACION
    oform.Show 1
    Set oform = Nothing
    Dim odeco As New clsDecodificadora
    odeco.cargar_mi_combo cmbDenominacion, SPECIMEN_ID_DENOMINACION
    Set odeco = Nothing

End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80FFFF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
