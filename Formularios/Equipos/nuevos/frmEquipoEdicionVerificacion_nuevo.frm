VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoEdicionVerificacion_nuevo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8685
   ClientLeft      =   2955
   ClientTop       =   2490
   ClientWidth     =   12615
   ClipControls    =   0   'False
   Icon            =   "frmEquipoEdicionVerificacion_nuevo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3870
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11550
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7815
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Verificación"
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
      Height          =   7275
      Left            =   45
      TabIndex        =   47
      Top             =   510
      Width           =   12570
      Begin VB.CommandButton cmdAnadirParametro 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Añadir accesorio"
         Top             =   4680
         Width           =   285
      End
      Begin VB.CommandButton cmdEliminarParametro 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   7170
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":0231
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Eliminar accesorio"
         Top             =   4680
         Width           =   285
      End
      Begin VB.TextBox txtFechaProxima 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10710
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "01/01/1900"
         Top             =   630
         Width           =   1785
      End
      Begin VB.TextBox txtEvaluacionResultado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   15
         Top             =   2775
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7620
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":03C5
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ver norma"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":061A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Buscar documento"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7245
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":088B
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Eliminar documento"
         Top             =   2760
         Width           =   360
      End
      Begin VB.CommandButton cmdEscanearEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":0A1F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Escanear documento"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdEscanearHoja 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":0DD9
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Escanear documento"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdEscanearCert 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":1193
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Escanear documento"
         Top             =   2400
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5580
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":154D
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Eliminar accesorio"
         Top             =   3240
         Width           =   285
      End
      Begin VB.CommandButton cmdAnadirLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5250
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":16E1
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Añadir accesorio"
         Top             =   3240
         Width           =   285
      End
      Begin VB.TextBox txtLimitacionesUso 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1650
         MaxLength       =   100
         TabIndex        =   20
         Top             =   3210
         Width           =   3555
      End
      Begin VB.Frame fraEstadoIntervencion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resultado Verificación "
         Height          =   1245
         Left            =   10350
         TabIndex        =   35
         Top             =   1470
         Width           =   2115
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Previsto"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cerrado Conforme"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   37
            Top             =   540
            Width           =   1605
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cerrado No Conforme"
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   38
            Top             =   840
            Width           =   1875
         End
      End
      Begin VB.ListBox lstLimitacionesUso 
         Appearance      =   0  'Flat
         Height          =   1395
         ItemData        =   "frmEquipoEdicionVerificacion_nuevo.frx":1906
         Left            =   1650
         List            =   "frmEquipoEdicionVerificacion_nuevo.frx":190D
         TabIndex        =   58
         Top             =   3540
         Width           =   4215
      End
      Begin VB.TextBox txtHojaVerificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         Top             =   2055
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarHojaCal 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7620
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":1925
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Ver norma"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarHojaCal 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":1B7A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar documento"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7245
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":1DEB
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Eliminar documento"
         Top             =   2400
         Width           =   360
      End
      Begin VB.CommandButton cmdAdjuntarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":1F7F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Buscar documento"
         Top             =   2400
         Width           =   405
      End
      Begin VB.CommandButton cmdMostrarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7620
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":21F0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ver norma"
         Top             =   2400
         Width           =   405
      End
      Begin VB.TextBox txtCertificado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2415
         Width           =   4770
      End
      Begin MSComCtl2.DTPicker txtFechaActual 
         Height          =   405
         Left            =   10710
         TabIndex        =   33
         Top             =   180
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   78249985
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFechaProxima_b 
         Height          =   405
         Left            =   10710
         TabIndex        =   46
         Top             =   1035
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   78249985
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoVerificacion 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   270
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPeriVerificacion 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   630
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbVerificador 
         Height          =   330
         Left            =   1650
         TabIndex        =   2
         Top             =   990
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbProcedimiento 
         Height          =   330
         Left            =   1650
         TabIndex        =   4
         Top             =   1710
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVerificadorExterno 
         Height          =   330
         Left            =   1650
         TabIndex        =   3
         Top             =   1350
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdEliminarHojaCal 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7245
         Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":2445
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Eliminar documento"
         Top             =   2040
         Width           =   360
      End
      Begin VB.Frame fraTipoParametro 
         BackColor       =   &H00C0C0C0&
         Height          =   2325
         Index           =   0
         Left            =   7530
         TabIndex        =   61
         Top             =   4890
         Visible         =   0   'False
         Width           =   4965
         Begin VB.TextBox txtDescripcion_Cualidad 
            Appearance      =   0  'Flat
            Height          =   975
            Left            =   60
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   540
            Width           =   4845
         End
         Begin VB.OptionButton optResultadoCualidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO CONFORME"
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
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   1890
            Width           =   3585
         End
         Begin MSDataListLib.DataCombo cmbTipoParametro 
            Height          =   315
            Index           =   0
            Left            =   2820
            TabIndex        =   23
            Top             =   180
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.OptionButton optResultadoCualidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CONFORME"
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
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   25
            Top             =   1560
            Value           =   -1  'True
            Width           =   3585
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   74
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   12
            Left            =   2370
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame fraTipoParametro 
         BackColor       =   &H00C0C0C0&
         Height          =   2325
         Index           =   2
         Left            =   7530
         TabIndex        =   68
         Top             =   4890
         Visible         =   0   'False
         Width           =   4965
         Begin VB.TextBox txtValor 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   44
            Top             =   1260
            Width           =   1605
         End
         Begin VB.TextBox txtNMedidas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   43
            Text            =   "1"
            Top             =   900
            Width           =   480
         End
         Begin pryCombo.miCombo cmbReactivos 
            Height          =   375
            Left            =   810
            TabIndex        =   32
            Top             =   150
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   661
         End
         Begin MSDataListLib.DataCombo cmbTipoParametro 
            Height          =   315
            Index           =   2
            Left            =   810
            TabIndex        =   42
            Top             =   540
            Visible         =   0   'False
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComctlLib.ListView lista_medidas 
            Height          =   1725
            Index           =   2
            Left            =   2790
            TabIndex        =   45
            Top             =   540
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   3043
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   12713983
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Resultados Medida"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label lblValor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   72
            Top             =   1320
            Width           =   450
         End
         Begin VB.Label Reactivo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reactivo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   71
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nº Medidas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   70
            Top             =   930
            Width           =   960
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   90
            TabIndex        =   69
            Top             =   600
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.Frame fraTipoParametro 
         BackColor       =   &H00C0C0C0&
         Height          =   2325
         Index           =   1
         Left            =   7530
         TabIndex        =   63
         Top             =   4890
         Visible         =   0   'False
         Width           =   4965
         Begin MSComctlLib.ListView lista_medidas 
            Height          =   1725
            Index           =   1
            Left            =   2790
            TabIndex        =   31
            Top             =   540
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   3043
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   12713983
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Resultados Medida"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox txtValor 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   1230
            Width           =   1635
         End
         Begin VB.TextBox txtNMedidas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   29
            Text            =   "1"
            Top             =   900
            Width           =   480
         End
         Begin pryCombo.miCombo cmbEquipos 
            Height          =   375
            Left            =   810
            TabIndex        =   27
            Top             =   150
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   661
         End
         Begin MSDataListLib.DataCombo cmbTipoParametro 
            Height          =   315
            Index           =   1
            Left            =   810
            TabIndex        =   28
            Top             =   540
            Visible         =   0   'False
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblValor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Valor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   73
            Top             =   1290
            Width           =   450
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   66
            Top             =   600
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nº Medidas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   65
            Top             =   930
            Width           =   960
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Equipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   64
            Top             =   195
            Width           =   570
         End
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2220
         Left            =   30
         TabIndex        =   77
         Top             =   5010
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   3916
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblParametro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7530
         TabIndex        =   67
         Top             =   4590
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Eval. Resultado"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   60
         Top             =   2835
         Width           =   1125
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Verificador Externo"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   59
         Top             =   1395
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Limitaciones uso"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   57
         Top             =   3285
         Width           =   1200
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hoja de Verificación"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   56
         Top             =   2130
         Width           =   1380
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cert. de verificación"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   55
         Top             =   2475
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo Verificación"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   54
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próx. Verificación"
         Height          =   195
         Index           =   0
         Left            =   9150
         TabIndex        =   53
         Top             =   780
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Actual Verificación"
         Height          =   195
         Index           =   10
         Left            =   9150
         TabIndex        =   52
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Ver. Interna"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   51
         Top             =   1050
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   49
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   48
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10470
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7815
      Width           =   1050
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12045
      Picture         =   "frmEquipoEdicionVerificacion_nuevo.frx":25D9
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Verificación de Equipo"
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
      TabIndex        =   50
      Top             =   120
      Width           =   2325
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmEquipoEdicionVerificacion_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarlngPK As Long
Public booSilencioso As Boolean
Private mvarobjEquipo As clsEquipos
Private mvarenuTipoEdicion As enumTipoEdicion
Private mvarstrId As String
Private bln_cambiando_tipo As Boolean

Private WithEvents TecladoNumerico As frmTecladoNumerico
Attribute TecladoNumerico.VB_VarHelpID = -1
Private blnEsTablet As Boolean
Private blnPrimeraVez As Boolean

Private bln_fecha_real_editable As Boolean

Private mvarobjVerificacion As New clsEquipoVerificacion
Private mvarblnResultado As Boolean
Private mvardtmFechaProximaInicial As Date
Private mvarlngidVerificadorInternoInicial As Long
Private mvarlngIdPeriodoInicial As Long
Private mvarlngIdTipoVerificacionIncial As Long
Private mvarblnVieneDeCuaderno As Boolean

Private mvarlngidEquipo As Long
Private mvardtmFechaPrevista As Date

Private mvarlngIdEvento As Long

Private xR As New XArrayDB
Private xM(1 To 2) As New XArrayDB
Private xUnidades As New XArrayDB

Const filasR As Integer = 50
Const ColR As Integer = 11
Const filasM As Integer = 50
Const ColM As Integer = 1

'Private Enum ColsR
'    DESCRIPCION = 0
'    RANGO_MIN = 1
'    RANGO_MAX = 2
'    unidad = 3
'    RESULTADO_CAL = 4
'    TOLERANCIA = 5
'    INCERTIDUMBRE = 6
'    CORRECCION = 7
'    Id_resultado = 8
'    ID_UNIDAD = 9
'End Enum

Private Enum ColsR
    Descripcion = 0
    Unidad = 1
    RANGO_MIN = 2
    RANGO_MAX = 3
    RESULTADO_MEDIA = 4
    id_tipo = 5
    Id_resultado = 6
    RESULTADO_CUALIDAD = 7
    RESULTADOS_PATRON = 8
    id_unidad = 9
    id_patron = 10
    n_medidas = 11
End Enum

Private mvarlngNumParametrosResultados As Long
Private mvarlngidProcedmientoInicial As Long
Private Sub cabecera()
  With lista.ColumnHeaders
        
        .Item(1).Text = "Pto. Verificacion"
        .Item(1).Width = lista.Width * 0.35
        .Item(1).Alignment = lvwColumnLeft
        .Add , , "Unidad", lista.Width * 0.19, lvwColumnLeft
        .Add , , "Rango Min", lista.Width * 0.15, lvwColumnCenter
        .Add , , "Rango Max", lista.Width * 0.15, lvwColumnCenter
        .Add , , "Resultado", lista.Width * 0.15, lvwColumnCenter
        
        ' ocultas
        
        .Add , , "id_tipo", 0
        .Add , , "id_resultado", 0
        .Add , , "RESULTADO_CUALIDAD", 0
        .Add , , "RESULTADOS_PATRON", 0
        .Add , , "id_unidad", 0
        .Add , , "ID_PATRON", 0
        .Add , , "n_medidas", 0
        
    End With
End Sub

Private Function comprobar_datos_parametros(ByRef resultado_conformidad As Boolean) As String
    Dim x As Long
    Dim tipo As Integer
    Dim cad As String
    Dim strTipo As String, rmin As Double, rmax As Double, res As Double
    
    If lista.ListItems.Count = 0 Then Exit Function
    
    resultado_conformidad = True
    
    For x = 1 To lista.ListItems.Count
        With lista.ListItems(x)
            tipo = CInt(.SubItems(ColsR.id_tipo))
            strTipo = IIf(tipo = 1, "Equipo", "Reactivo")
                
            If tipo <> 0 Then
                If CLng(.SubItems(ColsR.id_patron)) = 0 Then
                    cad = cad & vbCrLf & " - El parámetro " & .Text & ", del tipo Patrón-" & strTipo & " no tiene señalado el " & strTipo
                End If
                
                rmin = CDbl(.SubItems(ColsR.RANGO_MIN))
                rmax = CDbl(.SubItems(ColsR.RANGO_MAX))
                res = CDbl(.SubItems(ColsR.RESULTADO_MEDIA))
                If rmin > rmax Then
                    cad = cad & vbCrLf & " - En el parámetro " & .Text & ", del tipo Patrón-" & strTipo & ", el Rango Mínimo es mayor que el Rango Máximo" & strTipo
                End If
                
                If res < rmin Then
                    ' inferior al minimo
                    resultado_conformidad = False
                End If
                
                If res > rmax Then
                    ' superior al máximo
                    resultado_conformidad = False
                End If
            Else
                If optResultadoCualidad(0).value Then resultado_conformidad = False
            End If
        End With
        
    Next x
    comprobar_datos_parametros = cad
    
    
End Function

Private Sub ConfigurarTablet()
    Set TecladoNumerico = New frmTecladoNumerico
    
    
    TecladoNumerico.OcultarConformidad = True
    
    blnEsTablet = pc_es_tablet
    
    If blnEsTablet Then
        
        blnPrimeraVez = True
        'On Error Resume Next
        'grdResultados.Columns(ColsR.RESULTADO_MEDIA).Locked = True
        'On Error GoTo 0
        Me.Top = 0
        

    End If
End Sub

Private Sub CargarComboGridUnidad()
'    Dim rs As ADODB.RecordSet
'    Dim ote As New clsUnidades
'
'    Set rs = ote.Listado()
'    xUnidades.Clear
'    If rs.RecordCount > 0 Then
'        xUnidades.ReDim 0, rs.RecordCount, 0, 1
'        Dim i As Integer
'        i = 1
'        Do
'            xUnidades(i, 0) = CStr(rs("NOMBRE"))
'            xUnidades(i, 1) = CStr(rs("ID_UNIDAD"))
'            i = i + 1
'            rs.MoveNext
'        Loop Until rs.EOF
'    Else
'        xUnidades.ReDim 0, 0, 0, 1
'    End If
'    Set tUnidades.Array = xUnidades
'    tUnidades.Refresh
'    grdResultados.Refresh
End Sub

Private Function devolver_medidas_resultado(ByVal prm_id_resultado As Long, ByRef rs As ADODB.RecordSet) As String

    If rs.RecordCount = 0 Then Exit Function
    Dim cad As String
    
    cad = ""
    
    rs.Filter = "resultado_id = " & CStr(prm_id_resultado)
    
    If rs.RecordCount = 0 Then
        rs.MoveFirst
        While Not rs.EOF
            cad = cad & ";" & CStr(rs("resultado"))
            rs.MoveNext
        Wend
        cad = Mid(cad, 2)
    End If
    
    rs.Filter = ""
    
    devolver_medidas_resultado = cad

End Function

Private Sub modificar_parametro()

    Dim objfrm As New frmEquipoVerificacionAnadirParametro
    Dim tipo_id As Integer, tipo_actual As Integer

    tipo_actual = CInt(lista.SelectedItem.SubItems(ColsR.id_tipo))
    
    With objfrm
        .Descripcion = lista.SelectedItem.Text
        .id_unidad = lista.SelectedItem.SubItems(ColsR.id_unidad)
        .tipo = lista.SelectedItem.SubItems(ColsR.id_tipo)
        .rmax = lista.SelectedItem.SubItems(ColsR.RANGO_MAX)
        .rmin = lista.SelectedItem.SubItems(ColsR.RANGO_MIN)
        .medidas = lista.SelectedItem.SubItems(ColsR.n_medidas)
    End With
    
    objfrm.Show vbModal
    
    If Not objfrm.resultado Then Exit Sub
    
    tipo_id = objfrm.tipo
                
    With lista.ListItems(lista.SelectedItem.Index)
        .Text = objfrm.Descripcion
        .SubItems(ColsR.Unidad) = objfrm.Unidad
        If tipo_id = 0 Then
            .SubItems(ColsR.RESULTADO_MEDIA) = "CONFORME"
            .SubItems(ColsR.RANGO_MIN) = "N/A"
            .SubItems(ColsR.RANGO_MAX) = "N/A"
            .SubItems(ColsR.id_unidad) = "N/A"
            .SubItems(ColsR.n_medidas) = "1"
        Else
            .SubItems(ColsR.RANGO_MIN) = objfrm.rmin
            .SubItems(ColsR.RANGO_MAX) = objfrm.rmax
            .SubItems(ColsR.id_unidad) = objfrm.id_unidad
            .SubItems(ColsR.Unidad) = objfrm.Unidad
            If .SubItems(ColsR.n_medidas) <> objfrm.medidas Then
                Dim r As Double, str_total As String
                r = CDbl(.SubItems(ColsR.RESULTADO_MEDIA))
                .SubItems(ColsR.RESULTADOS_PATRON) = recalcular_resultados_patron(.SubItems(ColsR.RESULTADOS_PATRON), objfrm.medidas, r)
                If InStr(1, CStr(r), ",") Then
                    ' mide los decimales. Si son más de 6, los redondea a 6 decimales
                    str_total = Split(CStr(CDbl(r)), ",")(0) & "," & Left(Split(CStr(CDbl(r)), ",")(1), 6)
                Else
                    str_total = CStr(r)
                End If

                .SubItems(ColsR.RESULTADO_MEDIA) = str_total
            End If
            .SubItems(ColsR.n_medidas) = objfrm.medidas
        End If
        
        .SubItems(ColsR.id_tipo) = tipo_id
        If tipo_actual <> tipo_id Then
            ' solo si cambia el tipo se reinicializan estos parámetros
            .SubItems(ColsR.RESULTADO_MEDIA) = "0"
            .SubItems(ColsR.id_patron) = "0"
            .SubItems(ColsR.n_medidas) = "1"
            .SubItems(ColsR.RESULTADO_CUALIDAD) = "1"
            .SubItems(ColsR.RESULTADOS_PATRON) = IIf(tipo_id = 0, "", "0")
        End If
    End With

    Set objfrm = Nothing
    ' se va al añadido
    lista.SelectedItem = lista.ListItems(lista.ListItems.Count)
    lista_Click
    Set objfrm = Nothing
    
End Sub



Private Sub mostrar_datos_medidas()
'mostrar_medidas

Dim tipo_id As Integer, i As Integer, patron_id As Long
Dim cad As String
Dim arrRes() As String
Dim n_medidas As Integer, res_cual As Integer
Dim inicializar As Boolean
'Muestra los resultados segun tipo

If lista.ListItems.Count = 0 Then Exit Sub

patron_id = 0
tipo_id = -1
n_medidas = 1
res_cual = 1
inicializar = True

n_medidas = CInt(lista.SelectedItem.SubItems(ColsR.n_medidas))
tipo_id = CInt(lista.SelectedItem.SubItems(ColsR.id_tipo))
lblParametro = lista.SelectedItem
cad = lista.SelectedItem.SubItems(ColsR.RESULTADOS_PATRON)
res_cual = lista.SelectedItem.SubItems(ColsR.RESULTADO_CUALIDAD)
patron_id = lista.SelectedItem.SubItems(ColsR.id_patron)

If tipo_id < 0 Then Exit Sub

lblParametro.Visible = True

fraTipoParametro(0).Visible = False
fraTipoParametro(1).Visible = False
fraTipoParametro(2).Visible = False
fraTipoParametro(tipo_id).Visible = True

' limpia los resultados
If tipo_id <> 0 Then
    
    lista_medidas(tipo_id).ListItems.Clear
    For i = 1 To n_medidas
        lista_medidas(tipo_id).ListItems.Add , , "0"
    Next i
    
    If cad <> "" Then
        arrRes = Split(cad, ";")
        For i = 0 To UBound(arrRes)
            lista_medidas(tipo_id).ListItems(i + 1).Text = arrRes(i)
        Next i
    End If
    txtNMedidas(tipo_id).Text = n_medidas
    If tipo_id = 1 Then
        cmbEquipos.MostrarElemento patron_id
    Else
        cmbReactivos.MostrarElemento patron_id
    End If
    lista_medidas_Click tipo_id
Else
    txtDescripcion_Cualidad.Text = cad
    optResultadoCualidad(res_cual).value = True
End If

End Sub

Private Sub OpcionesEdicion()


    If mvarenuTipoEdicion = ALTA Then
        txtFechaActual.Enabled = True
    ElseIf mvarenuTipoEdicion = EDICION Then
        txtFechaActual.Enabled = bln_fecha_real_editable Or (mvarobjVerificacion.getESTADO = 0)
    ElseIf mvarenuTipoEdicion = visualizar Then
    
        cmbTipoVerificacion.Locked = True
        cmbPeriVerificacion.Locked = True
        cmbVerificador.desactivar
        cmbVerificadorExterno.desactivar
        cmbProcedimiento.desactivar
        txtHojaVerificacion.Locked = False
            cmdMostrarHojaCal.Left = cmdAdjuntarHojaCal.Left
            cmdAdjuntarHojaCal.Visible = False
            cmdEscanearHoja.Visible = False
            cmdEliminarHojaCal.Visible = False
        txtCertificado.Locked = False
            cmdMostrarCertificado.Left = cmdAdjuntarCertificado.Left
            cmdAdjuntarCertificado.Visible = False
            cmdEscanearCert.Visible = False
            cmdEliminarCertificado.Visible = False
        txtEvaluacionResultado.Locked = False
            cmdMostrarEvaluacion.Left = cmdAdjuntarEvaluacion.Left
            cmdAdjuntarEvaluacion.Visible = False
            cmdEscanearEvaluacion.Visible = False
            cmdEliminarEvaluacion.Visible = False
        txtLimitacionesUso.Locked = True
        cmdAnadirLimitacion.Enabled = False
        cmdEliminarLimitacion.Enabled = False
        lstLimitacionesUso.Enabled = False
        
        txtFechaProxima_b.Enabled = False
        fraEstadoIntervencion.Enabled = False
        
        txtCertificado.Locked = True
        txtHojaVerificacion.Locked = True
        txtEvaluacionResultado.Locked = True
    
        'grdResultados.Enabled = False
        cmdAnadirParametro.Enabled = False
        cmdEliminarParametro.Enabled = False
        fraTipoParametro(0).Enabled = False
        fraTipoParametro(1).Enabled = False
        fraTipoParametro(2).Enabled = False
        
        cmdok.Visible = False
    End If
End Sub

Private Function recalcular_resultados_patron(res_patron As String, n_medidas As Integer, resultado As Double) As String
Dim arrPatron() As String
Dim i As Integer
Dim res As String

    
    If Trim(res_patron) <> "" And n_medidas <> 0 Then
        arrPatron = Split(res_patron, ";")
        If (UBound(arrPatron) + 1) < n_medidas Then ' se aumentan las medidas
            res = res_patron
            For i = 0 To UBound(arrPatron)
                resultado = resultado + CDbl(arrPatron(i))
            Next i
            
            For i = (UBound(arrPatron) + 1) To n_medidas
                res = res = ";0"
            Next i
            res = Mid(res, 2)
        Else ' disminuyen las medidas
            For i = 0 To (n_medidas - 1)
                resultado = resultado + CDbl(arrPatron(i))
                res = res & ";" & arrPatron(i)
            Next i
            res = Mid(res, 2)
        End If
        
        'recalcula el resultado
        resultado = resultado / n_medidas
        
    End If


End Function

Private Sub cmbEquipos_change()
lista.SelectedItem.SubItems(ColsR.id_patron) = cmbEquipos.getPK_SALIDA
End Sub

Private Sub cmbPeriVerificacion_Click(AREA As Integer)

    Call txtFechaActual_Change

End Sub

Private Sub cmbReactivos_change()
xR(grdResultados.Row, ColsR.id_patron) = cmbReactivos.getPK_SALIDA
On Error Resume Next
grdResultados.Columns(ColsR.id_patron).RefetchCell grdResultados.Row
On Error GoTo 0
End Sub

Private Sub cmbTipoParametro_Change(Index As Integer)

Dim cad As String, arrRes() As String
Dim n_medidas As String
Dim tipo_id As Integer, i As Integer

If bln_cambiando_tipo Then Exit Sub


tipo_id = cmbTipoParametro(Index).BoundText

fraTipoParametro(0).Visible = False
fraTipoParametro(1).Visible = False
fraTipoParametro(2).Visible = False
fraTipoParametro(tipo_id).Visible = True
bln_cambiando_tipo = True
cmbTipoParametro(tipo_id).BoundText = tipo_id
bln_cambiando_tipo = False
xR(grdResultados.Row, ColsR.id_tipo) = tipo_id

cad = xR(grdResultados.Row, ColsR.RESULTADOS_PATRON)



n_medidas = xR(grdResultados.Row, ColsR.n_medidas)

If tipo_id > 0 Then
    If n_medidas = "" Then n_medidas = "1"
    
    lista_medidas(tipo_id).ListItems.Clear
    
    For i = 1 To n_medidas
        lista_medidas(tipo_id).ListItems.Add , , "0"
    Next i
    
    If cad <> "" Then
        arrRes = Split(cad, ";")
        If UBound(arrRes) = 0 And Not IsNumeric(arrRes(0)) Then
            lista_medidas(tipo_id).ListItems(1).Text = "0"
        Else
            For i = 0 To UBound(arrRes)
                If IsNumeric(arrRes(i)) Then
                    lista_medidas(tipo_id).ListItems(i + 1).Text = arrRes(i)
                End If
            Next i
        End If
    End If
    
    xR(grdResultados.Row, ColsR.id_patron) = "0"
    xR(grdResultados.Row, ColsR.n_medidas) = 1
    xR(grdResultados.Row, ColsR.RESULTADO_MEDIA) = "0"
    xR(grdResultados.Row, ColsR.RANGO_MIN) = ""
    xR(grdResultados.Row, ColsR.RANGO_MAX) = ""
    xR(grdResultados.Row, ColsR.id_unidad) = ""
    xR(grdResultados.Row, ColsR.Unidad) = ""
    
Else ' para el caso en que sea cualidad
    txtDescripcion_Cualidad.Text = cad
    
    xR(grdResultados.Row, ColsR.id_patron) = 0
    xR(grdResultados.Row, ColsR.n_medidas) = 1
    xR(grdResultados.Row, ColsR.RESULTADO_MEDIA) = "CONFORME"
    ' N/A
    xR(grdResultados.Row, ColsR.RANGO_MIN) = "N/A"
    xR(grdResultados.Row, ColsR.RANGO_MAX) = "N/A"
    xR(grdResultados.Row, ColsR.id_unidad) = "0"
    xR(grdResultados.Row, ColsR.Unidad) = "N/A"
    
End If

On Error Resume Next
grdResultados.Columns(ColsR.id_tipo).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.RANGO_MIN).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.RESULTADO_MEDIA).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.id_patron).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.n_medidas).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.RANGO_MAX).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.id_unidad).RefetchCell grdResultados.Row
grdResultados.Columns(ColsR.Unidad).RefetchCell grdResultados.Row
On Error GoTo 0
End Sub

Private Sub cmbTipoVerificacion_Change()


If cmbTipoVerificacion.BoundText = "1" Then ' Intera
    ' Es interna
    cmbVerificadorExterno.desactivar
Else
    ' Es externa
    cmbVerificadorExterno.activar
End If

End Sub

' botón que abre un cuadro de diálogo para seleccionar la plantilla excel de la verificación
Private Sub cmdAdjuntarCertificado_Click()

On Error GoTo cmdAdjuntarCertificado_Click_Error
    
    cd.ShowOpen
    
    If Trim(cd.FileName) = "" Then Exit Sub
    
    mvarobjVerificacion.Certificado.setRUTA_TEMPORAL = cd.FileName
    mvarobjVerificacion.Certificado.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
    txtCertificado.Text = cd.FileTitle

On Error GoTo 0
    Exit Sub
cmdAdjuntarCertificado_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarCertificado_Click of Formulario frmEquipoEdicionVerificacion_nuevo"
End Sub

Private Sub cmdAdjuntarEvaluacion_Click()

On Error GoTo cmdAdjuntarEvaluacion_Click_Error

    cd.ShowOpen
    
    If Trim(cd.FileName) = "" Then Exit Sub
    
    mvarobjVerificacion.Evaluacion.setRUTA_TEMPORAL = cd.FileName
    mvarobjVerificacion.Evaluacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
    txtEvaluacionResultado.Text = cd.FileTitle
   

On Error GoTo 0
    Exit Sub
cmdAdjuntarEvaluacion_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarEvaluacion_Click of Formulario frmEquipoEdicionVerificacion_nuevo"
End Sub


Private Sub cmdAdjuntarHojaCal_Click()


On Error GoTo cmdAdjuntarHojaCal_Click_Error


    cd.ShowOpen
    
    If Trim(cd.FileName) = "" Then Exit Sub
    
    mvarobjVerificacion.HojaVerificacion.setRUTA_TEMPORAL = cd.FileName
    mvarobjVerificacion.HojaVerificacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
    txtHojaVerificacion.Text = cd.FileTitle
    

On Error GoTo 0
    Exit Sub
cmdAdjuntarHojaCal_Click_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarHojaCal_Click of Formulario frmEquipoEdicionVerificacion_nuevo"
End Sub

Private Sub cmdAnadirLimitacion_Click()

    mvarobjEquipo.Anadir_limitacionuso_equipo txtLimitacionesUso.Text
           
    Call PresentarDatos_LimitacionesUso
End Sub

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide
End Sub

' botón que borra el documento de verificación
Private Sub cmdEliminarCertificado_Click()

txtCertificado.Text = ""
mvarobjVerificacion.Certificado.setID_AUX = enumIdAux.ID_AUX_ELIMINADO

End Sub

Private Sub cmdEliminarEvaluacion_Click()

txtEvaluacionResultado.Text = ""
mvarobjVerificacion.Evaluacion.setID_AUX = enumIdAux.ID_AUX_ELIMINADO

End Sub


Private Sub cmdEliminarHojaCal_Click()

txtHojaVerificacion.Text = ""

mvarobjVerificacion.HojaVerificacion.setID_AUX = enumIdAux.ID_AUX_ELIMINADO

End Sub

Private Sub cmdEliminarLimitacion_Click()
Dim lngid As Long

    If lstLimitacionesUso.ListIndex < 0 Then Exit Sub

    lngid = lstLimitacionesUso.ItemData(lstLimitacionesUso.ListIndex)

    mvarobjEquipo.Eliminar_LimitacionUso_equipo lngid
    
    Call PresentarDatos_LimitacionesUso
End Sub

Private Sub cmdEliminarParametro_Click()
If lista.ListItems.Count = 0 Then Exit Sub

If lista.SelectedItem.Index <= 0 Then
    MsgBox "Debe señalar el parámetro a eliminar", vbInformation, "Eliminar Parámetro"
    Exit Sub
End If

lista.ListItems.Remove lista.SelectedItem.Index


If lista.ListItems.Count = 0 Then
    fraTipoParametro(0).Visible = False
    fraTipoParametro(1).Visible = False
    fraTipoParametro(2).Visible = False
    lblParametro.Visible = False
Else
    lista.SelectedItem = lista.ListItems(1)
End If

End Sub

Private Sub cmdEscanearCert_Click()
Dim strArchivo As String
    
    strArchivo = EscanearATemp
    
    If Trim(strArchivo) = "" Then Exit Sub
        
    mvarobjVerificacion.Certificado.setRUTA_TEMPORAL = strArchivo
    mvarobjVerificacion.Certificado.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
    txtCertificado.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
End Sub

Private Sub cmdEscanearEvaluacion_Click()
Dim strArchivo As String
    
    strArchivo = EscanearATemp
    
    If Trim(strArchivo) = "" Then Exit Sub
        
    mvarobjVerificacion.Evaluacion.setRUTA_TEMPORAL = strArchivo
    mvarobjVerificacion.Evaluacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
    txtEvaluacionResultado.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    
End Sub


Private Sub cmdEscanearHoja_Click()
    
    Dim strArchivo As String
    
    strArchivo = EscanearATemp
    
    If Trim(strArchivo) = "" Then Exit Sub
    
    mvarobjVerificacion.HojaVerificacion.setRUTA_TEMPORAL = strArchivo
    mvarobjVerificacion.HojaVerificacion.setID_AUX = enumIdAux.ID_AUX_MODIFICADO
    txtHojaVerificacion.Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    
End Sub

' botón que permite imprimir la etiqueta de verificación
Private Sub cmdEtiqueta_Click()
  
    If cmbVerificador.getPK_SALIDA > 1 Then ' sólo si está seleccionada la verificación más actual
        Call imprimir_etiqueta(Format(txtFechaActual.value, "dd/mm/yyyy"), cmbVerificador.getPK_SALIDA)
    End If
    
End Sub

' botón que permite visualizar el archivo seleccionado
Private Sub cmdMostrarCertificado_Click()
    
    Dim objAI As New clsArchivoAdjunto
    Dim destino As String, r As Double
    
    Set objAI = mvarobjVerificacion.Certificado
    
    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
        destino = objAI.getRUTA_TEMPORAL
    ElseIf (objAI.getRUTA <> "") Then
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & mvarobjVerificacion.getID_VERIFICACION & "\CERT\" & objAI.getNOMBRE_ARCHIVO
    End If
    
    On Error GoTo fallo
    
    ' verificar si es hoja excel
    If UCase(Right(destino, 3) = "XLS") Then
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Open(destino, , True)
        Set XLS = XLW.Worksheets(1)
        XLA.Visible = True
    ElseIf Dir(destino, vbArchive) <> "" Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If
    
fallo:
End Sub

Private Sub cmdMostrarEvaluacion_Click()
    
    Dim objAI As New clsArchivoAdjunto
    Dim destino As String, r As Double
    
    Set objAI = mvarobjVerificacion.Evaluacion
    
    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
        destino = objAI.getRUTA_TEMPORAL
    ElseIf (objAI.getRUTA <> "") Then
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & mvarobjVerificacion.getID_VERIFICACION & "\EVAL\" & objAI.getNOMBRE_ARCHIVO
    End If
    
    On Error GoTo fallo
    
    ' verificar si es hoja excel
    If UCase(Right(destino, 3) = "XLS") Then
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Open(destino, , True)
        Set XLS = XLW.Worksheets(1)
        XLA.Visible = True
    ElseIf Dir(destino, vbArchive) <> "" Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If
    
fallo:
End Sub


Private Sub cmdMostrarHojaCal_Click()

    
    Dim objAI As New clsArchivoAdjunto
    Dim destino As String, r As Double
    Set objAI = mvarobjVerificacion.HojaVerificacion
    
    If Trim(objAI.getRUTA_TEMPORAL) <> "" Then
        destino = objAI.getRUTA_TEMPORAL
    ElseIf (objAI.getRUTA <> "") Then
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & mvarobjVerificacion.getID_VERIFICACION & "\HOJA\" & objAI.getNOMBRE_ARCHIVO
    End If
        
On Error GoTo fallo
    
    ' verificar si es hoja excel
    If UCase(Right(destino, 3) = "XLS") Then
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Open(destino, , True)
        Set XLS = XLW.Worksheets(1)
        XLA.Visible = True
    ElseIf Dir(destino, vbArchive) <> "" Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If
fallo:
End Sub


Private Sub cmdok_Click()
    ' Recoge los datos
    Dim lngId_Verificacion As Long
    Dim bln_conformidad As Boolean
    If Not ComprobarDatos() Then Exit Sub
    
    RecogerDatos
    
    If mvarenuTipoEdicion = ALTA Then
        mvarobjVerificacion.setEQUIPO_ID = mvarlngidEquipo
        lngId_Verificacion = mvarobjVerificacion.Insertar(True, lista)
    Else
        lngId_Verificacion = CLng(mvarstrId)
        Call mvarobjVerificacion.Modificar(lngId_Verificacion, True, , lista)
    End If
        
    'Call mvarobjVerificacion.GuardarParametrosVerificacion(mvarlngidEquipo, lngId_Verificacion, xR, filasR)
    
    'If Not mvarblnVieneDeCuaderno Then
        ' Si no viene del cuaderno de avisos, es decir, que viene de la gestion normal y corriente, recarga las calibraciones
    '    mvarobjEquipo.Carga_Verificaciones
    'End If
    
    mvarblnResultado = True
    Me.Hide

End Sub


Private Sub comprobar_fecha_real_modificable()

    Dim op As New clsParametros
    
    bln_fecha_real_editable = False
    
    If op.Carga(parametros.MODIFICACION_FECHAS_CALIBRACION_VERIFICACION, "") Then
        If op.getVALOR = "1" Then
            bln_fecha_real_editable = True
        End If
    End If
    

End Sub


Private Function ComprobarDatos() As Boolean
Dim strMs As String
Dim bln_conformidad As Boolean
On Error GoTo ComprobarDatos_Error

    ComprobarDatos = False

    strMs = ""

    If Not optEstado(0).value Then
        comprobarDatosResultados strMs
    End If
    
    If Trim(cmbTipoVerificacion.BoundText) = "" Or Trim(cmbTipoVerificacion.BoundText) = "0" Then
        strMs = strMs & vbCrLf & " - Debe indicar el Tipo de Verificación"
    End If

    If cmbVerificador.getPK_SALIDA = 0 Then
        strMs = strMs & vbCrLf & " - Debe indicar el Responsable Interno de Verificación"
    End If


    If Trim(cmbPeriVerificacion.BoundText) = "" Then
        strMs = strMs & vbCrLf & " - Debe indicar el Periodo para las Verificaciones"
    End If
    
    If getDataComboSel(cmbTipoVerificacion) = 1 Then
        If Trim(cmbProcedimiento.getPK_SALIDA) < 1 Then
            strMs = strMs & vbCrLf & " - Debe indicar el el Procedimiento"
        End If
    ElseIf getDataComboSel(cmbTipoVerificacion) = 2 Then
        If Trim(cmbVerificadorExterno.getPK_SALIDA) < 1 Then
            strMs = strMs & vbCrLf & " - Debe indicar el Verificador Externo"
        End If
    End If
    
    If CDate("01/01/1900") = txtFechaActual.value Then
        strMs = strMs & vbCrLf & " - Debe indicar una Fecha Actual de Verificación adecuada"
    End If
    
    If txtFechaActual.value >= txtFechaProxima_b.value Then
        strMs = strMs & vbCrLf & " - La fecha de la próxima verificación no puede ser anterior a la de la Verificación actual"
    End If

    
    strMs = comprobar_datos_parametros(bln_conformidad)
    
    If Trim(strMs) <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strMs
        Exit Function
    End If

    ' comprobar si se cierra conforme y no lo es
    If optEstado(1).value Or optEstado(2).value Then
        If optEstado(1).value And Not bln_conformidad Then
            MsgBox "ATENCION: No se puede cerrar esta verificación como CONFORME, dado que uno de los resultados de los parámetros está fuera de rango", vbInformation, "Verificación NO CONFORME"
            Exit Function
        ElseIf optEstado(2).value And bln_conformidad Then
            MsgBox "ATENCION: No se puede cerrar esta verificación como NO CONFORME, dado que todos los resultados de los parámetros está dentro de rango", vbInformation, "Verificación CONFORME"
            Exit Function
        End If
    End If
    


    ComprobarDatos = True

On Error GoTo 0
    Exit Function
ComprobarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmEquipoEdicionVerificacion_nuevo"
End Function


Private Sub comprobarDatosResultados(ByRef strMs As String)

    Dim i As Long
    i = 0
    Dim cad As String
    
    cad = ""
    'grdResultados.Refresh
    
    'For i = 0 To filasR
    '    cad = cad & xR(i, 0) & ", " & xR(i, 1) & ", " & xR(i, 2) & ", " & xR(i, 3) & ", " & xR(i, 4) & ", " & xR(i, 5) & ", " & xR(i, 6) & ", " & xR(i, 7) & ". " & vbCrLf
    'Next i
    
End Sub

Public Property Get Equipo() As clsEquipos

    Set Equipo = mvarobjEquipo

End Property

Public Property Set Equipo(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Public Property Get FechaPrevista() As Date

    FechaPrevista = mvardtmFechaPrevista

End Property

Public Property Let FechaPrevista(ByVal dtmFechaPrevista As Date)

    mvardtmFechaPrevista = dtmFechaPrevista

End Property

Public Property Get FechaProximaInicial() As Date

    FechaProximaInicial = mvardtmFechaProximaInicial

End Property

Public Property Let FechaProximaInicial(ByVal dtmFechaProximaInicial As Date)

    mvardtmFechaProximaInicial = dtmFechaProximaInicial

End Property

Private Sub cmdAnadirParametro_Click()
Dim objfrm As New frmEquipoVerificacionAnadirParametro
Dim tipo_id As Integer


    objfrm.Show vbModal
    
    If Not objfrm.resultado Then Exit Sub
    
    tipo_id = objfrm.tipo
    
    With lista.ListItems.Add(, , objfrm.Descripcion)
        .SubItems(ColsR.Unidad) = objfrm.Unidad
        If tipo_id = 0 Then
            .SubItems(ColsR.RESULTADO_MEDIA) = "CONFORME"
            .SubItems(ColsR.RANGO_MIN) = "N/A"
            .SubItems(ColsR.RANGO_MAX) = "N/A"
            .SubItems(ColsR.id_unidad) = "N/A"
            .SubItems(ColsR.n_medidas) = "1"
            .SubItems(ColsR.RESULTADOS_PATRON) = ""
        Else
            .SubItems(ColsR.RESULTADO_MEDIA) = "0"
            .SubItems(ColsR.RANGO_MIN) = objfrm.rmin
            .SubItems(ColsR.RANGO_MAX) = objfrm.rmax
            .SubItems(ColsR.id_unidad) = objfrm.id_unidad
            .SubItems(ColsR.n_medidas) = objfrm.medidas
            .SubItems(ColsR.RESULTADOS_PATRON) = "0"
        End If
        .SubItems(ColsR.id_tipo) = tipo_id
        .SubItems(ColsR.id_patron) = "0"
        .SubItems(ColsR.Id_resultado) = "0"
        
        .SubItems(ColsR.RESULTADO_CUALIDAD) = "1"
        
    End With

    Set objfrm = Nothing

    ' se va al añadido
    lista.SelectedItem = lista.ListItems(lista.ListItems.Count)
    lista_Click
    
    
End Sub

Private Sub Form_Activate()
    
    If blnPrimeraVez Then
        grdResultados_BeforeColEdit ColsR.RESULTADO_MEDIA, 0, 0
        blnPrimeraVez = False
    End If

End Sub

Private Sub Form_Load()

comprobar_fecha_real_modificable

If mvarblnVieneDeCuaderno Then
    Set mvarobjEquipo = New clsEquipos
    Call mvarobjEquipo.Carga(mvarlngidEquipo)
    
    'mvarlngIdTipoVerificacionIncial = mvarobjEquipo.getTIPO_VERIFICACION_ID
    'mvarlngIdPeriodoInicial = mvarobjEquipo.getPERIODICIDAD_VERIFICACION_ID
    'mvarlngidVerificadorInternoInicial = mvarobjEquipo.getVERIFICADOR_INTERNO_ID
    'mvardtmFechaProximaInicial = mvardtmFechaPrevista
    'mvarlngidProcedmientoInicial = mvarobjEquipo.getPROCEDIMIENTO_VERIFICACION_ID
        
    'If mvarlngIdEvento = 0 Then
    '    mvarenuTipoEdicion = ALTA
    'Else
        mvarenuTipoEdicion = EDICION
        mvarstrId = CStr(mvarlngIdEvento)
    'End If
    
End If

mvarlngidEquipo = mvarobjEquipo.getID_EQUIPO

Call PresentarDatos_LimitacionesUso


Call LlenarCombos
'Call inicializar_grid
Call cabecera
Call CargarComboGridUnidad

Call PresentarDatos_ParametrosResultados

blnPrimeraVez = False
    
Call ConfigurarTablet

If mvarenuTipoEdicion = ALTA Then

    'txtFechaActual.value = mvardtmFechaProximaInicial
    txtFechaActual.value = Now
    
    txtFechaActual.Enabled = bln_fecha_real_editable Or True
    'txtFechaProxima_b.value = calcularFechaProxima(mvardtmFechaProximaInicial, mvarlngIdPeriodoInicial)
    Set mvarobjVerificacion = New clsEquipoVerificacion
    cmbTipoVerificacion.BoundText = mvarlngIdTipoVerificacionIncial
    cmbPeriVerificacion.BoundText = mvarlngIdPeriodoInicial
    txtFechaActual_Change
    cmbVerificador.MostrarElemento mvarlngidVerificadorInternoInicial
    cmbProcedimiento.MostrarElemento mvarlngidProcedmientoInicial
    Exit Sub
End If

'Set mvarobjVerificacion = mvarobjEquipo.Verificaciones.Item(mvarstrId)
mvarobjVerificacion.Carga CLng(mvarstrId)
Call PresentarDatos

Call OpcionesEdicion

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub




Private Sub grdResultados_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If blnEsTablet And ColIndex = ColsR.RESULTADO_MEDIA Then
    grdResultados.Col = ColIndex
    TecladoNumerico.TextoInicial = grdResultados.Text
    TecladoNumerico.cabecera = xR(grdResultados.Row, 0)
    TecladoNumerico.Subcabecera = "Resultado" 'xP(gridP.Row, 1)
    
    TecladoNumerico.Show 1
    grdResultados.EditActive = False
    
End If

grdResultados_RowColChange 0, 0

End Sub





Private Sub grdResultados_KeyPress(KeyAscii As Integer)
    
    With grdResultados
        If .Col = 1 Or .Col = 2 Or .Col = 5 Or .Col = 1 Or .Col = 6 Or .Col = 7 Then
            KeyAscii = KeyAscii_SoloDecimal_tbgrid(.Text, KeyAscii, True)
        End If
        If .Col = 1 Then
            lblParametro.Caption = .Text
        End If
    End With
    

        
End Sub

Public Property Get id() As String

    id = mvarstrId

End Property

Public Property Let id(ByVal strId As String)

    mvarstrId = strId

End Property

Public Property Get idEquipo() As Long

    idEquipo = mvarlngidEquipo

End Property

Public Property Let idEquipo(ByVal lngidEquipo As Long)

    mvarlngidEquipo = lngidEquipo

End Property

Public Property Get IdEvento() As Long

    IdEvento = mvarlngIdEvento

End Property

Public Property Let IdEvento(ByVal lngIdEvento As Long)

    mvarlngIdEvento = lngIdEvento

End Property

Public Property Get IdPeriodoInicial() As Long

    IdPeriodoInicial = mvarlngIdPeriodoInicial

End Property

Public Property Let IdPeriodoInicial(ByVal lngIdPeriodoInicial As Long)

    mvarlngIdPeriodoInicial = lngIdPeriodoInicial

End Property

Public Property Get IdTipoVerificacionIncial() As Long

    IdTipoVerificacionIncial = mvarlngIdTipoVerificacionIncial

End Property

Public Property Let IdTipoVerificacionIncial(ByVal lngIdTipoVerificacionIncial As Long)

    mvarlngIdTipoVerificacionIncial = lngIdTipoVerificacionIncial

End Property

Public Property Get idVerificadorInternoInicial() As Long

    idVerificadorInternoInicial = mvarlngidVerificadorInternoInicial

End Property

Public Property Let idVerificadorInternoInicial(ByVal lngidVerificadorInternoInicial As Long)

    mvarlngidVerificadorInternoInicial = lngidVerificadorInternoInicial

End Property

Private Sub imprimir_etiqueta(strFecha_Verificacion As String, lngOperador_ID As Long)
On Error GoTo trataError
   
    With frmReport
        .iniciar
        .informe = "Equipos\rptEquipos_ETIQUETA_Verificacion"
        .CRITERIO = "{eq_verificacion_equipos.ID_VERIFICACION} = " & CLng(PK)
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    log ("Final impresion de etiqueta de verificación de equipo")
    
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir la etiqueta de verificación.", vbCritical, Err.Description
End Sub

Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error

    grdResultados.Col = 0
    grdResultados.Row = 0
    
    xR.Clear
    xR.ReDim 0, filasR, 0, ColR
    xR.Clear
    
    Set grdResultados.Array = xR
    grdResultados.Refresh
    

'    lista_medidas(1).Col = 0
'    lista_medidas(1).Row = 0
'    lista_medidas(2).Col = 0
'    lista_medidas(2).Row = 0
'
'    xM(1).Clear
'    xM(1).ReDim 0, filasM, 0, ColM
'    xM(1).Clear
'    xM(2).Clear
'    xM(2).ReDim 0, filasM, 0, ColM
'    xM(2).Clear
'
'    Set lista_medidas(1).Array = xM(1)
'    Set lista_medidas(2).Array = xM(2)
'    lista_medidas(1).Refresh
'    lista_medidas(2).Refresh
    
    

On Error GoTo 0
Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmEquipoEdicionVerificacion_nuevo"
End Sub

Private Sub grdResultados_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

'mostrar_medidas

Dim tipo_id As Integer, i As Integer
Dim cad As String
Dim arrRes() As String
Dim n_medidas As Integer, res_cual As Integer
Dim inicializar As Boolean
'Muestra los resultados segun tipo


Exit Sub
'If mvarlngNumParametrosResultados = 0 Then Exit Sub

tipo_id = -1
n_medidas = 1
res_cual = 1
inicializar = True

If (grdResultados.Row) <= xR.UpperBound(1) Then
    If Not IsEmpty(xR(grdResultados.Row, ColsR.id_tipo)) Then
        If Trim(xR(grdResultados.Row, ColsR.id_tipo)) <> "" Then
            n_medidas = CInt(xR(grdResultados.Row, ColsR.n_medidas))
            tipo_id = CInt(xR(grdResultados.Row, ColsR.id_tipo))
            lblParametro = xR(grdResultados.Row, ColsR.Descripcion)
            cad = xR(grdResultados.Row, ColsR.RESULTADOS_PATRON)
            res_cual = xR(grdResultados.Row, ColsR.RESULTADO_CUALIDAD)
            inicializar = False
        End If
    End If
End If
If inicializar Then
    tipo_id = 0
    On Error Resume Next
    xR(grdResultados.Row, ColsR.n_medidas) = CStr(n_medidas)
    xR(grdResultados.Row, ColsR.id_tipo) = CStr(tipo_id)
    xR(grdResultados.Row, ColsR.RESULTADO_CUALIDAD) = res_cual
    xR(grdResultados.Row, ColsR.id_patron) = "0"
    xR(grdResultados.Row, ColsR.Id_resultado) = "0"
    xR(grdResultados.Row, ColsR.RESULTADO_MEDIA) = "0"
    
    On Error Resume Next
    grdResultados.Columns(ColsR.n_medidas).RefetchCell grdResultados.Row
    grdResultados.Columns(ColsR.id_tipo).RefetchCell grdResultados.Row
    grdResultados.Columns(ColsR.RESULTADO_CUALIDAD).RefetchCell grdResultados.Row
    grdResultados.Columns(ColsR.id_patron).RefetchCell grdResultados.Row
    grdResultados.Columns(ColsR.Id_resultado).RefetchCell grdResultados.Row
    grdResultados.Columns(ColsR.RESULTADO_MEDIA).RefetchCell grdResultados.Row
    
    On Error GoTo 0
    
    lblParametro = xR(grdResultados.Row, ColsR.Descripcion)
    cad = ""
End If

If tipo_id < 0 Then Exit Sub

lblParametro.Visible = True

fraTipoParametro(0).Visible = False
fraTipoParametro(1).Visible = False
fraTipoParametro(2).Visible = False
fraTipoParametro(tipo_id).Visible = True
cmbTipoParametro(tipo_id).BoundText = tipo_id

' limpia los resultados
If tipo_id <> 0 Then
    
    lista_medidas(tipo_id).ListItems.Clear
    For i = 1 To n_medidas
        lista_medidas(tipo_id).ListItems.Add , , "0"
    Next i
    
    If cad <> "" Then
        arrRes = Split(cad, ";")
        For i = 0 To UBound(arrRes)
            lista_medidas(tipo_id).ListItems(i + 1).Text = arrRes(i)
        Next i
    End If
Else
    txtDescripcion_Cualidad.Text = cad
    optEstado(1).value = True
End If

End Sub



Private Sub lista_medidas_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Integer, Max As Integer
Dim total As Double, res As String

    ' Recalcula la media para esta fila
    total = 0
    res = ""
    ' toma todos los valores
    For i = 0 To filasM
        
        If IsEmpty(xM(Index)(i, 0)) Or Trim(xM(Index)(i, 0)) = "" Then ' mira se existe la descripción
            Max = i - 1
            Exit For
        End If
        total = total + xM(Index)(i, 0)
        res = res & ";" & Replace(Format(xM(Index)(i, 0), "#0.000000"), ",", ".")
    Next i
    
    If Max >= 0 Then
        total = total / (Max + 1)
        res = Mid(res, 2)
    End If
    
    xR(grdResultados.Row, ColsR.RESULTADO_MEDIA) = Format(total, "#0.000000")
    xR(grdResultados.Row, ColsR.RESULTADOS_PATRON) = res
    
    
    grdResultados.Rebind
End Sub


Private Sub lista_Click()
        mostrar_datos_medidas
    
End Sub


Private Sub lista_DblClick()
If cmdAnadirParametro.Enabled Then
    ' solo deja modificar si los no está cerrada.
    modificar_parametro
End If
End Sub

Private Sub lista_medidas_Click(Index As Integer)

    txtValor(Index).Text = lista_medidas(Index).SelectedItem.Text
    On Error Resume Next ' necesario porque lo ejecuta con el doble clic
    txtValor(Index).SetFocus
    On Error GoTo 0
    txtvalor_GotFocus (Index)
End Sub


Private Sub lstLimitacionesUso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdEliminarLimitacion_Click
End Sub

Private Sub LlenarCombos()
Dim oDeco As New clsDecodificadora

    oDeco.cargar_combo cmbPeriVerificacion, decodificadora.EQ_periodicidad
    oDeco.cargar_combo cmbTipoVerificacion, decodificadora.EQ_TIPO_CALIBRACION
    llenar_combo cmbVerificador, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmbProcedimiento, New clsCa_documentos, 0, frmCA_Documento, ""
    llenar_combo cmbVerificadorExterno, New clsProveedor, 0, frmProveedores, ""
    
    oDeco.cargar_combo cmbTipoParametro(0), decodificadora.EQ_TIPOS_PARAMETROS_RESULTADO
    oDeco.cargar_combo cmbTipoParametro(1), decodificadora.EQ_TIPOS_PARAMETROS_RESULTADO
    oDeco.cargar_combo cmbTipoParametro(2), decodificadora.EQ_TIPOS_PARAMETROS_RESULTADO
    
    llenar_combo cmbEquipos, New clsEquipos, 1, frmEquipoEdicion, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 1, frmREX_Bote, ""
    
    If mvarobjEquipo.getTIPO_VERIFICACION_ID = 2 Then ' es Externa
        cmbVerificadorExterno.activar
    Else
        cmbVerificadorExterno.desactivar
    End If

    
End Sub

' ----------------- Funciones auxiliares del formulario ----------------

Public Property Get PK() As Long

    PK = mvarlngPK

End Property

Public Property Let PK(ByVal lngPK As Long)

    mvarlngPK = lngPK

End Property



















Private Sub PresentarDatos()


On Error GoTo PresentarDatos_Error
    
    With mvarobjVerificacion
        cmbTipoVerificacion.BoundText = .getTIPO_ID
        cmbPeriVerificacion.BoundText = .getPERIODICIDAD_ID
        cmbVerificador.MostrarElemento .getVERIFICADOR_INTERNO_ID
        If .getVERIFICADOR_EXTERNO_ID > 0 Then
            cmbVerificadorExterno.MostrarElemento .getVERIFICADOR_EXTERNO_ID
        End If
        cmbProcedimiento.MostrarElemento .getPROCEDIMIENTO_ID
        txtHojaVerificacion.Text = .HojaVerificacion.getNOMBRE_ARCHIVO
        txtCertificado.Text = .Certificado.getNOMBRE_ARCHIVO
        txtEvaluacionResultado.Text = .Evaluacion.getNOMBRE_ARCHIVO
        
        If mvarenuTipoEdicion = ALTA Then
            'txtFechaActual.value = CDate(mvardtmFechaProximaInicial)
            txtFechaActual.value = Now
            txtFechaActual_Change
            cmbPeriVerificacion.BoundText = CStr(mvarlngIdPeriodoInicial)
            cmbTipoVerificacion.BoundText = CStr(mvarlngIdTipoVerificacionIncial)
            'txtFechaProxima_b.value = calcularFechaProxima(mvardtmFechaProximaInicial, mvarlngIdPeriodoInicial)
        Else
            'If .getESTADO = 0 Then
            '    txtFechaActual.value = Now
            '    txtFechaActual_Change
            'Else
                txtFechaActual.value = CDate(.getFECHA_ACTUAL)
                txtFechaActual_Change
                'txtFechaProxima_b.value = CDate(.getFECHA_PROXIMA)
                'txtFechaProxima.Text = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
            'End If
        End If
        
        optEstado(.getESTADO).value = True
    End With


    Call PresentarDatos_Adjuntos

On Error GoTo 0
    Exit Sub
PresentarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmEquipoEdicionVerificacion_nuevo"

End Sub


















Private Sub PresentarDatos_Adjuntos()

Dim obja As clsArchivoAdjunto

    Set obja = mvarobjVerificacion.HojaVerificacion
    If Not obja Is Nothing Then
        txtHojaVerificacion.Text = IIf(obja.getNOMBRE_ARCHIVO_TEMP <> "", obja.getNOMBRE_ARCHIVO_TEMP, obja.getNOMBRE_ARCHIVO)
    End If
    
    Set obja = mvarobjVerificacion.Certificado
    If Not obja Is Nothing Then
        txtCertificado.Text = IIf(obja.getNOMBRE_ARCHIVO_TEMP <> "", obja.getNOMBRE_ARCHIVO_TEMP, obja.getNOMBRE_ARCHIVO)
    End If
    
    
End Sub

Private Sub PresentarDatos_LimitacionesUso()
    Dim objItem As clsGenericClass

    lstLimitacionesUso.Clear
    txtLimitacionesUso.Text = ""
        
    For Each objItem In mvarobjEquipo.getLIMITACIONES_USO_COL.Iterator
        If objItem.getID_AUX <> enumIdAux.ID_AUX_ELIMINADO Then
            lstLimitacionesUso.AddItem objItem.getNOMBRE
            lstLimitacionesUso.ItemData(lstLimitacionesUso.ListCount - 1) = objItem.getID
        End If
    Next objItem

End Sub

Private Sub PresentarDatos_ParametrosResultados()

    Dim i As Integer
    Dim rs As ADODB.RecordSet
    Dim rs_medidas As ADODB.RecordSet
    On Error GoTo PresentarDatos_ParametrosResultados_Error

    i = 0
    mvarlngNumParametrosResultados = i
    lblParametro.Caption = ""
    lblParametro.Visible = False
    If mvarenuTipoEdicion <> ALTA Then
        ' Carga los Parametros de la verificacion
        
        Set rs = mvarobjVerificacion.DevolverParametrosResultados(mvarstrId)
        'Set rs_medidas = mvarobjVerificacion.DevolverParametrosResultados_medidas(mvarstrId)
'    Else
'        ' Carga los Parametros del Equipo
'        ' JONATHAN.2010.09.06 -> CUANDO ES ALTA, NO SE CREAN PARÁMETROS
'        'Set rs = mvarobjEquipo.DevolverParametrosResultadosEquipoVerificacion(CStr(mvarlngidEquipo))
'    End If
    
    
    
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            While Not rs.EOF
            
                With lista.ListItems.Add(, , rs!Descripcion)
                    .SubItems(ColsR.Unidad) = CStr(rs("unidad"))
                    If CInt(rs("tipo_id")) = 0 Then
                        .SubItems(ColsR.RESULTADO_MEDIA) = IIf(CInt(rs("resultado_cualidad")) = 0, "NO CONFORME", "CONFORME")
                        .SubItems(ColsR.RANGO_MIN) = "N/A"
                        .SubItems(ColsR.RANGO_MAX) = "N/A"
                        .SubItems(ColsR.id_unidad) = "N/A"
                    Else
                        .SubItems(ColsR.RESULTADO_MEDIA) = CStr(rs("resultado"))
                        .SubItems(ColsR.RANGO_MIN) = CStr(rs("rango_min"))
                        .SubItems(ColsR.RANGO_MAX) = CStr(rs("rango_max"))
                        .SubItems(ColsR.id_unidad) = CStr(rs("unidad_ID"))
                    End If
                    .SubItems(ColsR.id_tipo) = CStr(rs("tipo_id"))
                    .SubItems(ColsR.id_patron) = CStr(rs("patron_id"))
                    .SubItems(ColsR.Id_resultado) = CStr(rs("id_resultado"))
                    .SubItems(ColsR.n_medidas) = CStr(rs("n_medidas"))
                    .SubItems(ColsR.RESULTADO_CUALIDAD) = CStr(rs("resultado_cualidad"))
                    .SubItems(ColsR.RESULTADOS_PATRON) = CStr(rs("DATOS_PARAMETRO"))
                End With
                rs.MoveNext
            Wend
            ' se va al primero
            lista.SelectedItem = lista.ListItems(1)
            lista_Click
        End If
        
        
'        If rs.RecordCount > 0 Then
'            Do
'                xR(i, ColsR.Descripcion) = CStr(rs("descripcion"))
'                xR(i, ColsR.RANGO_MIN) = CStr(rs("rango_min"))
'                xR(i, ColsR.RANGO_MAX) = CStr(rs("rango_max"))
'                xR(i, ColsR.unidad) = CStr(rs("unidad"))
'                xR(i, ColsR.id_unidad) = CStr(rs("unidad_ID"))
'                If CInt(rs("tipo_id")) = EQUIPOS_TIPOS_PARAMETROS_RESULTADOS.PARAM_CUALIDAD Then
'                    If rs("resultado") = 0 Then
'
'                        xR(i, ColsR.RESULTADO_MEDIA) = IIf(CInt(rs("resultado_cualidad")) = 0, "NO CONFORME", "CONFORME")
'                    Else
'                        'para los anteriores que tenían resultado
'                        xR(i, ColsR.RESULTADO_MEDIA) = CStr(rs("resultado"))
'                    End If
'                Else
'                    xR(i, ColsR.RESULTADO_MEDIA) = CStr(rs("resultado"))
'                    xR(i, ColsR.RESULTADOS_PATRON) = devolver_medidas_resultado(rs("id_resultado"), rs_medidas)
'                End If
'                xR(i, ColsR.RESULTADO_CUALIDAD) = CStr(rs("resultado_CUALIDAD"))
'                xR(i, ColsR.id_tipo) = CStr(rs("tipo_id"))
'                xR(i, ColsR.ID_PATRON) = CStr(rs("PATRON_ID"))
'                xR(i, ColsR.n_medidas) = CStr(rs("N_MEDIDAS"))
'                xR(i, ColsR.Id_resultado) = CStr(rs("id_resultado"))
'                i = i + 1
'                rs.MoveNext
'            Loop Until rs.EOF
'
'            mvarlngNumParametrosResultados = i
'
'            grdResultados.Row = 0
'            grdResultados_RowColChange 0, 0
'
'        End If
        
        
        
        
    End If
    
    
    'grdResultados.Refresh
    'grdResultados.Enabled = True

    

On Error GoTo 0
    Exit Sub
PresentarDatos_ParametrosResultados_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_ParametrosResultados of Formulario frmEquipoEdicionVerificacion_nuevo"

End Sub

Private Sub RecogerDatos()

    With mvarobjVerificacion
    
    
        ' A patir del 02.09.2010, PROPUESTA
        ' Ahora que hay verificaciones previstas, la fecha se modifica siempre que sea prevista, nunca en el caso de cerrada.
        ' cuando se cierra, siempre es el momento en que se cierra.
        ' de no ser así, el usuario (no es el caso de automaticamente al cerrar una calibracion, que se crea la siguiente prevista)
        ' no se podrían crear previstas más allá del presente
        
        ' La fecha la establece solo si se cierra ahora
        If .getESTADO = 0 Then
            .setFECHA_ACTUAL = Format(txtFechaActual.value, "dd/mm/yyyy")
            .setFECHA_PROXIMA = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
        Else
            .setFECHA_ACTUAL = Format(Now, "dd/mm/yyyy")
            .setFECHA_PROXIMA = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
        End If
    
'        If .getESTADO = 0 Then
'            .setFECHA_ACTUAL = Format(txtFechaActual.value, "dd/mm/yyyy")
'            .setFECHA_PROXIMA = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
'        End If
'
        .setTIPO_ID = CLng(cmbTipoVerificacion.BoundText)
        .setPERIODICIDAD_ID = CLng(cmbPeriVerificacion.BoundText)
        .setVERIFICADOR_INTERNO_ID = cmbVerificador.getPK_SALIDA
        .setRESPONSABLE = cmbVerificador.getTEXTO
        If .getTIPO_ID = 2 Then
            .setVERIFICADOR_EXTERNO_ID = cmbVerificadorExterno.getPK_SALIDA
        Else
            .setVERIFICADOR_EXTERNO_ID = -1
        End If
        
        .setPROCEDIMIENTO_ID = cmbProcedimiento.getPK_SALIDA
        .setPROCEDIMIENTO = cmbProcedimiento.getTEXTO
        
        .setUNIDADES_ID = 0 'cmbUnidad.getPK_SALIDA
        .setRANGO_MIN = 0
        .setRANGO_MAX = 0
        
        .setESTADO = IIf(optEstado(0).value, 0, IIf(optEstado(1).value, 1, 2))
        
        If .getID_AUX = enumIdAux.ID_AUX_EXISTE Then
            .setID_AUX = enumIdAux.ID_AUX_MODIFICADO
        End If
        
    End With
    
    If mvarenuTipoEdicion = ALTA Then
        mvarobjVerificacion.setFECHA_PREVISTA = mvarobjVerificacion.getFECHA_ACTUAL
        Call mvarobjEquipo.Verificaciones.Add(mvarobjVerificacion)
    ElseIf mvarenuTipoEdicion = EDICION Then
        Call mvarobjEquipo.Verificaciones.Replace(mvarobjVerificacion.getID_VERIFICACION, mvarobjVerificacion)
    End If
    
    
End Sub

Public Property Get resultado() As Boolean

    resultado = mvarblnResultado

End Property

Public Property Let resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

Private Sub optEstado_Click(Index As Integer)

If fraEstadoIntervencion.Enabled = False Then Exit Sub

If Index = 0 Then
    txtFechaActual.Enabled = True
    If mvarobjVerificacion.getFECHA_ACTUAL <> "" Then
        txtFechaActual.value = mvarobjVerificacion.getFECHA_ACTUAL
    End If
Else
    If Not bln_fecha_real_editable Then
        txtFechaActual.Enabled = False
    End If
    txtFechaActual.value = Now
    txtFechaActual_Change
End If
End Sub

Private Sub optResultadoCualidad_Click(Index As Integer)
    lista.SelectedItem.SubItems(ColsR.RESULTADO_CUALIDAD) = Index
    lista.SelectedItem.SubItems(ColsR.RESULTADO_MEDIA) = IIf(Index = 0, "NO CONFORME", "CONFORME")
    lista.SelectedItem.SubItems(ColsR.n_medidas) = "1"
End Sub

Private Sub TecladoNumerico_Change(ByVal res As String)
    grdResultados.Text = res
End Sub

Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, resultado As String, fecha As String, Conforme As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
If grdResultados.Row + 1 > filasR Then
    TecladoNumerico.Hide
    grdResultados.EditActive = False
    Exit Sub
End If

' si existe siguiente Fila, edita la siguiente fila

If (grdResultados.Row + 1) <= xR.UpperBound(1) Then
    If Not IsEmpty(xR(grdResultados.Row + 1, 0)) Then
        If Trim(xR(grdResultados.Row + 1, 0)) <> "" Then
            grdResultados.EditActive = False
            grdResultados.Row = grdResultados.Row + 1
            resultado = grdResultados.Text
            cabecera = xR(grdResultados.Row, 0)
            fecha = xR(grdResultados.Row, 1)
            grdResultados.EditActive = True
        End If
    ElseIf mvarlngNumParametrosResultados = 1 Then
        grdResultados.Row = 1
        Cerrar = True
        grdResultados.EditActive = False
    ElseIf grdResultados.Row = mvarlngNumParametrosResultados - 1 Or mvarlngNumParametrosResultados = 0 Then
        'grdResultados.EditActive = False
        'Resultado = grdResultados.Text
        'cabecera = xP(grdResultados.Row, 0)
        'grdResultados.EditActive = True
        Cerrar = True
    End If
Else
    If mvarlngNumParametrosResultados = 1 Then
        grdResultados.Row = 1
    Else
        grdResultados.Row = 0
    End If
    
    Cerrar = True
    grdResultados.EditActive = False
End If
End Sub


Private Sub tUnidades_DropDownClose()
    On Error Resume Next
    grdResultados.Columns(ColsR.id_unidad) = tUnidades.Columns(1)
    On Error GoTo 0
    
    xR(grdResultados.Row, ColsR.id_unidad) = tUnidades.Columns(1)
    grdResultados.Col = 3
        
End Sub

Private Sub txtDescripcion_Cualidad_Change()
'If grdResultados.Row < 0 Then Exit Sub
If lista.ListItems.Count = 0 Then Exit Sub

lista.SelectedItem.SubItems(ColsR.RESULTADOS_PATRON) = Trim(txtDescripcion_Cualidad.Text)

'xR(grdResultados.Row, ColsR.RESULTADOS_PATRON) = Trim(txtDescripcion_Cualidad.Text)
'On Error Resume Next
'grdResultados.Columns(ColsR.RESULTADOS_PATRON).RefetchCell grdResultados.Row
'On Error GoTo 0
End Sub

Private Sub txtFechaActual_Change()

If IsDate(txtFechaActual.value) Then
    txtFechaProxima_b.value = calcularFechaProxima(txtFechaActual.value, getDataComboSel(cmbPeriVerificacion))
    txtFechaProxima.Text = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
End If

End Sub

Private Sub txtLimitacionesUso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAnadirLimitacion_Click
End Sub

Public Property Get VieneDeCuaderno() As Boolean

    VieneDeCuaderno = mvarblnVieneDeCuaderno

End Property

Public Property Let VieneDeCuaderno(ByVal blnVieneDeCuaderno As Boolean)

    mvarblnVieneDeCuaderno = blnVieneDeCuaderno

End Property


Public Property Get idProcedmientoInicial() As Long

    idProcedmientoInicial = mvarlngidProcedmientoInicial

End Property

Public Property Let idProcedmientoInicial(ByVal lngidProcedmientoInicial As Long)

    mvarlngidProcedmientoInicial = lngidProcedmientoInicial

End Property

Private Sub txtNMedidas_Change(Index As Integer)

txtNMedidas(Index).Locked = True

Dim n_medidas_act As Integer ' n_medidas_actuales
Dim n_medidas As Integer ' n medidas propuestas
Dim x As Integer

If Not IsNumeric(txtNMedidas(Index).Text) Then
    txtNMedidas(Index).Text = "1"
    txtNMedidas(Index).SelStart = 0
    txtNMedidas(Index).SelLength = 1
End If
If Trim(txtNMedidas(Index).Text) = "0" Then
    txtNMedidas(Index).Text = "1"
    txtNMedidas(Index).SelStart = 0
    txtNMedidas(Index).SelLength = 1
End If

n_medidas_act = lista_medidas(Index).ListItems.Count

n_medidas = CInt(txtNMedidas(Index).Text)

If n_medidas = n_medidas_act Then Exit Sub ' no se modifican

If n_medidas > n_medidas_act Then ' se añaden filas
    For x = (n_medidas_act + 1) To n_medidas
        lista_medidas(Index).ListItems.Add , , "0"
    Next x
End If

If n_medidas < n_medidas_act Then ' se quitan filas
    For x = n_medidas_act To (n_medidas + 1) Step -1
        lista_medidas(Index).ListItems.Remove lista_medidas(Index).ListItems.Count
    Next x
End If

' guarda las medidas en su fila correspondiente
lista.SelectedItem.SubItems(ColsR.n_medidas) = n_medidas
txtNMedidas(Index).Locked = False

End Sub



Private Sub txtNMedidas_GotFocus(Index As Integer)
    txtNMedidas(Index).SelStart = 0
    txtNMedidas(Index).SelLength = Len(txtNMedidas(Index).Text)
End Sub

Private Sub txtNMedidas_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtValor(Index).SetFocus
    Else
        KeyAscii = KeyAscii_SoloNumerico(txtNMedidas(Index), KeyAscii)
    End If
End Sub

'Private Sub txtNMedidas_LostFocus(Index As Integer)
'Dim n_medidas_act As Integer ' n_medidas_actuales
'Dim n_medidas As Integer ' n medidas propuestas
'Dim x As Integer
'
'If Not IsNumeric(txtNMedidas(Index).Text) Then txtNMedidas(Index).Text = "1"
'If Trim(txtNMedidas(Index).Text) = "0" Then txtNMedidas(Index).Text = "1"
'
'n_medidas_act = lista_medidas(Index).ListItems.Count
'
'n_medidas = CInt(txtNMedidas(Index).Text)
'
'If n_medidas = n_medidas_act Then Exit Sub ' no se modifican
'
'If n_medidas > n_medidas_act Then ' se añaden filas
'    For x = (n_medidas_act + 1) To n_medidas
'        lista_medidas(Index).ListItems.Add , , "0"
'    Next x
'End If
'
'If n_medidas < n_medidas_act Then ' se quitan filas
'    For x = n_medidas_act To (n_medidas + 1) Step -1
'        lista_medidas(Index).ListItems.Remove lista_medidas(Index).ListItems.Count
'    Next x
'End If
'
'' guarda las medidas en su fila correspondiente
'xR(grdResultados.Row, ColsR.n_medidas) = n_medidas
'
'End Sub
'

Private Sub txtValor_Change(Index As Integer)
    lista_medidas(Index).ListItems(lista_medidas(Index).SelectedItem.Index).Text = txtValor(Index).Text
    
    txtValor(Index).Locked = True
        preguardar_medidas Index
    txtValor(Index).Locked = False
End Sub


Private Sub txtvalor_GotFocus(Index As Integer)


If Trim(txtValor(Index).Text) = "" Then
    If lista_medidas(Index).ListItems.Count > 0 Then
        txtValor(Index).Text = lista_medidas(Index).SelectedItem.Text
    End If
End If

txtValor(Index).SelStart = 0
txtValor(Index).SelLength = Len(txtValor(Index).Text)
End Sub


Private Sub txtvalor_KeyPress(Index As Integer, KeyAscii As Integer)
Dim NUMERO As String

    If KeyAscii = 13 Then ' Tecla Intro
        KeyAscii = 0
        
        ' Comprueba que sea un valor correcto numérico
        If Trim(txtValor(Index).Text) = "" Then
            NUMERO = "0"
        ElseIf Not IsNumeric(Trim(txtValor(Index).Text)) Then
            NUMERO = "0"
        Else
            NUMERO = Trim(txtValor(Index).Text)
        End If

        lista_medidas(Index).ListItems(lista_medidas(Index).SelectedItem.Index).Text = NUMERO
        preguardar_medidas Index
        
        
        If lista_medidas(Index).SelectedItem.Index + 1 > lista_medidas(Index).ListItems.Count Then
            ' estamos al final de la lista
            ' Vuelve a la primera
            Set lista_medidas(Index).SelectedItem = lista_medidas(Index).ListItems(1)
            lista_medidas_Click Index
        Else
            ' avanza una linea
            Set lista_medidas(Index).SelectedItem = lista_medidas(Index).ListItems(lista_medidas(Index).SelectedItem.Index + 1)
            lista_medidas_Click Index
        End If
    Else
        KeyAscii = KeyAscii_SoloDecimal(txtValor(Index), KeyAscii, True)
    End If
End Sub


Private Sub preguardar_medidas(Index As Integer)
    
    ' recoge los datos modificados y los guarda en su celda de resultados
    Dim cad As String, str_total As String
    Dim fila As Long
    Dim total As Double, total_fila As Double
    Dim n_medidas As Integer
    
    total = 0
    
    For fila = 1 To lista_medidas(Index).ListItems.Count
        total_fila = 0
        If Trim(lista_medidas(Index).ListItems(fila)) <> "" Then
            If IsNumeric(lista_medidas(Index).ListItems(fila)) Then
                total_fila = CDbl(Replace(lista_medidas(Index).ListItems(fila), ".", ","))
            End If
        End If
        total = total + total_fila
        cad = cad & ";" & lista_medidas(Index).ListItems(fila)
    Next fila
    
    If cad <> "" Then cad = Mid(cad, 2)
    n_medidas = 1
    If Trim(txtNMedidas(Index).Text) <> "" Then
        If IsNumeric(txtNMedidas(Index).Text) Then
            n_medidas = txtNMedidas(Index).Text
        End If
    End If
    
    total = (total / n_medidas)
    
    If InStr(1, CStr(CDbl(total)), ",") Then
        ' mide los decimales. Si son más de 6, los redondea a 6 decimales
        str_total = Split(CStr(CDbl(total)), ",")(0) & "," & Left(Split(CStr(CDbl(total)), ",")(1), 6)
    Else
        str_total = CStr(total)
    End If
    
    lista.SelectedItem.SubItems(ColsR.RESULTADOS_PATRON) = cad
    lista.SelectedItem.SubItems(ColsR.RESULTADO_MEDIA) = str_total
    
End Sub



