VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmEquipoEdicionVerificacion_old 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8685
   ClientLeft      =   2955
   ClientTop       =   2490
   ClientWidth     =   11400
   ClipControls    =   0   'False
   Icon            =   "frmEquipoEdicionVerificacion_old.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   900
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7770
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   5
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
      Height          =   7215
      Left            =   45
      TabIndex        =   6
      Top             =   510
      Width           =   11340
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "01/01/1900"
         Top             =   690
         Width           =   1785
      End
      Begin VB.TextBox txtEvaluacionResultado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   46
         Top             =   2775
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7620
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Ver norma"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":0261
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Buscar documento"
         Top             =   2760
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7245
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":04D2
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Eliminar documento"
         Top             =   2760
         Width           =   360
      End
      Begin VB.CommandButton cmdEscanearEvaluacion 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":0666
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Escanear documento"
         Top             =   2760
         Width           =   405
      End
      Begin TrueDBGrid80.TDBDropDown tUnidades 
         Height          =   930
         Left            =   5250
         TabIndex        =   40
         Top             =   5460
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1640
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=265"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=185"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   2
         BorderStyle     =   1
         ColumnHeaders   =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   ""
         DataField       =   ""
         IntegralHeight  =   -1  'True
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   16777215
         ValueTranslate  =   0   'False
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=124,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0E0FF&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H8080FF&,.bold=0"
         _StyleDefs(11)  =   ":id=2,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(37)  =   ":id=32,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(38)  =   ":id=32,.fontname=MS Sans Serif"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin VB.CommandButton cmdEscanearHoja 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":0A20
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Escanear documento"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdEscanearCert 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6840
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":0DDA
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Escanear documento"
         Top             =   2400
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5580
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Eliminar accesorio"
         Top             =   3240
         Width           =   285
      End
      Begin VB.CommandButton cmdAnadirLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5250
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":1328
         Style           =   1  'Graphical
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   3210
         Width           =   3555
      End
      Begin VB.Frame fraEstadoIntervencion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resultado Verificación "
         Height          =   1245
         Left            =   9120
         TabIndex        =   29
         Top             =   1530
         Width           =   2115
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Previsto"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   540
            Width           =   1605
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cerrado No Conforme"
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   30
            Top             =   840
            Width           =   1875
         End
      End
      Begin VB.ListBox lstLimitacionesUso 
         Appearance      =   0  'Flat
         Height          =   1395
         ItemData        =   "frmEquipoEdicionVerificacion_old.frx":154D
         Left            =   1650
         List            =   "frmEquipoEdicionVerificacion_old.frx":1554
         TabIndex        =   28
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
         TabIndex        =   24
         Top             =   2055
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarHojaCal 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7620
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":156C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ver norma"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarHojaCal 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":17C1
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Buscar documento"
         Top             =   2040
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7245
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":1A32
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Eliminar documento"
         Top             =   2400
         Width           =   360
      End
      Begin VB.CommandButton cmdAdjuntarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   6435
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":1BC6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar documento"
         Top             =   2400
         Width           =   405
      End
      Begin VB.CommandButton cmdMostrarCertificado 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7620
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":1E37
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   2415
         Width           =   4770
      End
      Begin MSComCtl2.DTPicker txtFechaActual 
         Height          =   405
         Left            =   9480
         TabIndex        =   2
         Top             =   240
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
         Format          =   16515073
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFechaProxima_b 
         Height          =   405
         Left            =   9480
         TabIndex        =   3
         Top             =   1095
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
         Format          =   16515073
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
         TabIndex        =   14
         Top             =   990
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbProcedimiento 
         Height          =   330
         Left            =   1650
         TabIndex        =   25
         Top             =   1710
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbVerificadorExterno 
         Height          =   330
         Left            =   1650
         TabIndex        =   36
         Top             =   1350
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdEliminarHojaCal 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   7245
         Picture         =   "frmEquipoEdicionVerificacion_old.frx":208C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Eliminar documento"
         Top             =   2040
         Width           =   360
      End
      Begin TrueDBGrid80.TDBGrid grdResultados 
         Height          =   2220
         Left            =   30
         TabIndex        =   41
         Top             =   4950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   3916
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Descripción"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "Standard"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Rango Min."
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Rango Max."
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "General Number"
         Columns(2).DropDown=   "tResponsables"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   1
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Unidad"
         Columns(3).DataField=   ""
         Columns(3).DropDown=   "tUnidades"
         Columns(3).DropDown.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Resultado"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "General Number"
         Columns(4).ExternalEditor=   "TDBDate1"
         Columns(4).ExternalEditor.vt=   8
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Tolerancia"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "General Number"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Incertidumbre"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "General Number"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Corrección"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "General Number"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "ID_RESULTADO"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "ID_UNIDAD"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   1
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5292"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5212"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1931"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1852"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2540"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2461"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(25)=   "Column(3).Button=1"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(3).AutoDropDown=1"
         Splits(0)._ColumnProps(28)=   "Column(4).Width=1931"
         Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1852"
         Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(34)=   "Column(5).Width=1931"
         Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=1852"
         Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(40)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(41)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(43)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(46)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(49)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(52)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(53)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(55)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(56)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(57)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(58)=   "Column(9).Width=2566"
         Splits(0)._ColumnProps(59)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(9)._WidthInPix=2487"
         Splits(0)._ColumnProps(61)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(62)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(63)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   3
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Resultados Verificación"
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   2
         DirectionAfterTab=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
         _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H8000000A&,.bold=0"
         _StyleDefs(14)  =   ":id=3,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
         _StyleDefs(23)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
         _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1,.transparentBmp=0,.fgpicPosition=7,.bgpicMode=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.fgcolor=&HFF&,.bold=-1"
         _StyleDefs(37)  =   ":id=24,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
         _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
         _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=11,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=36,.parent=11,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=33,.parent=12"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=34,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=35,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=54,.parent=11,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
         _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=28,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
         _StyleDefs(57)  =   ":id=28,.locked=0"
         _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=12"
         _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=13"
         _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=58,.parent=11,.alignment=2"
         _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
         _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
         _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=62,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=12"
         _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=13"
         _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=66,.parent=11,.alignment=2,.locked=0"
         _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=12"
         _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=13"
         _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=74,.parent=11"
         _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=12"
         _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=13"
         _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=70,.parent=11"
         _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=12"
         _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=13"
         _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=15"
         _StyleDefs(81)  =   "Named:id=37:Normal"
         _StyleDefs(82)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(83)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(84)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(85)  =   "Named:id=38:Heading"
         _StyleDefs(86)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(87)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(88)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(89)  =   ":id=38,.fontname=MS Sans Serif"
         _StyleDefs(90)  =   "Named:id=39:Footing"
         _StyleDefs(91)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(92)  =   "Named:id=40:Selected"
         _StyleDefs(93)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(94)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(95)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(96)  =   "Named:id=41:Caption"
         _StyleDefs(97)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(98)  =   "Named:id=42:HighlightRow"
         _StyleDefs(99)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(100) =   "Named:id=43:EvenRow"
         _StyleDefs(101) =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(102) =   "Named:id=44:OddRow"
         _StyleDefs(103) =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(104) =   "Named:id=47:RecordSelector"
         _StyleDefs(105) =   ":id=47,.parent=38"
         _StyleDefs(106) =   "Named:id=50:FilterBar"
         _StyleDefs(107) =   ":id=50,.parent=37"
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Eval. Resultado"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   47
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
         TabIndex        =   37
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   15
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
         TabIndex        =   13
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próx. Verificación"
         Height          =   195
         Index           =   0
         Left            =   7920
         TabIndex        =   12
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Actual Verificación"
         Height          =   195
         Index           =   10
         Left            =   7920
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Ver. Interna"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   10
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7815
      Width           =   1050
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10905
      Picture         =   "frmEquipoEdicionVerificacion_old.frx":2220
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
      TabIndex        =   9
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
      Width           =   11415
   End
End
Attribute VB_Name = "frmEquipoEdicionVerificacion_old"
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
Private xUnidades As New XArrayDB

Const filasR As Integer = 50
Const ColR As Integer = 9
Private Enum ColsR
    Descripcion = 0
    RANGO_MIN = 1
    RANGO_MAX = 2
    Unidad = 3
    RESULTADO_CAL = 4
    TOLERANCIA = 5
    INCERTIDUMBRE = 6
    CORRECCION = 7
    Id_resultado = 8
    id_unidad = 9
End Enum

Private mvarlngNumParametrosResultados As Long
Private mvarlngidProcedmientoInicial As Long
Private Sub ConfigurarTablet()
    Set TecladoNumerico = New frmTecladoNumerico
    
    
    TecladoNumerico.OcultarConformidad = True
    
    blnEsTablet = pc_es_tablet
    
    If blnEsTablet Then
        
        blnPrimeraVez = True
        
        grdResultados.Columns(ColsR.RESULTADO_CAL).Locked = True
        Me.Top = 0
        

    End If
End Sub

Private Sub CargarComboGridUnidad()
    Dim rs As ADODB.RecordSet
    Dim ote As New clsUnidades

    Set rs = ote.Listado()
    xUnidades.Clear
    If rs.RecordCount > 0 Then
        xUnidades.ReDim 0, rs.RecordCount, 0, 1
        Dim i As Integer
        i = 1
        Do
            xUnidades(i, 0) = CStr(rs("NOMBRE"))
            xUnidades(i, 1) = CStr(rs("ID_UNIDAD"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xUnidades.ReDim 0, 0, 0, 1
    End If
    Set tUnidades.Array = xUnidades
    tUnidades.Refresh
    grdResultados.Refresh
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
    
        grdResultados.Enabled = False
        cmdok.Visible = False
    End If
End Sub

Private Sub cmbPeriVerificacion_Click(AREA As Integer)

    Call txtFechaActual_Change

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
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarCertificado_Click of Formulario frmEquipoEdicionVerificacion"
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
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarEvaluacion_Click of Formulario frmEquipoEdicionVerificacion"
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
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarHojaCal_Click of Formulario frmEquipoEdicionVerificacion"
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
    If Not ComprobarDatos Then Exit Sub
    
    RecogerDatos
    
    If mvarenuTipoEdicion = ALTA Then
        mvarobjVerificacion.setEQUIPO_ID = mvarlngidEquipo
        lngId_Verificacion = mvarobjVerificacion.Insertar(True, xR, filasR)
    Else
        lngId_Verificacion = CLng(mvarstrId)
        Call mvarobjVerificacion.Modificar(lngId_Verificacion, True, , xR, filasR)
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
    
    ' la fecha proxima no es modificable
    'If txtFechaActual.value >= txtFechaProxima_b.value Then
    '    strMs = strMs & vbCrLf & " - La fecha de la próxima verificación no puede ser anterior a la de la Verificación actual"
    'End If

    
    If Trim(strMs) <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strMs
        Exit Function
    End If

    ComprobarDatos = True

On Error GoTo 0
    Exit Function
ComprobarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmEquipoEdicionVerificacion"
End Function


Private Sub comprobarDatosResultados(ByRef strMs As String)

    Dim i As Long
    i = 0
    Dim cad As String
    
    cad = ""
    grdResultados.Refresh
    
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

Private Sub Form_Activate()
    
    If blnPrimeraVez Then
        grdResultados_BeforeColEdit ColsR.RESULTADO_CAL, 0, 0
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
Call inicializar_grid
Call CargarComboGridUnidad

Call PresentarDatos_ParametrosResultados

blnPrimeraVez = False
    
Call ConfigurarTablet

If mvarenuTipoEdicion = ALTA Then

    'txtFechaActual.value = mvardtmFechaProximaInicial
    txtFechaActual.value = Now
    txtFechaActual_Change
    txtFechaActual.Enabled = bln_fecha_real_editable Or True
    'txtFechaProxima_b.value = calcularFechaProxima(mvardtmFechaProximaInicial, mvarlngIdPeriodoInicial)
    Set mvarobjVerificacion = New clsEquipoVerificacion
    cmbTipoVerificacion.BoundText = mvarlngIdTipoVerificacionIncial
    cmbPeriVerificacion.BoundText = mvarlngIdPeriodoInicial
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
If blnEsTablet And ColIndex = ColsR.RESULTADO_CAL Then
    grdResultados.Col = ColIndex
    TecladoNumerico.TextoInicial = grdResultados.Text
    TecladoNumerico.cabecera = xR(grdResultados.Row, 0)
    TecladoNumerico.Subcabecera = "Resultado" 'xP(gridP.Row, 1)
    
    TecladoNumerico.Show 1
    grdResultados.EditActive = False
    
End If
End Sub

Private Sub grdResultados_KeyPress(KeyAscii As Integer)
    
    With grdResultados
        If .Col = 1 Or .Col = 2 Or .Col = 4 Or .Col = 5 Or .Col = 6 Or .Col = 7 Then
            KeyAscii = KeyAscii_SoloDecimal_tbgrid(.Text, KeyAscii, True)
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
    

On Error GoTo 0
Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmEquipoEdicionVerificacion"
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
                txtFechaProxima_b.value = CDate(.getFECHA_PROXIMA)
                txtFechaProxima.Text = Format(txtFechaProxima_b.value, "dd/mm/yyyy")
            'End If
        End If
        
        optEstado(.getESTADO).value = True
    End With


    Call PresentarDatos_Adjuntos

On Error GoTo 0
    Exit Sub
PresentarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmEquipoEdicionVerificacion"

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
On Error GoTo PresentarDatos_ParametrosResultados_Error

    i = 0
    
    If mvarenuTipoEdicion <> ALTA Then
        ' Carga los Parametros de la verificacion
        Set rs = mvarobjVerificacion.DevolverParametrosResultados(mvarstrId)
    Else
        ' Carga los Parametros del Equipo
        Set rs = mvarobjEquipo.DevolverParametrosResultadosEquipoVerificacion(CStr(mvarlngidEquipo))
    End If
    
    If rs.RecordCount > 0 Then
        Do
            xR(i, ColsR.Descripcion) = CStr(rs("descripcion"))
            xR(i, ColsR.RANGO_MIN) = CStr(rs("rango_min"))
            xR(i, ColsR.RANGO_MAX) = CStr(rs("rango_max"))
            xR(i, ColsR.Unidad) = CStr(rs("unidad"))
            xR(i, ColsR.id_unidad) = CStr(rs("unidad_ID"))
            
            xR(i, ColsR.RESULTADO_CAL) = CStr(rs("resultado"))
            xR(i, ColsR.TOLERANCIA) = CStr(rs("tolerancia_max"))
            xR(i, ColsR.INCERTIDUMBRE) = CStr(rs("incertidumbre"))
            xR(i, ColsR.CORRECCION) = CStr(rs("correccion"))
            xR(i, ColsR.Id_resultado) = CStr(rs("id_resultado"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    mvarlngNumParametrosResultados = i
    grdResultados.Refresh
    grdResultados.Enabled = True


On Error GoTo 0
    Exit Sub
PresentarDatos_ParametrosResultados_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_ParametrosResultados of Formulario frmEquipoEdicionVerificacion"

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
    grdResultados.Columns(ColsR.id_unidad) = tUnidades.Columns(1)
    xR(grdResultados.Row, ColsR.id_unidad) = tUnidades.Columns(1)
    grdResultados.Col = 3
        
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
