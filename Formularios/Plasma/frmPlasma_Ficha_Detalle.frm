VERSION 5.00
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmPlasma_Ficha_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Ficha de Plasma"
   ClientHeight    =   12345
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlasma_Ficha_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12345
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "DUREZA SHORE A"
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
      Height          =   1005
      Left            =   45
      TabIndex        =   38
      Top             =   10485
      Width           =   9510
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   39
         Top             =   630
         Width           =   6585
      End
      Begin pryCombo.miCombo cmbShoreA 
         Height          =   375
         Left            =   1440
         TabIndex        =   40
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REQUIREMENT:"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   42
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST:"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   41
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "THICKNESS"
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
      Height          =   1005
      Left            =   45
      TabIndex        =   33
      Top             =   9405
      Width           =   9510
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   34
         Top             =   630
         Width           =   6585
      End
      Begin pryCombo.miCombo cmbEspesor 
         Height          =   375
         Left            =   1440
         TabIndex        =   35
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   37
         Top             =   315
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REQUIREMENT:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   36
         Top             =   675
         Width           =   1245
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MACRO HARDNESS (ASTM E18:14)"
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
      Height          =   1050
      Left            =   45
      TabIndex        =   28
      Top             =   7245
      Width           =   9510
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   9
         Top             =   630
         Width           =   6585
      End
      Begin pryCombo.miCombo cmbMacro 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REQUIREMENT:"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   30
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST:"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   29
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MICRO HARDNESS (ASTM E 384:11)"
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
      Height          =   1005
      Left            =   45
      TabIndex        =   27
      Top             =   8325
      Width           =   9510
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   11
         Top             =   630
         Width           =   6585
      End
      Begin pryCombo.miCombo cmbmicro 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REQUIREMENT:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   32
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST:"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   31
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "TENSILE STRENGTH (ASTM D 638:10)"
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
      Height          =   1050
      Left            =   45
      TabIndex        =   24
      Top             =   6120
      Width           =   9510
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   255
         TabIndex        =   7
         Top             =   630
         Width           =   6585
      End
      Begin pryCombo.miCombo cmbTraccion 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   270
         Width           =   7620
         _ExtentX        =   13441
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REQUIREMENT:"
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   26
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST:"
         Height          =   195
         Index           =   24
         Left            =   90
         TabIndex        =   25
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "METALLOGRAPHIC EXAMINATION"
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
      Height          =   3525
      Left            =   45
      TabIndex        =   22
      Top             =   2565
      Width           =   9510
      Begin TrueDBGrid80.TDBGrid gridP 
         Height          =   2655
         Left            =   90
         TabIndex        =   5
         Top             =   765
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   4683
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "ID"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TEST"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "REQUIREMENT"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=7011"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6932"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=212"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=132"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         TabAction       =   2
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.bgcolor=&HD7D7D7&"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=11,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=15"
         _StyleDefs(48)  =   "Named:id=37:Normal"
         _StyleDefs(49)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(50)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(51)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(52)  =   "Named:id=38:Heading"
         _StyleDefs(53)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(55)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=38,.fontname=MS Sans Serif"
         _StyleDefs(57)  =   "Named:id=39:Footing"
         _StyleDefs(58)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=40:Selected"
         _StyleDefs(60)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(61)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(62)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(63)  =   "Named:id=41:Caption"
         _StyleDefs(64)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(65)  =   "Named:id=42:HighlightRow"
         _StyleDefs(66)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(67)  =   "Named:id=43:EvenRow"
         _StyleDefs(68)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=44:OddRow"
         _StyleDefs(70)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(71)  =   "Named:id=47:RecordSelector"
         _StyleDefs(72)  =   ":id=47,.parent=38"
         _StyleDefs(73)  =   "Named:id=50:FilterBar"
         _StyleDefs(74)  =   ":id=50,.parent=37"
      End
      Begin pryCombo.miCombo cmbMicroestructura 
         Height          =   375
         Left            =   1215
         TabIndex        =   4
         Top             =   315
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEST:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   11475
      Width           =   1365
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   11475
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   11475
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      Height          =   1815
      Left            =   45
      TabIndex        =   15
      Top             =   675
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   990
         Width           =   3225
      End
      Begin pryCombo.miCombo cmbFabricante 
         Height          =   375
         Left            =   1260
         TabIndex        =   1
         Top             =   630
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   661
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1260
         TabIndex        =   3
         Top             =   1395
         Width           =   3225
      End
      Begin pryCombo.miCombo cmbRecubrimiento 
         Height          =   375
         Left            =   1260
         TabIndex        =   0
         Top             =   270
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "METCO"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   19
         Top             =   1035
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricante"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   21
         Top             =   675
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recubrimiento"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre Fabric."
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   1425
         Width           =   1080
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Ficha de Plasma"
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
      TabIndex        =   18
      Top             =   45
      Width           =   2895
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Ficha de Plasma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   330
      Width           =   1935
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   9630
   End
End
Attribute VB_Name = "frmPlasma_Ficha_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Dim xP As New XArrayDB
Const filasP As Integer = 20
Const ColP As Integer = 2
Private Enum ColsP
    ENSAYO_ID = 0
    ENSAYO = 1
    REQUIREMENT = 2
End Enum

Private Sub cmbMicroestructura_change()
    If cmbMicroestructura.getTEXTO = "" Then
        inicializar_grid
    Else
        Dim oPFE As New clsPlasma_ficha_estructura
        Dim i As Integer
        i = 0
        Dim rs As ADODB.Recordset
        Set rs = oPFE.Listado(PK)
        If rs.RecordCount > 0 Then
            Do
                xP(i, ColsP.ENSAYO_ID) = CStr(rs(0))
                xP(i, ColsP.ENSAYO) = CStr(rs(1))
                If Not IsNull(rs(2)) Then
                    xP(i, ColsP.REQUIREMENT) = CStr(rs(2))
                Else
                    xP(i, ColsP.REQUIREMENT) = ""
                End If
                i = i + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oPFE = Nothing
    End If
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_PLASMA_ENSAYOS
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Tipo Ensayo Plasma " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim oPF As New clsPlasma_ficha
      Dim FICHA As Long
      With oPF
        .setRECUBRIMIENTO_ID = cmbRecubrimiento.getPK_SALIDA
        .setFABRICANTE_ID = cmbFabricante.getPK_SALIDA
        .setOMAT = txtDatos(0)
        .setMETCO = txtDatos(1)
        If cmbMicroestructura.getTEXTO = "" Then
            .setMICROESTRUCTURA = 0
        Else
            .setMICROESTRUCTURA = cmbMicroestructura.getPK_SALIDA
        End If
        If cmbTraccion.getTEXTO = "" Then
            .setTRACCION = 0
        Else
            .setTRACCION = cmbTraccion.getPK_SALIDA
        End If
        If cmbMacro.getTEXTO = "" Then
            .setMACRO_DUREZA = 0
        Else
            .setMACRO_DUREZA = cmbMacro.getPK_SALIDA
        End If
        If cmbmicro.getTEXTO = "" Then
            .setMICRO_DUREZA = 0
        Else
            .setMICRO_DUREZA = cmbmicro.getPK_SALIDA
        End If
        If cmbEspesor.getTEXTO = "" Then
            .setESPESOR = 0
        Else
            .setESPESOR = cmbEspesor.getPK_SALIDA
        End If
        If cmbShoreA.getTEXTO = "" Then
            .setSHOREA = 0
        Else
            .setSHOREA = cmbShoreA.getPK_SALIDA
        End If
        .setTRACCION_REQ = txtDatos(2)
        .setMACRO_DUREZA_REQ = txtDatos(3)
        .setMICRO_DUREZA_REQ = txtDatos(4)
        .setESPESOR_REQ = txtDatos(5)
        .setSHOREA_REQ = txtDatos(6)
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir una nueva ficha. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            FICHA = oPF.Insertar
            If FICHA > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_PLASMA_FICHAS
                    .setIDENTIFICADOR = FICHA
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar la ficha. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del ensayo."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            oPF.Modificar (PK)
            FICHA = PK
            With ohc
                .setTIPO = HC_TIPOS.HC_PLASMA_FICHAS
                .setIDENTIFICADOR = PK
                .setIDENTIFICADOR_TEXTO = txtDatos(0)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setMOTIVO = Trim(MOTIVO)
                .Insertar
            End With
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      ' Datos MicroEstructura
      Dim oPFE As New clsPlasma_ficha_estructura
      oPFE.Eliminar FICHA
      If cmbMicroestructura.getTEXTO <> "" Then
        For i = 0 To filasP
         If Not IsEmpty(xP(i, ColsP.ENSAYO_ID)) Then
          If Trim(xP(i, ColsP.ENSAYO_ID)) <> "" Then
            With oPFE
                .setFICHA_ID = FICHA
                .setENSAYO_ID = xP(i, ColsP.ENSAYO_ID)
                .setREQUIREMENT = xP(i, ColsP.REQUIREMENT)
                .Insertar
            End With
          End If
         End If
        Next
      End If
      Set oPFE = Nothing
      Me.MousePointer = 0
      If PK = 0 Then
          MsgBox "La ficha se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      Else
          MsgBox "La ficha se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Ficha_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo
    inicializar_grid
    If PK <> 0 Then
        lbltitulo = "Modificación de Ficha de Plasma"
        cargar_ficha
    Else
        lbltitulo = "Alta de Ficha de Plasma"
    End If
End Sub
Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbRecubrimiento, DECODIFICADORA.DECODIFICADORA_PLASMA_RECUBRIMIENTOS
    oDeco.cargar_mi_combo cmbFabricante, DECODIFICADORA.DECODIFICADORA_PLASMA_FABRICANTES
    
    llenar_combo cmbMicroestructura, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 1 "
    llenar_combo cmbTraccion, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 2 "
    llenar_combo cmbMacro, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 3 "
    llenar_combo cmbmicro, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 4 "
    llenar_combo cmbEspesor, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 5 "
    llenar_combo cmbShoreA, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 6 "
    
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_ficha()
    Dim i As Integer
    Dim oPF As New clsPlasma_ficha
    If oPF.Carga(PK) = True Then
        With oPF
            cmbRecubrimiento.MostrarElemento .getRECUBRIMIENTO_ID
            cmbFabricante.MostrarElemento .getFABRICANTE_ID
            txtDatos(0) = .getOMAT
            txtDatos(1) = .getMETCO
            
            cmbMicroestructura.MostrarElemento .getMICROESTRUCTURA
            cmbTraccion.MostrarElemento .getTRACCION
            cmbMacro.MostrarElemento .getMACRO_DUREZA
            cmbmicro.MostrarElemento .getMICRO_DUREZA
            cmbEspesor.MostrarElemento .getESPESOR
            cmbShoreA.MostrarElemento .getSHOREA
            
            txtDatos(2) = .getTRACCION_REQ
            txtDatos(3) = .getMACRO_DUREZA_REQ
            txtDatos(4) = .getMICRO_DUREZA_REQ
            txtDatos(5) = .getESPESOR_REQ
            txtDatos(6) = .getSHOREA_REQ
        End With
    End If
    Set oPF = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If cmbRecubrimiento.getTEXTO = "" Then
        MsgBox "Debe indicar el Recubrimiento de la Ficha.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe indicar el RR Omat.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe indicar el METCO.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbMicroestructura.getTEXTO = "" And cmbTraccion.getTEXTO = "" And cmbMacro.getTEXTO = "" And cmbmicro.getTEXTO = "" And cmbShoreA.getTEXTO = "" Then
        MsgBox "Debe indicar al menos un Ensayo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function
Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error

    gridP.Col = 0
    gridP.Row = 0
    xP.Clear
    xP.ReDim 0, filasP, 0, ColP
    xP.Clear
    Set gridP.Array = xP
    gridP.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

