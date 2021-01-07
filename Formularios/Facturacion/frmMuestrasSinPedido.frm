VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmMuestrasSinPedido 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Muestras Sin Pedido Asociado"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   13680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMuestrasSinPedido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   30
      TabIndex        =   4
      Top             =   630
      Width           =   13575
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   180
         TabIndex        =   16
         Top             =   960
         Width           =   285
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9750
         TabIndex        =   7
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   945
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
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
         Height          =   255
         Left            =   9750
         TabIndex        =   5
         Top             =   585
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   2160
         TabIndex        =   8
         Top             =   960
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
         Format          =   59244545
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4095
         TabIndex        =   9
         Top             =   960
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   59244545
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1530
         TabIndex        =   10
         Top             =   225
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTiposMuestra 
         Height          =   330
         Left            =   1530
         TabIndex        =   11
         Top             =   585
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionadas desde"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   14
         Top             =   1005
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   13
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   675
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7530
      Width           =   1050
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   2070
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   9551
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tArticulos"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cliente"
      Columns(1).DataField=   ""
      Columns(1).NumberFormat=   "Currency"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tipo Análisis/Solución"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "General Number"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Ref. Cliente"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Número"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Pedido"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2011"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=3810"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3731"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=5054"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4974"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=4657"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4577"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1746"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1667"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=6138"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6059"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.locked=-1,.bold=0"
      _StyleDefs(37)  =   ":id=24,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.alignment=0,.fgcolor=&HFF&"
      _StyleDefs(45)  =   ":id=28,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=0"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=36,.parent=11,.alignment=0"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=54,.parent=11,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=12"
      _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=58,.parent=11,.alignment=0"
      _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
      _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
      _StyleDefs(65)  =   "Named:id=37:Normal"
      _StyleDefs(66)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(67)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(68)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(69)  =   "Named:id=38:Heading"
      _StyleDefs(70)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   ":id=38,.wraptext=-1"
      _StyleDefs(72)  =   "Named:id=39:Footing"
      _StyleDefs(73)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   "Named:id=40:Selected"
      _StyleDefs(75)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(76)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(77)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(78)  =   "Named:id=41:Caption"
      _StyleDefs(79)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(80)  =   "Named:id=42:HighlightRow"
      _StyleDefs(81)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(82)  =   "Named:id=43:EvenRow"
      _StyleDefs(83)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(84)  =   "Named:id=44:OddRow"
      _StyleDefs(85)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(86)  =   "Named:id=47:RecordSelector"
      _StyleDefs(87)  =   ":id=47,.parent=38"
      _StyleDefs(88)  =   "Named:id=50:FilterBar"
      _StyleDefs(89)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Asignación de precios a códigos tarifarios"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   330
      Width           =   11805
      WordWrap        =   -1  'True
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13050
      Picture         =   "frmMuestrasSinPedido.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Muestras sin Pedido Asociado"
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
      TabIndex        =   2
      Top             =   30
      Width           =   4335
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   13635
   End
End
Attribute VB_Name = "frmMuestrasSinPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Dim x As New XArrayDB

Const filas As Integer = 1000
Const Col As Integer = 8
Private Enum COLS
    CODIGO = 0
    cliente = 1
    tipoanalisis = 2
    REFERENCIA = 3
    NUMERO = 4
    PEDIDO = 5
    ID_MUESTRA = 6
    ID_PEDIDO = 7
End Enum

Private Sub chkFecha_Click()
    If chkFecha.value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub chkTodas_Click()
    If chkTodas.value = Checked Then
        cmbTiposMuestra.Limpiar
        cmbTiposMuestra.desactivar
    Else
        cmbTiposMuestra.activar
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbClientes.Limpiar
        cmbClientes.desactivar
    Else
        cmbClientes.activar
    End If

End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    fdesde = Date
    fhasta = Date
    cargar_combos
    inicializar_grid filas
End Sub
Public Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
End Sub

Private Sub cargar_lista()
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strClientes As String
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim f_desde As String
    Dim f_hasta As String
    Dim REFERENCIA As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.value = Unchecked Then
        If cmbTiposMuestra.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA
    End If
    ' Clientes
    strClientes = ""
    If chkTodos.value = Unchecked Then
        If cmbClientes.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        strClientes = " AND mu.cliente_id = " & cmbClientes.getPK_SALIDA
    End If
    ' Fechas
    Dim FECHA_DESDE As String
    FECHA_DESDE = " AND mu.fecha_recepcion>='" & f_desde & "'"
    Dim FECHA_HASTA As String
    FECHA_HASTA = " AND mu.fecha_recepcion<='" & f_hasta & "'"

    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      FECHA_DESDE & FECHA_HASTA & _
                      strMuestra & _
                      strClientes & _
                      " order by mu.id_general desc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    lblsubtitulo = "Registros encontrados : " & rs.RecordCount
'    inicializar_grid rs.RecordCount
'    x.Clear
'    x.ReDim 0, rs.RecordCount, 0, Col
'    x.Clear
    If rs.RecordCount >= 1 Then
    
        i = 0
        While Not rs.EOF
            x(i, COLS.CODIGO) = CStr(rs(1))
            x(i, COLS.cliente) = CStr(rs(2))
            x(i, COLS.tipoanalisis) = CStr(rs(8))
            x(i, COLS.REFERENCIA) = CStr(rs(4))
            x(i, COLS.NUMERO) = CStr(Format(rs.Fields(9), "00000"))
            
            i = i + 1
            rs.MoveNext
        Wend
    End If
    Me.MousePointer = 0
'    Set grid.Array = x
    grid.Refresh
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
End Sub
Private Sub inicializar_grid(registros As Integer)
   On Error GoTo inicializar_grid_Error
   
    grid.Col = 0
    grid.Row = 0
    x.Clear
    x.ReDim 0, registros, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

