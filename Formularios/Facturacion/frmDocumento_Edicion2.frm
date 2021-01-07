VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmDocumento_Edicion2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo Documento"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13365
   Icon            =   "frmDocumento_Edicion2.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Muestra"
      Height          =   930
      Left            =   2340
      Picture         =   "frmDocumento_Edicion2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Recalcula el precio de la muestra seleccionada"
      Top             =   7515
      Width           =   1065
   End
   Begin VB.CommandButton cmdinsertar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inserta Línea"
      Height          =   930
      Left            =   90
      Picture         =   "frmDocumento_Edicion2.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7515
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   930
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7485
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   930
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7485
      Width           =   1155
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar &Línea"
      Height          =   930
      Left            =   1215
      Picture         =   "frmDocumento_Edicion2.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7515
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   7380
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   13018
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NºGeneral"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "General Date"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NºParticular"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tipo Análisis"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Referencia Cliente"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Código"
      Columns(6).DataField=   ""
      Columns(6).ConvertEmptyCell=   1
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Precio"
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Currency"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).ShowCollapseExpandIcons=   0   'False
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1826"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1746"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1).DropDownList=1"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1905"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1826"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2223"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2143"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=6615"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=6535"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=6509"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6429"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=2408"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2328"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=847"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=767"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=2"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.namedParent=38"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=72,.parent=6,.namedParent=40"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=71,.parent=7,.namedParent=40"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.namedParent=43"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.namedParent=44"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=45"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=78,.parent=48"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=67"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=67,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=68"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=69"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=71"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=90,.parent=67,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(45)  =   ":id=90,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(46)  =   ":id=90,.fontname=MS Sans Serif"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=71,.bold=0,.fontsize=975"
      _StyleDefs(50)  =   ":id=89,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(51)  =   ":id=89,.fontname=MS Sans Serif"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=94,.parent=67,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=68"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=69"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=71"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=68"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=69"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=71"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=102,.parent=67,.alignment=0"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=68"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=69"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=71"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=106,.parent=67,.alignment=2,.locked=0"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=103,.parent=68"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=104,.parent=69"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=105,.parent=71"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=110,.parent=67,.alignment=1"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=107,.parent=68"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=108,.parent=69"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=109,.parent=71"
      _StyleDefs(72)  =   "Named:id=37:Normal"
      _StyleDefs(73)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(74)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(75)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(76)  =   "Named:id=38:Heading"
      _StyleDefs(77)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=38,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=39:Footing"
      _StyleDefs(80)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=40:Selected"
      _StyleDefs(82)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(83)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(84)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(85)  =   "Named:id=41:Caption"
      _StyleDefs(86)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(87)  =   "Named:id=42:HighlightRow"
      _StyleDefs(88)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(89)  =   "Named:id=43:EvenRow"
      _StyleDefs(90)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(91)  =   "Named:id=44:OddRow"
      _StyleDefs(92)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(93)  =   "Named:id=47:RecordSelector"
      _StyleDefs(94)  =   ":id=47,.parent=38"
      _StyleDefs(95)  =   "Named:id=50:FilterBar"
      _StyleDefs(96)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lbliva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11160
      TabIndex        =   8
      Top             =   7785
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dto."
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
      Height          =   345
      Index           =   1
      Left            =   9675
      TabIndex        =   7
      Top             =   7785
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11160
      TabIndex        =   6
      Top             =   8115
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   315
      Index           =   2
      Left            =   9675
      TabIndex        =   5
      Top             =   8115
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   345
      Index           =   0
      Left            =   9675
      TabIndex        =   2
      Top             =   7455
      Width           =   1455
   End
   Begin VB.Label lblbase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11160
      TabIndex        =   1
      Top             =   7455
      Width           =   2160
   End
End
Attribute VB_Name = "frmDocumento_Edicion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_DOCUMENTO As Long
Dim x As New XArrayDB
Const filas As Integer = 2000
Const Col As Integer = 7
Private Enum Cols
    ID = 0
    general = 1
    fecha = 2
    particular = 3
    tipoanalisis = 4
    referencia = 5
    codigo = 6
    PRECIO = 7
End Enum
Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    Me.MousePointer = 11
    Dim odoc As New clsDocs_pago
    ' Informar las muestras a no facturadas
    Dim omuestra As New clsMuestra
    Dim oDocumento_Detalle As New clsDocs_pago_muestras
    Dim rs As ADODB.RecordSet
    Set rs = oDocumento_Detalle.MuestrasDocumento(CInt(PK_DOCUMENTO))
    Dim sgrupo As String
    If rs.RecordCount <> 0 Then
        Do
            sgrupo = sgrupo & rs("muestra_id") & ","
            rs.MoveNext
        Loop Until rs.EOF
        sgrupo = Left(sgrupo, Len(sgrupo) - 1)
        omuestra.Informar_Documentos_Pago sgrupo, 0
    End If
    ' Detalle del documento
    If PK_DOCUMENTO <> 0 Then
        oDocumento_Detalle.EliminarMuestras PK_DOCUMENTO
    End If
    ' Log completo salida
    Dim i As Integer
'    For i = x.LowerBound(1) To x.UpperBound(1)
'        If Trim(x.value(i, Cols.general)) <> "" Or Trim(x.value(i, Cols.tipoanalisis)) <> "" Then
'            log x.value(i, Cols.ID) & ";" & x.value(i, Cols.general) & ";" & x.value(i, Cols.fecha) & ";" & _
'                x.value(i, Cols.particular) & ";" & x.value(i, Cols.tipoanalisis) & ";" & _
'                x.value(i, Cols.REFERENCIA) & ";" & x.value(i, Cols.codigo) & ";" & _
'                x.value(i, Cols.PRECIO)
'        End If
'    Next
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.value(i, Cols.general)) <> "" Or Trim(x.value(i, Cols.tipoanalisis)) <> "" Then
            With oDocumento_Detalle
                .setDOC_ID = PK_DOCUMENTO
                .setORDEN = i
                .setABONADO = 0
                If Trim(x.value(i, Cols.ID)) = "" Then
                    .setMUESTRA_ID = 0
                Else
                    .setMUESTRA_ID = x.value(i, Cols.ID)
                End If
                If Trim(x.value(i, Cols.fecha)) = "" Then
                    .setFECHA = "1900-01-01"
                Else
                    .setFECHA = Format(x.value(i, Cols.fecha), "yyyy-mm-dd")
                End If
                .setTIPO_ANALISIS = x.value(i, Cols.tipoanalisis)
                .setREFERENCIA_CLIENTE = x.value(i, Cols.referencia)
                .setCODIGO = x.value(i, Cols.codigo)
                .setPRECIO = moneda_bd(x.value(i, Cols.PRECIO))
                .Insertar_linea
            End With
            ' Modificamos los datos referentes a la muestra
            If Not IsEmpty(x.value(i, Cols.ID)) And Trim(x.value(i, Cols.ID)) <> "" Then
                If CLng(x.value(i, Cols.ID)) <> 0 Then
                    With omuestra
                        .setPRECIO = moneda_bd(x.value(i, Cols.PRECIO))
                        If cmbpedido.Text = "" Then
                            .setPEDIDO_ID = 0
                        Else
                            .setPEDIDO_ID = cmbpedido.BoundText
                        End If
                        .setDOCUMENTO_PAGO = 2
                        .setCLIENTE_ID = cmbclientes.getPK_SALIDA
                        .Informar_Datos_documento CLng(x.value(i, Cols.ID))
                        ' Modificar el documento de pago de la muestra
'                        oMuestra.Informar_Documento_Pago CLng(x.value(i, Cols.ID)), 2
'                        If cmbpedido.Text = "" Then
'                            oMuestra.Informar_Pedido CLng(x.value(i, Cols.ID)), 0
'                        Else
'                            oMuestra.Informar_Pedido CLng(x.value(i, Cols.ID)), cmbpedido.BoundText
'                        End If
'                        oMuestra.Informar_Cliente CLng(x.value(i, Cols.ID)), cmbClientes.getPK_SALIDA
'                        oMuestra.actualizar_precio CLng(x.value(i, Cols.ID)), moneda_bd(x.value(i, Cols.PRECIO))
                    End With
                End If
            End If
        End If
    Next
    ' Informar el total de factura
    odoc.Informar_total_factura CInt(PK_DOCUMENTO)
    log ("Documento insertado correctamente.")
    Me.MousePointer = 0
    MsgBox "El documento se ha almacenado correctamente.", vbInformation, App.Title
    Unload Me
    Exit Sub
fallo:
    Me.MousePointer = 0

    MsgBox "Error al guardar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub cmdborrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Col
        grid.SelBookmarks.Add grid.Bookmark
        x(grid.Bookmark, i) = ""
        grid.SelBookmarks.Remove 0
    Next
    grid.Refresh
    calcular_total
    grid.SetFocus
End Sub
Private Sub cmdinsertar_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    For i = filas To grid.Bookmark Step -1
        For j = 0 To Col
            x(i + 1, j) = x(i, j)
        Next
    Next
    For j = 0 To Col
        x(grid.Bookmark, j) = ""
    Next
    grid.Refresh
    calcular_total
    grid.SetFocus
End Sub

Private Sub cmdMuestra_Click()
    
    If Not IsEmpty(x.value(grid.Bookmark, Cols.ID)) Then
        If CLng(x.value(grid.Bookmark, Cols.ID)) <> 0 Then
            gmuestra = CLng(x.value(grid.Bookmark, Cols.ID))
            frmVerMuestra.Show 1
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Unload Me
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' ESC
            cmdSalir_Click
        Case 121 ' F10
            cmdAceptar_Click
    End Select
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    inicializar_ventana
    cargar_documento
    ' Verificar si esta contabilidado
    Dim odoc As New clsDocs_pago
    If odoc.esta_contabilidado(CInt(PK_DOCUMENTO)) Then
        cmdAceptar.Enabled = False
        MsgBox "El documento se encuentra contabilizado. No se puede editar.", vbInformation, App.Title
    End If
'    Set Prospective = grid.Styles.Add("Prospective")
'    Prospective.Font.Italic = True
'    Prospective.ForeColor = vbBlue
End Sub

Public Sub calcular_total()
    Dim i As Integer
    On Error Resume Next
    Dim total As Single
    total = 0
    For i = 0 To filas
        If Trim(CStr(x(i, Cols.PRECIO))) <> "" And Trim(CStr(x(i, Cols.fecha))) <> "" And CInt(CStr(x(i, Cols.ID))) <> 0 Then
            total = total + CSng(CStr(x.value(i, Cols.PRECIO)))
        End If
    Next
    lblbase = Format(total, "#,##0.00")
    Dim dto As Currency
    If txtdescuento.Text <> "" Then
        dto = Format((CCur(lblbase) * CInt(txtdescuento.Text) / 100), "#,##0.00")
    Else
        dto = 0
    End If
'    dto = dto + Format(((CCur(lblbase) - dto) * CInt(txtdescuento.Text) / 100), "#,##0.00")
'    dto = Format((CCur(lblbase) - dto), "#,##0.00")
    lbliva = Format(dto, "#,##0.00")
    lbltotal = Format(CCur(lblbase) - CCur(lbliva), "#,##0.00")
End Sub

Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
   On Error GoTo grid_AfterColEdit_Error

    Select Case ColIndex
        Case Cols.codigo
        Case Cols.fecha
    End Select
'    grid.FetchRowStyle = True
'    grid.Refresh
    calcular_total
   On Error GoTo 0
   Exit Sub

grid_AfterColEdit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grid_AfterColEdit of Formulario frmDocumento_Edicion2"
End Sub
Private Sub grid_AfterUpdate()
    calcular_total
End Sub
Private Sub grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid80.StyleDisp)
'    RowStyle = Prospective
End Sub
Private Sub grid_KeyPress(KeyAscii As Integer)
    If (grid.Col = Cols.PRECIO) And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub
Private Sub cargar_documento()
    On Error GoTo fallo
    Dim oDOCUMENTO As New clsDocs_pago
    If oDOCUMENTO.CargarDocumento(CInt(PK_DOCUMENTO)) = True Then
       Me.Caption = "Modificación del documento : " & oDOCUMENTO.getNUMERO & "/" & Year(oDOCUMENTO.getFECHA_FACTURA)
       ' Cargar cabecera de factura
       fdesde.value = Format(oDOCUMENTO.getFECHA_FACTURA, "dd-mm-yyyy")
       txtdescuento = oDOCUMENTO.getDESCUENTO
       txtiva = oDOCUMENTO.getIVA
       cmbclientes.MostrarElemento oDOCUMENTO.getCLIENTE_ID
       cmbfp.BoundText = oDOCUMENTO.getFP_ID
       Dim ocliente As New clsCliente
       ocliente.CargaCliente oDOCUMENTO.getCLIENTE_ID
       cmbTarifa.MostrarElemento ocliente.getTARIFA_ID
       ' Cargamos los pedido del cliente
       cargar_pedidos CLng(oDOCUMENTO.getCLIENTE_ID), fdesde.value
       cmbpedido.BoundText = oDOCUMENTO.getPEDIDO_ID
       ' Cargar detalle del documento
       Dim oDocumento_Detalle As New clsDocs_pago_muestras
       Dim rs As ADODB.RecordSet
       Set rs = oDocumento_Detalle.lineas_factura(PK_DOCUMENTO)
       log ("CARGA DEL DOCUMENTO. ID_DOC : " & PK_DOCUMENTO)
       If rs.RecordCount > 0 Then
            Dim fila As Long
            fila = 0
            Do
                If rs(8) <> 0 Then
                    x(fila, Cols.ID) = CStr(rs(0))
                    x(fila, Cols.general) = CStr(rs(1))
                    x(fila, Cols.particular) = CStr(rs(3))
                Else
                    x(fila, Cols.ID) = "0"
'                    grid.FetchRowStyle = True
'                    grid.Refresh
                End If
                x(fila, Cols.fecha) = CStr(rs(2))
                x(fila, Cols.tipoanalisis) = CStr(rs(4))
                x(fila, Cols.referencia) = CStr(rs(5))
                x(fila, Cols.codigo) = CStr(rs(6))
                x(fila, Cols.PRECIO) = CStr(rs(7))
'                log CStr(rs(0)) & ";" & CStr(rs(1)) & ";" & _
'                    CStr(rs(2)) & ";" & CStr(rs(3)) & ";" & _
'                    CStr(rs(4)) & ";" & CStr(rs(5)) & ";" & _
'                    CStr(rs(6)) & ";" & CStr(rs(7))
                rs.MoveNext
                fila = fila + 1
            Loop Until rs.EOF
            grid.Row = 0
            grid.Col = 0
            grid.Refresh
        End If
        calcular_total
    Else
        MsgBox "Error al cargar el documento.", vbInformation, App.Title
    End If
    Set oDOCUMENTO = Nothing
    Set oDocumento_Detalle = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub

Public Sub inicializar_ventana()
    lbltotal = Format("0", "#,##0.00")
    lbliva = Format("0", "#,##0.00")
    lblbase = Format("0", "#,##0.00")
    cargar_combos
    inicializar_grid
End Sub

Public Sub cargar_combos()
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbfp, New clsFP
    llenar_combo cmbTarifa, New clsTarifas, 0, Me, ""
End Sub

Public Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim oPedido As New clsClientes_pedidos
    Set cmbpedido.RowSource = oPedido.Listado_en_fecha(CInt(cliente), CStr(fecha))
    cmbpedido.ListField = "CODIGO_LARGO"
    cmbpedido.DataField = "ID_PEDIDO"
    cmbpedido.BoundColumn = "ID_PEDIDO"
End Sub

Private Sub txtdescuento_Change()
    calcular_total
End Sub
Private Sub txtiva_LostFocus()
    If Trim(txtiva) <> "" Then
        If Not IsNumeric(txtiva) Then
            MsgBox "El IVA debe ser numérico.", vbCritical, App.Title
            txtiva.SetFocus
        End If
    End If
End Sub

