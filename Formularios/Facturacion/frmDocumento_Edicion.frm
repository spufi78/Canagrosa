VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmDocumento_Edicion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo Documento"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13365
   Icon            =   "frmDocumento_Edicion.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeterminaciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir Determinaciones"
      Height          =   930
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Recalcula el precio de la muestra seleccionada"
      Top             =   9315
      Width           =   1335
   End
   Begin VB.Frame frmDeter 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Añadir determinaciones a la Factura"
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
      Height          =   5565
      Left            =   1350
      TabIndex        =   30
      Top             =   2070
      Visible         =   0   'False
      Width           =   10635
      Begin VB.CommandButton cmdSalirDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   930
         Left            =   9390
         Picture         =   "frmDocumento_Edicion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4530
         Width           =   1155
      End
      Begin VB.CommandButton cmdAceptarDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   930
         Left            =   8190
         Picture         =   "frmDocumento_Edicion.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4530
         Width           =   1155
      End
      Begin MSComctlLib.ListView deter 
         Height          =   4275
         Left            =   90
         TabIndex        =   31
         Top             =   210
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   End
   Begin VB.CommandButton cmdMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Muestra"
      Height          =   930
      Left            =   2340
      Picture         =   "frmDocumento_Edicion.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Recalcula el precio de la muestra seleccionada"
      Top             =   9315
      Width           =   1305
   End
   Begin VB.CommandButton cmdinsertar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inserta Línea"
      Height          =   930
      Left            =   90
      Picture         =   "frmDocumento_Edicion.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9315
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la factura"
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
      Height          =   1170
      Left            =   6615
      TabIndex        =   15
      Top             =   45
      Width           =   6705
      Begin VB.TextBox txtdescuento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3645
         TabIndex        =   18
         Top             =   675
         Width           =   690
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   360
         Left            =   6255
         TabIndex        =   17
         Top             =   225
         Width           =   330
      End
      Begin VB.TextBox txtiva 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   330
         Left            =   5310
         TabIndex        =   16
         Top             =   675
         Width           =   960
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1125
         TabIndex        =   19
         Top             =   675
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
         Format          =   60424193
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbpedido 
         Height          =   315
         Left            =   1125
         TabIndex        =   20
         Top             =   270
         Width           =   4965
         _ExtentX        =   8758
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
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuento"
         Height          =   195
         Index           =   2
         Left            =   2655
         TabIndex        =   24
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   4410
         TabIndex        =   23
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   22
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
         Height          =   195
         Index           =   4
         Left            =   4815
         TabIndex        =   21
         Top             =   720
         Width           =   390
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del cliente"
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
      Height          =   1350
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   6525
      Begin MSDataListLib.DataCombo cmbfp 
         Height          =   315
         Left            =   1125
         TabIndex        =   11
         Top             =   585
         Width           =   4335
         _ExtentX        =   7646
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
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   1125
         TabIndex        =   12
         Top             =   225
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbTarifa 
         Height          =   345
         Left            =   1125
         TabIndex        =   26
         Top             =   945
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   27
         Top             =   990
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   16
         Left            =   135
         TabIndex        =   13
         Top             =   630
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   930
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9330
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   930
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9330
      Width           =   1155
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar &Línea"
      Height          =   930
      Left            =   1215
      Picture         =   "frmDocumento_Edicion.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9315
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   7830
      Left            =   60
      TabIndex        =   0
      Top             =   1440
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   13811
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
      Columns(1).Caption=   "DETERMINACION_ID"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NºGeneral"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fecha"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "General Date"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "NºParticular"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tipo Análisis"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Referencia Cliente"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Código"
      Columns(7).DataField=   ""
      Columns(7).ConvertEmptyCell=   1
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Precio"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "Currency"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=1905"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1826"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2223"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2143"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=6615"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6535"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=6509"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=6429"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=0"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2408"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2328"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=847"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=767"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=2"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=14,.parent=67"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=11,.parent=68"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=12,.parent=69"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=13,.parent=71"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=86,.parent=67,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=68"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=69"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=71"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=90,.parent=67,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(49)  =   ":id=90,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(50)  =   ":id=90,.fontname=MS Sans Serif"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=71,.bold=0,.fontsize=975"
      _StyleDefs(54)  =   ":id=89,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(55)  =   ":id=89,.fontname=MS Sans Serif"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=94,.parent=67,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=68"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=69"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=71"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=95,.parent=68"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=96,.parent=69"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=97,.parent=71"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=102,.parent=67,.alignment=0"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=68"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=69"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=71"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=106,.parent=67,.alignment=2,.locked=0"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=103,.parent=68"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=104,.parent=69"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=105,.parent=71"
      _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=110,.parent=67,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=68"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=69"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=71"
      _StyleDefs(76)  =   "Named:id=37:Normal"
      _StyleDefs(77)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(78)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(79)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(80)  =   "Named:id=38:Heading"
      _StyleDefs(81)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=38,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=39:Footing"
      _StyleDefs(84)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(85)  =   "Named:id=40:Selected"
      _StyleDefs(86)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(87)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(88)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(89)  =   "Named:id=41:Caption"
      _StyleDefs(90)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(91)  =   "Named:id=42:HighlightRow"
      _StyleDefs(92)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(93)  =   "Named:id=43:EvenRow"
      _StyleDefs(94)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(95)  =   "Named:id=44:OddRow"
      _StyleDefs(96)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(97)  =   "Named:id=47:RecordSelector"
      _StyleDefs(98)  =   ":id=47,.parent=38"
      _StyleDefs(99)  =   "Named:id=50:FilterBar"
      _StyleDefs(100) =   ":id=50,.parent=37"
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
      Top             =   9585
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
      Top             =   9585
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
      Top             =   9915
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
      Top             =   9915
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
      Top             =   9255
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
      Top             =   9255
      Width           =   2160
   End
End
Attribute VB_Name = "frmDocumento_Edicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_DOCUMENTO As Long
'Dim Prospective As New TrueDBGrid80.Style
Dim x As New XArrayDB
Dim xCodigos As New XArrayDB
Const filas As Integer = 2000
Const Col As Integer = 8
Private Enum COLS
    ID = 0
    DETERMINACION_ID = 1
    GENERAL = 2
    fecha = 3
    particular = 4
    tipoanalisis = 5
    REFERENCIA = 6
    CODIGO = 7
    PRECIO = 8
End Enum

Private Sub cmbClientes_change()
    If cmbClientes.getTEXTO <> "" Then
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbClientes.getPK_SALIDA
        cmbFP.BoundText = oCliente.getFP_ID
        ' Cargamos los pedido del cliente
        cargar_pedidos CLng(cmbClientes.getPK_SALIDA), fdesde.Value
        cmbPedido.Text = ""
    End If
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    If frmDeter.visible = True Then
        MsgBox "Cierre primero la ventana de Determinaciones.", vbExclamation, App.Title
        Exit Sub
    End If
    ' Lineas del documento
'    MsgBox "Pendiente de revisión. No esta implementado."
'    Exit Sub
    ' Log completo salida y VALIDACIONES
    Dim i As Integer
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.Value(i, COLS.GENERAL)) <> "" Or Trim(x.Value(i, COLS.tipoanalisis)) <> "" Then
            If Trim(x.Value(i, COLS.PRECIO)) <> "" Then
                If Not IsNumeric(CStr(Trim(x.Value(i, COLS.PRECIO)))) Then
                    MsgBox "El análisis " & x.Value(i, COLS.CODIGO) & " : " & x.Value(i, COLS.tipoanalisis) & " tiene el precio mal informado.", vbCritical, App.Title
                    Exit Sub
                End If
            End If
            log x.Value(i, COLS.ID) & ";" & x.Value(i, COLS.GENERAL) & ";" & x.Value(i, COLS.fecha) & ";" & _
                x.Value(i, COLS.particular) & ";" & x.Value(i, COLS.tipoanalisis) & ";" & _
                x.Value(i, COLS.REFERENCIA) & ";" & x.Value(i, COLS.CODIGO) & ";" & _
                x.Value(i, COLS.PRECIO)
        End If
    Next
    ''''' Validar que el precio de la muestra sea la suma de las determinaciones
    Dim sumatorio As Currency
    Dim precio_Analisis As Currency
    Dim diferencia As Boolean
    Dim salida As String
    Dim fila As Integer
    Dim descripcion_analisis As String
    diferencia = False
    sumatorio = 0
    precio_Analisis = 0
    i = x.LowerBound(1)
    salida = "Revise los siguientes anális, el precio total no coincide con el de sus determinaciones : " & vbNewLine & vbNewLine
    While i < x.UpperBound(1)
'        If Not IsEmpty(x.Value(i, COLS.GENERAL)) Then
'          If Trim(x.Value(i, COLS.GENERAL)) <> "" Or Trim(x.Value(i, COLS.tipoanalisis)) <> "" Then
        If Not IsEmpty(x.Value(i, COLS.ID)) Then
          If x.Value(i, COLS.ID) <> "" And Trim(x.Value(i, COLS.DETERMINACION_ID)) = "0" Then
            If precio_Analisis <> 0 And sumatorio <> 0 Then
                If precio_Analisis <> sumatorio Then
                    diferencia = True
                    salida = salida & descripcion_analisis & vbNewLine
'                    grid.Row = fila
'                    grid.Col = Cols.PRECIO
'                    grid.EditBackColor = vbRed
'                    grid.ForeColor = vbRed
                End If
                sumatorio = 0
            End If
            precio_Analisis = CCur(x.Value(i, COLS.PRECIO))
            descripcion_analisis = "* (" & x.Value(i, COLS.particular) & ") " & x.Value(i, COLS.tipoanalisis) & " - " & x.Value(i, COLS.REFERENCIA)
            fila = i
          Else
            sumatorio = sumatorio + CCur(x.Value(i, COLS.PRECIO))
          End If
        Else
            sumatorio = sumatorio + CCur(x.Value(i, COLS.PRECIO))
        End If
        i = i + 1
    Wend
    If precio_Analisis <> 0 And sumatorio <> 0 Then
       If precio_Analisis <> sumatorio Then
            diferencia = True
            salida = salida & descripcion_analisis & vbNewLine
       End If
    End If
    ''''' Fin Validar precio muestra
    If diferencia Then
        MsgBox salida, vbExclamation, App.Title
        MsgBox "Se modificará el documento aunque contenga errores.", vbInformation, App.Title
'        Exit Sub
    End If
    ' Informamos los datos del documento
    Me.MousePointer = 11
    Dim oDoc As New clsDocs_pago
    With oDoc
        .setCLIENTE_ID = cmbClientes.getPK_SALIDA
        .setDESCUENTO = numerico_bd(txtdescuento)
        .setFECHA_FACTURA = Format(fdesde.Value, "yyyy-mm-dd")
        .setFP_ID = cmbFP.BoundText
        If cmbPedido.BoundText <> "" Then
            .setPEDIDO_ID = cmbPedido.BoundText
        Else
            .setPEDIDO_ID = 0
        End If
        .setIVA = txtiva
        .Modificar (PK_DOCUMENTO)
    End With
    ' Informar las muestras a no facturadas
    Dim oMuestra As New clsMuestra
    Dim oDocumento_Detalle As New clsDocs_pago_muestras
    Dim rs As ADODB.Recordset
    Set rs = oDocumento_Detalle.MuestrasDocumento(PK_DOCUMENTO)
    Dim sgrupo As String
    If rs.RecordCount <> 0 Then
        Do
            sgrupo = sgrupo & rs("muestra_id") & ","
            rs.MoveNext
        Loop Until rs.EOF
        sgrupo = Left(sgrupo, Len(sgrupo) - 1)
        oMuestra.Informar_Documentos_Pago sgrupo, 0
    End If
    ' Detalle del documento
    If PK_DOCUMENTO <> 0 Then
        oDocumento_Detalle.EliminarMuestras PK_DOCUMENTO
    End If
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.Value(i, COLS.GENERAL)) <> "" Or Trim(x.Value(i, COLS.tipoanalisis)) <> "" Then
            With oDocumento_Detalle
                .setDOC_ID = PK_DOCUMENTO
                .setORDEN = i
                .setABONADO = 0
                .setMUESTRA_ID = 0
                If Trim(x.Value(i, COLS.ID)) <> "" Then
                    If IsNumeric(Trim(x.Value(i, COLS.ID))) Then
                        .setMUESTRA_ID = x.Value(i, COLS.ID)
                    End If
                End If
                .setDETERMINACION_ID = 0
                If Trim(x.Value(i, COLS.DETERMINACION_ID)) <> "" Then
                    If IsNumeric(Trim(x.Value(i, COLS.DETERMINACION_ID))) Then
                        .setDETERMINACION_ID = x.Value(i, COLS.DETERMINACION_ID)
                    End If
                End If
                If Trim(x.Value(i, COLS.fecha)) = "" Then
                    .setFECHA = Format(Date, "yyyy-mm-dd")
                Else
                    .setFECHA = Format(x.Value(i, COLS.fecha), "yyyy-mm-dd")
                End If
                .setTIPO_ANALISIS = x.Value(i, COLS.tipoanalisis)
                .setREFERENCIA_CLIENTE = x.Value(i, COLS.REFERENCIA)
                .setCODIGO = x.Value(i, COLS.CODIGO)
                .setPRECIO = moneda_bd(x.Value(i, COLS.PRECIO))
                .Insertar_linea
            End With
            ' Modificamos los datos referentes a la muestra
            If Not IsEmpty(x.Value(i, COLS.ID)) And Trim(x.Value(i, COLS.ID)) <> "" Then
                If CLng(x.Value(i, COLS.ID)) <> 0 Then
                    With oMuestra
                        .setPRECIO = moneda_bd(x.Value(i, COLS.PRECIO))
                        If cmbPedido.Text = "" Then
                            .setPEDIDO_ID = 0
                        Else
                            .setPEDIDO_ID = cmbPedido.BoundText
                        End If
                        .setDOCUMENTO_PAGO = 2
'                        .setCLIENTE_ID = cmbclientes.getPK_SALIDA
                        .Informar_Datos_documento CLng(x.Value(i, COLS.ID))
                    End With
                End If
            End If
        End If
    Next
    ' Informar el total de factura
    oDoc.Informar_total_factura PK_DOCUMENTO
    log ("Documento insertado correctamente.")
    Me.MousePointer = 0
    MsgBox "El documento se ha almacenado correctamente.", vbInformation, App.Title
    Unload Me
    Exit Sub
fallo:
    Me.MousePointer = 0

    MsgBox "Error al guardar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdAceptarDeter_Click()
    Dim i As Integer
    Dim f As Integer
    Dim r As Integer
    On Error Resume Next
    ' Añade una linea
'    For f = filas To grid.Bookmark Step -1
'        For r = 0 To Col
'            x(f + 1, r) = x(f, r)
'        Next
'    Next
'    x(grid.Bookmark, COLS.CODIGO) = ""
'    x(grid.Bookmark + 1, COLS.ID) = Empty
'    x(grid.Bookmark + 1, COLS.GENERAL) = Empty
'    x(grid.Bookmark + 1, COLS.particular) = ""
'    x(grid.Bookmark + 1, COLS.tipoanalisis) = x(grid.Bookmark, COLS.tipoanalisis) & " BASE"
'    x(grid.Bookmark + 1, COLS.REFERENCIA) = ""
'    grid.Bookmark = grid.Bookmark + 1
'    Dim finicio As Integer
'    finicio = grid.Bookmark
      
    Dim idMuestra As String
    Dim fecha As String
    idMuestra = CLng(x.Value(grid.Bookmark, COLS.ID))
    fecha = x.Value(grid.Bookmark, COLS.fecha)
    For i = 1 To deter.ListItems.Count
        If deter.ListItems(i).Checked = True Then
            For f = filas To grid.Bookmark Step -1
                For r = 0 To Col
                    x(f + 1, r) = x(f, r)
                Next
            Next
            x(grid.Bookmark + 1, COLS.ID) = idMuestra
            x(grid.Bookmark + 1, COLS.DETERMINACION_ID) = deter.ListItems(i).SubItems(7)
            x(grid.Bookmark + 1, COLS.GENERAL) = Empty
            x(grid.Bookmark + 1, COLS.fecha) = fecha
            x(grid.Bookmark + 1, COLS.particular) = ""
            x(grid.Bookmark + 1, COLS.tipoanalisis) = deter.ListItems(i).SubItems(1)
            x(grid.Bookmark + 1, COLS.REFERENCIA) = deter.ListItems(i).SubItems(2)
            x(grid.Bookmark + 1, COLS.CODIGO) = deter.ListItems(i).SubItems(3)
            If deter.ListItems(i).SubItems(5) <> "" Then
                x(grid.Bookmark + 1, COLS.PRECIO) = deter.ListItems(i).SubItems(5)
            Else
                x(grid.Bookmark + 1, COLS.PRECIO) = moneda("0")
            End If
            grid.Bookmark = grid.Bookmark + 1
        End If
    Next
    ' Recalcular muestra
    calcularPrecioMuestra CLng(idMuestra)
    
    grid.Refresh
    grid_AfterColEdit (COLS.PRECIO)
    frmDeter.visible = False
'    calcular_total
    grid.SetFocus
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    For i = grid.Bookmark To filas - 1
        For j = 0 To Col
            x(i, j) = x(i + 1, j)
        Next
    Next
    grid.Refresh
    grid_AfterColEdit (COLS.PRECIO)
    calcular_total
    grid.SetFocus
End Sub

Private Sub cmdDeterminaciones_Click()
    If Not IsEmpty(x.Value(grid.Bookmark, COLS.ID)) Then
        If x.Value(grid.Bookmark, COLS.ID) <> "" Then
            If CLng(x.Value(grid.Bookmark, COLS.ID)) <> 0 Then
                cargar_determinaciones CLng(x.Value(grid.Bookmark, COLS.ID)), cmbtarifa.getPK_SALIDA
                frmDeter.Caption = "Determinaciones muestra : " & x.Value(grid.Bookmark, COLS.GENERAL)
                frmDeter.visible = True
            End If
        End If
    End If
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
    If grid.Bookmark > 0 Then
        x(grid.Bookmark, COLS.ID) = x(grid.Bookmark - 1, COLS.ID)
        x(grid.Bookmark, COLS.DETERMINACION_ID) = "-1"
    End If
    grid.Refresh
    calcular_total
    grid.SetFocus
End Sub

Private Sub cmdMuestra_Click()
    
    If Not IsEmpty(x.Value(grid.Bookmark, COLS.ID)) Then
        If x.Value(grid.Bookmark, COLS.ID) <> "" Then
            If CLng(x.Value(grid.Bookmark, COLS.ID)) <> 0 Then
                gmuestra = CLng(x.Value(grid.Bookmark, COLS.ID))
                frmVerMuestra.Show 1
            End If
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSalirDeter_Click()
    frmDeter.visible = False
End Sub

Private Sub deter_DblClick()
    If deter.ListItems.Count > 0 Then
        frmTD_Detalle.PK = deter.ListItems(deter.selectedItem.Index).SubItems(6)
        frmTD_Detalle.Show 1
        actualizar_determinaciones CLng(x.Value(grid.Bookmark, COLS.ID)), cmbtarifa.getPK_SALIDA, deter.ListItems(deter.selectedItem.Index).SubItems(6)
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
    cabecera
    inicializar_ventana
    cargar_documento
    ' Verificar si esta contabilidado
    Dim oDoc As New clsDocs_pago
    If oDoc.esta_contabilidado(CLng(PK_DOCUMENTO)) Then
'        cmdaceptar.Enabled = False
'        MsgBox "El documento se encuentra contabilizado. No se puede editar.", vbInformation, App.Title
    End If
'    Set Prospective = grid.Styles.Add("Prospective")
'    Prospective.Font.Italic = True
'    Prospective.ForeColor = vbBlue
End Sub

Private Sub calcular_total()
    Dim i As Integer
    On Error Resume Next
    Dim total As Currency
    total = 0
    For i = 0 To filas
        If Not IsEmpty(x(i, COLS.ID)) Then
'            If Trim(CStr(x(i, COLS.PRECIO))) <> "" And Trim(CStr(x(i, COLS.fecha))) <> "" And CStr(x(i, COLS.ID)) <> "0" And CStr(x(i, COLS.ID)) <> "" Then
            If Trim(CStr(x(i, COLS.PRECIO))) <> "" And Trim(CStr(x(i, COLS.fecha))) <> "" And CStr(x(i, COLS.DETERMINACION_ID)) = "0" Then
                total = total + Format((CStr(x.Value(i, COLS.PRECIO))), "0.00")
            End If
        End If
    Next
    lblBase = Format(total, "#,##0.00")
    Dim dto As Currency
    If txtdescuento.Text <> "" Then
        dto = Format((CCur(lblBase) * CInt(txtdescuento.Text) / 100), "#,##0.00")
    Else
        dto = 0
    End If
'    dto = dto + Format(((CCur(lblbase) - dto) * CInt(txtdescuento.Text) / 100), "#,##0.00")
'    dto = Format((CCur(lblbase) - dto), "#,##0.00")
    lblIVA = Format(dto, "#,##0.00")
    lbltotal = Format(CCur(lblBase) - CCur(lblIVA), "#,##0.00")
End Sub

Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
   On Error GoTo grid_AfterColEdit_Error

'    Select Case ColIndex
'        Case COLS.CODIGO
'        Case COLS.fecha
'        Case COLS.PRECIO
'    End Select
'    calcular_total
   On Error GoTo 0
   Exit Sub

grid_AfterColEdit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grid_AfterColEdit of Formulario frmDocumento_Edicion"
End Sub
Private Sub grid_AfterUpdate()
    If x(grid.Bookmark, COLS.DETERMINACION_ID) <> "0" And x(grid.Row, COLS.DETERMINACION_ID) <> "" Then
        calcularPrecioMuestra x(grid.Bookmark, COLS.ID)
'        Dim idMuestra As String
'        Dim i As Integer
'        idMuestra = x(grid.Row, COLS.ID)
'        Dim encontrado As Boolean
'        Dim filaMuestra As Integer
'        encontrado = False
'        Dim Suma As Currency
'        For i = x.LowerBound(1) To x.UpperBound(1)
'            If x(i, COLS.ID) = idMuestra Then
'                encontrado = True
'                If x(i, COLS.DETERMINACION_ID) <> "0" Then
'                    Suma = Suma + CCur(x(i, COLS.PRECIO))
'                Else
'                    filaMuestra = i
'                End If
'            End If
'        Next
'        x(filaMuestra, COLS.PRECIO) = Suma
'        grid.Refresh
    End If
'    calcular_total
End Sub
Private Sub calcularPrecioMuestra(idMuestra As Long)
    Dim i As Integer
    Dim encontrado As Boolean
    Dim filaMuestra As Integer
   On Error GoTo calcularPrecioMuestra_Error

    encontrado = False
    Dim Suma As Currency
    For i = x.LowerBound(1) To x.UpperBound(1)
        If x(i, COLS.ID) = idMuestra Then
            encontrado = True
            If x(i, COLS.DETERMINACION_ID) <> "0" Then
                Suma = Suma + CCur(x(i, COLS.PRECIO))
            Else
                filaMuestra = i
            End If
'       Else
'           If Not encontrado Then Exit For
        End If
    Next
    x(filaMuestra, COLS.PRECIO) = Suma
    grid.Refresh
    calcular_total

   On Error GoTo 0
   Exit Sub

calcularPrecioMuestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularPrecioMuestra of Formulario frmDocumento_Edicion"
End Sub
Private Sub grid_DblClick()
    cmdMuestra_Click
End Sub

Private Sub grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid80.StyleDisp)
'    RowStyle = Prospective
End Sub
Private Sub grid_KeyPress(KeyAscii As Integer)
    If (grid.Col = COLS.PRECIO) And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub
Private Sub cargar_documento()
    On Error GoTo fallo
    Dim oDOCUMENTO As New clsDocs_pago
    If oDOCUMENTO.CargarDocumento(PK_DOCUMENTO) = True Then
       Me.Caption = "Modificación del documento : " & oDOCUMENTO.getNUMERO & "/" & Year(oDOCUMENTO.getFECHA_FACTURA)
       ' Cargar cabecera de factura
       fdesde.Value = Format(oDOCUMENTO.getFECHA_FACTURA, "dd-mm-yyyy")
       txtdescuento = Replace(oDOCUMENTO.getDESCUENTO, ",", ".")
       txtiva = oDOCUMENTO.getIVA
       cmbClientes.MostrarElemento oDOCUMENTO.getCLIENTE_ID
       cmbFP.BoundText = oDOCUMENTO.getFP_ID
       Dim oCliente As New clsCliente
       oCliente.CargaCliente oDOCUMENTO.getCLIENTE_ID
       cmbtarifa.MostrarElemento oCliente.getTARIFA_ID
       ' Cargamos los pedido del cliente
       cargar_pedidos CLng(oDOCUMENTO.getCLIENTE_ID), fdesde.Value
       cmbPedido.BoundText = oDOCUMENTO.getPEDIDO_ID
       ' Cargar detalle del documento
       Dim oDocumento_Detalle As New clsDocs_pago_muestras
       Dim rs As ADODB.Recordset
       Set rs = oDocumento_Detalle.lineas_factura(PK_DOCUMENTO)
       log ("CARGA DEL DOCUMENTO. ID_DOC : " & PK_DOCUMENTO)
       If rs.RecordCount > 0 Then
            Dim fila As Long
            fila = 0
            Do
                x(fila, COLS.ID) = CStr(rs(0))
                x(fila, COLS.DETERMINACION_ID) = CStr(rs(8))
                If rs(8) = 0 Then
                    x(fila, COLS.GENERAL) = CStr(rs(1))
                    x(fila, COLS.particular) = CStr(rs(3))
                End If
                x(fila, COLS.fecha) = CStr(rs(2))
                x(fila, COLS.tipoanalisis) = CStr(rs(4))
                x(fila, COLS.REFERENCIA) = CStr(rs(5))
                x(fila, COLS.CODIGO) = CStr(rs(6))
                x(fila, COLS.PRECIO) = CStr(rs(7))
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
'        cargar_muestra_sin_facturar oDOCUMENTO.getCLIENTE_ID
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
    lblIVA = Format("0", "#,##0.00")
    lblBase = Format("0", "#,##0.00")
    cargar_combos
    inicializar_grid
End Sub

Public Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbFP, New clsFP
    llenar_combo cmbtarifa, New clsTarifas, 0, Me, ""
End Sub

Public Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim oPedido As New clsClientes_pedidos
'    Set cmbPedido.RowSource = oPedido.Listado_en_fecha(CInt(cliente), CStr(fecha))
    Set cmbPedido.RowSource = oPedido.Listado_por_Cliente(CInt(cliente))
    cmbPedido.ListField = "CODIGO_LARGO"
    cmbPedido.DataField = "ID_PEDIDO"
    cmbPedido.BoundColumn = "ID_PEDIDO"
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

Private Sub cabecera()
    With deter.ColumnHeaders
        .Add , , "Pnt", 1200, lvwColumnLeft
        .Add , , "Nombre", 3400, lvwColumnLeft
        .Add , , "Referencia", 1500, lvwColumnCenter
        .Add , , "Código", 1200, lvwColumnCenter
        .Add , , "P.Base", 1200, lvwColumnRight
        .Add , , "P.Tarifa", 1200, lvwColumnRight
        .Add , , "ID_TIPO_DETERMINACION", 0, lvwColumnLeft
        .Add , , "ID_DETERMINACION", 0, lvwColumnLeft
    End With
End Sub
Private Sub cargar_determinaciones(MUESTRA As Long, TARIFA As Long)
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_determinaciones_Error

    deter.ListItems.Clear
    Dim oDeter As New clsDeterminaciones
    Set rs = oDeter.lista_determinaciones_para_factura(MUESTRA, TARIFA)
    While Not rs.EOF
       With deter.ListItems.Add(, , rs(0))
          .SubItems(1) = rs(1)
          .SubItems(2) = rs(2)
          If Not IsNull(rs(3)) Then
              .SubItems(3) = rs(3)
          End If
          If Not IsNull(rs(4)) Then
              .SubItems(4) = moneda(rs(4))
          End If
          If Not IsNull(rs(5)) Then
              .SubItems(5) = moneda(rs(5))
          End If
          .SubItems(6) = Trim(rs(6))
          .SubItems(7) = Trim(rs(7))
       End With
       deter.ListItems(deter.ListItems.Count).Checked = True
       rs.MoveNext
    Wend
    Set oDeter = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargar_determinaciones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_determinaciones of Formulario frmDocumento_Edicion"
End Sub
Private Sub actualizar_determinaciones(MUESTRA As Long, TARIFA As Long, TD_ID As Long)
    Dim rs As ADODB.Recordset
    Dim oDeter As New clsDeterminaciones
    Set rs = oDeter.lista_determinaciones_para_factura_determinacion(MUESTRA, TARIFA, TD_ID)
    While Not rs.EOF
        With deter.ListItems(deter.selectedItem.Index)
          .Text = rs(0)
          .SubItems(1) = rs(1)
          .SubItems(2) = rs(2)
          If Not IsNull(rs(3)) Then
              .SubItems(3) = rs(3)
          End If
          If Not IsNull(rs(4)) Then
              .SubItems(4) = moneda(rs(4))
          End If
          If Not IsNull(rs(5)) Then
              .SubItems(5) = moneda(rs(5))
          End If
          .SubItems(6) = Trim(rs(6))
       End With
       rs.MoveNext
    Wend
    Set oDeter = Nothing
    Set rs = Nothing
End Sub
