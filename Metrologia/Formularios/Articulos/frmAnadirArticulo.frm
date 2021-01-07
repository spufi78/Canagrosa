VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmAnadirArticulo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de artículos"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12870
   Icon            =   "frmAnadirArticulo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   12870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de almacen"
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
      Height          =   2775
      Left            =   7200
      TabIndex        =   16
      Top             =   600
      Width           =   5625
      Begin VB.CheckBox chkIne 
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.N.E"
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   4350
         TabIndex        =   10
         Top             =   2010
         Width           =   1155
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   1050
         TabIndex        =   7
         Top             =   2010
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   1050
         TabIndex        =   6
         Top             =   1650
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Index           =   3
         Left            =   1050
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   300
         Width           =   4230
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   1050
         TabIndex        =   8
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   5
         Left            =   3780
         TabIndex        =   3
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comisión"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   24
         Top             =   2100
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Peso (Kg.)"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   195
         Index           =   17
         Left            =   150
         TabIndex        =   18
         Top             =   810
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stock"
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   17
         Top             =   2430
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Portes"
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
      Height          =   5040
      Left            =   60
      TabIndex        =   25
      Top             =   3390
      Width           =   7125
      Begin TrueDBGrid80.TDBGrid grid 
         Height          =   4665
         Left            =   120
         TabIndex        =   27
         Top             =   270
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   8229
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tarifa"
         Columns(0).DataField=   ""
         Columns(0).DropDown=   "tArticulos"
         Columns(0).DropDown.vt=   8
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Precio"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "Currency"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ID_TARIFA"
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "General Number"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   1
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=8731"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8652"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8192"
         Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(0).AutoDropDown=1"
         Splits(0)._ColumnProps(9)=   "Column(0).AutoCompletion=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Width=2831"
         Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=2752"
         Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(16)=   "Column(2).Width=3016"
         Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=2937"
         Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=0,.bgcolor=&HD7D7FF&"
         _StyleDefs(37)  =   ":id=24,.locked=-1,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(38)  =   ":id=24,.charset=0"
         _StyleDefs(39)  =   ":id=24,.fontname=MS Sans Serif"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
         _StyleDefs(43)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(44)  =   ":id=23,.fontname=MS Sans Serif"
         _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.alignment=1,.locked=0"
         _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
         _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
         _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
         _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
         _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
         _StyleDefs(53)  =   "Named:id=37:Normal"
         _StyleDefs(54)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
         _StyleDefs(55)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(56)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(57)  =   "Named:id=38:Heading"
         _StyleDefs(58)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   ":id=38,.wraptext=-1"
         _StyleDefs(60)  =   "Named:id=39:Footing"
         _StyleDefs(61)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   "Named:id=40:Selected"
         _StyleDefs(63)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(64)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(65)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(66)  =   "Named:id=41:Caption"
         _StyleDefs(67)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(68)  =   "Named:id=42:HighlightRow"
         _StyleDefs(69)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(70)  =   "Named:id=43:EvenRow"
         _StyleDefs(71)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
         _StyleDefs(72)  =   "Named:id=44:OddRow"
         _StyleDefs(73)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
         _StyleDefs(74)  =   "Named:id=47:RecordSelector"
         _StyleDefs(75)  =   ":id=47,.parent=38"
         _StyleDefs(76)  =   "Named:id=50:FilterBar"
         _StyleDefs(77)  =   ":id=50,.parent=37"
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedor"
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
      Height          =   1155
      Left            =   45
      TabIndex        =   19
      Top             =   2220
      Width           =   7125
      Begin vb6projectpryComboBCA.miComboBCA cmbProveedor 
         Height          =   315
         Left            =   1290
         TabIndex        =   28
         Top             =   300
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   556
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   4
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio Costo"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   750
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7590
      Width           =   1035
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   855
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7590
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos generales"
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
      Height          =   1575
      Left            =   45
      TabIndex        =   13
      Top             =   600
      Width           =   7140
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1290
         TabIndex        =   0
         Text            =   "1234"
         Top             =   240
         Width           =   2610
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   1275
         TabIndex        =   2
         Top             =   1110
         Width           =   5715
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   1275
         TabIndex        =   1
         Top             =   705
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   9
         Left            =   135
         TabIndex        =   22
         Top             =   330
         Width           =   675
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
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   750
         Width           =   360
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
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   1200
         Width           =   990
      End
   End
   Begin VB.Label lbltitulo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Artículos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   26
      Top             =   90
      Width           =   4050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   12270
      Picture         =   "frmAnadirArticulo.frx":09EA
      Top             =   30
      Width           =   480
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   12825
   End
End
Attribute VB_Name = "frmAnadirArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk As Long

Dim x As New XArrayDB

Const filas As Integer = 50
Const Col As Integer = 3
Private Enum Cols
    TARIFA = 0
    PRECIO = 1
    ID_TARIFA = 2
End Enum

Public Sub cargar_combos()
    llenar_combo cmbProveedor, New clsProveedor, 0, frmProveedores, ""
    Cargar_Combo cmbtipos, New clsArticulos_Tipos
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      Dim oArticulo As New clsArticulos
      With oArticulo
        .setID_ARTICULO = txtcodigo
        .setTIPO_ARTICULO_ID = cmbtipos.BoundText
        .setDESCRIPCION = UCase(txtdatos(0))
        If cmbProveedor.getTEXTO = "" Then
            .setPROVEEDOR_ID = 0
        Else
            .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
        End If
        If Trim(txtdatos(1)) = "" Then
            .setPRECIO_COMPRA = moneda_bd("0")
        Else
            .setPRECIO_COMPRA = moneda_bd(txtdatos(1))
        End If
        If Trim(txtdatos(2)) = "" Then
            .setSTOCK = 0
        Else
            .setSTOCK = txtdatos(2)
        End If
        .setCOMENTARIO = txtdatos(3)
        .setPESO = Replace(txtdatos(4), ",", ".")
        If Trim(txtdatos(5)) = "" Then
            .setCOMISION = moneda_bd("0")
        Else
            .setCOMISION = moneda_bd(txtdatos(5))
        End If
        .setINE = 0
        .setINE_CODIGO = 0
        If chkIne.Value = Checked Then
            .setINE = chkIne.Value
            .setINE_CODIGO = txtdatos(6)
        End If
      End With
      Dim Insertar As Boolean
      Insertar = False
      If pk = 0 Then
        If MsgBox("Va a introducir el articulo, ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            pk = oArticulo.Insertar
            If pk = 0 Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el articulo, ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If oArticulo.Modificar(pk) = False Then
                Exit Sub
            End If
        End If
      End If
      ' PORTES
      If pk <> 0 Then
            Dim oTP As New clsTarifas_portes_articulos
            oTP.Eliminar (pk)
            Dim i As Integer
            For i = 0 To filas
                If Not IsEmpty(x(i, Cols.ID_TARIFA)) Then
                    oTP.setARTICULO_ID = pk
                    oTP.setTARIFA_PORTE_ID = x(i, Cols.ID_TARIFA)
                    oTP.setPRECIO = moneda_bd(x(i, Cols.PRECIO))
                    oTP.Insertar
                End If
            Next
            Set oTP = Nothing
      End If
      MsgBox "Artículo almacenado correctamente.", vbInformation, App.Title
      Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Error al añadir el artículo:" & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cargar_combos
    inicializar_grid
    If pk <> 0 Then
        lbltitulo = "Modificación de Artículo"
        cargar_articulo
        txtcodigo.Enabled = False
    Else
        nuevo_articulo
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If (grid.Col = Cols.PRECIO) And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii = 46 Then
         KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
    If Index = 1 Then
        If txtdatos(Index) <> "" Then
            txtdatos(Index) = moneda(txtdatos(Index))
        End If
    End If
End Sub
Public Sub cargar_articulo()
    Dim oArticulo As New clsArticulos
    With oArticulo
        .Carga (pk)
        txtcodigo = .getID_ARTICULO
        cmbtipos.BoundText = .getTIPO_ARTICULO_ID
        txtdatos(0) = .getDESCRIPCION
        txtdatos(1) = moneda(.getPRECIO_COMPRA)
        cmbProveedor.MostrarElemento .getPROVEEDOR_ID
        txtdatos(2) = .getSTOCK
        txtdatos(3) = .getCOMENTARIO
        txtdatos(4) = .getPESO
        txtdatos(5) = moneda(.getCOMISION)
        If .getINE = 1 Then
            chkIne.Value = .getINE
            txtdatos(6) = .getINE_CODIGO
        End If
        ' Tarifas de Porte
        Dim oTP As New clsTarifas_portes_articulos
        Dim rs As ADODB.Recordset
        Set rs = oTP.Listado_por_articulo(pk)
        Dim i As Integer
        If rs.RecordCount > 0 Then
            Do
                For i = 0 To filas - 1
                    If CInt(x(i, Cols.ID_TARIFA)) = CInt(rs("TARIFA_PORTE_ID")) Then
                        x(i, Cols.PRECIO) = moneda(rs("PRECIO"))
                        Exit For
                    End If
                Next
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        Set oTP = Nothing
        grid.Refresh
    End With
End Sub
Public Function validar() As Boolean
   On Error GoTo validar_Error
    validar = True
    If Trim(txtcodigo) = "" Then
        MsgBox "Debe introducir el código.", vbInformation, App.Title
        txtcodigo.SetFocus
        validar = False
        Exit Function
    End If
    If Not IsNumeric(txtcodigo) Then
        MsgBox "El código debe ser numérico.", vbInformation, App.Title
        txtcodigo.SetFocus
        validar = False
        Exit Function
    End If
    If cmbtipos.BoundText = "" Then
        MsgBox "Debe seleccionar un tipo de articulo.", vbInformation, App.Title
        cmbtipos.SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtdatos(0)) = "" Then
        MsgBox "Debe darle una descripción al artículo.", vbInformation, App.Title
        txtdatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtdatos(2)) = "" Then
        MsgBox "Debe indicar un stock.", vbInformation, App.Title
        txtdatos(2).SetFocus
        validar = False
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

validar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validar of Formulario frmAnadirArticulo"
End Function
Public Sub nuevo_articulo()
    Dim oart As New clsArticulos
    oart.CrearID
    txtcodigo = oart.getID_ARTICULO
End Sub
Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error
   
    grid.Col = 0
    grid.Row = 0
    x.Clear
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
    
    Dim rs As ADODB.Recordset
    Dim oTP As New clsTarifas_portes
    Set rs = oTP.Listado()
    If rs.RecordCount > 0 Then
        Dim i As Integer
        i = 0
        Do
            x(i, Cols.TARIFA) = CStr(rs("DESCRIPCION"))
            x(i, Cols.PRECIO) = moneda("0")
            x(i, Cols.ID_TARIFA) = CStr(rs("ID_TARIFA_PORTE"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    grid.Row = 0
    grid.Col = 0
    grid.Refresh
    

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

