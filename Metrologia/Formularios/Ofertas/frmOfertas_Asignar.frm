VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmOfertas_Asignar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Oferta a Obra"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   Icon            =   "frmOfertas_Asignar.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid80.TDBDropDown tArticulos 
      Height          =   4890
      Left            =   60
      TabIndex        =   8
      Top             =   3345
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   8625
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1693"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1588"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=291"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=185"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HC0E0FF&,.bold=0,.fontsize=975"
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
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8325
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Obra"
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
      Height          =   1125
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   9780
      Begin MSComCtl2.DTPicker FOFERTA 
         Height          =   345
         Left            =   1590
         TabIndex        =   7
         Top             =   690
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   51183617
         CurrentDate     =   40679
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   1620
         TabIndex        =   9
         Top             =   270
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   609
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de la Oferta"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8325
      Width           =   1155
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar &Línea"
      Height          =   885
      Left            =   60
      Picture         =   "frmOfertas_Asignar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8355
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid gridTarifa 
      Height          =   5895
      Left            =   60
      TabIndex        =   0
      Top             =   2385
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   10398
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Ref."
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tArticulos"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Artículo"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Pr. Millar Fábrica"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "Currency"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Pr. Millar Obra"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Currency"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1667"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
      Splits(0)._ColumnProps(8)=   "Column(0).DropDownList=1"
      Splits(0)._ColumnProps(9)=   "Column(0).AutoCompletion=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=9155"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=9049"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=3043"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2937"
      Splits(0)._ColumnProps(20)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=2619"
      Splits(0)._ColumnProps(26)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=975"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.bold=0,.fontsize=975"
      _StyleDefs(37)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=36,.parent=11,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
      _StyleDefs(56)  =   "Named:id=37:Normal"
      _StyleDefs(57)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(58)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(59)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(60)  =   "Named:id=38:Heading"
      _StyleDefs(61)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H0&,.wraptext=-1"
      _StyleDefs(62)  =   "Named:id=39:Footing"
      _StyleDefs(63)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   "Named:id=40:Selected"
      _StyleDefs(65)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(66)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(67)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(68)  =   "Named:id=41:Caption"
      _StyleDefs(69)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(70)  =   "Named:id=42:HighlightRow"
      _StyleDefs(71)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(72)  =   "Named:id=43:EvenRow"
      _StyleDefs(73)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(74)  =   "Named:id=44:OddRow"
      _StyleDefs(75)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(76)  =   "Named:id=47:RecordSelector"
      _StyleDefs(77)  =   ":id=47,.parent=38"
      _StyleDefs(78)  =   "Named:id=50:FilterBar"
      _StyleDefs(79)  =   ":id=50,.parent=37"
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "  Asignación de Oferta a Obra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccione la Obra a la que asignará la Oferta. La Tarifa de dicha obra se almacenará tal y como se muestre en esta pantalla."
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
      Height          =   645
      Index           =   0
      Left            =   45
      TabIndex        =   10
      Top             =   495
      Width           =   9780
   End
End
Attribute VB_Name = "frmOfertas_Asignar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long

Dim xTarifa As New XArrayDB
Dim xarticulos As New XArrayDB

Const filasTarifa As Integer = 50
Const ColTarifa As Integer = 4
Private Enum ColsTarifa
    ID = 0
    ARTICULO = 1
    PFABRICA = 2
    POBRA = 3
End Enum


Private Sub cmbObra_change()
    cargar_obra cmbObra.getPK_SALIDA
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    Dim algo As Boolean
   On Error GoTo cmdAceptar_Click_Error

    If cmbObra.getTEXTO = "" Then
        MsgBox "Debe indicar la Obra a la que asignar la oferta.", vbExclamation, App.Title
        Exit Sub
    End If
    algo = False
    For i = 0 To filasTarifa
        If Trim(xTarifa.Value(i, ColsTarifa.ID)) <> "" Then
            algo = True
        End If
    Next
    If algo = False Then
        MsgBox "La tarifa de la obra no contiene ninguna linea.", vbExclamation, App.Title
        gridTarifa.SetFocus
        Exit Sub
    End If
    ' Fecha de la Oferta
    Dim oObra As New clsObras
    oObra.ModificarFOferta cmbObra.getPK_SALIDA, Format(FOFERTA, "yyyy-mm-dd")
    Set oObra = Nothing
    ' Tarifa
    Dim oTarifa As New clsTarifas
    oTarifa.Eliminar cmbObra.getPK_SALIDA
    For i = xTarifa.LowerBound(1) To xTarifa.UpperBound(1)
        If Trim(xTarifa.Value(i, ColsTarifa.ID)) <> "" Then
            With oTarifa
                .setOBRA_ID = cmbObra.getPK_SALIDA
                .setARTICULO_ID = Trim(xTarifa.Value(i, ColsTarifa.ID))
                If Trim(xTarifa.Value(i, ColsTarifa.PFABRICA)) = "" Then
                    .setPRECIO_FABRICA = "0.00"
                Else
                    .setPRECIO_FABRICA = Replace(Format(xTarifa.Value(i, ColsTarifa.PFABRICA), "0.00"), ",", ".")
                End If
                If Trim(xTarifa.Value(i, ColsTarifa.POBRA)) = "" Then
                    .setPRECIO_OBRA = "0.00"
                Else
                    .setPRECIO_OBRA = Replace(Format(xTarifa.Value(i, ColsTarifa.POBRA), "0.00"), ",", ".")
                End If
                If .Insertar = 0 Then
                    Exit Sub
                Else
                End If
            End With
        End If
    Next
    MsgBox "La tarifa se ha almacenado correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmOfertas_Asignar"
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To ColTarifa
'        x(grid.Row, i) = ""
'        grid.RecordSelectors
        gridTarifa.SelBookmarks.Add gridTarifa.Bookmark
        xTarifa(gridTarifa.Bookmark, i) = ""
        gridTarifa.SelBookmarks.Remove 0
    Next
    gridTarifa.Refresh
    gridTarifa.SetFocus
End Sub

Private Sub cmdSalir_Click()
'    If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Unload Me
'    End If
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
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
    inicializar_grid
    cargar_articulos
    If pk > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.Carga pk
        FOFERTA = oOferta.getFECHA
        Set oOferta = Nothing
    End If
End Sub

Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error
   
    gridTarifa.Col = 0
    gridTarifa.Row = 0
    xTarifa.Clear
    xTarifa.ReDim 0, filasTarifa, 0, ColTarifa
    xTarifa.Clear
    Set gridTarifa.Array = xTarifa
    gridTarifa.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub


Private Sub cargar_obra(OBRA As Long)
    On Error GoTo fallo
    Dim oObra As New clsObras
    inicializar_grid
    If oObra.Carga(OBRA) = True Then
'       txtCliente(0) = oObra.getNOMBRE
'       If oObra.getFECHA_OFERTA <> "" Then
'            FOFERTA = oObra.getFECHA_OFERTA
'       End If
       ' Cargamos los datos de la tarifa
       Dim oTarifa As New clsTarifas
       Dim rs As ADODB.Recordset
       Set rs = oTarifa.Listado(OBRA)
       If rs.RecordCount > 0 Then
            Dim fila As Long
            fila = 0
            Do
                xTarifa(fila, ColsTarifa.ID) = CStr(rs(0))
                xTarifa(fila, ColsTarifa.ARTICULO) = CStr(rs(1))
                xTarifa(fila, ColsTarifa.PFABRICA) = CStr(rs(2))
                xTarifa(fila, ColsTarifa.POBRA) = CStr(rs(3))
                rs.MoveNext
                fila = fila + 1
            Loop Until rs.EOF
'            gridTarifa.SetFocus
        End If
        Set oTarifa = Nothing
        ' Cargamos los articulos de la oferta
        Dim oOD As New clsOfertas_detalle
        Set rs = oOD.Listado(pk)
        If rs.RecordCount > 0 Then
            Dim i As Integer
            Dim encontrado As Boolean
            Dim oArticulo As New clsArticulos
'            Dim S As New TrueDBGrid80.Style
'            S.Font.Bold = True
'            S.ForeColor = vbref
'            S.Font.Name = "Tahoma"
'            S.Font.Size = 20
            Do
                If rs(2) <> 0 Or rs(3) <> 0 Then
                    encontrado = False
                    For i = 0 To fila
                        If xTarifa(i, ColsTarifa.ID) = CStr(rs(0)) Then
                            encontrado = True
                            xTarifa(i, ColsTarifa.ARTICULO) = xTarifa(i, ColsTarifa.ARTICULO) & " (*)"
                            xTarifa(i, ColsTarifa.PFABRICA) = CStr(rs(2))
                            xTarifa(i, ColsTarifa.POBRA) = CStr(rs(3))
'                            gridTarifa.Row = i
'                            gridTarifa.Columns(ColsTarifa.PFABRICA).AddCellStyle 1, S
'                            gridTarifa.Columns(ColsTarifa.PFABRICA).AddCellStyle 1, S
'                            gridTarifa.Columns(ColsTarifa.POBRA).AddCellStyle 0, S
'                            gridTarifa.Columns(ColsTarifa.POBRA).AddCellStyle 1, S
                        End If
                    Next
                    If Not encontrado Then
                        xTarifa(fila, ColsTarifa.ID) = CStr(rs(0))
                        oArticulo.Carga rs(0)
                        xTarifa(fila, ColsTarifa.ARTICULO) = CStr(oArticulo.getDESCRIPCION) & " (*)"
                        xTarifa(fila, ColsTarifa.PFABRICA) = CStr(rs(2))
                        xTarifa(fila, ColsTarifa.POBRA) = CStr(rs(3))
'                        gridTarifa.Row = fila
'                        gridTarifa.Columns(ColsTarifa.PFABRICA).AddCellStyle 0, S
'                        gridTarifa.Columns(ColsTarifa.PFABRICA).AddCellStyle 1, S
'                        gridTarifa.Columns(ColsTarifa.POBRA).AddCellStyle 0, S
'                        gridTarifa.Columns(ColsTarifa.POBRA).AddCellStyle 1, S
                        fila = fila + 1
                    End If
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oOD = Nothing
        Set rs = Nothing
    Else
        MsgBox "Error al cargar la tarifa.", vbInformation, App.Title
    End If
    gridTarifa.Row = 0
    gridTarifa.Col = 0
    gridTarifa.Refresh
    Set oObra = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub cargar_articulos()
    Dim rs As ADODB.Recordset
    Dim oArt As New clsArticulos
    Set rs = oArt.ListadoTarifa()
    xarticulos.Clear
    If rs.RecordCount > 0 Then
        xarticulos.ReDim 1, rs.RecordCount, 1, 3
        Dim i As Integer
        i = 1
        Do
            xarticulos(i, 1) = CStr(rs(0))
            xarticulos(i, 2) = CStr(rs(1))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xarticulos.ReDim 1, 1, 1, 3
    End If
    Set tArticulos.Array = xarticulos
    tArticulos.Refresh
    gridTarifa.Refresh
End Sub

Private Sub gridTarifa_KeyPress(KeyAscii As Integer)
    If (gridTarifa.Col = ColsTarifa.PFABRICA Or gridTarifa.Col = ColsTarifa.POBRA) And KeyAscii = 46 Then
         KeyAscii = 44
    End If

End Sub

Private Sub tArticulos_DropDownClose()
    gridTarifa.Columns(ColsTarifa.ID) = tArticulos.Columns(0)
    gridTarifa.Columns(ColsTarifa.ARTICULO) = tArticulos.Columns(1)
    gridTarifa.Col = 2
End Sub
