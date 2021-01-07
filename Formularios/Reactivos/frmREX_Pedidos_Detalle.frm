VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmREX_Pedidos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Pedido a Proveedor"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14430
   Icon            =   "frmREX_Pedidos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   14430
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seguimiento"
      Height          =   885
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   9135
      Width           =   1365
   End
   Begin VB.CommandButton cmdFacturacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Facturas"
      Height          =   885
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   9135
      Width           =   1155
   End
   Begin TrueDBGrid80.TDBDropDown tFamilias 
      Height          =   4080
      Left            =   5490
      TabIndex        =   32
      Top             =   4275
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   7197
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=6562"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6456"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=291"
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
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   885
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   9135
      Width           =   1155
   End
   Begin TrueDBGrid80.TDBDropDown tBotes 
      Height          =   3540
      Left            =   90
      TabIndex        =   30
      Top             =   4005
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6244
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
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=9102"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8996"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3334"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3228"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=291"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=185"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=344"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=238"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(22)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2752"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
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
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=1,.bold=0,.fontsize=825"
      _StyleDefs(51)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(52)  =   ":id=28,.fontname=MS Sans Serif"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12645
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   8685
      Width           =   1725
   End
   Begin VB.CommandButton cmdCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enviar Correo"
      Height          =   885
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9135
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   10950
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9135
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   885
      Left            =   12090
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9135
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   885
      Left            =   13230
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9135
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   45
      TabIndex        =   16
      Top             =   585
      Width           =   14340
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   1
         Left            =   1665
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1440
         Width           =   10845
      End
      Begin VB.CheckBox chkFechaConfirmada 
         Caption         =   "Check1"
         Height          =   240
         Left            =   7290
         TabIndex        =   4
         Top             =   585
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CheckBox chkFechaEnvio 
         Caption         =   "Check1"
         Height          =   240
         Left            =   4275
         TabIndex        =   2
         Top             =   585
         Width           =   195
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   1665
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   930
         Width           =   10845
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   330
         Left            =   1665
         TabIndex        =   0
         Top             =   180
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   540
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
         CalendarTitleBackColor=   14737632
         Format          =   60293121
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaEnvio 
         Height          =   330
         Left            =   4560
         TabIndex        =   3
         Top             =   540
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
         CalendarTitleBackColor=   14737632
         Format          =   60293121
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaConfirmada 
         Height          =   330
         Left            =   7530
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
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
         CalendarTitleBackColor=   14737632
         Format          =   60293121
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbMoneda 
         Height          =   315
         Left            =   10095
         TabIndex        =   28
         Top             =   540
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "0"
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
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Pedidos_Detalle.frx":030A
         Height          =   315
         Left            =   1665
         TabIndex        =   8
         Top             =   2160
         Width           =   3225
         _ExtentX        =   5689
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   180
         TabIndex        =   31
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Moneda"
         Height          =   195
         Index           =   6
         Left            =   9405
         TabIndex        =   27
         Top             =   585
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   24
         Top             =   1635
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista"
         Height          =   195
         Index           =   10
         Left            =   6075
         TabIndex        =   23
         Top             =   585
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Envio"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   22
         Top             =   585
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Pedido"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   225
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Motivo"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1035
         Width           =   480
      End
   End
   Begin TrueDBGrid80.TDBGrid gridComponentes 
      Height          =   5520
      Left            =   45
      TabIndex        =   10
      Top             =   3150
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   9737
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID_PEDIDO"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "TIPO_BOTE_EX_ID"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Descripción"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "tBotes"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Código"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Familia"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Standard"
      Columns(4).DropDown=   "tFamilias"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Prec.Unit."
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "#,##0.00"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Cantidad"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Dto."
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Total"
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "#,##0.00"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "FAMILIA_ID"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "F.Prevista"
      Columns(10).DataField=   ""
      Columns(10).ExternalEditor=   "TDBDate1"
      Columns(10).ExternalEditor.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=953"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=9155"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=9049"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=3334"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3228"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=4551"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4445"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(30)=   "Column(5).Width=2064"
      Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=1958"
      Splits(0)._ColumnProps(33)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(36)=   "Column(6).Width=1826"
      Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=1720"
      Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(6).DropDownList=1"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2170"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2064"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(48)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=1852"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1746"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=8194"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(56)=   "Column(9).Width=1217"
      Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1111"
      Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(62)=   "Column(10).Width=873"
      Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=767"
      Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=1"
      Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=11"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=11"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=24,.parent=11,.alignment=0,.locked=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=11,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=11"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=11,.alignment=1,.bgcolor=&HDEEDFA&"
      _StyleDefs(57)  =   ":id=54,.locked=0"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=12"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=36,.parent=11,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=33,.parent=12"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=34,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=35,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=58,.parent=11,.alignment=2"
      _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=12"
      _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=62,.parent=11,.alignment=1,.bgcolor=&HC0C0C0&"
      _StyleDefs(70)  =   ":id=62,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=12"
      _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=32,.parent=11"
      _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=12"
      _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=74,.parent=11,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=12"
      _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=15"
      _StyleDefs(82)  =   "Named:id=37:Normal"
      _StyleDefs(83)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(84)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(85)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(86)  =   "Named:id=38:Heading"
      _StyleDefs(87)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(89)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(90)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(91)  =   "Named:id=39:Footing"
      _StyleDefs(92)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(93)  =   "Named:id=40:Selected"
      _StyleDefs(94)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(95)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(96)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(97)  =   "Named:id=41:Caption"
      _StyleDefs(98)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(99)  =   "Named:id=42:HighlightRow"
      _StyleDefs(100) =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(101) =   "Named:id=43:EvenRow"
      _StyleDefs(102) =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(103) =   "Named:id=44:OddRow"
      _StyleDefs(104) =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(105) =   "Named:id=47:RecordSelector"
      _StyleDefs(106) =   ":id=47,.parent=38"
      _StyleDefs(107) =   "Named:id=50:FilterBar"
      _StyleDefs(108) =   ":id=50,.parent=37"
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   285
      Left            =   3915
      TabIndex        =   33
      Top             =   9315
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   503
      Calendar        =   "frmREX_Pedidos_Detalle.frx":0350
      Caption         =   "frmREX_Pedidos_Detalle.frx":0468
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmREX_Pedidos_Detalle.frx":04D4
      Keys            =   "frmREX_Pedidos_Detalle.frx":04F2
      Spin            =   "frmREX_Pedidos_Detalle.frx":0550
      AlignHorizontal =   2
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   1
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "27/10/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40478
      CenturyMode     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse DOBLE-CLICK para ver el detalle y STOCK del producto seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   450
      TabIndex        =   29
      Top             =   8730
      Width           =   8790
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   11160
      TabIndex        =   26
      Top             =   8730
      Width           =   1365
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del Pedido a Proveedor"
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
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   45
      Width           =   11505
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   18
      Top             =   360
      Width           =   90
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frmREX_Pedidos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Public TIPO_BOTE_ID As Long
Dim xP As New XArrayDB
Dim xTB As New XArrayDB
Dim xTF As New XArrayDB

Const filasP As Integer = 100
Const ColP As Integer = 10
Private Enum ColsP
    ID_PEDIDO = 0
    TIPO_BOTE_EX = 1
    DESCRIPCION = 2
    CODIGO = 3
    familia = 4
    PRECIO_UNIDAD = 5
    CANTIDAD = 6
    DESCUENTO = 7
    IMPORTE = 8
    FAMILIA_ID = 9
    fPrevista = 10
End Enum


Private Sub cmdHistorialCambios_Click()
    frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_PEDIDOS
    frmHistorialCambios.PK_ID = PK
    frmHistorialCambios.PK_TITULO = "Seguimiento de Pedido"
    frmHistorialCambios.Show 1
End Sub


' botones
'M1271-I
Private Sub cmdFacturacion_Click()
'    frmProveedores_Facturas.TOBJETO = TOBJETO.TOBJETO_SC_DETERMINACIONES
'    frmProveedores_Facturas.COBJETO = PK
    frmProveedores_Facturas.PK = cmbProveedor.getPK_SALIDA
    frmProveedores_Facturas.Show 1
End Sub
'M1271-F


Private Sub cmdAdjuntos_Click()
    Dim oPed As New clsPedidos_bote_ex
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PEDIDO_BOTE_EX
        .COBJETO = oPed.calcularID_PEDIDO(PK, TIPO_BOTE_ID)
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    Set oPed = Nothing
End Sub

Private Sub chkFechaConfirmada_Click()
    If chkFechaConfirmada.Value = Checked Then
        fechaConfirmada.Enabled = True
    Else
        fechaConfirmada.Enabled = False
    End If
End Sub

Private Sub chkFechaEnvio_Click()
    If chkFechaEnvio.Value = Checked Then
        fechaEnvio.Enabled = True
    Else
        fechaEnvio.Enabled = False
    End If
End Sub

Private Sub cmbProveedor_change()
    cargarTiposBotes
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCorreo_Click()

    Dim oPed As New clsPedidos_bote_ex
    Dim destino As String
    
   On Error GoTo cmdCorreo_Click_Error

    If guardarPedido = False Then
        MsgBox "Error al guardar los cambios del pedido.", vbCritical, App.Title
        Exit Sub
    End If

    oPed.CARGAR PK
    destino = DIRECTORIO_TEMPORAL & "Pedido " & Format(oPed.getCODIGO_PEDIDO_PROVEEDOR, "000") & "." & Year(oPed.getFECHA_PEDIDO) & ".pdf"
    oPed.imprimir oPed.getCODIGO_PEDIDO_PROVEEDOR, Year(oPed.getFECHA_PEDIDO), destino
    ' Verificar si se ha creado
    If Dir(destino) = "" Then
        MsgBox "El documento no se ha generado correctamente.", vbCritical, App.Title
        Exit Sub
    End If
    ' Enviar por correo
    Dim oProveedor As New clsProveedor
    oProveedor.Carga oPed.getPROVEEDOR_ID
    Dim cuerpo As String
    Dim ASUNTO As String
    ASUNTO = "Pedido " & Format(oPed.getCODIGO_PEDIDO_PROVEEDOR, "000") & "." & Year(oPed.getFECHA_PEDIDO)
    cuerpo = "Adjunto el pedido número " & Format(oPed.getCODIGO_PEDIDO_PROVEEDOR, "000") & "." & Year(oPed.getFECHA_PEDIDO)
    If Trim(oProveedor.getEMAIL) <> "" Then
        enviar_correo oProveedor.getEMAIL, "", "", True, cuerpo, ASUNTO, destino
    Else
        enviar_correo "Introduzca destinatario", "", "", True, cuerpo, ASUNTO, destino
    End If

   On Error GoTo 0
   Exit Sub

cmdCorreo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCorreo_Click of Formulario frmREX_Pedidos_Detalle"

End Sub

Private Sub cmdImprimir_Click()
    Dim oPed As New clsPedidos_bote_ex
   On Error GoTo cmdImprimir_Click_Error
    If guardarPedido = False Then
        MsgBox "Error al guardar los cambios del pedido.", vbCritical, App.Title
        Exit Sub
    End If

    oPed.CARGAR PK
    oPed.imprimir oPed.getCODIGO_PEDIDO_PROVEEDOR, Year(oPed.getFECHA_PEDIDO), ""
    Set oPed = Nothing

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmREX_Pedidos_Detalle"
End Sub
'E0150-I
Private Sub cmdok_Click()
    If guardarPedido Then MsgBox "Pedido modificado correctamente.", vbInformation, App.Title
End Sub
Private Function guardarPedido() As Boolean
    Dim auxiliar As Long, i As Integer
    Dim strTipo As String
    
   On Error GoTo cmdok_Click_Error
    guardarPedido = False
    Dim oPedido_bote_ex As New clsPedidos_bote_ex
    Dim pedido As Long
    Me.MousePointer = 11
    If datos_correctos Then
        Dim oPed As New clsPedidos_bote_ex
        For i = 0 To filasP
            If xP(i, ColsP.ID_PEDIDO) <> vbNull Then
                If xP(i, ColsP.ID_PEDIDO) <> "" Then
                    With oPed
                        .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
                        .setCENTRO_ID = cmbCentro.BoundText
                        If xP(i, ColsP.FAMILIA_ID) = vbNull Then
                            .setFAMILIA_ID = 0
                        ElseIf Trim(xP(i, ColsP.FAMILIA_ID)) = "" Then
                            .setFAMILIA_ID = 0
                        Else
                            .setFAMILIA_ID = xP(i, ColsP.FAMILIA_ID)
                        End If
                        If chkFechaEnvio.Value = Checked Then
                            .setENVIADO = Format(fechaEnvio, "dd-mm-yyyy")
                        Else
                            .setENVIADO = ""
                        End If
'                        If chkFechaConfirmada.value = Checked Then
'                            .setCONFIRMADO = Format(fechaConfirmada, "dd-mm-yyyy")
'                        Else
'                            .setCONFIRMADO = ""
'                        End If
                        .setCONFIRMADO = xP(i, ColsP.fPrevista)
                        
                        .setFECHA_PEDIDO = Format(fecha, "yyyy-mm-dd")
                        .setMOTIVO = txtDatos(0)
                        .setOBSERVACIONES = txtDatos(1)
                        .setCANTIDAD = xP(i, ColsP.CANTIDAD)
                        .setPRECIO = moneda_bd(xP(i, ColsP.PRECIO_UNIDAD))
                        .setDTO = xP(i, ColsP.DESCUENTO)
                        'M1058-I
                        .setMONEDA = cmbMoneda.BoundText
                        'M1058-F
                        If xP(i, ColsP.ID_PEDIDO) = 0 Then
                            .CrearID
                            .setFECHA = Format(fecha, "yyyy-mm-dd")
                            .setRECIBIDO = 0
                            .setPRIORIDAD = 0
                            .setDTO = 0
                            .setUSUARIO = USUARIO.getID_EMPLEADO
                            .setMOTIVO = "Añadido manualmente"
                            .setTIPO_BOTE_EX_ID = xP(i, ColsP.TIPO_BOTE_EX)
                            Dim oPed2 As New clsPedidos_bote_ex
                            Dim numeroPedido As Long
                            oPed2.CARGAR PK
                            .setCODIGO_PEDIDO_PROVEEDOR = oPed2.getCODIGO_PEDIDO_PROVEEDOR
                            .setRECIBIDO = oPed2.getRECIBIDO
                            .setTRAMITADO_FECHA = oPed2.getTRAMITADO_FECHA
                            .setTRAMITADO_USUARIO_ID = oPed2.getTRAMITADO_USUARIO_ID
                            numeroPedido = .Insertar
                            xP(i, ColsP.ID_PEDIDO) = CStr(numeroPedido)
                            Set oPed2 = Nothing
                        Else
                            .Modificar xP(i, ColsP.ID_PEDIDO), xP(i, ColsP.TIPO_BOTE_EX)
                        End If
                    End With
                
                End If
            End If
        Next
        guardarPedido = True
    End If
    Me.MousePointer = 0
   On Error GoTo 0
   Exit Function

cmdok_Click_Error:
    guardarPedido = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Pedidos_Detalle"
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    llenar_combo cmbProveedor, New clsProveedor, 0, frmProveedores_Detalle, ""
    cargar_combo cmbCentro, New clsCentros
'    llenar_combo cmbRecepcionado, New clsUsuarios, 0, frmUsuarios, ""
    'M1058-I
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbMoneda, DECODIFICADORA.DECODIFICADORA_MONEDA
'    llenar_combo cmbCC, New clsFamilias, 0, Me, " PEDIDO = 1 "
    cargarFamilias
    'M1058-F
    fechaEnvio = Date
'    fechaConfirmada = Date
'    cmbRecepcionado.desactivar
    If Not USUARIO.getPER_TESORERIA_FP Then
        cmdFacturacion.visible = False
    End If

    If PK <> 0 Then
        cargarPedido
    End If
End Sub

Private Sub gridComponentes_AfterColEdit(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case ColsP.CANTIDAD, ColsP.PRECIO_UNIDAD, ColsP.DESCUENTO
            calcular_fila
    End Select
End Sub


Private Sub gridComponentes_AfterUpdate()
    calcular_total
End Sub

Private Sub gridComponentes_DblClick()
   On Error GoTo gridComponentes_DblClick_Error

    gridComponentes.Col = ColsP.TIPO_BOTE_EX
    If gridComponentes.Text <> "" Then
        frmREX_Bote.PK = gridComponentes.Text
        frmREX_Bote.Show 1
    End If

   On Error GoTo 0
   Exit Sub

gridComponentes_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gridComponentes_DblClick of Formulario frmREX_Pedidos_Detalle"
End Sub

Private Sub tBotes_DropDownClose()
    If tBotes.Columns(0) <> "" Then
        gridComponentes.Columns(ColsP.ID_PEDIDO) = 0
        gridComponentes.Columns(ColsP.CODIGO) = tBotes.Columns(1)
        gridComponentes.Columns(ColsP.TIPO_BOTE_EX) = tBotes.Columns(4)
        gridComponentes.Columns(ColsP.PRECIO_UNIDAD) = tBotes.Columns(2)
        gridComponentes.Columns(ColsP.CANTIDAD) = tBotes.Columns(3)
        gridComponentes.Columns(ColsP.IMPORTE) = tBotes.Columns(2) * tBotes.Columns(3)
        gridComponentes.Col = ColsP.familia
    End If
End Sub

Private Sub tFamilias_DropDownClose()
    If tFamilias.Columns(0) <> "" Then
        gridComponentes.Columns(ColsP.familia) = tFamilias.Columns(0)
        gridComponentes.Columns(ColsP.FAMILIA_ID) = tFamilias.Columns(1)
        gridComponentes.Col = ColsP.PRECIO_UNIDAD
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargarPedido()
    Dim oPed As New clsPedidos_bote_ex
'    oPed.CARGAR PK
    oPed.cargar_con_bote PK, TIPO_BOTE_ID
    Dim rs As ADODB.Recordset
    Dim lngNuevaFila As Long
'JGM    Set RS = oPed.CARGAR_POR_PEDIDO_PROVEEDOR(oPed.getCODIGO_PEDIDO_PROVEEDOR, Year(oPed.getFECHA))
    If oPed.getFECHA_PEDIDO = "" Then
        Set rs = oPed.CARGAR_POR_PEDIDO_PROVEEDOR(oPed.getCODIGO_PEDIDO_PROVEEDOR, Year(oPed.getFECHA))
    Else
        Set rs = oPed.CARGAR_POR_PEDIDO_PROVEEDOR(oPed.getCODIGO_PEDIDO_PROVEEDOR, Year(oPed.getFECHA_PEDIDO))
    End If
    
    If rs.RecordCount > 0 Then
        'M1076-I
        lbltitulo(0) = "Detalle del Pedido a Proveedor, Nº " & oPed.getCODIGO_PEDIDO_PROVEEDOR
        Me.Caption = lbltitulo(0)
        'M1076-F
        cmbProveedor.MostrarElemento rs("PROVEEDOR_ID")
        cmbCentro.BoundText = rs("CENTRO_ID")
'        cmbCC.MostrarElemento rs("FAMILIA_ID")
        If IsDate(rs("FECHA_PEDIDO")) Then
            fecha = rs("FECHA_PEDIDO")
        End If
        If rs("ENVIADO") = "" Then
            chkFechaEnvio.Value = Unchecked
            fechaEnvio.Enabled = False
        Else
            chkFechaEnvio.Value = Checked
            fechaEnvio = rs("ENVIADO")
            fechaEnvio.Enabled = True
        End If
'        If rs("CONFIRMADO") = "" Then
'            chkFechaConfirmada.value = Unchecked
'            fechaConfirmada.Enabled = False
'        Else
'            chkFechaConfirmada.value = Checked
'            fechaConfirmada = rs("CONFIRMADO")
'            fechaConfirmada.Enabled = True
'        End If
        txtDatos(0) = rs("MOTIVO")
        txtDatos(1) = rs("OBSERVACIONES")
        'M1058-I
        cmbMoneda.BoundText = rs("MONEDA")
        'M1058-F
'        If rs("RECEPCION_USUARIO") <> 0 Then
'            cmbRecepcionado.MostrarElemento rs("RECEPCION_USUARIO")
'        End If
'        If Not IsNull(rs("RECEPCION_FECHA")) Then
'            txtFecha = rs("RECEPCION_FECHA")
'        End If
        lngNuevaFila = 0
        Do
                
            xP(lngNuevaFila, ColsP.ID_PEDIDO) = CStr(rs("ID_PEDIDO_BOTE_EX"))
            xP(lngNuevaFila, ColsP.TIPO_BOTE_EX) = CStr(rs("TIPO_BOTE_EX_ID"))
            xP(lngNuevaFila, ColsP.CANTIDAD) = CStr(rs("CANTIDAD"))
            xP(lngNuevaFila, ColsP.DESCRIPCION) = CStr(rs("REACTIVO"))
            xP(lngNuevaFila, ColsP.CODIGO) = CStr(rs("CODIGO_PROVEEDOR"))
            If Not IsNull(rs("FAMILIA")) Then
                xP(lngNuevaFila, ColsP.familia) = CStr(rs("FAMILIA"))
            Else
                xP(lngNuevaFila, ColsP.familia) = ""
            End If
            xP(lngNuevaFila, ColsP.fPrevista) = CStr(rs("CONFIRMADO"))
            If Not IsNull(rs("FAMILIA_ID")) Then
                xP(lngNuevaFila, ColsP.FAMILIA_ID) = CStr(rs("FAMILIA_ID"))
            End If
            xP(lngNuevaFila, ColsP.PRECIO_UNIDAD) = moneda(rs("PRECIO"))
            xP(lngNuevaFila, ColsP.DESCUENTO) = CStr(rs("DTO"))
            Dim IMPORTE As Single
            Dim DESCUENTO As Single
            IMPORTE = moneda(rs("CANTIDAD") * rs("PRECIO"))
            DESCUENTO = (IMPORTE * rs("DTO")) / 100
            xP(lngNuevaFila, ColsP.IMPORTE) = Format(CCur(IMPORTE) - CCur(DESCUENTO), "currency")
            
            calcular_fila
            lngNuevaFila = lngNuevaFila + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPed = Nothing
    cargarTiposBotes
    calcular_total
End Sub
Private Function datos_correctos() As Boolean
    datos_correctos = True
    If cmbProveedor.getTEXTO = "" Then
        MsgBox "Debe indicar el Proveedor.", vbExclamation, App.Title
        cmbProveedor.SetFocus
        datos_correctos = False
        Exit Function
    End If
'    If cmbCC.getTEXTO = "" Then
'        MsgBox "Indique la familia del pedido.", vbCritical, App.Title
'        cmbCC.SetFocus
'        datos_correctos = False
'        Exit Function
'    End If
    
    If cmbCentro.Text = "" Then
        MsgBox "Debe indicar el CENTRO.", vbExclamation, App.Title
        cmbCentro.SetFocus
        datos_correctos = False
        Exit Function
    End If
    If total_filas_array = 0 Then
        MsgBox "Deben existir reactivos en el pedido.", vbExclamation, App.Title
        gridComponentes.SetFocus
        datos_correctos = False
        Exit Function
    End If
End Function
Private Sub inicializar_grid()
    On Error GoTo inicializar_grid_Error
    
    gridComponentes.Col = 0
    gridComponentes.Row = 0
    xP.Clear
    xP.ReDim 0, filasP, 0, ColP
    xP.Clear
    Set gridComponentes.Array = xP
    gridComponentes.Refresh
    
    On Error GoTo 0
    
    Exit Sub
    
inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmREX_Pedidos_Detalle"
End Sub
' Procedimiento que borra los datos del tdbgrid de la fila
Private Sub borrar_fila(lngFila As Long)
    Dim i As Integer
    For i = 0 To ColP
        xP(lngFila, i) = ""
    Next
End Sub
' Función que devuelve el número de filas (rellenas) que hay en el array
Private Function total_filas_array() As Long
    Dim lngFila As Long
    lngFila = 0
    While Not xP(lngFila, 0) = ""
        lngFila = lngFila + 1
    Wend
    total_filas_array = lngFila
End Function
Private Sub calcular_fila()
    If IsNumeric(gridComponentes.Columns(ColsP.CANTIDAD).Text) And IsNumeric(gridComponentes.Columns(ColsP.PRECIO_UNIDAD).Text) Then
        Dim dto As Single
        dto = 0
        If IsNumeric(gridComponentes.Columns(ColsP.DESCUENTO).Text) Then
            dto = gridComponentes.Columns(ColsP.DESCUENTO).Text
        End If
        Dim IMPORTE As Single
        Dim DESCUENTO As Single
        IMPORTE = gridComponentes.Columns(ColsP.CANTIDAD).Text * CSng(gridComponentes.Columns(ColsP.PRECIO_UNIDAD).Text)
        DESCUENTO = (IMPORTE * dto) / 100
        gridComponentes.Columns(ColsP.IMPORTE).Text = Format(CCur(IMPORTE) - CCur(DESCUENTO), "currency")
    End If
End Sub
Private Sub calcular_total()
    Dim i As Integer
    Dim total As Currency
    total = 0
    For i = 0 To filasP
        total = total + xP(i, ColsP.IMPORTE)
    Next
    'M1058-I
    ' txttotal = MONEDA(CStr(total))
    txttotal = Format(Replace(total, ".", ","), "#,##0.00")
    'M1058-F
End Sub

Private Sub cargarTiposBotes()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim PROVEEDOR_ID As Long
    If cmbProveedor.getTEXTO = "" Then
        PROVEEDOR_ID = 0
    Else
        PROVEEDOR_ID = cmbProveedor.getPK_SALIDA
    End If
    consulta = "SELECT TR.NOMBRE,TB.CODIGO_PROVEEDOR,TB.PRECIO,TB.CANTIDAD_MINIMA_PEDIDO,TB.ID_TIPO_BOTE_EX AS ID " & _
               "  FROM TIPOS_BOTE_EX AS TB, TIPOS_REACTIVO_EX AS TR" & _
               " WHERE TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX " & _
               "   AND TB.PROVEEDOR_ID = " & PROVEEDOR_ID & _
               " ORDER BY TR.NOMBRE "
    Set rs = datos_bd(consulta)
    xTB.Clear
    If rs.RecordCount > 0 Then
        xTB.ReDim 1, rs.RecordCount, 1, 5
        Dim i As Integer
        i = 1
        Do
            xTB(i, 1) = CStr(rs(0))
            xTB(i, 2) = CStr(rs(1))
            xTB(i, 3) = moneda(CStr(rs(2)))
            xTB(i, 4) = CStr(rs(3))
            xTB(i, 5) = CStr(rs(4))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xTB.ReDim 1, 1, 1, 5
    End If
    Set tBotes.Array = xTB
    tBotes.Refresh
End Sub
Private Sub cargarFamilias()
    Dim rs As ADODB.Recordset
    Dim ofam As New clsFamilias
    Set rs = ofam.ListadoPedido()
    xTF.Clear
    If rs.RecordCount > 0 Then
        xTF.ReDim 1, rs.RecordCount, 1, 2
        Dim i As Integer
        i = 1
        Do
            xTF(i, 1) = CStr(rs(0))
            xTF(i, 2) = CStr(rs(1))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xTF.ReDim 1, 1, 1, 2
    End If
    Set tFamilias.Array = xTF
    tFamilias.Refresh
End Sub

