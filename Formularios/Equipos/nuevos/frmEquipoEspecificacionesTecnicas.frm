VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmEquipoEspecificacionesTecnicas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Especificaciones Técnicas del Equipo"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12615
   Icon            =   "frmEquipoEspecificacionesTecnicas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Seleccionado"
      Height          =   930
      Left            =   45
      Picture         =   "frmEquipoEspecificacionesTecnicas.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6435
      Width           =   2280
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6435
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11505
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6435
      Width           =   1050
   End
   Begin TrueDBGrid80.TDBDropDown tParametros 
      Height          =   4755
      Left            =   45
      TabIndex        =   3
      Top             =   1125
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   8387
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
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AnchorRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=5133"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5054"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0E0E0&,.bold=0,.fontsize=825"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5745
      Left            =   45
      TabIndex        =   4
      Top             =   630
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   10134
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Parámetro Técnico"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tParametros"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Norma"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Especificación"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Requisito"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "tMetodos"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Observaciones"
      Columns(4).DataField=   ""
      Columns(4).ExternalEditor=   "TDBDate1"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).ValueItems(0)._DefaultItem=   0
      Columns(5).ValueItems(0).Value=   "0"
      Columns(5).ValueItems(0).Value.vt=   8
      Columns(5).ValueItems(0).DisplayValue=   "0"
      Columns(5).ValueItems(0).DisplayValue.vt=   8
      Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems(1)._DefaultItem=   0
      Columns(5).ValueItems(1).Value=   "1"
      Columns(5).ValueItems(1).Value.vt=   8
      Columns(5).ValueItems(1).DisplayValue=   "1"
      Columns(5).ValueItems(1).DisplayValue.vt=   8
      Columns(5).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(5).ValueItems.Count=   2
      Columns(5).Caption=   "Conforme"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "True/False"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "PARAMETRO_ID"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "NORMA_ID"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5106"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5027"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
      Splits(0)._ColumnProps(8)=   "Column(0).AutoCompletion=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=3889"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3810"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=2910"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2831"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=2646"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2566"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=5927"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=5847"
      Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(32)=   "Column(4).WrapText=1"
      Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(34)=   "Column(5).Width=344"
      Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=265"
      Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(40)=   "Column(5).AutoDropDown=1"
      Splits(0)._ColumnProps(41)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(47)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HC0E0FF&,.bold=0,.fontsize=825"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=11,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=12"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=11,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=11,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=12"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=11,.alignment=0,.wraptext=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=12"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=11,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=12"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=36,.parent=11"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=33,.parent=12"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=34,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=35,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=11"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=12"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=15"
      _StyleDefs(68)  =   "Named:id=37:Normal"
      _StyleDefs(69)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(70)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(71)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(72)  =   "Named:id=38:Heading"
      _StyleDefs(73)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(75)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(76)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(77)  =   "Named:id=39:Footing"
      _StyleDefs(78)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(79)  =   "Named:id=40:Selected"
      _StyleDefs(80)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(81)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(82)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(83)  =   "Named:id=41:Caption"
      _StyleDefs(84)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(85)  =   "Named:id=42:HighlightRow"
      _StyleDefs(86)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(87)  =   "Named:id=43:EvenRow"
      _StyleDefs(88)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=44:OddRow"
      _StyleDefs(90)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(91)  =   "Named:id=47:RecordSelector"
      _StyleDefs(92)  =   ":id=47,.parent=38"
      _StyleDefs(93)  =   "Named:id=50:FilterBar"
      _StyleDefs(94)  =   ":id=50,.parent=37"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre la linea y botónd derecho del ratón para ver la norma"
      Height          =   240
      Left            =   3465
      TabIndex        =   6
      Top             =   6435
      Width           =   4785
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especificaciones Técnicas del Equipo"
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
      Left            =   135
      TabIndex        =   2
      Top             =   180
      Width           =   4005
   End
   Begin VB.Image imagen 
      Height          =   420
      Left            =   12015
      Picture         =   "frmEquipoEspecificacionesTecnicas.frx":12B4
      Top             =   45
      Width           =   420
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12600
   End
End
Attribute VB_Name = "frmEquipoEspecificacionesTecnicas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Dim xP As New XArrayDB
Dim xParametros As New XArrayDB
Const filas As Integer = 100
Const Col As Integer = 7
Private Enum COLS
    PARAMETRO = 0
    NORMA = 1
    ESPECIFICACION = 2
    REQUISITO = 3
    OBSERVACIONES = 4
    CONFORME = 5
    PARAMETRO_ID = 6
    NORMA_ID = 7
End Enum
Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Col
        grid.SelBookmarks.Add grid.Bookmark
        xP(grid.Bookmark, i) = ""
        grid.SelBookmarks.Remove 0
    Next
    grid.Refresh
    grid.SetFocus
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    Dim oee As New clsEq_especificaciones_tecnicas
   On Error GoTo cmdok_Click_Error

    oee.Eliminar PK
    Dim i As Integer
    For i = xP.LowerBound(1) To xP.UpperBound(1)
        If Trim(xP.value(i, COLS.PARAMETRO)) <> "" Then
            With oee
                .setEQUIPO_ID = PK
                .setPARAMETRO_ID = xP.value(i, COLS.PARAMETRO_ID)
                .setESPECIFICACION = xP.value(i, COLS.ESPECIFICACION)
                .setREQUISITOS = xP.value(i, COLS.REQUISITO)
                .setOBSERVACIONES = xP.value(i, COLS.OBSERVACIONES)
                If xP.value(i, COLS.CONFORME) = "" Then
                    .setCONFORME = 0
                Else
                    .setCONFORME = 1
                End If
'                .setCONFORME = x.value(i, COLS.CONFORME)
                .Insertar
            End With
        End If
    Next
    Set oee = Nothing
    MsgBox "Parametros almacenados correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmEquipoEspecificacionesTecnicas"
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    inicializar_grid
    If PK <> 0 Then
        CARGAR
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEquipoEspecificacionesTecnicas = Nothing
End Sub

Private Sub CARGAR()
    On Error GoTo fallo
    Dim oee As New clsEq_especificaciones_tecnicas
    Dim RS As ADODB.RecordSet
    Set RS = oee.Listado(PK)
    If RS.RecordCount > 0 Then
        Dim i As Integer
        i = 0
        Do
            xP(i, COLS.PARAMETRO) = CStr(RS(0))
            If Not IsNull(RS(1)) Then
                xP(i, COLS.NORMA) = CStr(RS(1))
            End If
            xP(i, COLS.ESPECIFICACION) = CStr(RS(2))
            xP(i, COLS.REQUISITO) = CStr(RS(3))
            xP(i, COLS.OBSERVACIONES) = CStr(RS(4))
            If RS(4) = 1 Then
                xP(i, COLS.CONFORME) = CStr(RS(5))
            End If
            xP(i, COLS.PARAMETRO_ID) = CStr(RS(6))
            xP(i, COLS.NORMA_ID) = CStr(RS(7))
            i = i + 1
            RS.MoveNext
        Loop Until RS.EOF
    End If
    
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del proveedor.", vbCritical, Err.Description
End Sub

Private Sub cargar_combos()
    Dim RS As ADODB.RecordSet
    ' Parametros
    Dim oEP As New clsEq_parametros_tecnicos
    Set RS = oEP.Listado
    Dim i As Integer
    If RS.RecordCount > 0 Then
        xParametros.ReDim 1, RS.RecordCount, 1, 4
        i = 1
        Do
            xParametros(i, 1) = CStr(RS(0))
            xParametros(i, 2) = CStr(RS(1))
            If Not IsNull(RS(2)) Then
                xParametros(i, 3) = CStr(RS(2))
                xParametros(i, 4) = CStr(RS(3))
            Else
                xParametros(i, 3) = ""
                xParametros(i, 4) = "0"
            End If
            i = i + 1
            RS.MoveNext
        Loop Until RS.EOF
    Else
        xParametros.ReDim 1, 1, 1, 4
    End If
    Set tParametros.Array = xParametros
    tParametros.Refresh
End Sub
Private Sub inicializar_grid()
    xP.ReDim 0, filas, 0, Col
    xP.Clear
    Set grid.Array = xP
    grid.Refresh
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo grid_MouseDown_Error

    If Button And vbRightButton Then
        If Not IsEmpty(xP(grid.Row, COLS.NORMA_ID)) Then
            If xP(grid.Row, COLS.NORMA_ID) <> 0 Then
                Dim oNorma As New clsCa_normas
                oNorma.mostrar PK, True
                Set oNorma = Nothing
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

grid_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grid_MouseDown of Formulario frmEquipoEspecificacionesTecnicas"

End Sub

Private Sub tParametros_DropDownClose()
    grid.Columns(COLS.PARAMETRO) = tParametros.Columns(1)
    grid.Columns(COLS.PARAMETRO_ID) = tParametros.Columns(0)
    grid.Columns(COLS.NORMA_ID) = tParametros.Columns(3)
    grid.Columns(COLS.NORMA) = tParametros.Columns(2)
    grid.Col = COLS.ESPECIFICACION
End Sub
