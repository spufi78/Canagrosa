VERSION 5.00
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmCE_Ficha_Bano 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ficha de Control de Eficacia"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   15420
   Icon            =   "frmCE_Ficha_Bano.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid80.TDBDropDown tProducto 
      Height          =   2955
      Left            =   5625
      TabIndex        =   26
      Top             =   1710
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   5212
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame frmAreas 
      Height          =   5730
      Left            =   2475
      TabIndex        =   22
      Top             =   2340
      Visible         =   0   'False
      Width           =   10860
      Begin VB.CommandButton cmdCerrarCOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   510
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5130
         Width           =   1050
      End
      Begin TrueDBGrid80.TDBDropDown tmateriales 
         Height          =   4080
         Left            =   675
         TabIndex        =   24
         Top             =   810
         Width           =   10065
         _ExtentX        =   17754
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
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=8414"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8334"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin TrueDBGrid80.TDBGrid gridAreas 
         Height          =   4860
         Left            =   135
         TabIndex        =   23
         Top             =   225
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   8573
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Area"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Material/Pintura"
         Columns(1).DataField=   ""
         Columns(1).DropDown=   "tmateriales"
         Columns(1).DropDown.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Criterio"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=953"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=873"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=8193"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=8387"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=8308"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=0"
         Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(15)=   "Column(1).AutoDropDown=1"
         Splits(0)._ColumnProps(16)=   "Column(1).DropDownList=1"
         Splits(0)._ColumnProps(17)=   "Column(1).AutoCompletion=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
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
         Caption         =   "Detalle de las Áreas"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=11,.alignment=0"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=11"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
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
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   960
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9585
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cargar los datos de la ficha a partir del siguiente baño:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2700
      TabIndex        =   17
      Top             =   9630
      Width           =   8025
      Begin VB.CommandButton cmdCargar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cargar"
         Height          =   690
         Left            =   6930
         Picture         =   "frmCE_Ficha_Bano.frx":2AFA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1005
      End
      Begin pryCombo.miCombo cmbBano 
         Height          =   375
         Left            =   945
         TabIndex        =   18
         Top             =   315
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   19
         Top             =   405
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdNuevoMaterial 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nueva Pintura"
      Height          =   960
      Left            =   10890
      Picture         =   "frmCE_Ficha_Bano.frx":2D6B
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9585
      Width           =   1095
   End
   Begin VB.CommandButton cmdborrarensayo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Ensayo"
      Height          =   930
      Left            =   1350
      Picture         =   "frmCE_Ficha_Bano.frx":3635
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9585
      Width           =   1245
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Probeta"
      Height          =   930
      Left            =   45
      Picture         =   "frmCE_Ficha_Bano.frx":3EFF
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9585
      Width           =   1245
   End
   Begin TrueDBGrid80.TDBDropDown tProbetas 
      Height          =   3855
      Left            =   10755
      TabIndex        =   13
      Top             =   4950
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin TrueDBGrid80.TDBDropDown tAnalisis 
      Height          =   3630
      Left            =   45
      TabIndex        =   12
      Top             =   4950
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6403
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
      Columns.Count   =   3
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=10769"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=10689"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3784"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3704"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=318"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=238"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(17)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(47)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(48)  =   ":id=28,.fontname=MS Sans Serif"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9585
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   14265
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9585
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   45
      TabIndex        =   5
      Top             =   405
      Width           =   15300
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   12285
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   2700
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   4725
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   5040
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cadencia"
         Height          =   195
         Index           =   4
         Left            =   11520
         TabIndex        =   9
         Top             =   270
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   0
         Left            =   5895
         TabIndex        =   7
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   300
         Width           =   585
      End
   End
   Begin TrueDBGrid80.TDBGrid gridP 
      Height          =   3465
      Left            =   45
      TabIndex        =   0
      Top             =   1170
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   6112
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Designación"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Material/Pintura"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Des. Producto"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "tProducto"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Dimensión"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nº Probetas"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "General Number"
      Columns(4).ExternalEditor=   "TDBDate1"
      Columns(4).ExternalEditor.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Áreas"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AllowRowSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5609"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5530"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=6376"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6297"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=8467"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=8387"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=3704"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=3625"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1640"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1561"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=185"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=106"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
      Caption         =   "I. Probetas (Botón derecho para ver el detalle de las Areas)"
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=28,.parent=11,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=12"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=62,.parent=11,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=12"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=15"
      _StyleDefs(64)  =   "Named:id=37:Normal"
      _StyleDefs(65)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(66)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(67)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(68)  =   "Named:id=38:Heading"
      _StyleDefs(69)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(71)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(72)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(73)  =   "Named:id=39:Footing"
      _StyleDefs(74)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=40:Selected"
      _StyleDefs(76)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(77)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(78)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(79)  =   "Named:id=41:Caption"
      _StyleDefs(80)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(81)  =   "Named:id=42:HighlightRow"
      _StyleDefs(82)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(83)  =   "Named:id=43:EvenRow"
      _StyleDefs(84)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=44:OddRow"
      _StyleDefs(86)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(87)  =   "Named:id=47:RecordSelector"
      _StyleDefs(88)  =   ":id=47,.parent=38"
      _StyleDefs(89)  =   "Named:id=50:FilterBar"
      _StyleDefs(90)  =   ":id=50,.parent=37"
   End
   Begin TrueDBGrid80.TDBGrid gridA 
      Height          =   4815
      Left            =   45
      TabIndex        =   1
      Top             =   4680
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   8493
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tipo de Ensayo"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tAnalisis"
      Columns(0).DropDown.vt=   8
      Columns(0).ExternalEditor=   "TDBDate1"
      Columns(0).ExternalEditor.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Norma Asociada"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Probetas"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "tProbetas"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ID_ENSAYO"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "General Number"
      Columns(3).DropDown=   "tEstados"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=10769"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=10689"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
      Splits(0)._ColumnProps(8)=   "Column(0).AutoCompletion=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=8070"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=7990"
      Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=8193"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=1799"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=1720"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(2).AutoDropDown=1"
      Splits(0)._ColumnProps(22)=   "Column(2).DropDownList=1"
      Splits(0)._ColumnProps(23)=   "Column(2).AutoCompletion=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=4260"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=4180"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(29)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
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
      Caption         =   "II. Ensayos"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=0,.bold=0,.fontsize=825"
      _StyleDefs(37)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=32,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=36,.parent=11,.alignment=2,.fgcolor=&HFF&,.bold=-1"
      _StyleDefs(49)  =   ":id=36,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(50)  =   ":id=36,.fontname=MS Sans Serif"
      _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=33,.parent=12"
      _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=34,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=35,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=54,.parent=11,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
      _StyleDefs(58)  =   "Named:id=37:Normal"
      _StyleDefs(59)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(60)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(61)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(62)  =   "Named:id=38:Heading"
      _StyleDefs(63)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(65)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(66)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(67)  =   "Named:id=39:Footing"
      _StyleDefs(68)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=40:Selected"
      _StyleDefs(70)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(71)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(72)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(73)  =   "Named:id=41:Caption"
      _StyleDefs(74)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(75)  =   "Named:id=42:HighlightRow"
      _StyleDefs(76)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(77)  =   "Named:id=43:EvenRow"
      _StyleDefs(78)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=44:OddRow"
      _StyleDefs(80)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(81)  =   "Named:id=47:RecordSelector"
      _StyleDefs(82)  =   ":id=47,.parent=38"
      _StyleDefs(83)  =   "Named:id=50:FilterBar"
      _StyleDefs(84)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha de control de eficacia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   0
      Width           =   15060
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15390
   End
End
Attribute VB_Name = "frmCE_Ficha_Bano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Dim xP As New XArrayDB
Dim xA As New XArrayDB
Dim xM As New XArrayDB
Dim xAnalisis As New XArrayDB
Dim xProbetas As New XArrayDB
Dim xMateriales As New XArrayDB
Dim xProducto As New XArrayDB

Const filasM As Integer = 15
Private Enum ColsM
    area = 0
    MATERIAL = 1
    criterio = 2
End Enum
Const filasP As Integer = 100
Const ColP As Integer = 5
Private Enum ColsP
    DESIGNACION = 0
    MATERIAL = 1
    TT = 2
    DIMENSION = 3
    nProbetas = 4
    AREAS = 5
End Enum
Const filasA As Integer = 100
Const ColA As Integer = 3
Private Enum ColsA
    tipo_ensayo = 0
    NORMA = 1
    DESIGNACION = 2
    ID_TIPO_ENSAYO = 3
End Enum
Const ColM As Integer = 2

Private Sub cmdCerrarCOC_Click()
    frmAreas.visible = False
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_CE_FICHA_BANO
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Ficha Proceso " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = 0 To ColP
        gridP.SelBookmarks.Add gridP.Bookmark
        For j = 0 To filasA
            If Not IsEmpty(xA(j, ColsA.DESIGNACION)) Then
               If Trim(xA(j, ColsA.DESIGNACION)) = Trim(xP(gridP.Bookmark, ColsP.DESIGNACION)) Then
                    For k = 0 To ColA
                        xA(j, k) = ""
                    Next
                End If
            End If
        Next
        
        xP(gridP.Bookmark, i) = ""
        gridP.SelBookmarks.Remove 0
    Next
    gridP.Refresh
    gridP.SetFocus
End Sub

Private Sub cmdborrarensayo_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To ColA
        gridA.SelBookmarks.Add gridA.Bookmark
        xA(gridA.Bookmark, i) = ""
        gridA.SelBookmarks.Remove 0
    Next
    gridA.Refresh
    gridA.SetFocus
End Sub

Private Sub cmdCargar_Click()
    If cmbBano.getPK_SALIDA <> 0 Then
         cargar_ficha cmbBano.getPK_SALIDA
    End If
End Sub

Private Sub cmdNuevoMaterial_Click()
    frmCE_Materiales.PK = 0
    frmCE_Materiales.Show 1
    cargar_combo_materiales
End Sub

Private Sub cmdok_Click()
    Dim oCE_banos_probetas As New clsCe_banos_probetas
    Dim i As Integer
   On Error GoTo cmdok_Click_Error
   
    frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación/creación de la ficha del baño."
    frmMotivo.Show 1
    If Trim(MOTIVO) = "" Then
        MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
        Exit Sub
    End If
   
    oCE_banos_probetas.Eliminar PK
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
         If Trim(xP(i, ColsP.DESIGNACION)) <> "" Then
            With oCE_banos_probetas
                .setBANO_ID = PK
                .setORDEN = i
                .setDESIGNACION = xP(i, ColsP.DESIGNACION)
                .setMATERIAL = xP(i, ColsP.MATERIAL)
                .setTT = xP(i, ColsP.TT)
                .setDIMENSION = xP(i, ColsP.DIMENSION)
                If IsNumeric(xP(i, ColsP.nProbetas)) Then
                    .setCANTIDAD = xP(i, ColsP.nProbetas)
                Else
                    .setCANTIDAD = 0
                End If
                If IsNumeric(xP(i, ColsP.AREAS)) Then
                    .setAREAS = xP(i, ColsP.AREAS)
                Else
                    .setAREAS = 1
                End If
                .Insertar
            End With
          End If
        End If
    Next
    Dim oce_banos_ensayos As New clsCe_banos_ensayos
    oce_banos_ensayos.Eliminar PK
    For i = 0 To filasA
        If Not IsEmpty(xA(i, ColsA.ID_TIPO_ENSAYO)) Then
         If Trim(xA(i, ColsA.ID_TIPO_ENSAYO)) <> "" Then
            With oce_banos_ensayos
                .setBANO_ID = PK
                .setTIPO_ENSAYO_ID = xA(i, ColsA.ID_TIPO_ENSAYO)
                .setDESIGNACION = xA(i, ColsA.DESIGNACION)
                .setORDEN = i
                .Insertar
            End With
         End If
        End If
    Next
    Dim ohc As New clsHistorial_cambios
    With ohc
        .setTIPO = HC_TIPOS.HC_CE_FICHA_BANO
        .setIDENTIFICADOR = PK
        .setIDENTIFICADOR_TEXTO = "Proceso : " & txtDatos(0) & "/Baño : " & txtDatos(1)
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
        .setMOTIVO = Trim(MOTIVO)
        .Insertar
    End With
    Set ohc = Nothing
'    Dim oCE_Ficha As New clsCe_ficha
    
'    Set oCE_Ficha = Nothing
    MsgBox "Ficha almacenada correctamente.", vbInformation, App.Title
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Ficha_Bano"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    If PK <> 0 Then
        cargar_ficha PK
    End If
    cargar_combo_probetas
    cargar_combo_materiales
    cargar_combo_productos
    CARGAR_COMBO_BANOS
End Sub

Private Sub cargar_ficha(BANO As Long)
   On Error GoTo cargar_ficha_Error

    inicializar_grid
'    If BANO >= 10000 Then
'        Dim oPintura As New clsPinturas
'        If oPintura.Carga(BANO) Then
'            cargar_combo_tipos_ensayos oPintura.getFICHA_ID
'        End If
'    Else
        Dim oBANO As New clsBanos
        If oBANO.cargar_bano(BANO) Then
            cargar_combo_tipos_ensayos oBANO.getFICHA_ID
        End If
'    End If
    ' Cargar las probetas asociadas al baño
    Dim oCe_bano_probetas As New clsCe_banos_probetas
    Dim i As Integer
    i = 0
    Dim rs As ADODB.Recordset
    Set rs = oCe_bano_probetas.Listado(BANO)
    If rs.RecordCount > 0 Then
        Do
            xP(i, ColsP.DESIGNACION) = CStr(rs(0))
            xP(i, ColsP.MATERIAL) = CStr(rs(1))
            xP(i, ColsP.TT) = CStr(rs(2))
            xP(i, ColsP.DIMENSION) = CStr(rs(3))
            xP(i, ColsP.nProbetas) = CStr(rs(4))
            xP(i, ColsP.AREAS) = CStr(rs(5))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Dim oCe_bano_ensayos As New clsCe_banos_ensayos
    i = 0
    Set rs = oCe_bano_ensayos.Listado(BANO)
    If rs.RecordCount > 0 Then
        Do
            xA(i, ColsA.tipo_ensayo) = CStr(rs(0))
            xA(i, ColsA.NORMA) = CStr(rs(1))
            xA(i, ColsA.DESIGNACION) = CStr(rs(2))
            xA(i, ColsA.ID_TIPO_ENSAYO) = CStr(rs(3))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

cargar_ficha_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_ficha of Formulario frmCE_Ficha_Bano"
End Sub
Public Function validar() As Boolean
    validar = True
End Function
Private Sub inicializar_grid()
    ' Probetas
    xP.ReDim 0, filasP, 0, ColP
    xP.Clear
    Set gridP.Array = xP
    gridP.Refresh
    ' Ensayos
    xA.ReDim 0, filasA, 0, ColA
    xA.Clear
    Set gridA.Array = xA
    gridA.Refresh
    ' Materiales
    xM.ReDim 0, filasM, 0, ColM
    xM.Clear
    Set gridAreas.Array = xM
    gridAreas.Refresh
End Sub

Private Sub cargar_combo_tipos_ensayos(FICHA As Long)
    Dim rs As ADODB.Recordset
    ' Responsables
    Dim ote As New clsCe_ensayos
    Set rs = ote.Listado(FICHA)
    If rs.RecordCount > 0 Then
        xAnalisis.ReDim 1, rs.RecordCount, 1, 3
        Dim i As Integer
        i = 1
        Do
            xAnalisis(i, 1) = CStr(rs(0))
            xAnalisis(i, 2) = CStr(rs(6))
            xAnalisis(i, 3) = CStr(rs(2))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xAnalisis.ReDim 1, 1, 1, 3
    End If
    Set tAnalisis.Array = xAnalisis
    tAnalisis.Refresh
End Sub
Private Sub cargar_combo_probetas()
    tProbetas.Refresh
    Dim i As Integer
    Dim j As Integer
    Dim cont As Integer
    cont = 0
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
            cont = cont + 1
        End If
    Next
    If cont <> 0 Then
        xProbetas.ReDim 1, (cont + 1), 1, 1
        xProbetas.Clear
        xProbetas(1, 1) = "TODAS"
        j = 2
        For i = 0 To filasP
            If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
                xProbetas(j, 1) = CStr(xP(i, ColsP.DESIGNACION))
                j = j + 1
            End If
        Next
    Else
        xProbetas.ReDim 1, 1, 1, 1
    End If
    Set tProbetas.Array = xProbetas
    tProbetas.Refresh
End Sub

Private Sub gridA_Click()
'    cargar_areas_a_analizar
End Sub

Private Sub gridA_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        If Not IsEmpty(xA(gridA.Row, ColsA.ID_TIPO_ENSAYO)) Then
            frmCE_Tipo_Ensayo.PK = CLng(xA(gridA.Row, ColsA.ID_TIPO_ENSAYO))
            frmCE_Tipo_Ensayo.Show 1
        End If
    End If
End Sub

Private Sub gridP_KeyUp(KeyCode As Integer, Shift As Integer)
    cargar_areas
End Sub

Private Sub gridP_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    cargar_areas
    If Button And vbRightButton Then
        frmAreas.visible = True
    End If
End Sub
Private Sub tAnalisis_DropDownClose()
    gridA.Columns(ColsA.NORMA) = tAnalisis.Columns(1)
    gridA.Columns(ColsA.ID_TIPO_ENSAYO) = tAnalisis.Columns(2)
    gridA.Col = 2
    cargar_combo_productos
End Sub

Private Sub tmateriales_DropDownClose()
   On Error GoTo tmateriales_DropDownClose_Error

    gridAreas.SelBookmarks.Add gridAreas.Bookmark
    If xM(gridAreas.Bookmark, 0) = "" Then
        MsgBox "Haga doble click primero en el conjunto de probetas del que desea informar sus areas.", vbInformation, App.Title
    Else
        Dim oArea As New clsCe_banos_probetas_materiales
        With oArea
            .setBANO_ID = PK
            gridP.SelBookmarks.Add gridP.Bookmark
            .setDESIGNACION = xP(gridP.Bookmark, ColsP.DESIGNACION)
            .setAREA = CInt(xM(gridAreas.Bookmark, 0))
            .setMATERIAL_ID = tmateriales.Columns(1)
            .Insertar
            
            gridAreas.Columns(ColsM.criterio) = tmateriales.Columns(2)

            gridP.SelBookmarks.Clear
            gridAreas.Row = gridAreas.Row + 1
        End With
    End If
    gridAreas.SelBookmarks.Clear
   On Error GoTo 0
   Exit Sub

tmateriales_DropDownClose_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmateriales_DropDownClose of Formulario frmCE_Ficha_Bano"
End Sub

Private Sub tProbetas_DropDownClose()
    gridA.Col = 0
    gridA.Row = gridA.Row + 1
End Sub

Private Sub tProbetas_DropDownOpen()
    cargar_combo_probetas
End Sub

Private Sub cargar_combo_materiales()
    tmateriales.Refresh
    Dim rs As ADODB.Recordset
    Dim oCE_Mat As New clsCe_banos_materiales
    Set rs = oCE_Mat.Listado("", "")
    If rs.RecordCount > 0 Then
        xMateriales.ReDim 1, rs.RecordCount, 1, 3
        xMateriales.Clear
        Dim i As Integer
        i = 1
        Do
            xMateriales(i, 1) = CStr(rs("MATERIAL"))
            xMateriales(i, 2) = CStr(rs("ID_MATERIAL"))
            xMateriales(i, 3) = CStr(rs("CRITERIO"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xMateriales.ReDim 1, 1, 1, 3
    End If
    Set tmateriales.Array = xMateriales
    tmateriales.Refresh
End Sub
Private Sub cargar_areas()
    xM.Clear
    gridP.SelBookmarks.Add gridP.Bookmark
    gridAreas.Caption = "AREAS : " & xP(gridP.Bookmark, ColsP.DESIGNACION)
    If IsNumeric(xP(gridP.Bookmark, ColsP.AREAS)) Then
        Dim i As Integer
        Dim oArea As New clsCe_banos_probetas_materiales
        Dim oMaterial As New clsCe_banos_materiales
        For i = 0 To CInt(xP(gridP.Bookmark, ColsP.AREAS)) - 1
            xM(i, ColsM.area) = CStr(i + 1)
            If oArea.Carga(PK, xP(gridP.Bookmark, ColsP.DESIGNACION), i + 1) Then
                oMaterial.Carga oArea.getMATERIAL_ID
                xM(i, ColsM.MATERIAL) = oMaterial.getMATERIAL
                xM(i, ColsM.criterio) = oMaterial.getCRITERIO
            Else
                xM(i, ColsM.MATERIAL) = ""
                xM(i, ColsM.criterio) = ""
            End If
        Next
    End If
    gridP.SelBookmarks.Clear
    Set gridAreas.Array = xM
    gridAreas.Refresh
End Sub
Private Sub cabecera()
    Dim oBANO As New clsBanos
    If oBANO.cargar_bano(PK) Then
        Dim oProceso_base As New clsProceso_base
        oProceso_base.CARGAR oBANO.getID_PROCESO_BASE
        lbltitulo = oProceso_base.getNOMBRE
        Dim oSolucion As New clsSoluciones
        oSolucion.CARGAR oBANO.getID_SOLUCION
        txtDatos(0) = oSolucion.getNOMBRE
        txtDatos(1) = oBANO.getNOMBRE
        Dim oPeriodicidad As New clsTipos_Frecuencia
        oPeriodicidad.CARGAR oBANO.getTIPO_FRECUENCIA_ID
        txtDatos(2) = oPeriodicidad.getNOMBRE
    End If
End Sub

Private Sub CARGAR_COMBO_BANOS()
    Dim oBANO As New clsBanos
    Dim sPK As String
    Dim sCAMPO As String
    
   On Error GoTo CARGAR_COMBO_BANOS_Error

    Dim consulta As String
    If oBANO.cargar_bano(PK) Then
        consulta = "SELECT DISTINCT B.ID_BANO,CONCAT(B.NOMBRE, ' (',PB.NOMBRE,')') " & _
                   "  FROM BANOS B, PROCESOS_BASE PB " & _
                   " WHERE B.PROCESO_BASE_ID = PB.ID_PROCESO_BASE " & _
                   "   AND B.FICHA_ID =  " & oBANO.getFICHA_ID
    Else
        consulta = "SELECT DISTINCT B.ID_BANO,CONCAT(B.NOMBRE, ' (',PB.NOMBRE,')') " & _
                   "  FROM BANOS B, PROCESOS_BASE PB " & _
                   " WHERE B.PROCESO_BASE_ID = PB.ID_PROCESO_BASE " & _
                   "   AND B.FICHA_ID <> 0 "
    End If
    With cmbBano
        .setTABLA = "BANOS"
        .setDESCRIPCION = "Baños"
        .setPK = "B.ID_BANO"
        .setCAMPO = "B.NOMBRE"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = frmBANO_Detalle
    End With

   On Error GoTo 0
   Exit Sub

CARGAR_COMBO_BANOS_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CARGAR_COMBO_BANOS of Formulario frmCE_Ficha_Bano"
    
End Sub

Private Sub cargar_combo_productos()
    xProducto.Clear
    xProducto.ReDim 1, 1, 1, 1
    xProducto(1, 1) = " "
    Set tProducto.Array = xProducto
    tProducto.Refresh
    Dim i As Integer
    Dim join As String
    join = ""
    For i = 0 To filasA
        If Not IsEmpty(xA(i, ColsA.ID_TIPO_ENSAYO)) And xA(i, ColsA.ID_TIPO_ENSAYO) <> "" Then
            join = join & xA(i, ColsA.ID_TIPO_ENSAYO) & ","
        End If
    Next
    If join <> "" Then
        join = "(" & Left(join, Len(join) - 1) & ")"
    Else
        Exit Sub
    End If
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT DISTINCT B.TIPO_MUESTRA_ID " & _
               "  FROM CE_TIPOS_ENSAYOS A, TIPOS_ANALISIS B " & _
               " where A.ID_TIPO_ENSAYO IN " & join & _
               "   AND A.TIPO_ANALISIS_ID = B.ID_TIPO_ANALISIS"
    Set rs = datos_bd(consulta)
    join = ""
    If rs.RecordCount > 0 Then
        Do
            join = join & "'" & rs(0) & "',"
            rs.MoveNext
        Loop Until rs.EOF
    End If
    If join <> "" Then
        join = "(" & Left(join, Len(join) - 1) & ")"
    Else
        Exit Sub
    End If
'    tProducto.Refresh
    consulta = "SELECT DISTINCT TRIM(DESCRIPCION) " & _
               "  FROM decodificadora " & _
               " WHERE CODIGO = " & DECODIFICADORA.DESCRIPCION_PRODUCTO & _
               "   AND PARAMETROS in " & join
    Set rs = datos_bd(consulta)
    xProducto.Clear
    If rs.RecordCount > 0 Then
        xProducto.ReDim 1, rs.RecordCount, 1, 1
        i = 1
        Do
            xProducto(i, 1) = CStr(rs(0))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xProducto.ReDim 1, 1, 1, 1
    End If
    Set tProducto.Array = xProducto
    tProducto.Refresh
    gridA.Refresh
End Sub

