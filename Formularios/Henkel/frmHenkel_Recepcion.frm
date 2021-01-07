VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmHenkel_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Probetas HENKEL"
   ClientHeight    =   11535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11535
   ScaleWidth      =   14595
   Begin MSComDlg.CommonDialog cd 
      Left            =   7560
      Top             =   10710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDimension 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dimensiones"
      Height          =   930
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10575
      Width           =   1245
   End
   Begin VB.CommandButton cmdMaterial 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Materiales"
      Height          =   930
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10575
      Width           =   1245
   End
   Begin TrueDBGrid80.TDBDropDown tMaterial 
      Height          =   5430
      Left            =   7110
      TabIndex        =   25
      Top             =   4095
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   9578
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
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
   Begin Geslab.ControlPanelXP cpprobetas 
      Height          =   7755
      Left            =   45
      TabIndex        =   23
      Top             =   2745
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   13679
      Caption         =   "Probetas y Ensayos"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   7755
      Begin TrueDBGrid80.TDBDropDown tDimension 
         Height          =   5430
         Left            =   10440
         TabIndex        =   26
         Top             =   1350
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   9578
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=825"
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
      Begin TrueDBGrid80.TDBGrid gridP 
         Height          =   7185
         Left            =   90
         TabIndex        =   0
         Top             =   450
         Width           =   14325
         _ExtentX        =   25268
         _ExtentY        =   12674
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "F.PROCESADO"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "Short Date"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "CODIGO OP"
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "IDENTIFICACION"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "COC"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MATERIAL"
         Columns(4).DataField=   ""
         Columns(4).DropDown=   "tMaterial"
         Columns(4).DropDown.vt=   8
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DIMENSIONES"
         Columns(5).DataField=   ""
         Columns(5).DropDown=   "tDimension"
         Columns(5).DropDown.vt=   8
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "SET"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "LOTE"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "FACTURA"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "RAW"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   1
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0)._ColStyle=131585"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(0).AutoDropDown=1"
         Splits(0)._ColumnProps(9)=   "Column(0).DropDownList=1"
         Splits(0)._ColumnProps(10)=   "Column(0).AutoCompletion=1"
         Splits(0)._ColumnProps(11)=   "Column(0)._HeadDivider=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Width=2514"
         Splits(0)._ColumnProps(13)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(1)._WidthInPix=2434"
         Splits(0)._ColumnProps(15)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(1)._ColStyle=131585"
         Splits(0)._ColumnProps(17)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(18)=   "Column(2).Width=5768"
         Splits(0)._ColumnProps(19)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(2)._WidthInPix=5689"
         Splits(0)._ColumnProps(21)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(2)._ColStyle=131585"
         Splits(0)._ColumnProps(23)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(24)=   "Column(3).Width=1535"
         Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=1455"
         Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=131585"
         Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(30)=   "Column(3).AutoDropDown=1"
         Splits(0)._ColumnProps(31)=   "Column(3).DropDownList=1"
         Splits(0)._ColumnProps(32)=   "Column(3).AutoCompletion=1"
         Splits(0)._ColumnProps(33)=   "Column(4).Width=6006"
         Splits(0)._ColumnProps(34)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(4)._WidthInPix=5927"
         Splits(0)._ColumnProps(36)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(37)=   "Column(4)._ColStyle=131585"
         Splits(0)._ColumnProps(38)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(39)=   "Column(4).AutoDropDown=1"
         Splits(0)._ColumnProps(40)=   "Column(4).AutoCompletion=1"
         Splits(0)._ColumnProps(41)=   "Column(5).Width=6033"
         Splits(0)._ColumnProps(42)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(5)._WidthInPix=5953"
         Splits(0)._ColumnProps(44)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(45)=   "Column(5)._ColStyle=131585"
         Splits(0)._ColumnProps(46)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(47)=   "Column(5).AutoDropDown=1"
         Splits(0)._ColumnProps(48)=   "Column(5).AutoCompletion=1"
         Splits(0)._ColumnProps(49)=   "Column(6).Width=4313"
         Splits(0)._ColumnProps(50)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(6)._WidthInPix=4233"
         Splits(0)._ColumnProps(52)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(6)._ColStyle=131585"
         Splits(0)._ColumnProps(54)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(55)=   "Column(7).Width=4657"
         Splits(0)._ColumnProps(56)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(7)._WidthInPix=4577"
         Splits(0)._ColumnProps(58)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(7)._ColStyle=131585"
         Splits(0)._ColumnProps(60)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(61)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(62)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(64)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(8)._ColStyle=131585"
         Splits(0)._ColumnProps(66)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(67)=   "Column(9).Width=2249"
         Splits(0)._ColumnProps(68)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(9)._WidthInPix=2170"
         Splits(0)._ColumnProps(70)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(9)._ColStyle=131585"
         Splits(0)._ColumnProps(72)=   "Column(9).Order=10"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1,.bgcolor=&HFFFFFF&"
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
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=825"
         _StyleDefs(40)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(41)  =   ":id=23,.fontname=MS Sans Serif"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=11"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=12"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=62,.parent=11"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=12"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=36,.parent=11"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=11"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=12"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=28,.parent=11"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=12"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=15"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=11"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=12"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=15"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=66,.parent=11"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=12"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=70,.parent=11"
         _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=12"
         _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=15"
         _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=74,.parent=11"
         _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=12"
         _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=15"
         _StyleDefs(78)  =   "Named:id=37:Normal"
         _StyleDefs(79)  =   ":id=37,.parent=0,.alignment=2"
         _StyleDefs(80)  =   "Named:id=38:Heading"
         _StyleDefs(81)  =   ":id=38,.parent=37,.alignment=2,.valignment=2,.bgcolor=&H80000004&"
         _StyleDefs(82)  =   ":id=38,.fgcolor=&H80000012&,.wraptext=-1,.appearance=0,.ellipsis=0"
         _StyleDefs(83)  =   "Named:id=39:Footing"
         _StyleDefs(84)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(85)  =   "Named:id=40:Selected"
         _StyleDefs(86)  =   ":id=40,.parent=37,.alignment=2,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0"
         _StyleDefs(87)  =   ":id=40,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
   End
   Begin Geslab.ControlPanelXP cpDatos 
      Height          =   2145
      Left            =   45
      TabIndex        =   15
      Top             =   585
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   3784
      Caption         =   "Datos de recepción"
      BackColor       =   12632256
      HeaderColor     =   8421504
      Object.Height          =   2145
      Begin MSDataListLib.DataCombo cmbCentro 
         Height          =   315
         Left            =   12060
         TabIndex        =   5
         Top             =   1215
         Width           =   2280
         _ExtentX        =   4022
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
      Begin MSDataListLib.DataCombo cmbPedido 
         Height          =   315
         Left            =   1170
         TabIndex        =   3
         Top             =   1215
         Width           =   6855
         _ExtentX        =   12091
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
      Begin pryCombo.miCombo cmbbanos 
         Height          =   330
         Left            =   1170
         TabIndex        =   2
         Top             =   855
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbenvases 
         Height          =   315
         Left            =   10260
         TabIndex        =   8
         Top             =   1710
         Width           =   4080
         _ExtentX        =   7197
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
      Begin MSDataListLib.DataCombo cmbrealizada 
         Height          =   315
         Left            =   5670
         TabIndex        =   7
         Top             =   1710
         Width           =   3330
         _ExtentX        =   5874
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
      Begin MSDataListLib.DataCombo cmbentregada 
         Height          =   315
         Left            =   1170
         TabIndex        =   6
         Top             =   1710
         Width           =   3240
         _ExtentX        =   5715
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   12060
         TabIndex        =   4
         Top             =   495
         Width           =   1380
         _ExtentX        =   2434
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
         Format          =   60162049
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1170
         TabIndex        =   1
         Top             =   495
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   582
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   8100
         Picture         =   "frmHenkel_Recepcion.frx":0000
         Stretch         =   -1  'True
         Top             =   1215
         Width           =   255
      End
      Begin VB.Image imgPedidos 
         Height          =   300
         Left            =   8430
         Picture         =   "frmHenkel_Recepcion.frx":08CA
         Stretch         =   -1  'True
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   10935
         TabIndex        =   24
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   1260
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Recepción"
         Height          =   195
         Index           =   6
         Left            =   10935
         TabIndex        =   21
         Top             =   540
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   20
         Top             =   900
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   19
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizada por"
         Height          =   195
         Index           =   7
         Left            =   4545
         TabIndex        =   18
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   17
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   14400
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         Height          =   195
         Index           =   5
         Left            =   9675
         TabIndex        =   16
         Top             =   1755
         Width           =   540
      End
   End
   Begin VB.TextBox txtsolucion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4095
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   14
      Top             =   315
      Width           =   10065
   End
   Begin VB.TextBox txtproceso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5175
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   13
      Top             =   45
      Width           =   7140
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Probeta"
      Height          =   930
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10575
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10545
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   13470
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10545
      Width           =   1050
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cargar Excel"
      Height          =   930
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10575
      Width           =   1245
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Probetas HENKEL"
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
      TabIndex        =   12
      Top             =   120
      Width           =   3420
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   14545
   End
End
Attribute VB_Name = "frmHenkel_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xP As New XArrayDB
Dim xMaterial As New XArrayDB
Dim xDimension As New XArrayDB
Const filasP As Integer = 1000
Const ColP As Integer = 10
Private Enum ColsP
    F_PROCESADO = 0
    CODIGO_OP = 1
    DESIGNACION = 2
    COC = 3
    MATERIAL = 4
    DIMENSION = 5
    SET_D = 6
    LOTE = 7
    factura = 8
    RAW = 9
End Enum

Private Sub cmdBorrar_Click()
    Dim i As Integer
    Dim j As Integer
    For i = gridP.Bookmark To filasP - 1
        For j = 0 To ColP
            xP(i, j) = xP(i + 1, j)
        Next
    Next
    gridP.Refresh
End Sub

Private Sub cmbBanos_Change()
    If cmbbanos.getPK_SALIDA <> 0 Then
        Dim oBANO As New clsBanos
        oBANO.cargar_bano cmbbanos.getPK_SALIDA
        If oBANO.getFICHA_ID = 0 Then
            txtproceso = ""
            If MsgBox("El baño no tiene ficha asignada. ¿Desea crearla?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                frmBANO_Detalle.PK = cmbbanos.getPK_SALIDA
                frmBANO_Detalle.Show 1
            Else
                Exit Sub
            End If
        Else
            Dim oSolucion As New clsSoluciones
            oSolucion.CARGAR oBANO.getID_SOLUCION
            txtsolucion = oSolucion.getNOMBRE
            cargar_ficha cmbbanos.getPK_SALIDA, oBANO.getFICHA_ID
        End If
    End If
End Sub
Private Sub cmbClientes_change()
    cargar_banos
    cmdLimpiar_Click
End Sub
Private Sub cmdLimpiar_Click()
    cmbPedido.Text = ""
    cmbPedido.BoundText = ""
End Sub

Private Sub cmdDimension_Click()
    Dim oform As New frmDecodificadoraModal
    oform.CODIGO = DECODIFICADORA.DECODIFICADORA_DIMENSIONES
    oform.Show 1
    cargarCombosDimensiones
    Set oform = Nothing
End Sub

Private Sub cmdExcel_Click()
   On Error GoTo cmdExcel_Click_Error
    cd.DialogTitle = "Abrir fichero Excel de probetas"
    cd.ShowOpen
    If cd.FileName <> "" Then
        Dim fichero As String
        fichero = cd.FileName
        ' Cargar Excel
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Open(fichero)
        Set XLS = XLW.Worksheets(1)
        ' Cargar parametros
        Dim op As New clsParametros
        op.Carga parametros.HENKEL_HOJA_EXCEL, ""
        Dim l() As String
        l = Split(op.getVALOR, ";")
        'LINEA;FPROCESADO;OP;IDENTIFICACION;COC;MATERIAL;DIMENSION
        ' Leer contenido
        Dim linea As Integer
        Dim fila As Integer
        linea = l(0)
        fila = 0
        While Trim(XLS.Cells(linea, 1)) <> ""
            xP(fila, ColsP.F_PROCESADO) = CStr(XLS.Cells(linea, CInt(l(1))))
            xP(fila, ColsP.CODIGO_OP) = CStr(XLS.Cells(linea, CInt(l(2))))
            xP(fila, ColsP.DESIGNACION) = CStr(XLS.Cells(linea, CInt(l(3))))
            xP(fila, ColsP.COC) = CStr(XLS.Cells(linea, CInt(l(4))))
            xP(fila, ColsP.MATERIAL) = CStr(XLS.Cells(linea, CInt(l(5))))
            xP(fila, ColsP.DIMENSION) = CStr(XLS.Cells(linea, CInt(l(6))))
            xP(fila, ColsP.SET_D) = CStr(XLS.Cells(linea, CInt(l(7))))
            xP(fila, ColsP.LOTE) = CStr(XLS.Cells(linea, CInt(l(8))))
            xP(fila, ColsP.factura) = CStr(XLS.Cells(linea, CInt(l(9))))
            xP(fila, ColsP.RAW) = CStr(XLS.Cells(linea, CInt(l(10))))
            linea = linea + 1
            fila = fila + 1
        Wend
        XLA.Quit
        Set XLS = Nothing
        Set XLW = Nothing
        Set XLA = Nothing
        MsgBox "Carga finalizada correctamente.", vbOKOnly + vbInformation, App.Title
        gridP.Refresh
        gridP.Col = 0
        gridP.Row = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:
    Set XLS = Nothing
    Set XLW = Nothing
    Set XLA = Nothing

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmHenkel_Recepcion"

End Sub
Private Function insertarMuestra(numeroRecepcion As Long, listaProbetas As Collection) As Long
   On Error GoTo insertarMuestra_Error
    ' INSERTAR MUESTRAS
    Dim op As New clsParametros
    op.Carga parametros.HENKEL_ALTA_MUESTRA, ""
    'TIPO_MUESTRA;TIPO_ANALISIS;TIPO_ENSAYO
    Dim l() As String
    l = Split(op.getVALOR, ";")
    ' Recuperar la primera probeta
    Dim oCeResultados As clsCe_resultados
    Set oCeResultados = listaProbetas.Item(1)
    
    Dim muestra As Long
    Dim oMuestra As New clsMuestra
    With oMuestra
        .setTIPO_MUESTRA_ID = l(0)
        .setTIPO_ANALISIS_ID = l(1)
        .setANALISIS_MODIFICADO = 2 ' Para identificar que es un CE
        .setFECHA_MUESTREO = Format(fecha.Value, "yyyy-mm-dd")
        .setENTIDAD_MUESTREO_ID = cmbrealizada.BoundText
        .setDETALLE_MUESTREO = ""
        .setOBSERVACIONES_MUESTREO = ""
        .setFECHA_RECEPCION = Format(fecha.Value, "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFORMATO_ID = cmbenvases.BoundText
        .setENTIDAD_ENTREGA_ID = cmbentregada.BoundText
        .setDETALLE_ENTREGA = ""
        .setOBSERVACIONES_ENTREGA = ""
        .setCLIENTE_ID = cmbClientes.getPK_SALIDA
        .setCENTRO_ID = cmbCentro.BoundText
        .setREFERENCIA_CLIENTE = "PAA CoC-" & oCeResultados.getCOC & " OP-" & oCeResultados.getOP
        .setFECHA_PREV_FIN = Format(fecha, "yyyy-mm-dd")
        .setOBSERVACIONES = ""
        .setANULADA = 0
        .setPRECINTO = ""
        .setBANO_ID = cmbbanos.getPK_SALIDA
        .setFECHA_COMIENZO = "0000-00-00"
        .setFECHA_FINALIZACION = "000-00-00"
        .setFECHA_CIERRE = "0000-00-00"
        .setCERRADA = 0
        .setDOCUMENTO_PAGO = 0
        .setULT_EDICION_IMP = 0
        .setPRECIO = moneda_bd("0")
        .setANALISIS_DUPLICADO = 0
        .setPRODUCTO = ""
        If cmbPedido.Text <> "" Then
            .setPEDIDO_ID = cmbPedido.BoundText
        End If
        .setREPLACEMENT_ID = 0
        muestra = .guardarMuestra
        .informar_precio_muestra muestra
    End With
    'CE_RECEPCION
    insertarCe_recepcion muestra, numeroRecepcion, CLng(l(2)), oCeResultados.getCOC, oCeResultados.getFPROCESADO, listaProbetas.Count
    'PROBETAS
    Dim i As Integer
    For i = 1 To listaProbetas.Count
        Set oCeResultados = listaProbetas.Item(i)
        With oCeResultados
            .setMUESTRA_ID = muestra
            .setPROBETA = i
            .setDESIGNACION = "COC " & .getCOC
            .setAREA = 0
'            .setIDENTIFICACION_CANAGROSA = .getIDENTIFICACION_CLIENTE
            .Insertar
        End With
    Next
    insertarMuestra = muestra

   On Error GoTo 0
   Exit Function

insertarMuestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertarMuestra of Formulario frmHenkel_Recepcion"
End Function
Private Function insertarCe_recepcion(idMuestra As Long, numeroRecepcion As Long, tipoEnsayo As Long, COC As String, fechaProcesado As String, CANTIDAD As Integer) As Long
    ' Recepción del control de eficacia
    Dim oce_recepcion As New clsCe_recepcion
   On Error GoTo insertarCe_recepcion_Error

    With oce_recepcion
        .setNUMERO_RECEPCION = numeroRecepcion
        .setMUESTRA_ID = idMuestra
        .setTIPO_ENSAYO_ID = tipoEnsayo
        .setCANTIDAD = CANTIDAD
        .setDESIGNACION = "COC " & COC
        If fechaProcesado = "" Then
            .setFECHA_PROCESADO_PIEZAS = "null"
        Else
            .setFECHA_PROCESADO_PIEZAS = "'" & Format(fechaProcesado, "yyyy-mm-dd") & "'"
        End If
        .setESPESOR = "No requiere espesor."
        .setLOTE_PROBETA_ID = 0
        .setIDENTIFICACION_LABORATORIO = 0
        .setCONDICIONES_AMBIENTALES = ""
        .setMATERIAL = ""
        .setDIMENSION = "0"
        .setMAQUINA = ""
        .setREACTIVOS = ""
        .setREACTIVOS_PROPIOS = ""
       insertarCe_recepcion = .Insertar
    End With
    Set oce_recepcion = Nothing

   On Error GoTo 0
   Exit Function

insertarCe_recepcion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertarCe_recepcion of Formulario frmHenkel_Recepcion"
End Function

Private Sub cmdMaterial_Click()
    Dim oform As New frmDecodificadoraModal
    oform.CODIGO = DECODIFICADORA.DECODIFICADORA_MATERIALES
    oform.Show 1
    cargarCombosMateriales
    Set oform = Nothing
End Sub
Private Sub guardarDecodificadoras()
    On Error Resume Next
    Dim i As Integer
    Dim oDeco As New clsDecodificadora
    Dim lista As New Collection
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.MATERIAL)) Then
            If Trim(xP(i, ColsP.MATERIAL)) <> "" Then
                lista.Add Trim(xP(i, ColsP.MATERIAL)), Trim(xP(i, ColsP.MATERIAL))
            End If
        End If
    Next
    For i = 1 To lista.Count
        With oDeco
            .setCODIGO = DECODIFICADORA.DECODIFICADORA_MATERIALES
            .setIDIOMA = "ES"
            .setVALOR = 0
            .setDESCRIPCION = lista.Item(i)
            .Insertar
        End With
    Next
    cargarCombosMateriales
    Set lista = Nothing
    Set lista = New Collection
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.DIMENSION)) Then
            If Trim(xP(i, ColsP.DIMENSION)) <> "" Then
                lista.Add Trim(xP(i, ColsP.DIMENSION)), Trim(xP(i, ColsP.DIMENSION))
            End If
        End If
    Next
    For i = 1 To lista.Count
        With oDeco
            .setCODIGO = DECODIFICADORA.DECODIFICADORA_DIMENSIONES
            .setIDIOMA = "ES"
            .setVALOR = 0
            .setDESCRIPCION = lista.Item(i)
            .Insertar
        End With
    Next
    cargarCombosDimensiones
End Sub
Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    gridP.Col = 0
    gridP.Row = 0

    If validar = True Then
        ' Contar lo que se va a registrar
        Dim i As Integer
        Dim cocAux As String
        Dim opAux As String
        Dim idenAux As String
        Dim salida As String
        Dim cont As Integer
        For i = 0 To filasP
            If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
                If Trim(xP(i, ColsP.DESIGNACION)) <> "" Then
'                    If opAux <> Trim(xP(i, ColsP.CODIGO_OP)) And opAux <> "" Then
                    If opAux <> Trim(xP(i, ColsP.COC)) And opAux <> "" Then
                        salida = salida & "COC : " & opAux & ", PROBETAS : " & idenAux & vbNewLine
                        idenAux = ""
                        cont = cont + 1
                    End If
                    If idenAux <> "" Then
                        idenAux = idenAux & ","
                    End If
                    idenAux = idenAux & Trim(xP(i, ColsP.DESIGNACION))
'                    opAux = Trim(xP(i, ColsP.CODIGO_OP))
                    opAux = Trim(xP(i, ColsP.COC))
                End If
            End If
        Next
        cont = cont + 1
        salida = salida & "COC : " & opAux & ", PROBETAS : " & idenAux & vbNewLine
        salida = "Se van a registrar " & cont & " muestras : " & vbNewLine & vbNewLine & salida & vbNewLine & "¿Desea continuar?"
        If MsgBox(salida, vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
        
        Me.MousePointer = 11
        Dim oce_recepcion As New clsCe_recepcion
        Dim RECEPCION As Long
        oce_recepcion.Calcular_Numero_Recepcion
        RECEPCION = oce_recepcion.getNUMERO_RECEPCION
        ' Generamos el registro de las muestras
        Dim idMuestra As Long
        Dim listaProbetas As New Collection
        opAux = ""
        For i = 0 To filasP
            If Not IsEmpty(xP(i, ColsP.COC)) Then
                If Trim(xP(i, ColsP.COC)) <> "" Then
                            
'                    If opAux <> Trim(xP(i, ColsP.CODIGO_OP)) And opAux <> "" Then
                    If opAux <> Trim(xP(i, ColsP.COC)) And opAux <> "" Then
                        idMuestra = insertarMuestra(RECEPCION, listaProbetas)
                        Set listaProbetas = Nothing
                        Set listaProbetas = New Collection
                        opAux = ""
                    End If
                    Dim oCe_resultados As New clsCe_resultados
                    With oCe_resultados
                        .setFPROCESADO = ""
                        If CStr(xP(i, ColsP.F_PROCESADO)) <> "" Then
                            If IsDate(CStr(xP(i, ColsP.F_PROCESADO))) Then
                                .setFPROCESADO = Format(CStr(xP(i, ColsP.F_PROCESADO)), "yyyy-mm-dd")
                            End If
                        End If
                        .setOP = CStr(xP(i, ColsP.CODIGO_OP))
                        .setCOC = CStr(xP(i, ColsP.COC))
'                        .setIDENTIFICACION_CLIENTE = CStr(xP(i, ColsP.DESIGNACION))
                        .setIDENTIFICACION_CANAGROSA = CStr(xP(i, ColsP.DESIGNACION))
                        .setIDENTIFICACION_CLIENTE = CStr(xP(i, ColsP.LOTE))
                        .setMATERIAL = Trim(CStr(xP(i, ColsP.MATERIAL)))
                        .setDIMENSION = Trim(CStr(xP(i, ColsP.DIMENSION)))
                        .setSETD = Trim(CStr(xP(i, ColsP.SET_D)))
                        .setLOTE = Trim(CStr(xP(i, ColsP.LOTE)))
                        .setFACTURA = Trim(CStr(xP(i, ColsP.factura)))
                        .setRAW = Trim(CStr(xP(i, ColsP.RAW)))
                        .setAREA = 0
                    End With
                    listaProbetas.Add oCe_resultados
                    Set oCe_resultados = Nothing
                    
'                    opAux = Trim(xP(i, ColsP.CODIGO_OP))
                    opAux = Trim(xP(i, ColsP.COC))
                End If
            End If
        Next
        idMuestra = insertarMuestra(RECEPCION, listaProbetas)
        Me.MousePointer = 0
        guardarDecodificadoras
        MsgBox "La recepción se ha realizado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmHenkel_Recepcion")
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cpDatos_Expand(State As Boolean)
    gridP.Refresh
End Sub

Private Sub Form_Initialize()
    Me.SetFocus
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Me.Left = 50
    Me.top = 50
    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    Call cargar_combos
    cargarCombosDimensiones
    cargarCombosMateriales
    ' Datos inicio
    fecha = Date
    cmbClientes.MostrarElemento 3641
    cmbbanos.MostrarElemento 3412
    cmbentregada.BoundText = 2
    cmbrealizada.BoundText = 3
    cmbenvases.BoundText = 2
    cmbCentro.BoundText = 1
    
    gridP.Col = 0
    gridP.Row = 0
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmHenkel_Recepcion"
End Sub
Public Function validar() As Boolean
    On Error GoTo validar_Error

    validar = True
    If cmbbanos.getTEXTO = "" Then
        MsgBox "Debe asignar un baño a la selección.", vbExclamation, App.Title
        cmbbanos.SetFocus
        validar = False
        Exit Function
    End If
    If cmbCentro.Text = "" Then
        MsgBox "El CENTRO no puede estar en blanco.", vbExclamation, "Validación"
        cmbCentro.SetFocus
        validar = False
        Exit Function
    End If
    If cmbentregada.BoundText = "" Then
        MsgBox "Debe indicar quien entrega el control de eficacia.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbrealizada.BoundText = "" Then
        MsgBox "Debe indicar quien realiza el control de eficacia.", vbExclamation, App.Title
        validar = False
        cmbrealizada.SetFocus
        Exit Function
    End If
    If cmbenvases.BoundText = "" Then
        MsgBox "Debe indicar en envase.", vbExclamation, App.Title
        validar = False
        cmbenvases.SetFocus
        Exit Function
    End If
    ' Verificar que existe algún dato en la lista
    Dim i As Integer
    Dim algo As Boolean
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
            algo = True
        End If
    Next
    If Not algo Then
        validar = False
        MsgBox "Debe insertar al menos alguna probeta.", vbCritical, App.Title
        Exit Function
    End If
   On Error GoTo 0
   Exit Function

validar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validar of Formulario frmHenkel_Recepcion"
End Function
Private Sub cargar_combos()
    cargar_clientes
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbbanos, New clsBanos, 0, frmBANO_Detalle, " ANULADO = 0 "
    cargar_combo cmbenvases, New clsformatos
    cargar_combo cmbentregada, New clsEntidades_Entrega
    cargar_combo cmbrealizada, New clsEntidades_muestreo
End Sub
Private Sub cargar_ficha(BANO As Long, FICHA As Long)
   On Error GoTo cargar_ficha_Error

    inicializar_grid
    Dim oCe_Ficha As New clsCe_ficha
    If oCe_Ficha.Carga(FICHA) Then
        txtproceso = oCe_Ficha.getPROCESO
    End If

   On Error GoTo 0
   Exit Sub

cargar_ficha_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_ficha of Formulario frmHenkel_Recepcion"
End Sub

Public Sub cargar_clientes()
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
               "  FROM CLIENTES C, BANOS B " & _
               " WHERE B.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND C.ANULADO = 0 " & _
               "   AND B.FICHA_ID <> 0 "
    With cmbClientes
        .setTABLA = "CLIENTES"
        .setDESCRIPCION = "Clientes"
        .setPK = "C.ID_CLIENTE"
        .setCAMPO = "C.NOMBRE"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = frmClientes
    End With
End Sub
Private Sub cargar_banos()
    If cmbClientes.getPK_SALIDA <> 0 Then
        Dim consulta As String
        cmbbanos.limpiar
        consulta = "SELECT ID_BANO, NOMBRE FROM BANOS " & _
                   " WHERE CLIENTE_ID = " & cmbClientes.getPK_SALIDA & _
                   "   AND FICHA_ID <> 0 " & _
                   "   AND ANULADO = 0 "
        With cmbbanos
            .setTABLA = "BANOS"
            .setDESCRIPCION = "Baños"
            .setPK = "ID_BANO"
            .setCAMPO = "NOMBRE"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmBANO_Detalle
        End With
        cargar_pedidos CLng(cmbClientes.getPK_SALIDA), fecha.Value
    End If
End Sub
Private Sub Image1_Click()
    cmbPedido.Text = ""
    cmbPedido.BoundText = ""
End Sub

Private Sub imgPedidos_Click()
    If cmbClientes.getTEXTO <> "" Then
        frmClientes_Pedidos.PK = cmbClientes.getPK_SALIDA
        frmClientes_Pedidos.Show 1
        cargar_pedidos CLng(cmbClientes.getPK_SALIDA), fecha.Value
    End If
End Sub

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmHenkel_Recepcion_Nuevo2"
End Sub
Private Sub cargarCombosDimensiones()
    Dim rs As ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set rs = oDeco.Listado_por_Codigo(DECODIFICADORA.DECODIFICADORA_DIMENSIONES)
    xDimension.Clear
    If rs.RecordCount > 0 Then
        xDimension.ReDim 1, rs.RecordCount, 1, 1
        Dim i As Integer
        i = 1
        Do
            xDimension(i, 1) = CStr(rs("DESCRIPCION"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xDimension.ReDim 1, 1, 1, 1
    End If
    Set tDimension.Array = xDimension
    tDimension.Refresh
End Sub
Private Sub cargarCombosMateriales()
    Dim rs As ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set rs = oDeco.Listado_por_Codigo(DECODIFICADORA.DECODIFICADORA_MATERIALES)
    xMaterial.Clear
    If rs.RecordCount > 0 Then
        xMaterial.ReDim 1, rs.RecordCount, 1, 1
        Dim i As Integer
        i = 1
        Do
            xMaterial(i, 1) = CStr(rs("DESCRIPCION"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xMaterial.ReDim 1, 1, 1, 1
    End If
    Set tMaterial.Array = xMaterial
    tMaterial.Refresh
End Sub
Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim oPedido As New clsClientes_pedidos
    Set cmbPedido.RowSource = oPedido.Listado_en_fecha(CInt(cliente), CStr(fecha))
    cmbPedido.ListField = "CODIGO_LARGO"
    cmbPedido.DataField = "ID_PEDIDO"
    cmbPedido.BoundColumn = "ID_PEDIDO"
End Sub
