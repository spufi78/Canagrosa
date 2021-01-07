VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#34.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmCE_Recepcion_Nuevo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Control de Eficacia"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCE_Recepcion_Nuevo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   12885
   Begin TrueDBGrid80.TDBDropDown tAnalisis 
      Height          =   1845
      Left            =   45
      TabIndex        =   33
      Top             =   6120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3254
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=9472"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=9393"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=265"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=185"
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
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Probeta"
      Height          =   930
      Left            =   45
      Picture         =   "frmCE_Recepcion_Nuevo.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7830
      Width           =   1245
   End
   Begin VB.CommandButton cmdborrarensayo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar Ensayo"
      Height          =   930
      Left            =   1350
      Picture         =   "frmCE_Recepcion_Nuevo.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7830
      Width           =   1245
   End
   Begin TrueDBGrid80.TDBDropDown tProbetas 
      Height          =   1605
      Left            =   7425
      TabIndex        =   24
      Top             =   6120
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   2831
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
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   10725
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7800
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   11805
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Comúnes del Control de Eficacia"
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
      Height          =   2715
      Left            =   45
      TabIndex        =   13
      Top             =   540
      Width           =   12795
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   27
         Text            =   "Realizar análisis"
         Top             =   2295
         Width           =   3810
      End
      Begin VB.CheckBox chkSinEspecificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3375
         TabIndex        =   9
         Top             =   1890
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   10665
         MaxLength       =   255
         TabIndex        =   5
         Top             =   945
         Width           =   1965
      End
      Begin VB.CheckBox chkRutinario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rutinario"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9540
         TabIndex        =   4
         Top             =   675
         Visible         =   0   'False
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cmbproceso 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   990
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Left            =   11025
         TabIndex        =   3
         Top             =   270
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   58654721
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbbanos 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   630
         Width           =   7905
         _ExtentX        =   13944
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
      Begin MSDataListLib.DataCombo cmbClientes 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   270
         Width           =   7905
         _ExtentX        =   13944
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
         Left            =   1980
         TabIndex        =   6
         Top             =   1485
         Width           =   3825
         _ExtentX        =   6747
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
         Left            =   7605
         TabIndex        =   7
         Top             =   1485
         Width           =   4365
         _ExtentX        =   7699
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
      Begin MSComCtl2.DTPicker fprocesado 
         Height          =   330
         Left            =   1980
         TabIndex        =   8
         Top             =   1890
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   58654721
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbenvases 
         Height          =   315
         Left            =   7605
         TabIndex        =   10
         Top             =   1890
         Width           =   4350
         _ExtentX        =   7673
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
      Begin pryCombo.miCombo cmbLote 
         Height          =   330
         Left            =   7605
         TabIndex        =   26
         Top             =   2295
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos de Espesor"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   29
         Top             =   2295
         Width           =   1260
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote Probetas"
         Height          =   195
         Index           =   18
         Left            =   6480
         TabIndex        =   28
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         Height          =   195
         Index           =   5
         Left            =   6480
         TabIndex        =   23
         Top             =   1950
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procesado de las piezas"
         Height          =   195
         Index           =   16
         Left            =   135
         TabIndex        =   21
         Top             =   1935
         Width           =   1755
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   12735
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizada por"
         Height          =   195
         Index           =   7
         Left            =   6480
         TabIndex        =   19
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orden Compra"
         Height          =   195
         Index           =   9
         Left            =   9540
         TabIndex        =   18
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   675
         Width           =   375
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Recepción"
         Height          =   195
         Index           =   6
         Left            =   9540
         TabIndex        =   15
         Top             =   315
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
         Top             =   1035
         Width           =   585
      End
   End
   Begin TrueDBGrid80.TDBGrid gridA 
      Height          =   2205
      Left            =   0
      TabIndex        =   25
      Top             =   5535
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   3889
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
      Columns(3).Caption=   "Ref. Cliente"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ID_ENSAYO"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "General Number"
      Columns(4).DropDown=   "tEstados"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=9181"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=9102"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AutoDropDown=1"
      Splits(0)._ColumnProps(8)=   "Column(0).DropDownList=1"
      Splits(0)._ColumnProps(9)=   "Column(0).AutoCompletion=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Width=3572"
      Splits(0)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._WidthInPix=3493"
      Splits(0)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(14)=   "Column(1)._ColStyle=8193"
      Splits(0)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=3651"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=3572"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(2).AutoDropDown=1"
      Splits(0)._ColumnProps(23)=   "Column(2).DropDownList=1"
      Splits(0)._ColumnProps(24)=   "Column(2).AutoCompletion=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=4260"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=4180"
      Splits(0)._ColumnProps(33)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(35)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=0,.bold=0,.fontsize=825"
      _StyleDefs(37)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.namedParent=40,.bold=0"
      _StyleDefs(42)  =   ":id=23,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=54,.parent=11,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=12"
      _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=15"
      _StyleDefs(62)  =   "Named:id=37:Normal"
      _StyleDefs(63)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(64)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(65)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(66)  =   "Named:id=38:Heading"
      _StyleDefs(67)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(69)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(70)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(71)  =   "Named:id=39:Footing"
      _StyleDefs(72)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=40:Selected"
      _StyleDefs(74)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(75)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(76)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(77)  =   "Named:id=41:Caption"
      _StyleDefs(78)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(79)  =   "Named:id=42:HighlightRow"
      _StyleDefs(80)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(81)  =   "Named:id=43:EvenRow"
      _StyleDefs(82)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(83)  =   "Named:id=44:OddRow"
      _StyleDefs(84)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(85)  =   "Named:id=47:RecordSelector"
      _StyleDefs(86)  =   ":id=47,.parent=38"
      _StyleDefs(87)  =   "Named:id=50:FilterBar"
      _StyleDefs(88)  =   ":id=50,.parent=37"
   End
   Begin TrueDBGrid80.TDBGrid gridP 
      Height          =   2250
      Left            =   45
      TabIndex        =   30
      Top             =   3285
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   3969
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Designación"
      Columns(0).DataField=   ""
      Columns(0).NumberFormat=   "Standard"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Material"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "T.T."
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "tResponsables"
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
      Columns(5).Caption=   "Recibidas"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "General Number"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   4
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Identificadas"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "IDEN_CLIENTE"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "IDEN_CANAGROSA"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3731"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3651"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4260"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4180"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=3678"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3598"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=4471"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4392"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=3122"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=3043"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=8193"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2487"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2408"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=926"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=847"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=8193"
      Splits(0)._ColumnProps(43)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(45)=   "Column(7).Width=3784"
      Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=3704"
      Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(49)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(51)=   "Column(8).Width=3784"
      Splits(0)._ColumnProps(52)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(8)._WidthInPix=3704"
      Splits(0)._ColumnProps(54)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(55)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
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
      Caption         =   "I. Probetas"
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
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
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=28,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=12"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=62,.parent=11,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=12"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=66,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=12"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=70,.parent=11"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=12"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=74,.parent=11"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=12"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=15"
      _StyleDefs(76)  =   "Named:id=37:Normal"
      _StyleDefs(77)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
      _StyleDefs(78)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(79)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(80)  =   "Named:id=38:Heading"
      _StyleDefs(81)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(83)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(84)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(85)  =   "Named:id=39:Footing"
      _StyleDefs(86)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=40:Selected"
      _StyleDefs(88)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=975"
      _StyleDefs(89)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(90)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(91)  =   "Named:id=41:Caption"
      _StyleDefs(92)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(93)  =   "Named:id=42:HighlightRow"
      _StyleDefs(94)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(95)  =   "Named:id=43:EvenRow"
      _StyleDefs(96)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(97)  =   "Named:id=44:OddRow"
      _StyleDefs(98)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(99)  =   "Named:id=47:RecordSelector"
      _StyleDefs(100) =   ":id=47,.parent=38"
      _StyleDefs(101) =   "Named:id=50:FilterBar"
      _StyleDefs(102) =   ":id=50,.parent=37"
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12105
      Picture         =   "frmCE_Recepcion_Nuevo.frx":3C8E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Control de Eficacia"
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
      TabIndex        =   22
      Top             =   75
      Width           =   3495
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   -45
      Width           =   12825
   End
End
Attribute VB_Name = "frmCE_Recepcion_Nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xP As New XArrayDB
Dim xA As New XArrayDB
Dim xAnalisis As New XArrayDB
Dim xProbetas As New XArrayDB
Const filasP As Integer = 50
Const ColP As Integer = 8
Private Enum ColsP
    DESIGNACION = 0
    MATERIAL = 1
    TT = 2
    DIMENSION = 3
    NPROBETAS = 4
    RECIBIDAS = 5
    IDENTIFICADA = 6
    IDEN_CLIENTE = 7
    IDEN_CANAGROSA = 8
End Enum
Const filasA As Integer = 50
Const ColA As Integer = 4
Private Enum ColsA
    TIPO_ENSAYO = 0
    NORMA = 1
    DESIGNACION = 2
    REFCLIENTE = 3
    ID_TIPO_ENSAYO = 4
End Enum
Private Sub cmdborrar_Click()
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    For i = 0 To ColP
        gridP.SelBookmarks.Add gridP.Bookmark
        For j = 0 To filasA
            If Not IsEmpty(xA(j, ColsA.DESIGNACION)) Then
'J100-I
'                If Trim(xA(j, ColsA.DESIGNACION)) = Trim(xP(gridP.Row, ColsP.DESIGNACION)) Then
                If Trim(xA(j, ColsA.DESIGNACION)) = Trim(xP(gridP.Bookmark, ColsP.DESIGNACION)) Then
'J100-F
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
    gridA.Refresh
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

Private Sub cmbBanos_Change()
    If cmbbanos.Text <> "" Then
        Dim oBANO As New clsBanos
        oBANO.cargar_bano cmbbanos.BoundText
        If oBANO.getFICHA_ID = 0 Then
            cmbproceso.Text = ""
            If MsgBox("El baño no tiene ficha asignada. ¿Desea crearla?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                frmBANO_Detalle.PK = cmbbanos.BoundText
                frmBANO_Detalle.Show 1
            Else
                Exit Sub
            End If
        Else
            cmbproceso.BoundText = oBANO.getFICHA_ID
            cargar_ficha cmbbanos.BoundText, oBANO.getFICHA_ID
        End If
    End If
End Sub

Private Sub chkSinEspecificar_Click()
    If chkSinEspecificar.value = Checked Then
        fprocesado.value = "01/01/1900"
        fprocesado.Enabled = False
    Else
        fprocesado.value = Date
        fprocesado.Enabled = True
    End If
End Sub
Private Sub cmbClientes_change()
    cargar_banos
End Sub
Private Sub cmbLote_change()
'    If ensayos.ListItems.Count > 0 Then
'        If cmbLote.getPK_SALIDA <> 0 Then
'            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(12) = cmbLote.getPK_SALIDA
'        End If
'    End If
End Sub

Private Sub cmbproceso_change()
    If cmbproceso.Text <> "" Then
    End If
End Sub
Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    gridP.Col = 0
    gridP.Row = 0
    gridA.Col = 0
    gridA.Row = 0

    If validar = True Then
        Me.MousePointer = 11
        Dim oce_recepcion As New clsCe_recepcionX
        Dim RECEPCION As Long
        Dim i As Integer
        oce_recepcion.CrearID
        ' Generamos el registro de las muestras
        Dim omuestra As New clsMuestra
        Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
        Dim oce_tipo_ensayo_detalle As New clsCe_tipos_ensayos_detalle
        Dim oTipo_analisis As New clsTipos_analisis
        Dim oDatos_especificos As New clsDatos_valores
        Dim oTDA As New clsTipos_datos_analisis
        Dim oBANO As New clsBanos
        Dim MUESTRA As Long
        Dim rs As ADODB.RecordSet
        Dim indice As Integer
        For i = 0 To filasA
         If Not IsEmpty(xA(i, ColsA.ID_TIPO_ENSAYO)) Then
          If Trim(xA(i, ColsA.ID_TIPO_ENSAYO)) <> "" Then
            oce_tipo_ensayo.Carga (CLng(xA(i, ColsA.ID_TIPO_ENSAYO)))
            oTipo_analisis.CARGAR (oce_tipo_ensayo.getTIPO_ANALISIS_ID)
            With omuestra
                .setTIPO_MUESTRA_ID = oTipo_analisis.getTIPO_MUESTRA_ID
                .setTIPO_ANALISIS_ID = oce_tipo_ensayo.getTIPO_ANALISIS_ID
                .setANALISIS_MODIFICADO = 2 ' Para identificar que es un CE
                .setFECHA_MUESTREO = Format(fecha.value, "yyyy-mm-dd")
                .setENTIDAD_MUESTREO_ID = cmbrealizada.BoundText
                .setDETALLE_MUESTREO = ""
                .setOBSERVACIONES_MUESTREO = ""
                .setFECHA_RECEPCION = Format(fecha.value, "yyyy-mm-dd")
                .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                .setFORMATO_ID = cmbenvases.BoundText
                .setENTIDAD_ENTREGA_ID = cmbentregada.BoundText
                .setDETALLE_ENTREGA = ""
                .setOBSERVACIONES_ENTREGA = ""
                .setCLIENTE_ID = cmbClientes.BoundText
                .setREFERENCIA_CLIENTE = xA(i, ColsA.REFCLIENTE)
                .setFECHA_PREV_FIN = Format(fecha.value, "yyyy-mm-dd")
                .setOBSERVACIONES = ""
                .setANULADA = 0
                .setPRECINTO = ""
                .setBANO_ID = cmbbanos.BoundText
'J51
                .setFECHA_COMIENZO = "0000-00-00"
                .setFECHA_CIERRE = "0000-00-00"
                .setCERRADA = 0
                .setDOCUMENTO_PAGO = 0
                .setULT_EDICION_IMP = 0
                .setPRECIO = moneda_bd("0")
                MUESTRA = .guardarMuestra
                .informar_precio_muestra MUESTRA
            End With
            ' Datos específicos de la muestra
            Set rs = oTDA.Listado_por_tipo_analisis(oce_tipo_ensayo.getTIPO_ANALISIS_ID)
            indice = 1
            If rs.RecordCount > 0 Then
                Do
                    With oDatos_especificos
                        .setMUESTRA_ID = MUESTRA
                        .setBANO_ID = cmbbanos.BoundText
                        .setTIPO_DATO_ID = rs(0)
                        If rs(0) = 28 Then ' Orden de compra
                            .setVALOR = txtdatos(0)
                        Else
                            .setVALOR = ""
                        End If
                        .setORDEN = indice
                        .Insertar
                        indice = indice + 1
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Recepción del control de eficacia
            oce_tipo_ensayo_detalle.Carga CLng(xA(i, ColsA.ID_TIPO_ENSAYO))
            With oce_recepcion
                .setID_RECEPCION = .getID_RECEPCION
                .setMUESTRA_ID = MUESTRA
                .setTIPO_ENSAYO_ID = CLng(xA(i, ColsA.ID_TIPO_ENSAYO))
                .setORDEN = i
                .setFECHA = Format(fecha.value, "yyyy-mm-dd")
                .setENSAYO = oce_tipo_ensayo_detalle.getENSAYO
                .setIDENTIFICACION = ""
                .setIDENTIFICACION_CANAGROSA = ""
                .setDESIGNACION = CStr(xA(i, ColsA.DESIGNACION))
                ' Informar campo PROBETAS, DIMENSION y CANTIDAD
                Dim j As Integer
                If xA(i, ColsA.DESIGNACION) = "TODAS" Then
                    .setPROBETA = ""
                    .setDIMENSION = ""
                    .setCANTIDAD = 0
'JGM-I Verificar si exiten materiales o dimensiones distintas
                    Dim material_aux As String
                    Dim dimension_aux As String
                    Dim sw_material As Boolean
                    sw_material = False
                    Dim sw_dimension As Boolean
                    sw_dimension = False
                    
                    For j = 0 To filasP
                        If Not IsEmpty(xP(j, ColsP.DESIGNACION)) Then
                         If Trim(xP(j, ColsP.DESIGNACION)) <> "" Then
                            If material_aux = "" Then
                                material_aux = Trim(CStr(xP(j, ColsP.MATERIAL)))
                            End If
                            If dimension_aux = "" Then
                                dimension_aux = Trim(CStr(xP(j, ColsP.DIMENSION)))
                            End If
                            If Trim(CStr(xP(j, ColsP.MATERIAL))) <> material_aux Then
                                sw_material = True
                            End If
                            If Trim(CStr(xP(j, ColsP.DIMENSION))) <> dimension_aux Then
                                sw_dimension = True
                            End If
                         End If
                        End If
                    Next
'JGM-F
                    For j = 0 To filasP
                        If Not IsEmpty(xP(j, ColsP.DESIGNACION)) Then
                         If Trim(xP(j, ColsP.DESIGNACION)) <> "" Then
'JGM-I Si es todas, hay que informar una vez por cada probeta recibida
                            Dim k As Integer
                            If sw_material Then
                                For k = 1 To CInt(xP(j, ColsP.RECIBIDAS))
                                    .setPROBETA = .getPROBETA & Trim(CStr(xP(j, ColsP.MATERIAL))) & ";"
                                Next
                            Else
                                .setPROBETA = Trim(CStr(xP(j, ColsP.MATERIAL)))
                            End If
                            If sw_dimension Then
                                For k = 1 To CInt(xP(j, ColsP.RECIBIDAS))
                                    .setDIMENSION = .getDIMENSION & Trim(CStr(xP(j, ColsP.DIMENSION))) & ";"
                                Next
                            Else
                                .setDIMENSION = Trim(CStr(xP(j, ColsP.DIMENSION)))
                            End If
'JGM-F
                            .setCANTIDAD = .getCANTIDAD + CInt(xP(j, ColsP.RECIBIDAS))
                         End If
                        End If
                    Next
                Else
                    For j = 0 To filasP
                        If Trim(CStr(xP(j, ColsP.DESIGNACION))) = Trim(CStr(xA(i, ColsA.DESIGNACION))) Then
                            .setPROBETA = Trim(CStr(xP(j, ColsP.MATERIAL)))
                            .setDIMENSION = Trim(CStr(xP(j, ColsP.DIMENSION))) & ";"
                            .setCANTIDAD = CInt(xP(j, ColsP.RECIBIDAS))
                        End If
                    Next
                End If
                .setUNIDAD_ID = oce_tipo_ensayo_detalle.getUNIDAD_ID
                If chkSinEspecificar.value = Unchecked Then
                    .setFECHA_PROCESADO_PIEZAS = Format(fprocesado.value, "yyyy-mm-dd")
                'J100-I
                Else
                    .setFECHA_PROCESADO_PIEZAS = "1900-01-01"
                'J100-F
                End If
                ' Espesor
                If oce_tipo_ensayo.getINCLUYE_ESPESOR = 1 Then
                    .setESPESOR = txtdatos(1)
                Else
                    .setESPESOR = "No requiere espesor."
                End If
                .setLOTE_PROBETA_ID = 0
                If oce_tipo_ensayo.getLOTE_PROBETAS = 1 Then
                    If cmbLote.getTEXTO <> "" Then
                        .setLOTE_PROBETA_ID = cmbLote.getPK_SALIDA
                    End If
                End If
               .Insertar
            End With
           End If
          End If
        Next
        Me.MousePointer = 0
        MsgBox "La recepción se ha realizado correctamente. Proceda ahora a informar los datos de las probetas.", vbInformation, App.Title
        frmCE_Recepcion_Nuevo_Detalle_Probetas.lDESIGNACION = ""
        frmCE_Recepcion_Nuevo_Detalle_Probetas.lProbetas = ""
        For i = 0 To filasP
          If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
           If Trim(xP(i, ColsP.DESIGNACION)) <> "" Then
              frmCE_Recepcion_Nuevo_Detalle_Probetas.lDESIGNACION = frmCE_Recepcion_Nuevo_Detalle_Probetas.lDESIGNACION & Trim(xP(i, ColsP.DESIGNACION)) & ";"
              frmCE_Recepcion_Nuevo_Detalle_Probetas.lProbetas = frmCE_Recepcion_Nuevo_Detalle_Probetas.lProbetas & CInt(xP(i, ColsP.RECIBIDAS)) & ";"
           End If
          End If
        Next
        frmCE_Recepcion_Nuevo_Detalle_Probetas.lRecepcion = oce_recepcion.getID_RECEPCION
        frmCE_Recepcion_Nuevo_Detalle_Probetas.Show 1
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Recepcion")
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Initialize()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Me.Left = 100
    Me.Top = 50
    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    Call cargar_combos
    fecha = Date
    fprocesado = Date
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbbanos.BoundText = "" Then
        MsgBox "Debe asignar un baño a la selección.", vbExclamation, App.Title
        cmbbanos.SetFocus
        validar = False
        Exit Function
    End If
    If cmbproceso.BoundText = "" Then
        MsgBox "Debe asignar un proceso a la selección.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If txtdatos(0) = "" Then
        MsgBox "Informe la orden de compra.", vbExclamation, App.Title
        validar = False
        txtdatos(0).SetFocus
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
    ' Verificar número de probetas
    Dim i As Integer
    Dim algo As Boolean
    Dim numero_probetas As Boolean
    
    algo = False
    numero_probetas = False
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
            algo = True
            If IsEmpty(xP(i, ColsP.RECIBIDAS)) Then
                numero_probetas = True
            End If
        End If
    Next
    If algo = False Then
        MsgBox "Debe indicar las probetas a recepcionar.", vbExclamation, App.Title
        validar = False
        gridP.SetFocus
        Exit Function
    End If
    If numero_probetas = True Then
        MsgBox "Debe indicar el numero de probetas recibidas.", vbExclamation, App.Title
        validar = False
        gridP.SetFocus
        Exit Function
    End If
    ' Validar que la designación venga informada
    algo = False
    For i = 0 To filasP
        If Not IsEmpty(xP(i, ColsP.RECIBIDAS)) Then
'JGM-I
            If xP(i, ColsP.DESIGNACION) = "" And xP(i, ColsP.RECIBIDAS) <> "" Then
'JGM-f
                algo = True
            End If
        End If
    Next
    If algo Then
        MsgBox "Debe indicar las designaciones de las probetas.", vbExclamation, App.Title
        validar = False
        gridP.SetFocus
        Exit Function
    End If
    
    ' Verificar número de ensayos
    algo = False
    Dim ref_cliente As Boolean
    Dim DESIG As Boolean
    ref_cliente = False
    DESIG = False
    For i = 0 To filasA
        If Not IsEmpty(xA(i, ColsA.TIPO_ENSAYO)) Then
'JGM-I
         If Trim(xA(i, ColsA.TIPO_ENSAYO)) <> "" Then
'JGM-F
            algo = True
            If IsEmpty(xA(i, ColsA.REFCLIENTE)) Then
                ref_cliente = True
            End If
            If IsEmpty(xA(i, ColsA.DESIGNACION)) Then
                DESIG = True
            End If
'JGM-I
         End If
'JGM-F
            
        End If
    Next
    If algo = False Then
        MsgBox "Debe indicar los ensayos a recepcionar.", vbExclamation, App.Title
        validar = False
        gridA.SetFocus
        Exit Function
    End If
    If ref_cliente = True Then
        MsgBox "Debe indicar las referencias del cliente.", vbExclamation, App.Title
        validar = False
        gridA.SetFocus
        Exit Function
    End If
    If DESIG = True Then
        MsgBox "Debe indicar las probetas de los análisis.", vbExclamation, App.Title
        validar = False
        gridA.SetFocus
        Exit Function
    End If
    
End Function
Private Sub cargar_combos()
    cargar_clientes
    Cargar_Combo cmbproceso, New clsCe_ficha
    Cargar_Combo cmbenvases, New clsformatos
    Cargar_Combo cmbentregada, New clsEntidades_Entrega
    Cargar_Combo cmbrealizada, New clsEntidades_muestreo
    llenar_combo cmbLote, New clsCe_lotes_probetas, 0, frmCE_Lote_Probeta, ""
    cmbLote.desactivar
End Sub
Private Sub cargar_ficha(BANO As Long, ficha As Long)
    inicializar_grid
    ' Recuperamos el secuencial de recepción del control de eficacia
'    Dim oce_recepcion As New clsCe_recepcionx
'    oce_recepcion.CrearID
'    txtrecepcion = oce_recepcion.getID_RECEPCION
    ' Recuperamos los datos de la ficha de proceso
    Dim oCe_bano_probetas As New clsCe_banos_probetas
    Dim i As Integer
    i = 0
    Dim rs As ADODB.RecordSet
    Set rs = oCe_bano_probetas.Listado(BANO)
    If rs.RecordCount > 0 Then
        Do
            xP(i, ColsP.DESIGNACION) = CStr(rs(0))
            xP(i, ColsP.MATERIAL) = CStr(rs(1))
            xP(i, ColsP.TT) = CStr(rs(2))
            xP(i, ColsP.DIMENSION) = CStr(rs(3))
            xP(i, ColsP.NPROBETAS) = CStr(rs(4))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Dim oCe_bano_ensayos As New clsCe_banos_ensayos
    i = 0
'JGM-I
    txtdatos(1) = ""
    txtdatos(1).Enabled = False
    cmbLote.Limpiar
    cmbLote.desactivar
'JGM-F
    Set rs = oCe_bano_ensayos.Listado(BANO)
    If rs.RecordCount > 0 Then
        Do
            xA(i, ColsA.TIPO_ENSAYO) = CStr(rs(0))
            xA(i, ColsA.NORMA) = CStr(rs(1))
            xA(i, ColsA.DESIGNACION) = CStr(rs(2))
            xA(i, ColsA.ID_TIPO_ENSAYO) = CStr(rs(3))
            i = i + 1
            If rs(4) = 1 Then
                txtdatos(1).Enabled = True
            End If
            If rs(5) = 1 Then
                cmbLote.activar
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    gridP.Refresh
    gridA.Refresh
    cargar_combo_tipos_ensayos ficha
    cargar_combo_probetas
End Sub

Public Sub cargar_clientes()
    'Clientes
    Dim obanos As New clsBanos
    Set cmbClientes.RowSource = obanos.Listado_Clientes_CE
    cmbClientes.ListField = "C2"
    cmbClientes.DataField = "C1" 'campo asociado
    cmbClientes.BoundColumn = "C1" 'lo que realmente
    Set obanos = Nothing
End Sub
Public Sub cargar_banos()
    'Clientes
    If cmbClientes.BoundText <> "" Then
        Dim obanos As New clsBanos
        Set cmbbanos.RowSource = obanos.Listado_por_Cliente_con_CE(cmbClientes.BoundText)
        cmbbanos.ListField = "NOMBRE"
        cmbbanos.DataField = "ID_BANO" 'campo asociado
        cmbbanos.BoundColumn = "ID_BANO" 'lo que realmente
        Set obanos = Nothing
    End If
End Sub

Private Sub gridA_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then
        If Not IsEmpty(xA(gridA.Row, ColsA.ID_TIPO_ENSAYO)) Then
            frmCE_Tipo_Ensayo.PK = CLng(xA(gridA.Row, ColsA.ID_TIPO_ENSAYO))
            frmCE_Tipo_Ensayo.Show 1
        End If
    End If
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80FFFF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error

    gridP.Col = 0
    gridP.Row = 0
    gridA.Col = 0
    gridA.Row = 0
    xP.Clear
    xP.ReDim 0, filasP, 0, ColP
    xP.Clear
    Set gridP.Array = xP
    gridP.Refresh
    xA.Clear
    xA.ReDim 0, filasA, 0, ColA
    xA.Clear
    Set gridA.Array = xA
    gridA.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo"
End Sub

Private Sub cargar_combo_tipos_ensayos(ficha As Long)
    Dim rs As ADODB.RecordSet
    Dim ote As New clsCe_tipos_ensayos
'    Set rs = ote.Listado(ficha)
    Set rs = ote.Listado("", "")
    xAnalisis.Clear
    If rs.RecordCount > 0 Then
        xAnalisis.ReDim 1, rs.RecordCount, 1, 3
        Dim i As Integer
        i = 1
        Do
            xAnalisis(i, 1) = CStr(rs(0))
            xAnalisis(i, 2) = CStr(rs(3))
            xAnalisis(i, 3) = CStr(rs(2))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xAnalisis.ReDim 1, 1, 1, 3
    End If
    Set tAnalisis.Array = xAnalisis
    tAnalisis.Refresh
    gridA.Refresh
End Sub
Private Sub cargar_combo_probetas()
'    xProbetas.ReDim 1, 1, 1, 1
'    xProbetas.Clear
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
'        xProbetas(1, 2) = "0"
        j = 2
        For i = 0 To filasP
            If Not IsEmpty(xP(i, ColsP.DESIGNACION)) Then
                xProbetas(j, 1) = CStr(xP(i, ColsP.DESIGNACION))
'                xProbetas(j, 2) = CStr(j)
                j = j + 1
            End If
        Next
    Else
        xProbetas.ReDim 1, 1, 1, 1
    End If
    Set tProbetas.Array = xProbetas
    tProbetas.Refresh
End Sub

Private Sub tAnalisis_DropDownClose()
    gridA.Columns(ColsA.NORMA) = tAnalisis.Columns(1)
    gridA.Columns(ColsA.ID_TIPO_ENSAYO) = tAnalisis.Columns(2)
    gridA.Col = 2
'    gridA.Row = gridA.Row + 1
End Sub

Private Sub tProbetas_DropDownClose()
    gridA.Col = 0
    gridA.Row = gridA.Row + 1
End Sub

Private Sub tProbetas_DropDownOpen()
    cargar_combo_probetas
End Sub

