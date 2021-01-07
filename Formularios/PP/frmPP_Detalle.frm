VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmPP_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos a Proveedor"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15795
   Icon            =   "frmPP_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15795
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid80.TDBDropDown tDescripcion 
      Height          =   6240
      Left            =   4005
      TabIndex        =   24
      Top             =   2205
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   11007
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "DESCRIPCIÓN"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DESC. (%)"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PVP (Ud.)"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=12277"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=12171"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1746"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1640"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1667"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   2
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HC0C0C0&,.bold=0"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
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
   Begin TrueDBGrid80.TDBDropDown tFamilias 
      Height          =   4620
      Left            =   1215
      TabIndex        =   23
      Top             =   2925
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   8149
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "FAMILIA"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ID_FAMILIA"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=9975"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=9869"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3810"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3704"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   2
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HC0C0C0&,.bold=0"
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
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.TextBox txtDatos 
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
      ForeColor       =   &H80000002&
      Height          =   345
      Index           =   2
      Left            =   13725
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   19
      Top             =   7860
      Width           =   1740
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   900
      Left            =   11925
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8250
      Width           =   1275
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar Concepto"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   45
      TabIndex        =   12
      Top             =   630
      Width           =   15675
      Begin VB.CheckBox chkFechaRecepcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Recepción"
         Height          =   285
         Left            =   10170
         TabIndex        =   21
         Top             =   990
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   3
         Left            =   7065
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   540
         Width           =   5880
      End
      Begin VB.CheckBox chkFechaEnvio 
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Envío"
         Height          =   285
         Left            =   5940
         TabIndex        =   2
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox txtFactura 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3555
         MaxLength       =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   315
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker datFecha 
         Height          =   315
         Left            =   810
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
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
         Format          =   51838977
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbUsuario 
         Height          =   330
         Left            =   810
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   585
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaEnvio 
         Height          =   315
         Left            =   7065
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
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
         Format          =   51838977
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmdProveedor 
         Height          =   330
         Left            =   7065
         TabIndex        =   0
         Top             =   180
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaRecepcion 
         Height          =   315
         Left            =   11610
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
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
         Format          =   51838977
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmPP_Detalle.frx":08CA
         Height          =   315
         Left            =   810
         TabIndex        =   26
         Top             =   945
         Width           =   4035
         _ExtentX        =   7117
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
         Left            =   90
         TabIndex        =   27
         Top             =   990
         Width           =   465
      End
      Begin VB.Label lblObservaciones 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   285
         Left            =   5940
         TabIndex        =   18
         Top             =   585
         Width           =   1140
      End
      Begin VB.Shape fondo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   645
         Left            =   0
         Top             =   -1125
         Width           =   13500
      End
      Begin VB.Image imagen 
         Height          =   480
         Left            =   12825
         Picture         =   "frmPP_Detalle.frx":0910
         Top             =   -1080
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   1
         Left            =   6255
         TabIndex        =   13
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   900
      Left            =   13230
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Modificar paquete"
      Top             =   8250
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   900
      Left            =   14535
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   8250
      Width           =   1230
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5805
      Left            =   45
      TabIndex        =   5
      Top             =   2025
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   10239
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "REF."
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "FAMILIA"
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "tFamilias"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "FAMILIA_ID"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DESCRIPCIÓN"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "tDescripcion"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "UDs."
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "General Number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DESC. (%)"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "PVP (Ud.)"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "Currency"
      Columns(6).ConvertEmptyCell=   1
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "IMPORTE"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerStyle=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=131585"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4948"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4842"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=131585"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1).AutoDropDown=1"
      Splits(0)._ColumnProps(14)=   "Column(1).AutoCompletion=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=291"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=185"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=131585"
      Splits(0)._ColumnProps(20)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=11271"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=11165"
      Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=131585"
      Splits(0)._ColumnProps(27)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(28)=   "Column(3).AutoDropDown=1"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=1984"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1879"
      Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=131585"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=1746"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=1640"
      Splits(0)._ColumnProps(38)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._ColStyle=131585"
      Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(41)=   "Column(6).Width=2170"
      Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=2064"
      Splits(0)._ColumnProps(44)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(45)=   "Column(6)._ColStyle=131586"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(47)=   "Column(7).Width=344"
      Splits(0)._ColumnProps(48)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._WidthInPix=238"
      Splits(0)._ColumnProps(50)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(7)._ColStyle=131586"
      Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      Appearance      =   2
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   0
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HC0E0FF&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41,.alignment=0,.fgcolor=&H80000001&"
      _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
      _StyleDefs(13)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(14)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H80000009&"
      _StyleDefs(16)  =   ":id=3,.fgcolor=&H80000001&,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(17)  =   ":id=3,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43,.alignment=3"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
      _StyleDefs(26)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
      _StyleDefs(27)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43,.alignment=2"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=62,.parent=11"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=12"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=66,.parent=11"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=12"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=32,.parent=11"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=12"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=11"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=12"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=36,.parent=11,.alignment=1"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=33,.parent=12"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=34,.parent=13"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=35,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=58,.parent=11,.alignment=1"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=12"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=13"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=15"
      _StyleDefs(71)  =   "Named:id=37:Normal"
      _StyleDefs(72)  =   ":id=37,.parent=0,.alignment=2,.bgcolor=&H80000018&,.appearance=0,.borderType=0"
      _StyleDefs(73)  =   ":id=37,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(74)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(75)  =   "Named:id=38:Heading"
      _StyleDefs(76)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   ":id=38,.wraptext=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(78)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(79)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(80)  =   "Named:id=39:Footing"
      _StyleDefs(81)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(82)  =   "Named:id=40:Selected"
      _StyleDefs(83)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(84)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(85)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(86)  =   "Named:id=41:Caption"
      _StyleDefs(87)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(88)  =   "Named:id=42:HighlightRow"
      _StyleDefs(89)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(90)  =   "Named:id=43:EvenRow"
      _StyleDefs(91)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(92)  =   "Named:id=44:OddRow"
      _StyleDefs(93)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(94)  =   "Named:id=47:RecordSelector"
      _StyleDefs(95)  =   ":id=47,.parent=38"
      _StyleDefs(96)  =   "Named:id=50:FilterBar"
      _StyleDefs(97)  =   ":id=50,.parent=37"
   End
   Begin VB.CommandButton cmdRecepcionar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recepcionar"
      Height          =   900
      Left            =   10620
      Picture         =   "frmPP_Detalle.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8250
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe Total"
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
      Index           =   0
      Left            =   12240
      TabIndex        =   20
      Top             =   7920
      Width           =   1395
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Pedido a Proveedor"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   15165
      Picture         =   "frmPP_Detalle.frx":20A4
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Pedidos a Proveedores"
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
      TabIndex        =   16
      Top             =   45
      Width           =   3645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "frmPP_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private x As New XArrayDB
Private fila As Integer
Dim xFamilias As New XArrayDB
Dim xDescripcion As New XArrayDB

Const filas As Integer = 100
Const Col As Integer = 7
Const cReferencia As Integer = 0
Const cFamilia As Integer = 1
Const cFamiliaId As Integer = 2
Const cDescripcion As Integer = 3
Const cUnidades As Integer = 4
Const cDescuento As Integer = 5
Const cPrecio As Integer = 6
Const cImporte As Integer = 7

Private Sub chkFechaEnvio_Click()
    fechaEnvio.Enabled = chkFechaEnvio.value
End Sub

Private Sub chkFechaRecepcion_Click()
    fechaRecepcion.Enabled = chkFechaRecepcion.value
End Sub

Private Sub cmdRecepcionar_Click()
    frmPP_Detalle_Recepcion.PK = PK
    frmPP_Detalle_Recepcion.Show 1
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    
    inicializar_ventana
    cargar_combos
    
    cmbUsuario.desactivar
    
    fechaEnvio.value = Date
    fechaRecepcion.value = Date

    If PK <> 0 Then
        MODIFICACION
    Else
        cmdAdjuntos.Enabled = False
        Alta
    End If
End Sub

Private Sub inicializar_ventana()
    Dim i As Integer
    log (Me.Name)
    Me.top = 1700
    Me.Left = 300
    fila = 0
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
End Sub
Private Sub cmdAdjuntos_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PEDIDO_PROVEEDOR
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub

Private Sub cmdEliminar_Click()
    Dim f As Integer
    Dim c As Integer
    Dim linea As Integer
    linea = grid.Bookmark
    ' Movemos las lineas al final
    For f = linea To filas - 1
        For c = 0 To Col
            x(f, c) = x(f + 1, c)
        Next
    Next
    grid.Refresh
    grid.SetFocus
    SumarImportes
End Sub
Private Sub cmdok_Click()
    guardarCambios
End Sub
Private Sub guardarCambios()

   On Error GoTo modificarPaquete_Error
    Dim strMensaje As String
    If datos_correctos Then
                
        Dim oPP As New clsPP
        Dim lngPP As Long
        
        With oPP
              .setPRESUPUESTO = txtDatos(2)
              .setOBSERVACIONES = txtDatos(3)
              .setPROVEEDOR_ID = cmdProveedor.getPK_SALIDA
              .setCENTRO_ID = cmbCentro.BoundText
    
'              If Trim(txtFactura.Text) <> "" Then
'                 .setFACTURA_RECIBIDA = 1
'                 .setFFACTURA = Format(datFechaFactura.value, "yyyy-mm-dd")
'                 .setNFACTURA = Trim(txtFactura)
'              Else
                 .setFACTURA_RECIBIDA = 0
                 .setFFACTURA = "0000-00-00"
                 .setNFACTURA = 0
'              End If
              If chkFechaEnvio.value = Checked Then
                .setFECHA_ENVIO = "'" & Format(fechaEnvio, "yyyy-mm-dd") & "'"
              Else
                .setFECHA_ENVIO = "NULL"
              End If
              If chkFechaRecepcion.value = Checked Then
                .setFECHA_RECEPCION = "'" & Format(fechaRecepcion, "yyyy-mm-dd") & "'"
              Else
                .setFECHA_RECEPCION = "NULL"
              End If
              
              If PK = 0 Then
                 strMensaje = "Se va a crear un nuevo pedido a proveedor. ¿Está seguro?"
                 .setFECHA_CREACION = Left(Format(Date, "yyyy-mm-dd hh:nn:ss"), 10)
                 .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                 .setESTADO = SC_ESTADO_PENDIENTE
                 .setTIPO = TOBJETO_PEDIDO_PROVEEDOR
              Else
                 strMensaje = "Va a modificar el pedido a proveedor. ¿Está seguro?"
              End If
              
              If MsgBox(strMensaje, vbQuestion + vbYesNo, App.Title) = vbYes Then
                If PK <> 0 Then
                   lngPP = PK
                   
                   If .Modificar(lngPP) = False Then
                    MsgBox "Se ha producido un error al modificar el Pedido a Proveedor.", vbCritical, App.Title
                    Exit Sub
                   End If
                Else
                   lngPP = .Insertar
                   If lngPP = 0 Then
                    MsgBox "Se ha producido un error al insertar el Pedido a Proveedor.", vbCritical, App.Title
                    Exit Sub
                   End If
                End If
              Else
                Unload Me
              End If
        End With
        ' CONCEPTOS
        Dim oPP_Detalle As New clsPP_Detalle
        Dim i As Long
        oPP_Detalle.Eliminar_Pedido lngPP
        For i = 0 To filas - 1
           If filaCargada(i) Then
            With oPP_Detalle
               .setPP_ID = lngPP
               .setREFERENCIA = Trim(x(i, cReferencia))
               .setFAMILIA_ID = Trim(x(i, cFamiliaId))
               .setDESCRIPCION = Trim(x(i, cDescripcion))
               .setUNIDADES = moneda_bd(x(i, cUnidades))
               .setDESCUENTO = CLng(x(i, cDescuento))
               .setPRECIO = moneda_bd(Trim(x(i, cPrecio)))
               .setIMPORTE = moneda_bd(Trim(x(i, cImporte)))
               .Insertar
            End With
           End If
        Next i

        If PK = 0 Then
           MsgBox "El pedido a proveedor nº " & oPP.getNUMERO & "/" & oPP.getANNO & " se ha creado correctamente.", vbOKOnly + vbInformation, App.Title
        Else
           MsgBox "El pedido a proveedor se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
        End If
        Unload Me
        
      End If

   On Error GoTo 0
   Exit Sub

modificarPaquete_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure modificarPaquete of Formulario frmPP_Detalle"
End Sub
 
Private Function recorrer_filas() As Double
    
    Dim indice As Integer
    Dim encontrado As Boolean
    encontrado = True
    indice = 0
    recorrer_filas = 0
    
    Do
        If x(indice, cPrecio) <> "" Then
           recorrer_filas = recorrer_filas + CDbl(Trim(x(indice, cPrecio)))
        Else
           encontrado = False
        End If
        
        indice = indice + 1
    Loop Until Not encontrado Or indice > filas
    
End Function

Private Sub cmdcancel_Click()
    Unload Me
End Sub
' --------------------------

Private Function calcularNumeroFilas() As Integer
    Dim i As Long
    i = 0
    Do While filaCargada(i)
        i = i + 1
    Loop
    calcularNumeroFilas = i  'sale con i incrementada sobre la última fila real
End Function

Private Function filaCargada(fila As Long) As Boolean
    Dim i As Integer
    filaCargada = False
    For i = 0 To Col - 1
        If Trim(x(fila, i)) <> "" Then
            filaCargada = True
        End If
    Next i
End Function
Private Sub tDescripcion_DropDownClose()
'    grid.Columns(cDescripcion) = tDescripcion.Columns(0)
    grid.Columns(cDescuento) = tDescripcion.Columns(1)
    grid.Columns(cPrecio) = tDescripcion.Columns(2)
    grid.Col = cUnidades
End Sub

Private Sub tFamilias_DropDownClose()
    grid.Columns(cFamilia) = tFamilias.Columns(0)
    grid.Columns(cFamiliaId) = tFamilias.Columns(1)
    grid.Col = cDescripcion
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 2, 3:
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

' funciones auxiliares del formulario
Private Sub Alta()
    Me.MousePointer = vbHourglass
    cmdok.Caption = "Alta"
    lblsubtitulo = "Creación de nuevo pedido a proveedor"
    cmdProveedor.activar
    cmbUsuario.activar
    cmbUsuario.MostrarElemento USUARIO.getID_EMPLEADO
    txtFactura = ""
    Me.MousePointer = vbNormal
End Sub
Private Sub MODIFICACION()
    Dim oPP As New clsPP
    Dim usu As New clsUsuarios
    Dim rs As ADODB.Recordset
    Dim lngTotalConceptosPaquete As Long
   On Error GoTo MODIFICACION_Error
    cmdRecepcionar.Visible = True
    datFecha.Enabled = False
    Me.MousePointer = vbHourglass
    cmdProveedor.activar
    If oPP.Carga(PK) = True Then
        With oPP
            lblsubtitulo = "Detalle del Pedido a Proveedor: " & .getNUMERO & "/" & .getANNO
            txtDatos(2) = .getPRESUPUESTO
            txtDatos(3) = .getOBSERVACIONES
            cmdProveedor.MostrarElemento .getPROVEEDOR_ID
            cmbCentro.BoundText = .getCENTRO_ID
            cmbUsuario.MostrarElemento .getUSUARIO_ID
            txtFactura = .getNFACTURA
            If .getFECHA_ENVIO <> "" Then
                fechaEnvio.value = .getFECHA_ENVIO
                chkFechaEnvio.value = Checked
            Else
                chkFechaEnvio.value = Unchecked
            End If
            If .getFECHA_RECEPCION <> "" And IsDate(.getFECHA_RECEPCION) Then
                fechaRecepcion.value = .getFECHA_RECEPCION
                chkFechaRecepcion.value = Checked
            Else
                chkFechaRecepcion.value = Unchecked
            End If
            'Carga del GRID con listado de conceptos
            Dim oPP_Detalle As New clsPP_Detalle
            Set rs = oPP_Detalle.Listado(PK)
            lngTotalConceptosPaquete = rs.RecordCount
            
            If rs.RecordCount <> 0 Then
                Dim i As Integer
                Dim impTxt As String
                i = 0
                Do
                    x(i, cReferencia) = CStr(rs("REFERENCIA"))
                    If IsNull(rs("FAMILIA")) Then
                        x(i, cFamilia) = ""
                        x(i, cFamiliaId) = "0"
                    Else
                        x(i, cFamilia) = CStr(rs("FAMILIA"))
                        x(i, cFamiliaId) = CStr(rs("FAMILIA_ID"))
                    End If
                    x(i, cDescripcion) = CStr(rs("DESCRIPCION"))
                    x(i, cUnidades) = CStr(rs("UNIDADES"))
                    x(i, cDescuento) = CStr(rs("DESCUENTO"))
                    x(i, cPrecio) = moneda(Trim(rs("PRECIO")))
                    x(i, cImporte) = moneda(Trim(rs("IMPORTE")))
                    i = i + 1
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            grid.Row = 0
            grid.Col = 0
            grid.Refresh
        End With
    End If
    Me.MousePointer = vbNormal
    Set oPP = Nothing
    SumarImportes

   On Error GoTo 0
   Exit Sub

MODIFICACION_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MODIFICACION of Formulario frmPP_Detalle"
End Sub

Private Function calcularImporte(unidades As Double, desc As Double, PRECIO As Double) As String
    Dim importeTotal As Double
    Dim DESCUENTO As Double
    
    If unidades = 0 Then
        calcularImporte = moneda(0)
    End If
    
    DESCUENTO = 0
    If desc > 0 Then
       If desc > 100 Then
          desc = 100
       End If
       DESCUENTO = (unidades * PRECIO * desc) / 100
    End If
    
    importeTotal = (unidades * PRECIO) - DESCUENTO
    calcularImporte = moneda(CStr(importeTotal))
End Function

Private Function datos_correctos() As Boolean
    datos_correctos = True

    If cmbCentro.BoundText = "" Then
        MsgBox "Debe indicar el CENTRO antes de generar el pedido.", vbExclamation, App.Title
        datos_correctos = False
        cmbCentro.SetFocus
        Exit Function
   
    End If
    If cmdProveedor.getPK_SALIDA = 0 Then
        MsgBox "Debe indicar el proveedor antes de generar el pedido.", vbExclamation, App.Title
        datos_correctos = False
        cmdProveedor.SetFocus
        Exit Function
    End If
    If cmbUsuario.getTEXTO = "" Then
        MsgBox "Debe indicar el usuario antes de generar el pedido", vbExclamation, App.Title
        datos_correctos = False
        cmbUsuario.SetFocus
        Exit Function
    End If
End Function

Private Sub cargar_combos()
    llenar_combo cmdProveedor, New clsProveedor, 0, frmProveedores_Detalle, ""
    llenar_combo cmbUsuario, New clsUsuarios, 0, frmUsuarios, ""
    cargar_combo cmbCentro, New clsCentros
    cargar_combo_familias
    cargar_combo_descripcion
End Sub

Private Sub SumarImportes()
   Dim indice As Integer
   Dim Suma As Double
   On Error GoTo SumarImportes_Error

   Suma = 0

   For indice = 0 To filas - 1
        If IsNumeric(x(indice, cImporte)) Then
            Suma = Suma + CDbl(x(indice, cImporte))
        End If
   Next indice
   txtDatos(2) = moneda(CStr(Suma))

   On Error GoTo 0
   Exit Sub

SumarImportes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SumarImportes of Formulario frmPP_Detalle"
End Sub

Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Select Case LastCol
    Case cUnidades To cImporte
        If Not IsNumeric(x(LastRow, cPrecio)) Then
            x(LastRow, cPrecio) = "0"
        End If

        If Not IsNumeric(x(LastRow, cUnidades)) Then
            x(LastRow, cUnidades) = "0"
        End If
        
        If Not IsNumeric(x(LastRow, cDescuento)) Then
            x(LastRow, cDescuento) = "0"
        End If
        
'        x(LastRow, cImporte) = calcularImporte(CInt(x(LastRow, cUnidades)), CInt(x(LastRow, cDescuento)), CDbl(x(LastRow, cPrecio)))
        x(LastRow, cImporte) = calcularImporte(CDbl(x(LastRow, cUnidades)), CDbl(x(LastRow, cDescuento)), CDbl(x(LastRow, cPrecio)))
        SumarImportes
    End Select
    grid.Refresh
End Sub
Private Sub cargar_combo_familias()
    Dim rs As ADODB.Recordset
    Dim oFamilias As New clsFamilias
    Set rs = oFamilias.Listado_completo()
    xFamilias.Clear
    If rs.RecordCount > 0 Then
        xFamilias.ReDim 1, rs.RecordCount, 1, 2
        Dim i As Integer
        i = 1
        Do
            xFamilias(i, 1) = CStr(rs(2))
            xFamilias(i, 2) = CStr(rs(0))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xFamilias.ReDim 1, 1, 1, 2
    End If
    Set tFamilias.Array = xFamilias
    tFamilias.Refresh
End Sub
Private Sub cargar_combo_descripcion()
    Dim rs As ADODB.Recordset
    Dim opd As New clsPP_Detalle
    Set rs = opd.ListadoDescripciones
    xDescripcion.Clear
    If rs.RecordCount > 0 Then
        xDescripcion.ReDim 1, rs.RecordCount, 1, 3
        Dim i As Integer
        i = 1
        Do
            xDescripcion(i, 1) = CStr(rs(0))
            xDescripcion(i, 2) = CStr(rs(1))
            xDescripcion(i, 3) = CStr(rs(2))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xDescripcion.ReDim 1, 1, 1, 3
    End If
    Set tDescripcion.Array = xDescripcion
    tDescripcion.Refresh
End Sub

