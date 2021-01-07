VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProveedores_Calidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluación de Calidad del Proveedor"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14955
   Icon            =   "frmProveedores_Calidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   14955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13860
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9540
      Width           =   1050
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   15
      Left            =   3690
      TabIndex        =   1
      Top             =   11205
      Width           =   645
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8925
      Left            =   45
      TabIndex        =   2
      Top             =   540
      Width           =   14865
      _Version        =   851970
      _ExtentX        =   26220
      _ExtentY        =   15743
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      ItemCount       =   6
      SelectedItem    =   1
      Item(0).Caption =   "Autoevaluación"
      Item(0).ControlCount=   0
      Item(1).Caption =   "Evaluación Inicial"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "Label2"
      Item(1).Control(1)=   "txtDatos(0)"
      Item(1).Control(2)=   "grid"
      Item(1).Control(3)=   "tServicios"
      Item(1).Control(4)=   "tMetodos"
      Item(1).Control(5)=   "TDBDate1"
      Item(1).Control(6)=   "lblServicios"
      Item(1).Control(7)=   "txtEvaluacionInicialObservaciones"
      Item(1).Control(8)=   "cmdEvaluacionInicialBorrarServicio(0)"
      Item(2).Caption =   "Proc. NC"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Riesgo"
      Item(3).ControlCount=   0
      Item(4).Caption =   "Evaluación"
      Item(4).ControlCount=   0
      Item(5).Caption =   "KPI's"
      Item(5).ControlCount=   0
      Begin TrueDBGrid80.TDBDropDown tServicios 
         Height          =   3630
         Left            =   90
         TabIndex        =   7
         Top             =   1620
         Width           =   4845
         _ExtentX        =   8546
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
         Columns.Count   =   2
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
      Begin TrueDBGrid80.TDBDropDown tMetodos 
         Height          =   3405
         Left            =   4905
         TabIndex        =   8
         Top             =   2025
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   6006
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
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   -68335
         TabIndex        =   3
         Top             =   5760
         Width           =   5625
      End
      Begin TrueDBGrid80.TDBGrid grid 
         Height          =   5205
         Left            =   90
         TabIndex        =   6
         Top             =   675
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   9181
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Servicio/Producto a Evaluar"
         Columns(0).DataField=   ""
         Columns(0).DropDown=   "tServicios"
         Columns(0).DropDown.vt=   8
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Tipo"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Metodo de Evaluación"
         Columns(2).DataField=   ""
         Columns(2).DropDown=   "tMetodos"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Fecha Inicial"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "Short Date"
         Columns(3).ExternalEditor=   "TDBDate1"
         Columns(3).ExternalEditor.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "F. Aprobación"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Short Date"
         Columns(4).DropDown=   "tUnidades"
         Columns(4).DropDown.vt=   8
         Columns(4).ExternalEditor=   "TDBDate1"
         Columns(4).ExternalEditor.vt=   8
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Observaciones"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "ID_SERVICIO"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "ID_METODO"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4577"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4471"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3942"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3836"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=3942"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3836"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2090"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1984"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2223"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2117"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(30)=   "Column(4).AutoDropDown=1"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=1799"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1693"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(36)=   "Column(6).Width=2752"
         Splits(0)._ColumnProps(37)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(42)=   "Column(7).Width=2752"
         Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(45)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(46)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(47)=   "Column(7).Order=8"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HFFFFFF&,.bold=0,.fontsize=975"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=11"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=11,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=11,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=12"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=11"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=12"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=36,.parent=11"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=33,.parent=12"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=34,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=35,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=11"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=12"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=15"
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
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   300
         Left            =   11340
         TabIndex        =   9
         Top             =   8370
         Visible         =   0   'False
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   529
         Calendar        =   "frmProveedores_Calidad.frx":6852
         Caption         =   "frmProveedores_Calidad.frx":696A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmProveedores_Calidad.frx":69D6
         Keys            =   "frmProveedores_Calidad.frx":69F4
         Spin            =   "frmProveedores_Calidad.frx":6A52
         AlignHorizontal =   0
         AlignVertical   =   0
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
         Text            =   "14/06/2009"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39978
         CenturyMode     =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtEvaluacionInicialObservaciones 
         Height          =   1455
         Left            =   135
         TabIndex        =   11
         Top             =   6300
         Width           =   14550
         _Version        =   851970
         _ExtentX        =   25665
         _ExtentY        =   2566
         _StockProps     =   77
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdEvaluacionInicialBorrarServicio 
         Height          =   480
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   8190
         Width           =   2865
         _Version        =   851970
         _ExtentX        =   5054
         _ExtentY        =   847
         _StockProps     =   79
         Caption         =   "Borrar Servicio Seleccionado"
         Appearance      =   5
         Picture         =   "frmProveedores_Calidad.frx":6A7A
      End
      Begin VB.Label lblServicios 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Height          =   330
         Left            =   4140
         TabIndex        =   10
         Top             =   8235
         Width           =   4560
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   6030
         Width           =   2955
         _Version        =   851970
         _ExtentX        =   5212
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Observaciones sobre la Evaluación Inicial"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FICHA DE PROVEEDOR"
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
      TabIndex        =   0
      Top             =   90
      Width           =   2535
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   15210
   End
End
Attribute VB_Name = "frmProveedores_Calidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Dim x As New XArrayDB

Dim xServicios As New XArrayDB
Dim xMetodos As New XArrayDB
Const filas As Integer = 10
Const Col As Integer = 7
Private Enum COLS
    servicio = 0
    tipo = 1
    METODO = 2
    finicial = 3
    fAprobacion = 4
    OBSERVACIONES = 5
    ID_SERVICIO = 6
    ID_METODO = 7
End Enum
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEvaluacionInicialBorrarServicio_Click(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Col
        grid.SelBookmarks.Add grid.Bookmark
        x(grid.Bookmark, i) = ""
        grid.SelBookmarks.Remove 0
    Next
    evaluacionInicial_NumServicios
    grid.Refresh
    grid.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    inicializar_grid
    cargar_combos
    If PK <> 0 Then
        carga
    End If
End Sub
Private Sub carga()
    Dim oProveedor As New clsProveedor
    If oProveedor.carga(PK) Then
        lbltitulo = "PROVEEDOR : " & oProveedor.getNOMBRE
        cargar_servicios
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmProveedores_Calidad = Nothing
End Sub
Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
   On Error GoTo grid_AfterColEdit_Error
       evaluacionInicial_NumServicios
   On Error GoTo 0
   Exit Sub
grid_AfterColEdit_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grid_AfterColEdit of Formulario frmDocumento_Edicion"
End Sub
Private Sub insertar_servicios()
    ' Evidencias
    Dim oPS As New clsProveedores_servicios
   On Error GoTo insertar_servicios_Error

    oPS.Eliminar PK
    Dim i As Integer
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.Value(i, COLS.servicio)) <> "" Then
            With oPS
                .setPROVEEDOR_ID = PK
                .setORDEN = i
                .setSERVICIO_ID = x.Value(i, COLS.ID_SERVICIO)
                .setTIPO = x.Value(i, COLS.tipo)
                .setMETODO_ID = x.Value(i, COLS.ID_METODO)
                .setFECHA_INICIAL = Format(x.Value(i, COLS.finicial), "yyyy-mm-dd")
                If IsEmpty(x.Value(i, COLS.fAprobacion)) Or (x.Value(i, COLS.fAprobacion) = "") Then
                    .setFECHA_APROBACION = "9999-12-31"
                Else
                    .setFECHA_APROBACION = Format(x.Value(i, COLS.fAprobacion), "yyyy-mm-dd")
                End If
                .setOBSERVACIONES = x.Value(i, COLS.OBSERVACIONES)
                .Insertar
            End With
        End If
    Next
    Set oPS = Nothing

   On Error GoTo 0
   Exit Sub

insertar_servicios_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_servicios of Formulario frmProveedores_Calidad"
End Sub
Private Sub cargar_combos()
    Dim rs As ADODB.Recordset
    ' Servicios
    Dim oDeco As New clsDecodificadora
    Set rs = oDeco.Listado(DECODIFICADORA.PROVEEDORES_SERVICIOS)
    If rs.RecordCount > 0 Then
        xServicios.ReDim 1, rs.RecordCount, 1, 2
        Dim i As Integer
        i = 1
        Do
            xServicios(i, 1) = CStr(rs("DESCRIPCION"))
            xServicios(i, 2) = CStr(rs("VALOR"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xServicios.ReDim 1, 1, 1, 2
    End If
    Set tServicios.Array = xServicios
    tServicios.Refresh
    ' Metodos
    Set rs = oDeco.Listado(DECODIFICADORA.PROVEEDORES_METODOS)
    If rs.RecordCount > 0 Then
        xMetodos.ReDim 1, rs.RecordCount, 1, 2
        i = 1
        Do
            xMetodos(i, 1) = CStr(rs("DESCRIPCION"))
            xMetodos(i, 2) = CStr(rs("VALOR"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xMetodos.ReDim 1, 1, 1, 2
    End If
    Set tMetodos.Array = xMetodos
    tMetodos.Refresh
End Sub
Private Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub tServicios_DropDownClose()
    grid.Columns(COLS.ID_SERVICIO) = tServicios.Columns(1)
    grid.Col = COLS.servicio + 1
End Sub

Private Sub tMetodos_DropDownClose()
    grid.Columns(COLS.ID_METODO) = tMetodos.Columns(1)
    grid.Col = COLS.METODO + 1
End Sub
Private Sub cargar_servicios()
    Dim oPS As New clsProveedores_servicios
    Dim rs As ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set rs = oPS.Listado(PK)
    If rs.RecordCount > 0 Then
        Dim i As Integer
        i = 0
        Do
            oDeco.Carga_valor DECODIFICADORA.PROVEEDORES_SERVICIOS, rs("SERVICIO_ID")
            x(i, COLS.servicio) = CStr(oDeco.getDESCRIPCION)
            x(i, COLS.tipo) = CStr(rs("TIPO"))
            oDeco.Carga_valor DECODIFICADORA.PROVEEDORES_METODOS, rs("METODO_ID")
            x(i, COLS.METODO) = CStr(oDeco.getDESCRIPCION)
            x(i, COLS.finicial) = CStr(Format(rs("FECHA_INICIAL"), "dd-mm-yyyy"))
            If Format(rs("FECHA_APROBACION"), "yyyy-mm-dd") <> "9999-12-31" Then
                x(i, COLS.fAprobacion) = CStr(Format(rs("FECHA_APROBACION"), "dd-mm-yyyy"))
            End If
            x(i, COLS.OBSERVACIONES) = CStr(rs("OBSERVACIONES"))
            x(i, COLS.ID_SERVICIO) = CStr(rs("SERVICIO_ID"))
            x(i, COLS.ID_METODO) = CStr(rs("METODO_ID"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    evaluacionInicial_NumServicios
    Set rs = Nothing
    Set oDeco = Nothing
End Sub
Private Sub evaluacionInicial_NumServicios()
    Dim cont As Integer
    cont = 0
    Dim i As Integer
    For i = 0 To filas
        If Trim(x.Value(i, COLS.servicio)) <> "" Then
            cont = cont + 1
        End If
    Next
    lblServicios.Caption = "Número de Servicios : " & cont
End Sub
