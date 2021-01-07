VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmEquipoEdicionCalibracion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10770
   ClientLeft      =   2790
   ClientTop       =   2010
   ClientWidth     =   14355
   ClipControls    =   0   'False
   Icon            =   "frmEquipoEdicionCalibracion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMuestra 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Muestra Asociada"
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
      Height          =   1050
      Left            =   2700
      TabIndex        =   70
      Top             =   9675
      Width           =   5325
      Begin VB.CheckBox chkAjuste 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AJUSTE"
         Height          =   285
         Left            =   3600
         TabIndex        =   77
         Top             =   585
         Width           =   1365
      End
      Begin VB.CheckBox chkURGENTE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "URGENTE"
         Height          =   285
         Left            =   3600
         TabIndex        =   76
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtMuestra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         MaxLength       =   255
         TabIndex        =   72
         Top             =   450
         Width           =   1845
      End
      Begin VB.CommandButton cmdMuestraConsulta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consulta"
         Height          =   765
         Left            =   2160
         Picture         =   "frmEquipoEdicionCalibracion.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Generar etiqueta"
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRevisiones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Revisiones"
      Height          =   900
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   9810
      Width           =   1530
   End
   Begin VB.CommandButton cmdCrearEtiquetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta Solicitud"
      Height          =   870
      Left            =   10395
      Picture         =   "frmEquipoEdicionCalibracion.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   9825
      Width           =   1710
   End
   Begin VB.CommandButton cmdCrearDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documento Solicitud"
      Height          =   870
      Left            =   8550
      Picture         =   "frmEquipoEdicionCalibracion.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   9825
      Width           =   1800
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   900
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9810
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1500
      Top             =   9885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13260
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9825
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   9135
      Left            =   45
      TabIndex        =   14
      Top             =   495
      Width           =   14295
      Begin VB.TextBox txtIncertidumbre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   9090
         MaxLength       =   255
         TabIndex        =   8
         Top             =   5760
         Width           =   4995
      End
      Begin VB.CheckBox chkEtiquetado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ETIQUETADO"
         Height          =   285
         Left            =   11700
         TabIndex        =   74
         Top             =   4185
         Width           =   1365
      End
      Begin VB.CheckBox chkSegregado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SEGREGADO"
         Height          =   285
         Left            =   11700
         TabIndex        =   73
         Top             =   3915
         Width           =   1365
      End
      Begin VB.TextBox txtCalibradoEn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   9090
         MaxLength       =   255
         TabIndex        =   7
         Top             =   5400
         Width           =   4995
      End
      Begin VB.Frame frmResultado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resultado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   9585
         TabIndex        =   63
         Top             =   3870
         Width           =   1950
         Begin VB.OptionButton optResultado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "REQ. AJUSTE"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   66
            Top             =   1065
            Width           =   1560
         End
         Begin VB.OptionButton optResultado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO CONFORME"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   65
            Top             =   705
            Width           =   1650
         End
         Begin VB.OptionButton optResultado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CONFORME"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   64
            Top             =   315
            Value           =   -1  'True
            Width           =   1410
         End
      End
      Begin VB.Frame fraEstadoIntervencion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   8055
         TabIndex        =   35
         Top             =   3870
         Width           =   1500
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Anulada"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   60
            Top             =   1125
            Width           =   1140
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prevista"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   38
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Realizada"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   37
            Top             =   570
            Width           =   1020
         End
         Begin VB.OptionButton optEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Revisada"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   36
            Top             =   840
            Width           =   1065
         End
      End
      Begin VB.TextBox txtFechaProxima 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "01/01/1900"
         Top             =   600
         Width           =   1785
      End
      Begin VB.CommandButton cmdAnadirReactivo2 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   13515
         Picture         =   "frmEquipoEdicionCalibracion.frx":169C
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Añadir"
         Top             =   450
         Width           =   285
      End
      Begin VB.CommandButton cmdEliminarReactivo2 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   13845
         Picture         =   "frmEquipoEdicionCalibracion.frx":18C1
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Eliminar"
         Top             =   450
         Width           =   285
      End
      Begin VB.TextBox txtAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   3
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   50
         Top             =   3405
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   7305
         Picture         =   "frmEquipoEdicionCalibracion.frx":1A55
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Ver Evaluación"
         Top             =   3390
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarP 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   6525
         Picture         =   "frmEquipoEdicionCalibracion.frx":1CAA
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Buscar Evaluación"
         Top             =   3390
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   6930
         Picture         =   "frmEquipoEdicionCalibracion.frx":1F1B
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Eliminar Evaluación"
         Top             =   3390
         Width           =   360
      End
      Begin VB.CommandButton cmdEscanearAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   3
         Left            =   6930
         Picture         =   "frmEquipoEdicionCalibracion.frx":20AF
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Escanear documento"
         Top             =   4695
         Visible         =   0   'False
         Width           =   405
      End
      Begin TrueDBGrid80.TDBDropDown tUnidades 
         Height          =   2280
         Left            =   6975
         TabIndex        =   45
         Top             =   6165
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   4022
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=265"
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
      Begin VB.CommandButton cmdEscanearAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   6930
         Picture         =   "frmEquipoEdicionCalibracion.frx":2469
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Escanear documento"
         Top             =   4335
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdEscanearAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   6930
         Picture         =   "frmEquipoEdicionCalibracion.frx":2823
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Escanear documento"
         Top             =   3975
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6120
         Picture         =   "frmEquipoEdicionCalibracion.frx":2BDD
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Eliminar"
         Top             =   3780
         Width           =   285
      End
      Begin VB.CommandButton cmdAnadirLimitacion 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   5790
         Picture         =   "frmEquipoEdicionCalibracion.frx":2D71
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Añadir"
         Top             =   3780
         Width           =   285
      End
      Begin VB.TextBox txtLimitacionesUso 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1650
         MaxLength       =   100
         TabIndex        =   5
         Top             =   3750
         Width           =   4095
      End
      Begin VB.ListBox lstLimitacionesUso 
         Appearance      =   0  'Flat
         Height          =   1395
         ItemData        =   "frmEquipoEdicionCalibracion.frx":2F96
         Left            =   1650
         List            =   "frmEquipoEdicionCalibracion.frx":2F9D
         TabIndex        =   6
         Top             =   4065
         Width           =   4755
      End
      Begin VB.ListBox lstAccesorios 
         Appearance      =   0  'Flat
         Height          =   1380
         Left            =   8070
         Style           =   1  'Checkbox
         TabIndex        =   31
         Top             =   2415
         Width           =   6090
      End
      Begin VB.TextBox txtAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   1
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   30
         Top             =   2685
         Width           =   4770
      End
      Begin VB.CommandButton cmdMostrarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   7305
         Picture         =   "frmEquipoEdicionCalibracion.frx":2FB5
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Ver Hoja de Calibración"
         Top             =   2670
         Width           =   405
      End
      Begin VB.CommandButton cmdAdjuntarP 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   6525
         Picture         =   "frmEquipoEdicionCalibracion.frx":320A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Buscar Hoja de Calibración"
         Top             =   2670
         Width           =   405
      End
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   6930
         Picture         =   "frmEquipoEdicionCalibracion.frx":347B
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Eliminar Certificado"
         Top             =   3030
         Width           =   360
      End
      Begin VB.CommandButton cmdAdjuntarP 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   6525
         Picture         =   "frmEquipoEdicionCalibracion.frx":360F
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Buscar Certificado"
         Top             =   3030
         Width           =   405
      End
      Begin VB.CommandButton cmdMostrarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   2
         Left            =   7305
         Picture         =   "frmEquipoEdicionCalibracion.frx":3880
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Ver Certificado"
         Top             =   3030
         Width           =   405
      End
      Begin VB.TextBox txtAdjunto 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   2
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   23
         Top             =   3045
         Width           =   4770
      End
      Begin MSComCtl2.DTPicker txtFechaActual 
         Height          =   405
         Left            =   1650
         TabIndex        =   10
         Top             =   150
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51380225
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFechaProxima_b 
         Height          =   405
         Left            =   3450
         TabIndex        =   11
         Top             =   570
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51380225
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipoCalibracion 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   1020
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPeriCalibracion 
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   1350
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbCalibrador 
         Height          =   330
         Left            =   1650
         TabIndex        =   2
         Top             =   1680
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbProcedimiento 
         Height          =   330
         Left            =   1650
         TabIndex        =   4
         Top             =   2340
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCalibradorExterno 
         Height          =   330
         Left            =   1650
         TabIndex        =   3
         Top             =   2010
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdEliminarAdjunto 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Index           =   1
         Left            =   6930
         Picture         =   "frmEquipoEdicionCalibracion.frx":3AD5
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Eliminar Hoja de Calibración"
         Top             =   2670
         Width           =   360
      End
      Begin TrueDBGrid80.TDBGrid grdResultados 
         Height          =   2925
         Left            =   30
         TabIndex        =   44
         Top             =   6120
         Width           =   14085
         _ExtentX        =   24844
         _ExtentY        =   5159
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Descripción"
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "Standard"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Rango Min."
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "General Number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Rango Max."
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "General Number"
         Columns(2).DropDown=   "tResponsables"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   1
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Unidad"
         Columns(3).DataField=   ""
         Columns(3).DropDown=   "tUnidades"
         Columns(3).DropDown.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Resultado"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "General Number"
         Columns(4).ExternalEditor=   "TDBDate1"
         Columns(4).ExternalEditor.vt=   8
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Tolerancia"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "General Number"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Incertidumbre"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "General Number"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Corrección"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "General Number"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "ID_RESULTADO"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "ID_UNIDAD"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=8361"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=8281"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1931"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1852"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2540"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2461"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(25)=   "Column(3).Button=1"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(3).AutoDropDown=1"
         Splits(0)._ColumnProps(28)=   "Column(4).Width=1931"
         Splits(0)._ColumnProps(29)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._WidthInPix=1852"
         Splits(0)._ColumnProps(31)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(34)=   "Column(5).Width=1931"
         Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=1852"
         Splits(0)._ColumnProps(37)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(40)=   "Column(6).Width=1931"
         Splits(0)._ColumnProps(41)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(6)._WidthInPix=1852"
         Splits(0)._ColumnProps(43)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(45)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(46)=   "Column(7).Width=1931"
         Splits(0)._ColumnProps(47)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._WidthInPix=1852"
         Splits(0)._ColumnProps(49)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(52)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(53)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(54)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(55)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(56)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(57)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(58)=   "Column(9).Width=2566"
         Splits(0)._ColumnProps(59)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(60)=   "Column(9)._WidthInPix=2487"
         Splits(0)._ColumnProps(61)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(62)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(63)=   "Column(9).Order=10"
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
         Caption         =   "Resultados Calibracion"
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
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=28,.parent=11,.alignment=2,.bgcolor=&HC1FFFF&"
         _StyleDefs(57)  =   ":id=28,.locked=0"
         _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=12"
         _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=13"
         _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=58,.parent=11,.alignment=2"
         _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=12"
         _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=13"
         _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=62,.parent=11,.alignment=2,.bgcolor=&HDEEDFA&"
         _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=12"
         _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=13"
         _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=66,.parent=11,.alignment=2,.locked=0"
         _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=12"
         _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=13"
         _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=74,.parent=11"
         _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=12"
         _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=13"
         _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=70,.parent=11"
         _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=12"
         _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=13"
         _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=15"
         _StyleDefs(81)  =   "Named:id=37:Normal"
         _StyleDefs(82)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
         _StyleDefs(83)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(84)  =   ":id=37,.fontname=MS Sans Serif"
         _StyleDefs(85)  =   "Named:id=38:Heading"
         _StyleDefs(86)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
         _StyleDefs(87)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(88)  =   ":id=38,.strikethrough=0,.charset=0"
         _StyleDefs(89)  =   ":id=38,.fontname=MS Sans Serif"
         _StyleDefs(90)  =   "Named:id=39:Footing"
         _StyleDefs(91)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(92)  =   "Named:id=40:Selected"
         _StyleDefs(93)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(94)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(95)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(96)  =   "Named:id=41:Caption"
         _StyleDefs(97)  =   ":id=41,.parent=38,.alignment=2"
         _StyleDefs(98)  =   "Named:id=42:HighlightRow"
         _StyleDefs(99)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
         _StyleDefs(100) =   "Named:id=43:EvenRow"
         _StyleDefs(101) =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
         _StyleDefs(102) =   "Named:id=44:OddRow"
         _StyleDefs(103) =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
         _StyleDefs(104) =   "Named:id=47:RecordSelector"
         _StyleDefs(105) =   ":id=47,.parent=38"
         _StyleDefs(106) =   "Named:id=50:FilterBar"
         _StyleDefs(107) =   ":id=50,.parent=37"
      End
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   330
         Left            =   8070
         TabIndex        =   53
         Top             =   450
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   582
      End
      Begin MSComctlLib.ListView listaReactivos 
         Height          =   1380
         Left            =   8070
         TabIndex        =   56
         Top             =   780
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   2434
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
      Begin XtremeSuiteControls.PushButton cmdTendencia 
         Height          =   840
         Left            =   11700
         TabIndex        =   61
         Top             =   4545
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   1482
         _StockProps     =   79
         Caption         =   "Tendencia"
         Appearance      =   5
         Picture         =   "frmEquipoEdicionCalibracion.frx":3C69
      End
      Begin pryCombo.miCombo cmbUBICACION_ID 
         Height          =   330
         Left            =   1665
         TabIndex        =   69
         Top             =   5490
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incertidumbre"
         Height          =   195
         Index           =   15
         Left            =   8055
         TabIndex        =   75
         Top             =   5805
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incidencia"
         Height          =   195
         Index           =   14
         Left            =   8055
         TabIndex        =   68
         Top             =   5445
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calibrado En"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   67
         Top             =   5550
         Width           =   900
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivos Utilizados"
         Height          =   225
         Index           =   12
         Left            =   8070
         TabIndex        =   52
         Top             =   210
         Width           =   1860
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Eval. Resultado"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   51
         Top             =   3465
         Width           =   1125
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calibrador Externo"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   41
         Top             =   2115
         Width           =   1290
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Limitaciones uso"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   34
         Top             =   3825
         Width           =   1200
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hoja de Calibración"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   33
         Top             =   2760
         Width           =   1380
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accesorios de Equipo Utilizados"
         Height          =   225
         Index           =   7
         Left            =   8070
         TabIndex        =   32
         Top             =   2190
         Width           =   2670
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cert. de calibración"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   3105
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodo Calibración"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   1410
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Próx. Calibración"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Actual Calibración"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Cal. Interna"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   18
         Top             =   1770
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Top             =   2430
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Top             =   1080
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9825
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Calibración de Equipo"
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
      TabIndex        =   17
      Top             =   120
      Width           =   2325
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   14495
   End
End
Attribute VB_Name = "frmEquipoEdicionCalibracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarlngPK As Long
Public booSilencioso As Boolean
Private mvarobjEquipo As clsEquipos
Private mvarenuTipoEdicion As enumTipoEdicion
Private mvarstrId As String

Private bln_fecha_real_editable As Boolean

Private mvarobjCalibracion As New clsEquipoCalibracion
Private mvarblnResultado As Boolean
Private mvardtmFechaProximaInicial As Date
Private mvarlngidCalibradorInternoInicial As Long
Private mvarlngIdPeriodoInicial As Long
Private mvarlngIdTipoCalibracionIncial As Long
Private mvarblnVieneDeCuaderno As Boolean

Private mvarlngidEquipo As Long
Private mvardtmFechaPrevista As Date
Private mvarblnPresentandoDatos As Boolean
Private mvarlngIdEvento As Long

Private xR As New XArrayDB
Private xUnidades As New XArrayDB

Const filasR As Integer = 50
Const ColR As Integer = 9
Private Enum ColsR
    DESCRIPCION = 0
    RANGO_MIN = 1
    RANGO_MAX = 2
    Unidad = 3
    RESULTADO_CAL = 4
    TOLERANCIA = 5
    INCERTIDUMBRE = 6
    CORRECCION = 7
    Id_resultado = 8
    id_unidad = 9
End Enum

Private mvarlngNumParametrosResultados As Long
Private mvarlngidProcedmientoInicial As Long

Private Sub chkAjuste_Click()
   On Error GoTo chkAjuste_Click_Error

    If txtMuestra = "" Then Exit Sub
    Dim oEQC As New clsEquipoCalibracion
    oEQC.Carga CLng(mvarstrId)
    Dim oMuestra As New clsMuestra
    oMuestra.informar_ajuste oEQC.getMUESTRA_ID, chkAjuste.Value
    Set oMuestra = Nothing

   On Error GoTo 0
   Exit Sub

chkAjuste_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkAjuste_Click of Formulario frmEquipoEdicionCalibracion"
End Sub

Private Sub chkURGENTE_Click()
   On Error GoTo chkURGENTE_Click_Error

    If txtMuestra = "" Then Exit Sub
    Dim oEQC As New clsEquipoCalibracion
    oEQC.Carga CLng(mvarstrId)
    Dim oMuestra As New clsMuestra
    oMuestra.informar_urgente oEQC.getMUESTRA_ID, chkURGENTE.Value
    Set oMuestra = Nothing

   On Error GoTo 0
   Exit Sub

chkURGENTE_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkURGENTE_Click of Formulario frmEquipoEdicionCalibracion"
End Sub

Private Sub cmdMuestraConsulta_Click()
   On Error GoTo cmdMuestraConsulta_Click_Error

    If txtMuestra = "" Then Exit Sub
    Dim oEQC As New clsEquipoCalibracion
    oEQC.Carga CLng(mvarstrId)
    gmuestra = oEQC.getMUESTRA_ID
    frmVerMuestra.Show 1

   On Error GoTo 0
   Exit Sub

cmdMuestraConsulta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMuestraConsulta_Click of Formulario frmEquipoEdicionCalibracion"
End Sub

Private Sub cmdAdjuntarP_Click(Index As Integer)
   On Error GoTo cmdAdjuntarP_Click_Error
    If mvarenuTipoEdicion = Alta Then
        MsgBox "Guarde primero la calibración para poder asignar documentos.", vbCritical, App.Title
        Exit Sub
    End If

    cd.ShowOpen
    If Trim(cd.FileName) = "" Then Exit Sub
    Dim oD As New clsDocumentacion
    Dim salida As String
    salida = oD.SubirEquipo(mvarlngidEquipo, 0, CLng(mvarstrId), Index, cd.FileName, cd.FileTitle)
    If salida <> "" Then
        MsgBox "Se ha producido un error al subir el documento : " & salida, vbCritical, App.Title
    Else
        txtAdjunto(Index) = cd.FileTitle
        Dim c As String
        c = "UPDATE eq_calibracion_equipos " & _
           "   set ruta_plantilla = '" & txtAdjunto(1) & "'" & _
           "      ,ruta_certificado = '" & txtAdjunto(2) & "'" & _
           "      ,ruta_evaluacion = '" & txtAdjunto(3) & "'" & _
           " where id_calibracion = " & CLng(mvarstrId)
        execute_bd c
    End If

   On Error GoTo 0
   Exit Sub

cmdAdjuntarP_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntarP_Click of Formulario frmEquipoEdicionCalibracion"

End Sub

Private Sub cmdAnadirReactivo2_Click()
    If cmbReactivos.getPK_SALIDA <> 0 Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex, x As Long
        Dim oTR As New clsTipos_reactivo_ex
        oBote.CARGAR cmbReactivos.getPK_SALIDA
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        
        If listaReactivos.ListItems.Count > 0 Then
            For x = 1 To listaReactivos.ListItems.Count
                If CLng(listaReactivos.ListItems(x)) = oBote.getID_BOTE_EX Then
                    Exit Sub
                End If
            Next x
            
        End If
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
        cmbReactivos.limpiar
    End If

End Sub

Private Sub cmdEliminarAdjunto_Click(Index As Integer)
    Dim oD As New clsDocumentacion
   On Error GoTo cmdEliminarAdjunto_Click_Error

    If oD.EliminarEquipo(mvarlngidEquipo, 0, CLng(mvarstrId), Index) = "" Then
        
        Dim c As String
        Dim s As String
        Select Case Index
        Case 1
            s = " ruta_plantilla = '' "
        Case 2
            s = " ruta_certificado = '' "
        Case 3
            s = " ruta_evaluacion = '' "
        End Select
        c = "UPDATE eq_calibracion_equipos set " & _
           s & _
           " where id_calibracion = " & CLng(mvarstrId)
        execute_bd c
        txtAdjunto(Index) = ""
    End If
    Set oD = Nothing

   On Error GoTo 0
   Exit Sub

cmdEliminarAdjunto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarAdjunto_Click of Formulario frmEquipoEdicionCalibracion"
End Sub

Private Sub cmdEliminarReactivo2_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
    End If
End Sub

Private Sub cmdMostrarAdjunto_Click(Index As Integer)
    Dim oD As New clsDocumentacion
   On Error GoTo cmdMostrarAdjunto_Click_Error

    oD.CargarEquipo mvarlngidEquipo, 0, CLng(mvarstrId), Index, True
    Set oD = Nothing

   On Error GoTo 0
   Exit Sub

cmdMostrarAdjunto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrarAdjunto_Click of Formulario frmEquipoEdicionCalibracion"
End Sub

Private Sub cmdRevisiones_Click()
    With frmRevisiones
        .TOBJETO = TOBJETO_REV_EQ_CALIBRACION
        .COBJETO = mvarlngidEquipo
        .Show 1
    End With
End Sub

Private Sub cmdTendencia_Click()
    If xR(0, 0) <> "" Then
        frmEquiposTendencias.PK_EQUIPO_ID = mvarlngidEquipo
        frmEquiposTendencias.PK_PERIODICIDAD = cmbPeriCalibracion.BoundText
        grdResultados.Col = ColsR.DESCRIPCION
        frmEquiposTendencias.PK_PARAMETRO = grdResultados.Text
        grdResultados.Col = ColsR.RANGO_MIN
        If IsNumeric(grdResultados.Text) Then
            frmEquiposTendencias.PK_RANGO_MIN = grdResultados.Text
        Else
            frmEquiposTendencias.PK_RANGO_MIN = "0"
        End If
        grdResultados.Col = ColsR.RANGO_MIN
        If IsNumeric(grdResultados.Text) Then
            frmEquiposTendencias.PK_RANGO_MAX = grdResultados.Text
        Else
            frmEquiposTendencias.PK_RANGO_MAX = "0"
        End If
        frmEquiposTendencias.PK_TIPO = 1 ' Calibracion
        frmEquiposTendencias.Show 1
    End If
End Sub


Private Sub cmdCrearDocumentos_Click()
    On Error GoTo fallo
    Me.MousePointer = vbHourglass
    With frmReport
        .iniciar
        .informe = "\Equipos\rptEquipos_SC_Calibracion_Informe"
        .criterio = "{eq_calibracion_equipos.ID_CALIBRACION} = " & mvarstrId
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
    Me.MousePointer = vbNormal
    Exit Sub
fallo:
    MsgBox "Error al generar el documento." & Err.Description, vbCritical, App.Title
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdCrearEtiquetas_Click()
    Me.MousePointer = vbHourglass
    With frmReport
        .iniciar
        .informe = "\Equipos\rptEquipos_SC_Calibracion_Etiqueta"
        .criterio = "{proveedores.ID_PROVEEDOR} = " & cmbCalibradorExterno.getPK_SALIDA
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
    Me.MousePointer = vbNormal
End Sub

Private Sub comprobar_fecha_real_modificable()
    Dim op As New clsParametros
    bln_fecha_real_editable = False
    
    If op.Carga(parametros.MODIFICACION_FECHAS_CALIBRACION_VERIFICACION, "") Then
        If op.getVALOR = "1" Then
            bln_fecha_real_editable = True
        End If
    End If
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
        bln_fecha_real_editable = False
    End If
End Sub
Private Sub CargarComboGridUnidad()
    Dim rs As ADODB.Recordset
    Dim ote As New clsUnidades

    Set rs = ote.Listado("")
    xUnidades.Clear
    If rs.RecordCount > 0 Then
        xUnidades.ReDim 0, rs.RecordCount, 0, 1
        Dim i As Integer
        i = 1
        Do
            xUnidades(i, 0) = CStr(rs("NOMBRE"))
            xUnidades(i, 1) = CStr(rs("ID_UNIDAD"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xUnidades.ReDim 0, 0, 0, 1
    End If
    Set tUnidades.Array = xUnidades
    tUnidades.Refresh
    grdResultados.Refresh
End Sub

Private Function PresentarDatos_Reactivos()
        
    Dim arrRex() As String, x As Integer
    listaReactivos.ListItems.Clear
    
    If Trim(mvarobjCalibracion.getREACTIVOS) <> "" Then
        arrRex() = Split(mvarobjCalibracion.getREACTIVOS, ",")
    Else
        Exit Function
    End If
    
    Dim oBote As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    For x = 0 To UBound(arrRex())
        oBote.CARGAR CLng(arrRex(x))
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
        End With
    Next x
    
    listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    
End Function

Private Sub cmbPeriCalibracion_Click(area As Integer)

    Call txtFechaActual_Change

End Sub


Private Sub cmbTipoCalibracion_Change()
    If cmbTipoCalibracion.BoundText = "1" Then ' Intera
        ' Es interna
        cmbCalibradorExterno.desactivar
    Else
        ' Es externa
        cmbCalibradorExterno.activar
    End If
End Sub

Private Sub cmdAnadirLimitacion_Click()
    mvarobjEquipo.Anadir_limitacionuso_equipo txtLimitacionesUso.Text
    Call PresentarDatos_LimitacionesUso
End Sub

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide
End Sub

Private Sub cmdEliminarLimitacion_Click()
    Dim lngid As Long
    If lstLimitacionesUso.ListIndex < 0 Then Exit Sub
    lngid = lstLimitacionesUso.ItemData(lstLimitacionesUso.ListIndex)
    mvarobjEquipo.Eliminar_LimitacionUso_equipo lngid
    Call PresentarDatos_LimitacionesUso
End Sub

Private Sub cmdetiqueta_Click()
    If optEstado(CAL_ESTADOS.CAL_ESTADO_REALIZADA).Value = True Or optEstado(CAL_ESTADOS.CAL_ESTADO_REVISADA).Value = True Then
        Dim oEC As New clsEquipoCalibracion
        oEC.imprimir_etiqueta CLng(mvarstrId)
        Set oEC = Nothing
    Else
        MsgBox "La calibración tiene que estar cerrada para poder generar la etiqueta.", vbExclamation, App.Title
    End If
End Sub
Private Sub cmdok_Click()
    ' Recoge los datos
    Dim lngId_Calibracion As Long
    If Not ComprobarDatos Then Exit Sub
    
    RecogerDatos
        
    ' VERIFICAR SI EL EMPLEADO ESTA CUALIFICADO. SI NO LO ESTA, MENSAJE Y CORREO A RRHH
    If UCase(USUARIO.getUSUARIO) <> "JULIO" Then
        If optEstado(1) = True And cmbProcedimiento.getPK_SALIDA <> 0 Then
            Dim oEmpleados_cualificaciones As New clsEmpleados_cualificaciones
            If oEmpleados_cualificaciones.estaCualificadoPNT(USUARIO.getID_EMPLEADO, cmbProcedimiento.getPK_SALIDA) = False Then
                MsgBox "ATENCIÓN : No esta CUALIFICADO para realizar esta calibración. Contacte con RRHH, procedimiento : " & cmbProcedimiento.getTEXTO, vbCritical + vbOKOnly, App.Title
                oEmpleados_cualificaciones.enviarCorreoNoCualificado cmbProcedimiento.getPK_SALIDA, "Calibración Equipo Nº" & mvarlngidEquipo
            End If
            Set oEmpleados_cualificaciones = Nothing
        End If
    End If
    If mvarenuTipoEdicion = Alta Then
        mvarobjCalibracion.setEQUIPO_ID = mvarlngidEquipo
        lngId_Calibracion = mvarobjCalibracion.Insertar(True, xR, filasR)
    Else
        lngId_Calibracion = CLng(mvarstrId)
        Call mvarobjCalibracion.Modificar(lngId_Calibracion, True, , xR, filasR)
    End If
    mvarblnResultado = True
    Me.Hide
End Sub

Private Function ComprobarDatos() As Boolean
    Dim strMs As String
    On Error GoTo ComprobarDatos_Error
    ComprobarDatos = False
    strMs = ""

    If Not optEstado(CAL_ESTADOS.CAL_ESTADO_PREVISTA).Value Then
        comprobarDatosResultados strMs
    End If
    
    If Trim(cmbTipoCalibracion.BoundText) = "" Or Trim(cmbTipoCalibracion.BoundText) = "0" Then
        strMs = strMs & vbCrLf & " - Debe indicar el Tipo de Calibración"
    End If

    If cmbCalibrador.getPK_SALIDA = 0 Then
        strMs = strMs & vbCrLf & " - Debe indicar el Responsable Interno de Calibración"
    End If

    If Trim(cmbPeriCalibracion.BoundText) = "" Then
        strMs = strMs & vbCrLf & " - Debe indicar el Periodo para las Calibraciones"
    End If
    
    
    If getDataComboSel(cmbTipoCalibracion) = 1 Then
        If Trim(cmbProcedimiento.getPK_SALIDA) < 1 Then
            strMs = strMs & vbCrLf & " - Debe indicar el Procedimiento"
        End If
    ElseIf getDataComboSel(cmbTipoCalibracion) = 2 Then
        If Trim(cmbCalibradorExterno.getPK_SALIDA) < 1 Then
            strMs = strMs & vbCrLf & " - Debe indicar el Calibrador Externo"
        End If
    End If
    
    If CDate("01/01/1900") = txtFechaActual.Value Then
        strMs = strMs & vbCrLf & " - Debe indicar una Fecha Actual de Calibración adecuada"
    End If
    
    ' la fecha proxima no es modificable
    'If txtFechaActual.value >= txtFechaProxima_b.value Then
    '    strMs = strMs & vbCrLf & " - La fecha de la próxima calibración no puede ser anterior a la de la calibración actual"
    'End If
    
    If Trim(strMs) <> "" Then
        MsgBox "Se han detectado los siguientes errores: " & strMs
        Exit Function
    End If

    ComprobarDatos = True

On Error GoTo 0
    Exit Function
ComprobarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ComprobarDatos of Formulario frmEquipoEdicionCalibracion"
End Function

Private Sub comprobarDatosResultados(ByRef strMs As String)

    Dim i As Long
    i = 0
    Dim cad As String
    
    cad = ""
    grdResultados.Refresh
    
    'For i = 0 To filasR
    '    cad = cad & xR(i, 0) & ", " & xR(i, 1) & ", " & xR(i, 2) & ", " & xR(i, 3) & ", " & xR(i, 4) & ", " & xR(i, 5) & ", " & xR(i, 6) & ", " & xR(i, 7) & ". " & vbCrLf
    'Next i
    
End Sub

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Public Property Get FechaPrevista() As Date

    FechaPrevista = mvardtmFechaPrevista

End Property

Public Property Let FechaPrevista(ByVal dtmFechaPrevista As Date)

    mvardtmFechaPrevista = dtmFechaPrevista

End Property

Public Property Get FechaProximaInicial() As Date

    FechaProximaInicial = mvardtmFechaProximaInicial

End Property

Public Property Let FechaProximaInicial(ByVal dtmFechaProximaInicial As Date)

    mvardtmFechaProximaInicial = dtmFechaProximaInicial

End Property
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    
    comprobar_fecha_real_modificable
    
    mvarblnPresentandoDatos = True
    If mvarblnVieneDeCuaderno Then
        Set mvarobjEquipo = New clsEquipos
        Call mvarobjEquipo.Carga(mvarlngidEquipo)
        
            mvarenuTipoEdicion = EDICION
            mvarstrId = CStr(mvarlngIdEvento)
    End If
    
    mvarlngidEquipo = mvarobjEquipo.getID_EQUIPO
    
    Call PresentarDatos_LimitacionesUso
    
    Call Presentar_Accesorios
    
    Call LlenarCombos
    Call inicializar_grid
    Call CargarComboGridUnidad
    
    Call PresentarDatos_ParametrosResultados
   
    lbltitulo.Caption = "Calibración del Equipo : " & CStr(mvarobjEquipo.getID_EQUIPO) & ": " & mvarobjEquipo.getNOMBRE
    
    If mvarenuTipoEdicion = Alta Then
        txtFechaActual.Value = Now
        txtFechaActual.Enabled = bln_fecha_real_editable Or True
        Set mvarobjCalibracion.AccesoriosEquipo = mvarobjEquipo.getACCESORIOS_COL
        cmbTipoCalibracion.BoundText = mvarlngIdTipoCalibracionIncial
        cmbPeriCalibracion.BoundText = mvarlngIdPeriodoInicial
        txtFechaActual_Change
        cmbCalibrador.MostrarElemento mvarlngidCalibradorInternoInicial
        cmbProcedimiento.MostrarElemento mvarlngidProcedmientoInicial
        mvarblnPresentandoDatos = False
        Exit Sub
    End If
    
    mvarobjCalibracion.Carga CLng(mvarstrId)
    
    Call PresentarDatos
    
    Call PresentarDatos_Accesorios
    
    mvarblnPresentandoDatos = False
    
    Call OpcionesEdicion

    If cmbCalibradorExterno.getTEXTO = "" Then
        cmdCrearDocumentos.Enabled = False
        cmdCrearEtiquetas.Enabled = False
'        optEstado(3).Enabled = False
    Else
        cmdCrearDocumentos.Enabled = True
        cmdCrearEtiquetas.Enabled = True
'        optEstado(3).Enabled = True
    End If
    
    ' Si no esta pendiente, ocultamos icono Hoja de certificado
    If optEstado(0).Value = False Then
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
            cmdMostrarAdjunto(1).visible = False
        End If
    End If
    
End Sub

Private Sub OpcionesEdicion()

    If mvarenuTipoEdicion = Alta Then
        txtFechaActual.Enabled = True
    ElseIf mvarenuTipoEdicion = EDICION Then
        txtFechaActual.Enabled = bln_fecha_real_editable
    ElseIf mvarenuTipoEdicion = visualizar Then
        
        cmbTipoCalibracion.Enabled = False
        cmbPeriCalibracion.Enabled = False
        cmbCalibrador.desactivar
        cmbCalibradorExterno.desactivar
        cmbProcedimiento.desactivar
        cmdMostrarAdjunto(1).Left = cmdAdjuntarP(1).Left
        cmdAdjuntarP(1).visible = False
        cmdEscanearAdjunto(1).visible = False
        cmdEliminarAdjunto(1).visible = False
        cmdMostrarAdjunto(2).Left = cmdAdjuntarP(2).Left
        cmdAdjuntarP(2).visible = False
        cmdEscanearAdjunto(2).visible = False
        cmdEliminarAdjunto(2).visible = False
        cmdMostrarAdjunto(3).Left = cmdAdjuntarP(3).Left
        cmdAdjuntarP(3).visible = False
        cmdEscanearAdjunto(3).visible = False
        cmdEliminarAdjunto(3).visible = False
        txtLimitacionesUso.Locked = True
        cmdAnadirLimitacion.Enabled = False
        cmdEliminarLimitacion.Enabled = False
        lstLimitacionesUso.Enabled = False
        lstAccesorios.Enabled = False
        
        txtFechaActual.Enabled = False
        fraEstadoIntervencion.Enabled = False
        frmResultado.Enabled = False
        chkSegregado.Enabled = False
        chkEtiquetado.Enabled = False
        grdResultados.EditActive = False
        cmdok.visible = False
        
        cmdAnadirReactivo2.visible = False
        cmdEliminarReactivo2.visible = False
        
        chkURGENTE.Enabled = False
        chkAjuste.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub

Private Sub grdResultados_KeyPress(KeyAscii As Integer)
    
    With grdResultados
        If .Col = 1 Or .Col = 2 Or .Col = 4 Or .Col = 5 Or .Col = 6 Or .Col = 7 Then
            KeyAscii = KeyAscii_SoloDecimal_tbgrid(.Text, KeyAscii, True)
        End If
    End With
        
End Sub

Public Property Get ID() As String

    ID = mvarstrId

End Property

Public Property Let ID(ByVal strId As String)

    mvarstrId = strId

End Property

Public Property Get idCalibradorInternoInicial() As Long

    idCalibradorInternoInicial = mvarlngidCalibradorInternoInicial

End Property

Public Property Let idCalibradorInternoInicial(ByVal lngidCalibradorInternoInicial As Long)

    mvarlngidCalibradorInternoInicial = lngidCalibradorInternoInicial

End Property

Public Property Get idEquipo() As Long

    idEquipo = mvarlngidEquipo

End Property

Public Property Let idEquipo(ByVal lngidEquipo As Long)

    mvarlngidEquipo = lngidEquipo

End Property

Public Property Get IdEvento() As Long

    IdEvento = mvarlngIdEvento

End Property

Public Property Let IdEvento(ByVal lngIdEvento As Long)

    mvarlngIdEvento = lngIdEvento
    
End Property

Public Property Get IdPeriodoInicial() As Long

    IdPeriodoInicial = mvarlngIdPeriodoInicial

End Property

Public Property Let IdPeriodoInicial(ByVal lngIdPeriodoInicial As Long)

    mvarlngIdPeriodoInicial = lngIdPeriodoInicial

End Property

Public Property Get IdTipoCalibracionIncial() As Long

    IdTipoCalibracionIncial = mvarlngIdTipoCalibracionIncial

End Property

Public Property Let IdTipoCalibracionIncial(ByVal lngIdTipoCalibracionIncial As Long)

    mvarlngIdTipoCalibracionIncial = lngIdTipoCalibracionIncial

End Property
Private Sub inicializar_grid()
   On Error GoTo inicializar_grid_Error

    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
    End With
    
    
    grdResultados.Col = 0
    grdResultados.Row = 0
    
    xR.Clear
    xR.ReDim 0, filasR, 0, ColR
    xR.Clear
    
    Set grdResultados.Array = xR
    grdResultados.Refresh
    

On Error GoTo 0
Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmEquipoEdicionCalibracion"
End Sub

Private Sub lstAccesorios_ItemCheck(Item As Integer)
    Dim lngFila As Long
    Dim objItem As clsEquipoAccesorios
    Dim objCol As clsGenericCollection

    If mvarblnPresentandoDatos Then Exit Sub
    If Item < 0 Then Exit Sub
    If lstAccesorios.ListCount = 0 Then Exit Sub
    
    lngFila = Item
    
    Set objCol = mvarobjCalibracion.AccesoriosCalibracion
    Set objItem = objCol.Item(CStr(lstAccesorios.ItemData(lngFila)))
            
    If lstAccesorios.Selected(lngFila) Then
        ' lo ha señalado. Lo busca por si estuviera ya dentro
        If Not objItem Is Nothing Then
            ' si está eliminado, lo pone como existente
            objItem.setID_AUX = enumIdAux.ID_AUX_EXISTE
        Else
            ' no existe, lo inserta en la coleccion
            Set objItem = New clsEquipoAccesorios
            objItem.setNOMBRE = lstAccesorios.List(lngFila)
            objItem.setID_ACCESORIO = lstAccesorios.ItemData(lngFila)
            objItem.setEQUIPO_ID = mvarobjEquipo.getID_EQUIPO
            objCol.Add objItem, CStr(objItem.getID_ACCESORIO), enumIdAux.ID_AUX_NUEVO
        End If
    Else
        ' cuando lo deselecciona, no tiene más qe existir en la lista
        If Not objItem Is Nothing Then
            objCol.Remove (CStr(objItem.getID_ACCESORIO))
        Else
        End If
        
    End If

    
    Set mvarobjCalibracion.AccesoriosCalibracion = objCol
    
End Sub

Private Sub lstLimitacionesUso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdEliminarLimitacion_Click
End Sub

Private Sub LlenarCombos()
    Dim oDeco As New clsDecodificadora

    oDeco.cargar_combo cmbPeriCalibracion, DECODIFICADORA.EQ_periodicidad
    oDeco.cargar_combo cmbTipoCalibracion, DECODIFICADORA.EQ_TIPO_CALIBRACION
    
    llenar_combo cmbCalibrador, New clsUsuarios, 0, frmUsuarios, ""
    llenar_combo cmbProcedimiento, New clsCa_documentos, 0, frmCA_Documento, ""
    llenar_combo cmbCalibradorExterno, New clsProveedor, 0, Me, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, "AND ABIERTO = 1"
    
    oDeco.cargar_mi_combo cmbUBICACION_ID, DECODIFICADORA.EQ_UBICACION_ID
    
    If mvarobjEquipo.getTIPO_CALIBRACION_ID = 2 Then ' es Externa
        cmbCalibradorExterno.activar
    Else
        cmbCalibradorExterno.desactivar
    End If

    
End Sub
Public Property Get PK() As Long
    PK = mvarlngPK
End Property
Public Property Let PK(ByVal lngPK As Long)
    mvarlngPK = lngPK
End Property

Private Sub Presentar_Accesorios()
    Dim objCol As clsGenericCollection, objItem As clsEquipoAccesorios
    
    Set objCol = mvarobjEquipo.getACCESORIOS_COL
    
    If objCol Is Nothing Then Exit Sub
    
    For Each objItem In objCol.Iterator
        If objItem.getID_AUX <> enumIdAux.ID_AUX_ELIMINADO Then
            Call lstAccesorios.AddItem(objItem.getNOMBRE)
            lstAccesorios.ItemData(lstAccesorios.ListCount - 1) = objItem.getID_ACCESORIO
            
        End If
    Next objItem
End Sub

Private Sub PresentarDatos()
    On Error GoTo PresentarDatos_Error
    With mvarobjCalibracion
        cmbTipoCalibracion.BoundText = .getTIPO_ID
        cmbPeriCalibracion.BoundText = .getPERIODICIDAD_ID
        cmbCalibrador.MostrarElemento .getCALIBRADOR_INTERNO_ID
        If .getCALIBRADOR_EXTERNO_ID > 0 Then
            cmbCalibradorExterno.MostrarElemento .getCALIBRADOR_EXTERNO_ID
        End If
        cmbProcedimiento.MostrarElemento .getPROCEDIMIENTO_ID
        optEstado(.getESTADO).Value = True
        optResultado(.getRESULTADO).Value = True
        txtCalibradoEn = .getINCIDENCIAS
        txtIncertidumbre = .getINCERTIDUMBRE
        cmbUBICACION_ID.MostrarElemento .getUBICACION_ID
        chkSegregado.Value = .getSEGREGADO
        chkEtiquetado.Value = .getETIQUETADO
        If mvarenuTipoEdicion = Alta Then
            txtFechaActual.Value = Now
            cmbPeriCalibracion.BoundText = CStr(mvarlngIdPeriodoInicial)
            txtFechaActual_Change
            cmbTipoCalibracion.BoundText = CStr(mvarlngIdTipoCalibracionIncial)
        Else
            If IsDate(.getFECHA_ACTUAL) Then
                txtFechaActual.Value = CDate(.getFECHA_ACTUAL)
                txtFechaActual_Change
            End If
            txtAdjunto(1) = .getRUTA_PLANTILLA
            txtAdjunto(2) = .getRUTA_CERTIFICADO
            txtAdjunto(3) = .getRUTA_EVALUACION
        End If
        frmMuestra.Enabled = False
        If .getMUESTRA_ID <> 0 Then
            frmMuestra.Enabled = True
            Dim oMuestra As New clsMuestra
            oMuestra.CargaMuestra .getMUESTRA_ID
            txtMuestra = oMuestra.getID_GENERAL & " (" & oMuestra.CodigoParticular(.getMUESTRA_ID) & ")"
            chkURGENTE.Value = oMuestra.getURGENTE
            chkAjuste.Value = oMuestra.getAJUSTE
            Set oMuestra = Nothing
        End If
    End With
    Call PresentarDatos_Reactivos

On Error GoTo 0
    Exit Sub
PresentarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmEquipoEdicionCalibracion"

End Sub

Private Sub PresentarDatos_Accesorios()
    Dim objCol As clsGenericCollection, objItem As clsEquipoAccesorios
    Dim lngCont As Long

    mvarobjCalibracion.Carga_AccesoriosCalibracion

    Set objCol = mvarobjCalibracion.AccesoriosCalibracion

    For lngCont = 0 To lstAccesorios.ListCount - 1
        Set objItem = objCol.Item(CStr(lstAccesorios.ItemData(lngCont)))
        If Not objItem Is Nothing Then
            lstAccesorios.Selected(lngCont) = True
        End If
    Next
End Sub
Private Sub PresentarDatos_LimitacionesUso()
    Dim objItem As clsGenericClass

    lstLimitacionesUso.Clear
    txtLimitacionesUso.Text = ""
        
    For Each objItem In mvarobjEquipo.getLIMITACIONES_USO_COL.Iterator
        If objItem.getID_AUX <> enumIdAux.ID_AUX_ELIMINADO Then
            lstLimitacionesUso.AddItem objItem.getNOMBRE
            lstLimitacionesUso.ItemData(lstLimitacionesUso.ListCount - 1) = objItem.getID
        End If
    Next objItem

End Sub

Private Sub PresentarDatos_ParametrosResultados()

    Dim i As Integer
    Dim rs As ADODB.Recordset
On Error GoTo PresentarDatos_ParametrosResultados_Error

    i = 0
    
    
    If mvarenuTipoEdicion <> Alta Then
        ' Carga los Parametros de la calibracion
        Set rs = mvarobjCalibracion.DevolverParametrosResultados(mvarstrId)
    Else
        ' Carga los Parametros del Equipo
        Set rs = mvarobjEquipo.DevolverParametrosResultadosEquipoCalibracion(CStr(mvarlngidEquipo))
    End If
    
    If rs.RecordCount > 0 Then
        Do
            xR(i, ColsR.DESCRIPCION) = CStr(rs("descripcion"))
            xR(i, ColsR.RANGO_MIN) = CStr(rs("rango_min"))
            xR(i, ColsR.RANGO_MAX) = CStr(rs("rango_max"))
            xR(i, ColsR.Unidad) = CStr(rs("unidad"))
            xR(i, ColsR.id_unidad) = CStr(rs("unidad_ID"))
            
            xR(i, ColsR.RESULTADO_CAL) = CStr(rs("resultado"))
            xR(i, ColsR.TOLERANCIA) = CStr(rs("tolerancia_max"))
            xR(i, ColsR.INCERTIDUMBRE) = CStr(rs("incertidumbre"))
            xR(i, ColsR.CORRECCION) = CStr(rs("correccion"))
            xR(i, ColsR.Id_resultado) = CStr(rs("id_resultado"))
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    mvarlngNumParametrosResultados = i
    
    grdResultados.Refresh
    grdResultados.Enabled = True


On Error GoTo 0
    Exit Sub
PresentarDatos_ParametrosResultados_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_ParametrosResultados of Formulario frmEquipoEdicionCalibracion"

End Sub

Private Sub RecogerDatos()
    Dim Reactivo As String, i As Integer
    
    With mvarobjCalibracion
        ' A patir del 18.05.2010, las fechas no se podran modificar
        '.setFECHA_ACTUAL = Format(txtFechaActual.value, "dd/mm/yyyy")
        
        ' A patir del 02.09.2010, PROPUESTA
        ' Ahora que hay calibraciones previstas, la fecha se modifica siempre que sea prevista.
        ' cuando se cierra, siempre es el momento en que se cierra.
        ' de no ser así, el usuario (no es el caso de automaticamente al cerrar una calibracion, que se crea la siguiente prevista)
        ' no se podrían crear previstas más allá del presente
        
        ' La fecha la establece solo si se cierra ahora
        If .getESTADO = 0 Then
            .setFECHA_ACTUAL = Format(txtFechaActual.Value, "dd/mm/yyyy")
            .setFECHA_PROXIMA = Format(txtFechaProxima_b.Value, "dd/mm/yyyy")
        Else
            .setFECHA_ACTUAL = Format(Now, "dd/mm/yyyy")
            .setFECHA_PROXIMA = Format(txtFechaProxima_b.Value, "dd/mm/yyyy")
        End If
    
        .setTIPO_ID = CLng(cmbTipoCalibracion.BoundText)
        .setPERIODICIDAD_ID = CLng(cmbPeriCalibracion.BoundText)
        .setCALIBRADOR_INTERNO_ID = cmbCalibrador.getPK_SALIDA
        .setRESPONSABLE = cmbCalibrador.getTEXTO
        If .getTIPO_ID = 2 Then
            .setCALIBRADOR_EXTERNO_ID = cmbCalibradorExterno.getPK_SALIDA
        Else
            .setCALIBRADOR_EXTERNO_ID = -1
        End If
        .setUBICACION_ID = cmbUBICACION_ID.getPK_SALIDA
        .setSEGREGADO = chkSegregado.Value
        .setETIQUETADO = chkEtiquetado.Value
        .setINCIDENCIAS = txtCalibradoEn
        .setINCERTIDUMBRE = txtIncertidumbre
        .setPROCEDIMIENTO_ID = cmbProcedimiento.getPK_SALIDA
        .setPROCEDIMIENTO = cmbProcedimiento.getTEXTO
        
        .setUNIDADES_ID = 0 'cmbUnidad.getPK_SALIDA
        .setRANGO_MIN = 0
        .setRANGO_MAX = 0
        ' Estado
        If optEstado(0).Value = True Then
            .setESTADO = 0
        ElseIf optEstado(1).Value = True Then
            .setESTADO = 1
        ElseIf optEstado(2).Value = True Then
            .setESTADO = 2
        Else
            .setESTADO = 3
        End If
        ' Resultado
        If optResultado(0).Value = True Then
            .setRESULTADO = 0
        ElseIf optResultado(1).Value = True Then
            .setRESULTADO = 1
        Else
            .setRESULTADO = 2
        End If
        .setRUTA_PLANTILLA = txtAdjunto(1)
        .setRUTA_CERTIFICADO = txtAdjunto(2)
        .setRUTA_EVALUACION = txtAdjunto(3)
        
        If .getID_AUX = enumIdAux.ID_AUX_EXISTE Then
            .setID_AUX = enumIdAux.ID_AUX_MODIFICADO
        End If
    End With
    
    If mvarenuTipoEdicion = Alta Then
        mvarobjCalibracion.setFECHA_PREVISTA = mvarobjCalibracion.getFECHA_ACTUAL
    End If
    
    
    If listaReactivos.ListItems.Count > 0 Then
        For i = 1 To listaReactivos.ListItems.Count
            Reactivo = Reactivo & listaReactivos.ListItems(i).Text & ","
        Next i
        Reactivo = Left(Reactivo, Len(Reactivo) - 1)
    Else
        Reactivo = ""
    End If
        
    mvarobjCalibracion.setREACTIVOS = Reactivo
    
    
End Sub

Public Property Get resultado() As Boolean

    resultado = mvarblnResultado

End Property

Public Property Let resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

Private Sub optEstado_Click(Index As Integer)
    If fraEstadoIntervencion.Enabled = False Then Exit Sub
    If mvarenuTipoEdicion = visualizar Then Exit Sub
    
'    If Index = 0 Or Index = 3 Then
    If Index = 0 Then
        If mvarobjCalibracion.getFECHA_ACTUAL <> "" Then
            txtFechaActual.Value = mvarobjCalibracion.getFECHA_ACTUAL
            txtFechaActual_Change
        End If
    Else
        txtFechaActual.Value = Now
        txtFechaActual_Change
        MsgBox "La Fecha de Calibración al Cerrar se Establecerá a la de Hoy (" & Format(Now, "dd/mm/yyyy") & ")." & vbCrLf & "La fecha de Próxima Calibración se recalcula a " & txtFechaProxima.Text, vbInformation, "Calibración"
    End If
End Sub
Private Sub tUnidades_DropDownClose()
    grdResultados.Columns(ColsR.id_unidad) = tUnidades.Columns(1)
    xR(grdResultados.Row, ColsR.id_unidad) = tUnidades.Columns(1)
    grdResultados.Col = 3
End Sub

Private Sub txtFechaActual_Change()
    If IsDate(txtFechaActual.Value) Then
        txtFechaProxima_b.Value = calcularFechaProxima(txtFechaActual.Value, getDataComboSel(cmbPeriCalibracion))
        txtFechaProxima.Text = Format(txtFechaProxima_b.Value, "dd/mm/yyyy")
    End If
End Sub

Private Sub txtLimitacionesUso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdAnadirLimitacion_Click
End Sub

Public Property Get VieneDeCuaderno() As Boolean
    VieneDeCuaderno = mvarblnVieneDeCuaderno
End Property

Public Property Let VieneDeCuaderno(ByVal blnVieneDeCuaderno As Boolean)
    mvarblnVieneDeCuaderno = blnVieneDeCuaderno
End Property

Public Property Get idProcedmientoInicial() As Long
    idProcedmientoInicial = mvarlngidProcedmientoInicial
End Property

Public Property Let idProcedmientoInicial(ByVal lngidProcedmientoInicial As Long)
    mvarlngidProcedmientoInicial = lngidProcedmientoInicial
End Property
