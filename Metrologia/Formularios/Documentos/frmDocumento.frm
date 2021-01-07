VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmDocumento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Albaran/Factura"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   Icon            =   "frmDocumento.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   13410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContratos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contratos de la Obra"
      Height          =   885
      Left            =   5490
      Picture         =   "frmDocumento.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8340
      Width           =   1695
   End
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar todos los artículos"
      Height          =   285
      Left            =   1260
      TabIndex        =   14
      Top             =   8325
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3195
      Left            =   6300
      TabIndex        =   43
      Top             =   480
      Width           =   7065
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Servido en "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   150
         TabIndex        =   0
         Top             =   210
         Width           =   1875
         Begin VB.OptionButton opServido 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fábrica"
            Height          =   345
            Index           =   0
            Left            =   180
            TabIndex        =   1
            Top             =   210
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton opServido 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Obra"
            Height          =   345
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   540
            Width           =   1185
         End
      End
      Begin VB.CheckBox chkValoracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir Importe Albaran"
         Height          =   225
         Left            =   2190
         TabIndex        =   3
         Top             =   630
         Width           =   2145
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2580
         Width           =   5475
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   150
         TabIndex        =   44
         Top             =   1200
         Width           =   6735
         Begin vb6projectpryComboBCA.miComboBCA cmbvehiculos 
            Height          =   375
            Left            =   930
            TabIndex        =   7
            Top             =   510
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   661
         End
         Begin VB.TextBox txtCliente 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   11
            Left            =   5130
            TabIndex        =   10
            Top             =   900
            Width           =   1560
         End
         Begin VB.CheckBox chkSusMedios 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sus medios"
            Height          =   225
            Left            =   150
            TabIndex        =   6
            Top             =   270
            Width           =   2145
         End
         Begin VB.TextBox txtCliente 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   7
            Left            =   930
            TabIndex        =   8
            Top             =   900
            Width           =   1320
         End
         Begin VB.TextBox txtCliente 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   8
            Left            =   2820
            TabIndex        =   9
            Top             =   900
            Width           =   1290
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Remolque"
            Height          =   195
            Index           =   12
            Left            =   4320
            TabIndex        =   51
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vehículo"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   47
            Top             =   600
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Matrícula"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   46
            Top             =   960
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "N.I.F."
            Height          =   195
            Index           =   9
            Left            =   2340
            TabIndex        =   45
            Top             =   960
            Width           =   390
         End
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   9
         Left            =   5640
         TabIndex        =   4
         Top             =   360
         Width           =   1230
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   10
         Left            =   5640
         TabIndex        =   5
         Top             =   690
         Width           =   1230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones "
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   50
         Top             =   2745
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bultos"
         Height          =   195
         Index           =   10
         Left            =   4620
         TabIndex        =   49
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Peso Aprox."
         Height          =   195
         Index           =   11
         Left            =   4620
         TabIndex        =   48
         Top             =   750
         Width           =   855
      End
   End
   Begin TrueDBGrid80.TDBDropDown tArticulos 
      Height          =   3540
      Left            =   60
      TabIndex        =   31
      Top             =   4185
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   6244
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod.Artículo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Prec. Fábrica"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "Currency"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Prec. Obra"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Currency"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Prec. Porte"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Currency"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=11086"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=10980"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2381"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2275"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2408"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2302"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=582"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=476"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(37)  =   ":id=32,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=32,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(58)  =   "Named:id=33:Normal"
      _StyleDefs(59)  =   ":id=33,.parent=0"
      _StyleDefs(60)  =   "Named:id=34:Heading"
      _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=34,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=35:Footing"
      _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=36:Selected"
      _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=37:Caption"
      _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(69)  =   "Named:id=38:HighlightRow"
      _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=39:EvenRow"
      _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=40:OddRow"
      _StyleDefs(74)  =   ":id=40,.parent=33"
      _StyleDefs(75)  =   "Named:id=41:RecordSelector"
      _StyleDefs(76)  =   ":id=41,.parent=34"
      _StyleDefs(77)  =   "Named:id=42:FilterBar"
      _StyleDefs(78)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8340
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Cliente/Obra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3195
      Left            =   60
      TabIndex        =   19
      Top             =   480
      Width           =   6180
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Obra"
         Height          =   765
         Left            =   4590
         Picture         =   "frmDocumento.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2340
         Width           =   1425
      End
      Begin VB.CommandButton cmdUltimoAlbaran 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver último Albaran"
         Height          =   765
         Left            =   2980
         Picture         =   "frmDocumento.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2340
         Width           =   1575
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1560
         Width           =   4890
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1230
         Width           =   4890
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   900
         Width           =   4890
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   570
         Width           =   4890
      End
      Begin VB.CommandButton cmdCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos del Cliente"
         Height          =   765
         Left            =   1520
         Picture         =   "frmDocumento.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2340
         Width           =   1425
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1890
         Width           =   750
      End
      Begin VB.CommandButton cmdObra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos de la Obra"
         Height          =   765
         Left            =   120
         Picture         =   "frmDocumento.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2340
         Width           =   1365
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   4890
      End
      Begin MSDataListLib.DataCombo cmbTarifa 
         Height          =   315
         Left            =   2910
         TabIndex        =   28
         Top             =   1890
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   14737632
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contacto"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   42
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Top             =   1245
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dirección"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   915
         Width           =   675
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   585
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   1950
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarifa Porte"
         Height          =   195
         Index           =   4
         Left            =   1980
         TabIndex        =   27
         Top             =   1950
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   255
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8340
      Width           =   1155
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borrar &Línea"
      Height          =   885
      Left            =   90
      Picture         =   "frmDocumento.frx":35DC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8340
      Width           =   1065
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   4020
      Left            =   60
      TabIndex        =   12
      Top             =   3690
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   7091
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Cod.Artículo"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tArticulos"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Precio Millar"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "Currency"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Cantidad"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "General Number"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Total"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Currency"
      Columns(4).ConvertEmptyCell=   1
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Portes"
      Columns(5).DataField=   ""
      Columns(5).NumberFormat=   "Currency"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Portes Precio Millar"
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "Currency"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Fecha Albaran"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Nº Albaran"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "SERVIDO"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(0).AutoCompletion=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=8255"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=8149"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2117"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2011"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=514"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1588"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1482"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2275"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2170"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2302"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2196"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=3810"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=3704"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(7).Width=2487"
      Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2381"
      Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=8705"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=1879"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1773"
      Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=8705"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=3810"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=3704"
      Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(59)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
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
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41,.alignment=2"
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
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4,.alignment=2"
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
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=36,.parent=11,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=66,.parent=11,.alignment=1,.locked=0"
      _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=12,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=54,.parent=11,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=12,.alignment=2"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=58,.parent=11"
      _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=12"
      _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=15"
      _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=62,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=12,.alignment=2"
      _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=15"
      _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=70,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=12,.alignment=2"
      _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=15"
      _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=74,.parent=11"
      _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=12"
      _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=15"
      _StyleDefs(80)  =   "Named:id=37:Normal"
      _StyleDefs(81)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(82)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(83)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(84)  =   "Named:id=38:Heading"
      _StyleDefs(85)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=38,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=39:Footing"
      _StyleDefs(88)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=40:Selected"
      _StyleDefs(90)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(91)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(92)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(93)  =   "Named:id=41:Caption"
      _StyleDefs(94)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(95)  =   "Named:id=42:HighlightRow"
      _StyleDefs(96)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(97)  =   "Named:id=43:EvenRow"
      _StyleDefs(98)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(99)  =   "Named:id=44:OddRow"
      _StyleDefs(100) =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(101) =   "Named:id=47:RecordSelector"
      _StyleDefs(102) =   ":id=47,.parent=38"
      _StyleDefs(103) =   "Named:id=50:FilterBar"
      _StyleDefs(104) =   ":id=50,.parent=37"
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   390
      Left            =   11820
      TabIndex        =   54
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   14737632
      Format          =   51576833
      CurrentDate     =   38002
   End
   Begin VB.Label lblDescuento 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11205
      TabIndex        =   58
      Top             =   8370
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Descuento"
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
      Height          =   345
      Index           =   4
      Left            =   9720
      TabIndex        =   57
      Top             =   8370
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11160
      TabIndex        =   55
      Top             =   60
      Width           =   570
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo Albaran/Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Width           =   13400
   End
   Begin VB.Label lblportes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11205
      TabIndex        =   34
      Top             =   7740
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Portes"
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
      Height          =   345
      Index           =   3
      Left            =   9720
      TabIndex        =   33
      Top             =   7740
      Width           =   1455
   End
   Begin VB.Label lbliva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11205
      TabIndex        =   25
      Top             =   8700
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Iva"
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
      Height          =   345
      Index           =   1
      Left            =   9720
      TabIndex        =   24
      Top             =   8700
      Width           =   1455
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11205
      TabIndex        =   22
      Top             =   9030
      Width           =   2160
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total con IVA"
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
      Height          =   315
      Index           =   2
      Left            =   9720
      TabIndex        =   21
      Top             =   9030
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Base"
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
      Height          =   345
      Index           =   0
      Left            =   9720
      TabIndex        =   18
      Top             =   8055
      Width           =   1455
   End
   Begin VB.Label lblbase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   11205
      TabIndex        =   17
      Top             =   8055
      Width           =   2160
   End
End
Attribute VB_Name = "frmDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_OBRA As Long
Public PK_CLIENTE As Long
Public PK_DOCUMENTO As Long

Dim x As New XArrayDB
Dim xarticulos As New XArrayDB

Const filas As Integer = 80
Const Col As Integer = 9
Private Enum Cols
    cARTICULO_ID = 0
    cDescripcion = 1
    cPrecio = 2
    cCantidad = 3
    cTotal = 4
    cPORTES = 5
    cPORTES_PRECIO = 6
    cFECHA_ALBARAN = 7
    cNUMERO_ALBARAN = 8
    cSERVIDO = 9
End Enum
Private albaranModificado As Boolean
Private descuentoAlbaran As Double


Private Sub cmdContratos_Click()
    If PK_OBRA > 0 Then
        frmObras_Contratos.pk = PK_OBRA
        frmObras_Contratos.Show 1
    End If
End Sub

Private Sub chkSusMedios_Click()
    txtCliente(7) = ""
    txtCliente(8) = ""
    txtCliente(11) = ""
    If chkSusMedios.Value = Checked Then
        cmbvehiculos.MostrarElemento 1
        cmbvehiculos.desactivar
        txtCliente(7).BackColor = vbWhite
        txtCliente(8).BackColor = vbWhite
        txtCliente(11).BackColor = vbWhite
        txtCliente(7).Enabled = True
        txtCliente(8).Enabled = True
        txtCliente(11).Enabled = True
        On Error Resume Next
        txtCliente(7).SetFocus
    Else
        cmbvehiculos.activar
        txtCliente(7).BackColor = &HE0E0E0
        txtCliente(8).BackColor = &HE0E0E0
        txtCliente(11).BackColor = &HE0E0E0
        txtCliente(7).Enabled = False
        txtCliente(8).Enabled = False
        txtCliente(11).Enabled = False
        On Error Resume Next
        cmbvehiculos.SetFocus
        
    End If
    recalcular_portes
End Sub

Private Sub chkTodos_Click()
    cargar_tarifa
    grid.SetFocus
End Sub

Private Sub cmbVehiculos_change()
    txtCliente(7).BackColor = &HE0E0E0
    If cmbvehiculos.getTEXTO <> "" Then
        If cmbvehiculos.getPK_SALIDA = 1 Then
            chkSusMedios.Value = Checked
'            txtCliente(7).BackColor = vbWhite
'            txtCliente(8).BackColor = vbWhite
'            txtCliente(11).BackColor = vbWhite
'            txtCliente(7) = ""
'            txtCliente(8) = ""
'            txtCliente(11) = ""
'            On Error Resume Next
'            txtCliente(7).SetFocus
        Else
            chkSusMedios.Value = Unchecked
            Dim oVeh As New clsVehiculos
            oVeh.Carga cmbvehiculos.getPK_SALIDA
            txtCliente(7) = oVeh.getMATRICULA
            txtCliente(8) = oVeh.getNIF
            txtCliente(11) = oVeh.getREMOLQUE
            Set oVeh = Nothing
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    Dim modificacion As Boolean
    modificacion = False
    If PK_OBRA = 0 Then
        MsgBox "Debe seleccionar una obra.", vbInformation, App.Title
        Exit Sub
    End If
    ' Validar vehiculo
    If cmbvehiculos.getTEXTO = "" Then
        MsgBox "Debe indicar el vehículo que retira la mercancía.", vbExclamation, App.Title
        cmbvehiculos.SetFocus
        Exit Sub
    End If
    
    If PK_DOCUMENTO = 0 Then
        gTipo_Documento = 1
        frmDocumento_Seleccion.Show 1
        If gTipo_Documento = 0 Then
            MsgBox "Debe seleccionar un tipo de documento.", vbInformation, App.Title
            Exit Sub
        End If
    End If
    Dim i As Integer
    Dim algo As Boolean
    algo = False
    For i = 0 To filas
        If Trim(x.Value(i, Cols.cARTICULO_ID)) <> "" Or Trim(x.Value(i, Cols.cDescripcion)) <> "" Then
            algo = True
        End If
    Next
    If algo = False Then
        MsgBox "El documento no contiene ninguna linea.", vbExclamation, App.Title
        grid.SetFocus
        Exit Sub
    End If
    ' Documento
    Dim oDOCUMENTO As New clsDocumentos
    Dim ocliente As New clsCliente
    Dim DOCUMENTO As Long
    With oDOCUMENTO
        If PK_DOCUMENTO = 0 Then
            .setTIPO_DOCUMENTO_ID = gTipo_Documento
            .setANNO = Format(txtfecha, "yyyy")
            .setFACTURADO = 0
            .setDOCUMENTO_ID_REL = 0
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setHORA = Format(Time, "HH:MM:SS")
        End If
        .setFECHA = Format(txtfecha, "yyyy-mm-dd")
        .setOBRA_ID = PK_OBRA
        Dim oobra As New clsObras
        ' Forma de Pago
        oobra.Carga PK_OBRA
        .setFP_ID = oobra.getFORMA_PAGO_ID
        Set oobra = Nothing
        
        .setVALORACION = chkValoracion.Value
        .setOBSERVACIONES = txtCliente(2)
        If opServido(0).Value = True Then
            .setSERVIDO = ENUM_SERVIDO_FABRICA
        Else
            .setSERVIDO = ENUM_SERVIDO_OBRA
        End If
        .setVALORACION = chkValoracion.Value
        
        .setVEHICULO_ID = cmbvehiculos.getPK_SALIDA
        
        If cmbvehiculos.getPK_SALIDA = 1 Then
            .setMATRICULA = txtCliente(7)
            .setNIF = txtCliente(8)
            .setREMOLQUE = txtCliente(11)
        Else
            .setMATRICULA = ""
            .setNIF = ""
            .setREMOLQUE = ""
        End If
        
        If txtCliente(9) = "" Then
            .setBULTOS = 0
        Else
            .setBULTOS = txtCliente(9)
        End If
        If txtCliente(10) = "" Then
            .setPESO = 0
        Else
            .setPESO = txtCliente(10)
        End If
        ' Descuento
'        .setDESCUENTO = 0
'        .setDESCUENTO_PORCENTAJE = 0
'        Dim oObra As New clsObras
'        If oObra.Carga(PK_OBRA) = True Then
'            If oObra.getDESCUENTO <> 0 Then
'                .setDESCUENTO_PORCENTAJE = oObra.getDESCUENTO
'            End If
'        End If
'        Set oObra = Nothing
        .setIVA = txtCliente(1)
        .setTOTAL = moneda_bd(lblbase)
        .setPORTES = moneda_bd(lblportes)
        .setDESCUENTO = moneda_bd(lblDescuento)
        
        log ("Insertando documento...")
        If PK_DOCUMENTO = 0 Then
            DOCUMENTO = .Insertar
        Else
            modificacion = True
            .Modificar (PK_DOCUMENTO)
            DOCUMENTO = PK_DOCUMENTO
        End If
    End With
    ' Lineas del documento
    Dim oDocumento_Detalle As New clsDocumentos_detalle
    If PK_DOCUMENTO <> 0 Then
        oDocumento_Detalle.Eliminar (PK_DOCUMENTO)
    End If
    For i = x.LowerBound(1) To x.UpperBound(1)
        If Trim(x.Value(i, Cols.cDescripcion)) <> "" Then
            With oDocumento_Detalle
                .setDOCUMENTO_ID = DOCUMENTO
                .setORDEN = i
                .setARTICULO_ID = 0
                If Not IsEmpty(x.Value(i, Cols.cARTICULO_ID)) Then
                    If Trim(x.Value(i, Cols.cARTICULO_ID)) <> "" Then
                        .setARTICULO_ID = x.Value(i, Cols.cARTICULO_ID)
                    End If
                End If
                .setDESCRIPCION = x.Value(i, Cols.cDescripcion)
                If Trim(x.Value(i, Cols.cPrecio)) = "" Then
                    .setPRECIO = "0.00"
                Else
                    .setPRECIO = moneda_bd(x.Value(i, Cols.cPrecio))
                End If
                If Trim(x.Value(i, Cols.cCantidad)) = "" Then
                    .setCANTIDAD = 0
                Else
                    .setCANTIDAD = x.Value(i, Cols.cCantidad)
                End If
                If Trim(x.Value(i, Cols.cTotal)) = "" Then
                    .setTOTAL = "0.00"
                Else
                    .setTOTAL = moneda_bd(x.Value(i, Cols.cTotal))
                End If
                If Trim(x.Value(i, Cols.cPORTES)) = "" Then
                    .setPORTES = "0.00"
                Else
                    .setPORTES = moneda_bd(x.Value(i, Cols.cPORTES))
                End If
                If Trim(x.Value(i, Cols.cFECHA_ALBARAN)) = "" Then
                    .setFECHA_ALBARAN = ""
                Else
                    .setFECHA_ALBARAN = Format(x.Value(i, Cols.cFECHA_ALBARAN), "dd-mm-yyyy")
                End If
                If Trim(x.Value(i, Cols.cNUMERO_ALBARAN)) = "" Then
                    .setNUMERO_ALBARAN = "0.00"
                Else
                    .setNUMERO_ALBARAN = x.Value(i, Cols.cNUMERO_ALBARAN)
                End If
                If Trim(x.Value(i, Cols.cSERVIDO)) = "" Then
                    .setSERVIDO = ""
                Else
                    .setSERVIDO = x.Value(i, Cols.cSERVIDO)
                End If

                If .Insertar = 0 Then
                    If PK_DOCUMENTO = 0 Then
                        oDOCUMENTO.Eliminar (DOCUMENTO)
                    End If
                    Exit Sub
                End If
                ' Cantidades del Contrato
                If cmdContratos.Visible = True Then
                    If x(i, Cols.cARTICULO_ID) <> "" And x(i, Cols.cCantidad) <> "" Then
                        Dim oOCC As New clsObras_contratos_cantidades
                        oOCC.Descontar_Cantidad PK_OBRA, CLng(x(i, Cols.cARTICULO_ID)), CLng(x(i, Cols.cCantidad))
                        Set oOCC = Nothing
                    End If
                End If
            End With
        End If
    Next
    log ("Documento insertado correctamente.")
    ' Recibos
    Dim oRecibo As New clsDocumentos_Recibos
    oRecibo.Generar_Recibos DOCUMENTO
    Set oRecibo = Nothing
'    If PK_DOCUMENTO = 0 Then
'        MsgBox "El documento se ha almacenado correctamente.", vbInformation, App.Title
'    End If

    'se llama a la funcionalidad de modificar el detalle de la factura si se cumple que
    ' estamos en modificacion y el tipo de documento es un albaran.
    'Si la factura a modificar está facturada lo avisará
    
    oDOCUMENTO.Carga DOCUMENTO
    
    If modificacion And oDOCUMENTO.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN And _
       oDOCUMENTO.getFACTURADO = 1 Then
    
        
        Dim oFactura As New clsDocumentos
        oFactura.Carga oDOCUMENTO.getDOCUMENTO_ID_REL
        
        If MsgBox("Se ha modificado el albaran, ¿desea actualizar la factura " & _
             oFactura.getNUMERO & "/" & oFactura.getANNO & "?", vbInformation + vbYesNo, App.Title) = vbYes Then
                       
                        
            If oFactura.getESTADO_ID = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA Then
                If MsgBox("La factura " & oFactura.getNUMERO & "/" & oFactura.getANNO & " esta cobrada." & Chr(13) & _
                    "¿Desea actualizar el detalle de la factura?", vbInformation + vbYesNo, App.Title) = vbYes Then
                    
                    If Not oFactura.actualizar_Albaran_en_Factura(PK_DOCUMENTO) Then
                        MsgBox "Error en la actualización de la factura " & oFactura.getNUMERO & "/" & oFactura.getANNO, vbExclamation, App.Title
                    End If
                End If
            Else
                If Not oFactura.actualizar_Albaran_en_Factura(PK_DOCUMENTO) Then
                        MsgBox "Error en la actualización de la factura " & oFactura.getNUMERO & "/" & oFactura.getANNO, vbExclamation, App.Title
                End If
            End If
    
        End If
    End If
        
    
    frmimprimir.pk = DOCUMENTO
    frmimprimir.Show 1
    Unload Me
    Exit Sub
fallo:
    MsgBox "Error al guardar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To Col
        grid.SelBookmarks.Add grid.Bookmark
        x(grid.Bookmark, i) = ""
        grid.SelBookmarks.Remove 0
    Next
    grid.Refresh
    calcular_total
    grid.SetFocus
End Sub

Private Sub cmdCliente_Click()
    If PK_CLIENTE > 0 Then
        frmClientes.pk = PK_CLIENTE
        frmClientes.Show 1
        cargar_obra PK_OBRA
    End If
    cargar_tarifa
End Sub

Private Sub cmdObra_Click()
    If PK_OBRA > 0 Then
        frmObras.pk = PK_OBRA
        frmObras.Show 1
        cargar_obra PK_OBRA
    Else
        frmObras_Buscar.Show 1
        If gobra <> 0 Then
            PK_OBRA = gobra
            cargar_obra PK_OBRA
        End If
    End If
    cargar_tarifa
End Sub

Private Sub cmdSalir_Click()
    If PK_DOCUMENTO <> 0 Then
        Unload Me
    Else
        If MsgBox("¿Esta seguro que desea salir?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub cmdUltimoAlbaran_Click()
    If cmdUltimoAlbaran.Tag <> "" Then
        Dim frm As New frmDocumento
        frm.PK_DOCUMENTO = cmdUltimoAlbaran.Tag
        frm.Show 1
        Set frm = Nothing
    End If
End Sub

Private Sub Command1_Click()
    If PK_CLIENTE <> 0 Then
        frmObras_Buscar.PK_CLIENTE = PK_CLIENTE
    Else
        frmObras_Buscar.PK_CLIENTE = 0
    End If
    gobra = 0
    frmObras_Buscar.Show 1
    If gobra <> 0 Then
        PK_OBRA = gobra
        cargar_obra PK_OBRA
        cargar_tarifa
        albaranModificado = True
    End If
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
    albaranModificado = False
    cargar_botones Me
    inicializar_ventana
    If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
        lbltitulo.BackColor = vbYellow
        
        Label3(1).Visible = False
        Label3(2).Visible = False
        lbliva.Visible = False
        lbltotal.Visible = False
    End If
End Sub

Private Sub calcular_total()
    Dim i As Integer
    On Error Resume Next
    Dim totalBase As Single
    Dim totalPortes As Single
    Dim totalDescuento As Single
    Dim iva As Single
    totalBase = 0
    totalPortes = 0
    iva = 0
    For i = 0 To filas
        If Trim(CStr(x(i, Cols.cTotal))) <> "" Then
            totalBase = totalBase + CSng(CStr(x.Value(i, Cols.cTotal)))
        End If
        If Trim(CStr(x(i, Cols.cPORTES))) <> "" Then
            totalPortes = totalPortes + CSng(CStr(x.Value(i, Cols.cPORTES)))
        End If
    Next
    lblbase = Format(totalBase, "#,##0.00")
    lblportes = Format(totalPortes, "#,##0.00")
    
    totalDescuento = ((-1) * totalBase * descuentoAlbaran) / 100
    
    lblDescuento = Format(totalDescuento, "#,##0.00")
    
'    IVA = IVA + (((totalBase + totalPortes) * CStr(txtCliente(1))) / 100)
    iva = iva + (((totalBase + totalDescuento) * CStr(txtCliente(1))) / 100)
    lbliva = Format(iva, "#,##0.00")
    
'    lbltotal = Format(CCur(lblbase) + CCur(lblportes) + CCur(lbliva), "#,##0.00")
    lbltotal = Format(CCur(lblbase) + CCur(lbliva) + CCur(lblDescuento), "#,##0.00")
End Sub
Private Sub recalcular_portes()
    Dim i As Integer
    For i = 0 To filas
      If Not IsEmpty(x.Value(i, Cols.cCantidad)) And Not IsEmpty(x.Value(i, Cols.cPORTES_PRECIO)) Then
        If IsNumeric(x.Value(i, Cols.cCantidad)) And IsNumeric(x.Value(i, Cols.cPORTES_PRECIO)) Then
            If opServido(0).Value = True Or chkSusMedios.Value = Checked Then
                x.Value(i, Cols.cPORTES) = moneda("0")
            Else
                x.Value(i, Cols.cPORTES) = Format(CInt(x.Value(i, Cols.cCantidad)) * CSng(x.Value(i, Cols.cPORTES_PRECIO)) / 1000, "currency")
            End If
        End If
      End If
    Next
    calcular_total
    grid.Refresh
End Sub
Private Sub calcular_fila()
   On Error GoTo calcular_fila_Error

    ' Total = Cantidad * Precio / 1000
    If IsNumeric(grid.Columns(Cols.cCantidad).Text) And IsNumeric(grid.Columns(Cols.cPrecio).Text) Then
        grid.Columns(Cols.cTotal).Text = Format(CInt(grid.Columns(Cols.cCantidad).Text) * CSng(grid.Columns(Cols.cPrecio).Text) / 1000, "currency")
    End If
    ' Portes = Cantidad * Precio Porte Millar / 1000
    If IsNumeric(grid.Columns(Cols.cCantidad).Text) And IsNumeric(grid.Columns(Cols.cPORTES_PRECIO).Text) Then
        If opServido(0).Value = True Or chkSusMedios.Value = Checked Then
            grid.Columns(Cols.cPORTES).Text = moneda("0")
        Else
            grid.Columns(Cols.cPORTES).Text = Format(CInt(grid.Columns(Cols.cCantidad).Text) * CSng(grid.Columns(Cols.cPORTES_PRECIO).Text) / 1000, "currency")
        End If
    End If
    calcular_total

   On Error GoTo 0
   Exit Sub

calcular_fila_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcular_fila of Formulario frmDocumento"
End Sub

Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
   On Error GoTo grid_AfterColEdit_Error
    Select Case ColIndex
'        Case Cols.cCantidad
'            If grid.Columns(ColIndex) = "" Then
'                If MsgBox("No ha introducido la cantidad. ¿Es correcto?", vbQuestion + vbYesNo, App.Title) = vbNo Then
'                    grid.Col = cCantidad
'                End If
'            End If
        Case Cols.cPORTES
            Exit Sub
    End Select
    calcular_fila
   On Error GoTo 0
   Exit Sub

grid_AfterColEdit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure grid_AfterColEdit of Formulario frmDocumento"
End Sub
Private Sub grid_AfterUpdate()
    calcular_total
    If cmdContratos.Visible = True Then
        verificar_cantidades_contratadas
    End If
End Sub
Private Sub grid_KeyPress(KeyAscii As Integer)
    If (grid.Col = Cols.cPrecio Or grid.Col = Cols.cTotal Or grid.Col = Cols.cPORTES) And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub cargar_documento()
    On Error GoTo fallo
    Dim oDOCUMENTO As New clsDocumentos
    If oDOCUMENTO.Carga(PK_DOCUMENTO) = True Then
       If oDOCUMENTO.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.factura Then
        Frame3.Visible = False
        chkValoracion.Visible = False
        txtCliente(9).Visible = False
        Label7(10).Visible = False
        txtCliente(10).Visible = False
        Label7(11).Visible = False
        Me.Caption = "Modificación de la FACTURA : " & oDOCUMENTO.getNUMERO & "/" & oDOCUMENTO.getANNO
       ElseIf oDOCUMENTO.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN Then
        Me.Caption = "Modificación del ALBARAN : " & oDOCUMENTO.getNUMERO & "/" & oDOCUMENTO.getANNO
       Else
        Me.Caption = "Modificación del DOCUMENTO : " & oDOCUMENTO.getNUMERO & "/" & oDOCUMENTO.getANNO
       End If
       
       
        If oDOCUMENTO.getTIPO_DOCUMENTO_ID <> ENUM_TIPOS_DOCUMENTOS.factura Then
            grid.Splits(0).Columns(7).Visible = False
            grid.Splits(0).Columns(8).Visible = False
            grid.Splits(0).Columns(1).Width = 4680 + 1409 + 1066
            
        End If

       
       txtfecha = oDOCUMENTO.getFECHA
       lbltitulo = Me.Caption
       ' Cargamos la obra y el cliente
       PK_OBRA = oDOCUMENTO.getOBRA_ID
       cargar_obra PK_OBRA
       cmbTarifa.BoundText = oDOCUMENTO.getTARIFA_ID
       chkValoracion.Value = oDOCUMENTO.getVALORACION
       txtCliente(2) = oDOCUMENTO.getOBSERVACIONES
       txtCliente(9) = oDOCUMENTO.getBULTOS
       txtCliente(10) = oDOCUMENTO.getPESO
       
       If oDOCUMENTO.getSERVIDO = ENUM_SERVIDO_FABRICA Then
        opServido(0).Value = True
       ElseIf oDOCUMENTO.getSERVIDO = ENUM_SERVIDO_OBRA Then
        opServido(1).Value = True
       Else
        opServido(0).Value = True
       End If
       
       If oDOCUMENTO.getVEHICULO_ID = 1 Then
        chkSusMedios.Value = Checked
        cmbvehiculos.MostrarElemento oDOCUMENTO.getVEHICULO_ID
        txtCliente(7) = oDOCUMENTO.getMATRICULA
        txtCliente(8) = oDOCUMENTO.getNIF
        txtCliente(11) = oDOCUMENTO.getREMOLQUE
       Else
        chkSusMedios.Value = Unchecked
        cmbvehiculos.MostrarElemento oDOCUMENTO.getVEHICULO_ID
        Dim oVehiculo As New clsVehiculos
        oVehiculo.Carga oDOCUMENTO.getVEHICULO_ID
        txtCliente(7) = oVehiculo.getMATRICULA
        txtCliente(8) = oVehiculo.getNIF
        txtCliente(11) = oVehiculo.getREMOLQUE
        Set oVehiculo = Nothing
       End If
       
       
       ' Cargamos los datos del documento
       Dim oDocumento_Detalle As New clsDocumentos_detalle
       Dim rs As ADODB.Recordset
       Set rs = oDocumento_Detalle.Detalle_Documento(PK_DOCUMENTO)
       If rs.RecordCount > 0 Then
            Dim fila As Long
            fila = 0
            Do
                x(fila, Cols.cARTICULO_ID) = CStr(rs(0))
                x(fila, Cols.cDescripcion) = Trim(CStr(rs(1)))
                x(fila, Cols.cPrecio) = CStr(rs(2))
                x(fila, Cols.cCantidad) = CStr(rs(3))
                x(fila, Cols.cTotal) = CStr(rs(4))
                x(fila, Cols.cPORTES) = CStr(rs(5))
                If CInt(rs(5)) = 0 Then
                    x(fila, Cols.cPORTES_PRECIO) = moneda("0")
                Else
                    x(fila, Cols.cPORTES_PRECIO) = moneda(rs(5) / (rs(3) / 1000)) ' precio porte millar = precio porte / cantidad * 1000
                End If
                x(fila, Cols.cFECHA_ALBARAN) = Format(CStr(rs(7)), "dd/mm/yyyy")
                x(fila, Cols.cNUMERO_ALBARAN) = CStr(rs(8))
                x(fila, Cols.cSERVIDO) = CStr(rs(9))
                ' Observaciones
                rs.MoveNext
                fila = fila + 1
            Loop Until rs.EOF
            grid.Row = 0
            grid.Col = 0
            grid.Refresh
'            grid.SetFocus
        End If
        calcular_total
    Else
        MsgBox "Error al cargar el documento.", vbInformation, App.Title
    End If
    Set oDOCUMENTO = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub inicializar_ventana()
    txtfecha = Date
    lblbase = Format("0", "#,##0.00")
    lblportes = Format("0", "#,##0.00")
    lbliva = Format("0", "#,##0.00")
    lbltotal = Format("0", "#,##0.00")
    lblDescuento = Format("0", "#,##0.00")
    cargar_combos
    inicializar_grid
    If PK_DOCUMENTO <> 0 Then
        cargar_documento
        cargar_tarifa
    Else
        If PK_CLIENTE <> 0 Then
            frmObras_Buscar.PK_CLIENTE = PK_CLIENTE
        Else
            frmObras_Buscar.PK_CLIENTE = 0
        End If
        frmObras_Buscar.Show 1
        If gobra <> 0 Then
            PK_OBRA = gobra
            cargar_obra PK_OBRA
            cargar_tarifa
        End If
    End If
End Sub

Private Sub cargar_combos()
    Cargar_Combo cmbTarifa, New clsTarifas_portes
    llenar_combo cmbvehiculos, New clsVehiculos, 0, frmVehículos_Detalle, ""
End Sub

Private Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    If PK_DOCUMENTO = 0 Then
        grid.Splits(0).Columns(7).Visible = False
        grid.Splits(0).Columns(8).Visible = False
        grid.Splits(0).Columns(1).Width = 4680 + 1409 + 1066
    End If
    grid.Refresh
End Sub
Private Sub cargar_tarifa()
    Dim rs As ADODB.Recordset
    Dim oTarifa As New clsTarifas
    If chkTodos.Value = Unchecked Then
        Set rs = oTarifa.Listado(PK_OBRA)
    Else
        Set rs = oTarifa.Listado_Completo()
    End If
    xarticulos.Clear
    If rs.RecordCount > 0 Then
        xarticulos.ReDim 1, rs.RecordCount, 1, 5
        Dim i As Integer
        i = 1
        Do
            xarticulos(i, 1) = CStr(rs(0))
            xarticulos(i, 2) = CStr(rs(1))
            xarticulos(i, 3) = CStr(rs(2))
            xarticulos(i, 4) = CStr(rs(3))
            If IsNull(rs(4)) Then
                xarticulos(i, 5) = moneda("0")
            Else
                xarticulos(i, 5) = CStr(rs(4))
            End If
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    Else
        xarticulos.ReDim 1, 1, 1, 5
    End If
    Set tArticulos.Array = xarticulos
    tArticulos.Refresh
    grid.Refresh
End Sub
Private Sub tArticulos_DropDownClose()
    If IsNull(tArticulos.SelectedItem) Then
        Exit Sub
    End If
    grid.Columns(Cols.cARTICULO_ID) = tArticulos.Columns(0)
    grid.Columns(Cols.cDescripcion) = tArticulos.Columns(1)
    If opServido(0).Value = True Then
        grid.Columns(Cols.cPrecio) = tArticulos.Columns(2)
    Else
        grid.Columns(Cols.cPrecio) = tArticulos.Columns(3)
    End If
    grid.Columns(Cols.cPORTES_PRECIO) = tArticulos.Columns(4)
    grid.Col = Cols.cPrecio
End Sub
Private Sub cargar_obra(OBRA As Long)
    Dim ocliente As New clsCliente
    Dim oobra As New clsObras
    If oobra.Carga(OBRA) = True Then
        
        descuentoAlbaran = oobra.getDESCUENTO
        
        PK_CLIENTE = oobra.getCLIENTE_ID
        If ocliente.CargaCliente(oobra.getCLIENTE_ID) Then
            txtCliente(0) = ocliente.getNOMBRE
            txtCliente(3) = oobra.getNOMBRE
            txtCliente(4) = ocliente.getDIRECCION
            txtCliente(5) = ocliente.getTELEFONO
            txtCliente(6) = ocliente.getRAZON
            If oobra.getTIPO_IVA = 2 Then
                txtCliente(1) = 18
            Else
                txtCliente(1) = 0
            End If
'            txtCliente(1) = ocliente.getIVA
            
            cmbTarifa.BoundText = oobra.getTARIFA_PORTE_ID
            ' Ultimo Albaran
            Dim oDOC As New clsDocumentos
            If oDOC.Carga_Ultimo_Albaran(OBRA, PK_DOCUMENTO) Then
                cmdUltimoAlbaran.Caption = "Ver Albaran Anterior: " & oDOC.getNUMERO & "/" & oDOC.getANNO
                cmdUltimoAlbaran.Tag = oDOC.getID_DOCUMENTO
            Else
                cmdUltimoAlbaran.Visible = False
            End If
            Set oDOC = Nothing
            ' Avisos
            If Trim(oobra.getAVISOS) <> "" And PK_DOCUMENTO = 0 Then
                frmAvisos.pk = OBRA
                frmAvisos.Show 1
            End If
            ' Contratos
            Dim oOC As New clsObras_contratos
            If oOC.ContratosEnVigor(OBRA) = True Then
                MsgBox "La obra tiene contratos en vigor. Pulse sobre el botón Contratos si desea consultarlos.", vbInformation, App.Title
                cmdContratos.Visible = True
            Else
                cmdContratos.Visible = False
            End If
            Set oOC = Nothing
        End If
    End If
    Set ocliente = Nothing
    Set oobra = Nothing
    
End Sub
Private Sub txtCliente_LostFocus(Index As Integer)
    If Index = 9 Or Index = 10 Then
        If txtCliente(Index) <> "" Then
            If Not IsNumeric(txtCliente(Index)) Then
                txtCliente(Index) = ""
            End If
        End If
    End If
End Sub
Private Sub verificar_cantidades_contratadas()
    If x(grid.Row, Cols.cARTICULO_ID) <> "" And x(grid.Row, Cols.cARTICULO_ID) <> "" Then
        Dim oOCC As New clsObras_contratos_cantidades
    '    MsgBox x(grid.Row, Cols.cARTICULO_ID) & "-" & x(grid.Row, Cols.cCantidad)
        oOCC.Comprobar_Cantidad PK_OBRA, CLng(x(grid.Row, Cols.cARTICULO_ID)), CLng(x(grid.Row, Cols.cCantidad))
        Set oOCC = Nothing
    End If
End Sub
