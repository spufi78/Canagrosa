VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmREX_Evaluacion_Parametros 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Evaluación de certificado de material de Referencia"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReevaluar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reevaluar"
      Height          =   780
      Left            =   7335
      Picture         =   "frmREX_Evaluacion_Parametros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   9090
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CheckBox chkConforme 
      Caption         =   "Conforme"
      Height          =   240
      Left            =   9045
      TabIndex        =   72
      Top             =   9180
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Frame frmAviso 
      Height          =   735
      Left            =   360
      TabIndex        =   68
      Top             =   9000
      Visible         =   0   'False
      Width           =   3795
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         Caption         =   "Solo el responsable del tipo de reactivo podrá certificarlo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   510
         Left            =   90
         TabIndex        =   69
         Top             =   180
         Width           =   3555
      End
   End
   Begin VB.TextBox texto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   64
      Top             =   8280
      Width           =   5145
   End
   Begin VB.TextBox texto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Index           =   13
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   63
      Top             =   8550
      Width           =   1950
   End
   Begin VB.CommandButton cmdParametros 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisitos"
      Height          =   780
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   9000
      Width           =   870
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   39
      Top             =   180
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Certificado Externo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8250
      Left            =   7335
      TabIndex        =   33
      Top             =   810
      Width           =   6270
      Begin AcroPDFLibCtl.AcroPDF pdf1 
         Height          =   7890
         Left            =   45
         TabIndex        =   34
         Top             =   225
         Width           =   6180
         _cx             =   5080
         _cy             =   5080
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   690
      Left            =   11835
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9180
      Width           =   870
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   690
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9180
      Width           =   870
   End
   Begin TabDlg.SSTab Hojas 
      Height          =   8970
      Left            =   45
      TabIndex        =   3
      Top             =   900
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   15822
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "M.R. / M.R.C."
      TabPicture(0)   =   "frmREX_Evaluacion_Parametros.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(1)=   "Label2(17)"
      Tab(0).Control(2)=   "Label2(0)"
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(6)=   "Label1(4)"
      Tab(0).Control(7)=   "Label1(5)"
      Tab(0).Control(8)=   "Label1(6)"
      Tab(0).Control(9)=   "lblValor(3)"
      Tab(0).Control(10)=   "lblValor(2)"
      Tab(0).Control(11)=   "lblValor(1)"
      Tab(0).Control(12)=   "lblValor(0)"
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(14)=   "lblValor(6)"
      Tab(0).Control(15)=   "lblValor(5)"
      Tab(0).Control(16)=   "lblValor(4)"
      Tab(0).Control(17)=   "lblValor(10)"
      Tab(0).Control(18)=   "lblValor(9)"
      Tab(0).Control(19)=   "lblValor(8)"
      Tab(0).Control(20)=   "lblValor(7)"
      Tab(0).Control(21)=   "Label14"
      Tab(0).Control(22)=   "Label2(1)"
      Tab(0).Control(23)=   "Label2(2)"
      Tab(0).Control(24)=   "op(12)"
      Tab(0).Control(25)=   "texto(10)"
      Tab(0).Control(26)=   "texto(9)"
      Tab(0).Control(27)=   "texto(8)"
      Tab(0).Control(28)=   "texto(7)"
      Tab(0).Control(29)=   "texto(6)"
      Tab(0).Control(30)=   "texto(5)"
      Tab(0).Control(31)=   "texto(4)"
      Tab(0).Control(32)=   "texto(3)"
      Tab(0).Control(33)=   "texto(2)"
      Tab(0).Control(34)=   "texto(1)"
      Tab(0).Control(35)=   "texto(0)"
      Tab(0).Control(36)=   "op(11)"
      Tab(0).Control(37)=   "op(10)"
      Tab(0).Control(38)=   "op(9)"
      Tab(0).Control(39)=   "op(8)"
      Tab(0).Control(40)=   "op(7)"
      Tab(0).Control(41)=   "op(6)"
      Tab(0).Control(42)=   "op(5)"
      Tab(0).Control(43)=   "op(4)"
      Tab(0).Control(44)=   "op(3)"
      Tab(0).Control(45)=   "op(2)"
      Tab(0).Control(46)=   "op(0)"
      Tab(0).Control(47)=   "texto(11)"
      Tab(0).Control(48)=   "Command1"
      Tab(0).Control(49)=   "op(1)"
      Tab(0).Control(50)=   "texto(14)"
      Tab(0).Control(51)=   "cmdCertificar"
      Tab(0).ControlCount=   52
      TabCaption(1)   =   "Reactivos Externos / P.C."
      TabPicture(1)   =   "frmREX_Evaluacion_Parametros.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(9)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(10)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fechaCertificado"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "frmParametros"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdCertificar2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdCertificar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Certificar"
         Height          =   780
         Left            =   5355
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   8100
         Width           =   870
      End
      Begin VB.CommandButton cmdCertificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Certificar"
         Height          =   780
         Left            =   -69600
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   8100
         Width           =   870
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   -73245
         MaxLength       =   100
         TabIndex        =   67
         Top             =   6255
         Width           =   5145
      End
      Begin VB.Frame frmParametros 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parámetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5820
         Left            =   135
         TabIndex        =   55
         Top             =   990
         Width           =   6945
         Begin TrueDBGrid80.TDBGrid grid 
            Height          =   5490
            Left            =   90
            TabIndex        =   56
            Top             =   225
            Width           =   6780
            _ExtentX        =   11959
            _ExtentY        =   9684
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Parámetro"
            Columns(0).DataField=   ""
            Columns(0).DropDown=   "tServicios"
            Columns(0).DropDown.vt=   8
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tolerancia/Valor Límite"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Unidades"
            Columns(2).DataField=   ""
            Columns(2).DropDown=   "tMetodos"
            Columns(2).DropDown.vt=   8
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   1
            Columns(3)._MaxComboItems=   5
            Columns(3).ValueItems(0)._DefaultItem=   0
            Columns(3).ValueItems(0).Value=   "Sí"
            Columns(3).ValueItems(0).Value.vt=   8
            Columns(3).ValueItems(0).DisplayValue=   "0"
            Columns(3).ValueItems(0).DisplayValue.vt=   8
            Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems(1)._DefaultItem=   0
            Columns(3).ValueItems(1).Value=   "No"
            Columns(3).ValueItems(1).Value.vt=   8
            Columns(3).ValueItems(1).DisplayValue=   "1"
            Columns(3).ValueItems(1).DisplayValue.vt=   8
            Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems.Count=   2
            Columns(3).Caption=   "Conforme"
            Columns(3).DataField=   ""
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2328"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2249"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=4895"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4815"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2805"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2725"
            Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(18)=   "Column(3).Width=847"
            Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=767"
            Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Button=1"
            Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=11"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=12"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=15"
            _StyleDefs(52)  =   "Named:id=37:Normal"
            _StyleDefs(53)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
            _StyleDefs(54)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(55)  =   ":id=37,.fontname=MS Sans Serif"
            _StyleDefs(56)  =   "Named:id=38:Heading"
            _StyleDefs(57)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
            _StyleDefs(58)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(59)  =   ":id=38,.strikethrough=0,.charset=0"
            _StyleDefs(60)  =   ":id=38,.fontname=MS Sans Serif"
            _StyleDefs(61)  =   "Named:id=39:Footing"
            _StyleDefs(62)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=40:Selected"
            _StyleDefs(64)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
            _StyleDefs(65)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(66)  =   ":id=40,.fontname=MS Sans Serif"
            _StyleDefs(67)  =   "Named:id=41:Caption"
            _StyleDefs(68)  =   ":id=41,.parent=38,.alignment=2"
            _StyleDefs(69)  =   "Named:id=42:HighlightRow"
            _StyleDefs(70)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
            _StyleDefs(71)  =   "Named:id=43:EvenRow"
            _StyleDefs(72)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=44:OddRow"
            _StyleDefs(74)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
            _StyleDefs(75)  =   "Named:id=47:RecordSelector"
            _StyleDefs(76)  =   ":id=47,.parent=38"
            _StyleDefs(77)  =   "Named:id=50:FilterBar"
            _StyleDefs(78)  =   ":id=50,.parent=37"
         End
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   -68475
         TabIndex        =   43
         Top             =   1350
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informe"
         Height          =   780
         Left            =   -70635
         Picture         =   "frmREX_Evaluacion_Parametros.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   8100
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   -73245
         MaxLength       =   100
         TabIndex        =   32
         Top             =   5985
         Width           =   5145
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   -68475
         TabIndex        =   26
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   -68475
         TabIndex        =   25
         Top             =   1620
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   -68475
         TabIndex        =   24
         Top             =   1890
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   -68475
         TabIndex        =   23
         Top             =   2655
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estabilidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   -68475
         TabIndex        =   22
         Top             =   2925
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento de fabricación:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   -68475
         TabIndex        =   21
         Top             =   3195
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor de la propiedad:"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   7
         Left            =   -68475
         TabIndex        =   20
         Top             =   3960
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incertidumbre máxima expandida:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   -68475
         TabIndex        =   19
         Top             =   4230
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento de asignación del valor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   -68475
         TabIndex        =   18
         Top             =   4500
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Validez de datos e incertidumbre:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   -68475
         TabIndex        =   17
         Top             =   4770
         Width           =   195
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESULTADO (Marque si es Conforme) "
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
         Height          =   195
         Index           =   11
         Left            =   -68475
         TabIndex        =   16
         Top             =   5220
         Width           =   195
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   15
         Top             =   1035
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1305
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1575
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1845
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2610
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2880
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   9
         Top             =   3150
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   8
         Top             =   3915
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   7
         Top             =   4185
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   6
         Top             =   4455
         Width           =   3435
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   -72345
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   5
         Top             =   4725
         Width           =   3435
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   12
         Left            =   -68475
         TabIndex        =   4
         Top             =   6660
         Width           =   195
      End
      Begin MSComCtl2.DTPicker fechaCertificado 
         Height          =   330
         Left            =   1755
         TabIndex        =   74
         Top             =   7650
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Certificado por:"
         Height          =   195
         Index           =   4
         Left            =   405
         TabIndex        =   66
         Top             =   7425
         Width           =   1320
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   3
         Left            =   405
         TabIndex        =   65
         Top             =   7695
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Fecha:"
         Height          =   195
         Index           =   2
         Left            =   -74595
         TabIndex        =   62
         Top             =   7695
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Certificado por:"
         Height          =   195
         Index           =   1
         Left            =   -74595
         TabIndex        =   61
         Top             =   7425
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "1. Lista de parámetros requeridos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   45
         TabIndex        =   58
         Top             =   495
         Width           =   7170
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "3. Usuario que certifica y fecha de certificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   45
         TabIndex        =   57
         Top             =   6975
         Width           =   7170
      End
      Begin VB.Label Label14 
         Caption         =   "RESULTADO (Marque si es Conforme) "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72300
         TabIndex        =   54
         Top             =   5220
         Width           =   3345
      End
      Begin VB.Label lblValor 
         Caption         =   "Valor de la propiedad:"
         Height          =   195
         Index           =   7
         Left            =   -74640
         TabIndex        =   53
         Top             =   3960
         Width           =   1950
      End
      Begin VB.Label lblValor 
         Caption         =   "Incertidumbre máxima expand.:"
         Height          =   240
         Index           =   8
         Left            =   -74640
         TabIndex        =   52
         Top             =   4230
         Width           =   2265
      End
      Begin VB.Label lblValor 
         Caption         =   "Proceso de asignación :"
         Height          =   195
         Index           =   9
         Left            =   -74640
         TabIndex        =   51
         Top             =   4500
         Width           =   1995
      End
      Begin VB.Label lblValor 
         Caption         =   "Validez de los datos:"
         Height          =   240
         Index           =   10
         Left            =   -74640
         TabIndex        =   50
         Top             =   4815
         Width           =   1995
      End
      Begin VB.Label lblValor 
         Caption         =   "Homogeneidad:"
         Height          =   240
         Index           =   4
         Left            =   -74640
         TabIndex        =   49
         Top             =   2655
         Width           =   1905
      End
      Begin VB.Label lblValor 
         Caption         =   "Estabilidad:"
         Height          =   195
         Index           =   5
         Left            =   -74640
         TabIndex        =   48
         Top             =   2925
         Width           =   1995
      End
      Begin VB.Label lblValor 
         Caption         =   "Sistema de Producción:"
         Height          =   330
         Index           =   6
         Left            =   -74640
         TabIndex        =   47
         Top             =   3195
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Conforme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   -68880
         TabIndex        =   46
         Top             =   765
         Width           =   960
      End
      Begin VB.Label lblValor 
         Caption         =   "Analito:"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   45
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label lblValor 
         Caption         =   "Matriz:"
         Height          =   240
         Index           =   1
         Left            =   -74640
         TabIndex        =   44
         Top             =   1350
         Width           =   1905
      End
      Begin VB.Label lblValor 
         Caption         =   "Dispone de certificado:"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   42
         Top             =   1620
         Width           =   1995
      End
      Begin VB.Label lblValor 
         Caption         =   "Tamaño de la muestra:"
         Height          =   240
         Index           =   3
         Left            =   -74640
         TabIndex        =   41
         Top             =   1890
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Valor esperado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   -72345
         TabIndex        =   38
         Top             =   765
         Width           =   3435
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Valor e incertidumbre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   -74685
         TabIndex        =   37
         Top             =   3600
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Producción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   -74640
         TabIndex        =   36
         Top             =   2295
         Width           =   2310
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Definición del material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   -74640
         TabIndex        =   35
         Top             =   765
         Width           =   2310
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "1. Definición de los requisitos analíticos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -74955
         TabIndex        =   31
         Top             =   450
         Width           =   7170
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "2. En cuanto a la propiedad certificada y su uso en el laboratorio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   -74955
         TabIndex        =   30
         Top             =   5715
         Width           =   7170
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Uso previsto:"
         Height          =   195
         Index           =   0
         Left            =   -74595
         TabIndex        =   29
         Top             =   6030
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "¿Es conforme la propiedad certificada a este uso? Marque si es Conforme"
         Height          =   195
         Index           =   17
         Left            =   -74640
         TabIndex        =   28
         Top             =   6660
         Width           =   5280
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "3. Usuario que certifica y fecha de certificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   -74955
         TabIndex        =   27
         Top             =   6975
         Width           =   7170
      End
   End
   Begin VB.Label lblresultado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "NO CONFORME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   9675
      TabIndex        =   60
      Top             =   225
      Width           =   3255
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Evaluación del certificado de material"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4725
      TabIndex        =   2
      Top             =   225
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   13635
   End
End
Attribute VB_Name = "frmREX_Evaluacion_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BOTE_EX_ID As Long

Const SEPARADOR As String = " ; "
Const HOJA_MR As Integer = 0
Const HOJA_PC As Integer = 1
Dim x As New XArrayDB
Const filas As Integer = 20
Const Col As Integer = 3
Private ELEMENTOS_LISTA As Integer
Private Enum COLS
    PARAMETRO = 0
    TOLERANCIA = 1
    unidades = 2
    CONFORME = 3
End Enum

Private tipo As Long
Private USUARIO_CERTIFICADOR As Long
Private CERTIFICADO As Integer

Private Sub cmdReevaluar_Click()
   On Error GoTo cmdReevaluar_Click_Error

    If MsgBox("¿Esta seguro que desea reevaluar el reactivo?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oBote As New clsBotes_ex
        If oBote.ReCertificar(BOTE_EX_ID) Then
            cargarFormulario
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdReevaluar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdReevaluar_Click of Formulario frmREX_Evaluacion_Parametros"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 116 ' F5 Datos especiales
            texto(13).visible = Not texto(13).visible
            fechaCertificado.visible = Not fechaCertificado.visible
    End Select
End Sub

'-------------------------------------------------------------------------------------
' FORM_LOAD
'-------------------------------------------------------------------------------------
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargarFormulario
End Sub
Private Sub inicializar_grid(nfilas As Integer)
    x.ReDim 0, nfilas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargarFormulario()
'CARGA CADA UNA DE LAS PESTAÑAS EN FUNCIÓN DEL TIPO
    Dim oBote As New clsBotes_ex
    Dim oTIPO As New clsTipos_bote_ex
    oBote.CARGAR BOTE_EX_ID
    oTIPO.CARGAR oBote.getTIPO_BOTE_EX_ID
    fechaCertificado = Date
    'PROVISIONAL
    If oBote.getCERTIFICADO <> 0 Then
        USUARIO_CERTIFICADOR = oBote.getUSUARIO_CERTIFICADOR
    Else
        USUARIO_CERTIFICADOR = USUARIO.getID_EMPLEADO
    End If
    
    CERTIFICADO = oBote.getCERTIFICADO
    txtID = "General: " & BOTE_EX_ID
    tipo = CLng(oTIPO.getTIPO_M_REFERENCIA_ID)
    Select Case tipo
    Case 1
        Hojas.TabEnabled(HOJA_MR) = False
        Hojas.TabEnabled(HOJA_PC) = True
        Hojas.Tab = HOJA_PC
        cmdok.Enabled = True
        cargarParametros
    Case 2 To 3
        Hojas.TabEnabled(HOJA_MR) = True
        Hojas.TabEnabled(HOJA_PC) = False
        Hojas.Tab = HOJA_MR
        cargarCertificadoMR
        cmdok.Enabled = False
        cmdok.visible = False
    Case 4 To 8
        Hojas.TabEnabled(HOJA_MR) = False
        Hojas.TabEnabled(HOJA_PC) = True
        Hojas.Tab = HOJA_PC
        cargarParametros
        cmdok.Enabled = True
    End Select
    
    cargarUsuario
    certificadoPDF
    revisarPermisos
    
    Set oBote = Nothing
    Set oTIPO = Nothing
End Sub
Private Sub revisarPermisos()
    Dim i As Integer
'PROVISIONAL PARA QUE TODO EL MUNDO PUEDA CERTIFICAR DE MOMENTO
'    If USUARIO.getID_EMPLEADO = USUARIO_CERTIFICADOR And CERTIFICADO = 0 Then
    If CERTIFICADO = 0 Then
        grid.Enabled = True
    
        lblresultado(2).visible = False
        cmdok.Enabled = True
        cmdCertificar.Enabled = True
        cmdCertificar2.Enabled = True
        frmAviso.visible = False
        For i = 0 To op.Count - 1
            op(i).Enabled = True
        Next i
        cmdReevaluar.visible = False
    Else
        grid.Enabled = False
        
        lblresultado(2).visible = True
        cmdok.Enabled = False
        cmdCertificar.Enabled = False
        cmdCertificar2.Enabled = False
        frmAviso.visible = True
        For i = 0 To op.Count - 1
            op(i).Enabled = False
        Next i
        If CERTIFICADO > 0 Then
            lblMensaje.Caption = "El bote ya ha sido certificado"
            lblMensaje.ForeColor = &H800000
            Label2(2).Enabled = True
            If USUARIO.getID_EMPLEADO = USUARIO_CERTIFICADOR Then
                cmdReevaluar.visible = True
            End If
        Else
            Label2(1).Caption = "Debe certificar el usuario:"
            Label2(2).Enabled = False
        End If
    End If
End Sub
Private Sub cargarParametros()
' Carga de la LISTA
' (Parametros)
' Por BOTE
    Dim rs As ADODB.Recordset
    Dim oConforme As New clsRex_botes_certificados_conf
    Set rs = oConforme.ListadoBote(BOTE_EX_ID)
    ELEMENTOS_LISTA = rs.RecordCount
    If ELEMENTOS_LISTA > 0 Then
        inicializar_grid CInt(ELEMENTOS_LISTA)
        Dim i As Integer
        i = 0
        Do
            x(i, COLS.PARAMETRO) = CStr(rs(2))
            x(i, COLS.TOLERANCIA) = CStr(rs(3))
            x(i, COLS.unidades) = CStr(rs(4))
            If CStr(rs(5)) = "No" Or CStr(rs(5)) = "0" Then
                 x.Value(i, COLS.CONFORME) = "No"
            Else
                 x.Value(i, COLS.CONFORME) = "Sí"
            End If
            rs.MoveNext
            i = i + 1
        Loop Until rs.EOF
        'Marcadores
        Dim oCertif As New clsRex_botes_certificados
        Set oConforme = Nothing
        Set oCertif = Nothing
    Else
        cargarParametrosTipo
    End If
    Set rs = Nothing
End Sub
Private Sub cargarParametrosTipo()
' Carga de la lista
' Parametros
' Por TIPO
    Dim rs As ADODB.Recordset
    Dim oBote As New clsBotes_ex
    Dim oParams As New clsTipos_bote_ex_parametros
    oBote.CARGAR BOTE_EX_ID
    Set rs = oParams.Listado(oBote.getTIPO_BOTE_EX_ID)
    ELEMENTOS_LISTA = rs.RecordCount
    If ELEMENTOS_LISTA > 0 Then
    inicializar_grid CInt(ELEMENTOS_LISTA)
    Dim i As Integer
    i = 0
        Do
            x(i, COLS.PARAMETRO) = CStr(rs(0))
            x(i, COLS.TOLERANCIA) = CStr(rs(1))
            x(i, COLS.unidades) = CStr(rs(2))
            x.Value(i, COLS.CONFORME) = "No"
            rs.MoveNext
            i = i + 1
        Loop Until rs.EOF
    End If
    Set oTBP = Nothing
    Set rs = Nothing
End Sub
Private Sub cargarCertificadoMR()
    'VALORES M.R. y M.R.C.
    Dim strUso() As String
    Dim intCount As Integer
    
    Dim oBote As New clsBotes_ex
    Dim oParam As New clsRex_botes_certificados_conf
    Dim rs As ADODB.Recordset
    Dim VALOR As Long
    
    oBote.CARGAR BOTE_EX_ID
    Set rs = oParam.ListadoBote(BOTE_EX_ID)
    If rs.RecordCount <> 0 Then
        intCount = 0
        Do
           lblValor(CInt(rs("ORDEN"))).Caption = rs("VALOR")
           texto(CInt(rs("ORDEN"))).Text = rs("TOLERANCIA")
           op(CInt(rs("ORDEN"))) = rs("CONFORME")
           intCount = intCount + 1
           rs.MoveNext
        Loop Until rs.EOF Or intCount > 10
    Else
        cargarCertificadoMRParametrico
    End If
    valoresComunes
    
    'TERCER BLOQUE DE VALORES: resultados en caso de existir (oCertif y oConforme)
    '-----------------------------------------------------------------------------
    'Marcadores
    Dim oCertif As New clsRex_botes_certificados
    If oCertif.Carga(BOTE_EX_ID) Then
        op(11) = oCertif.getC20_RESULTADO
        op(12) = oCertif.getC22_CONFORME_PROPIEDAD
    End If
     
    Set oCertif = Nothing
    Set oParam = Nothing
    Set oBote = Nothing
    Set rs = Nothing
End Sub

Private Sub cargarCertificadoMRParametrico()
    Dim strUso() As String
    Dim intCount As Integer
    
    Dim oBote As New clsBotes_ex
    Dim oMR As New clsTipos_bote_ex_req_analiticos 'Parámetros M.R. específicos
    Dim oBoteMR As New clsTipos_bote_ex_mr       'Características adicionales para BOTES M.R.
    Dim oDeco As New clsDecodificadora           'Descodificadora
    Dim VALOR As Long
    
    oBote.CARGAR BOTE_EX_ID
    oBoteMR.CargaTipo CLng(oBote.getTIPO_BOTE_EX_ID) 'Carga por tipo (relación 1:1)
    oMR.CargaTipo CLng(oBote.getTIPO_BOTE_EX_ID)
    
    'PRIMER BLOQUE DE VALORES:
    '-----------------------------------------------------------------------
    'DEFINICIÓN DEL MATERIAL
    texto(0) = oMR.getANALITO
    texto(1) = oMR.getMATRIZ
    texto(2) = oMR.getCERTIFICADO
    texto(3) = oMR.getTAMANYO
    
    'HOMOGENEIDAD,ESTABILIDAD Y SISTEMA DE PRODUCCIÓN
    '------------------------------------------------
    For intCount = 4 To 6
        texto(intCount) = ""
    Next intCount
    strUso = Split(oMR.getHOMOGENEIDAD, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then
            VALOR = CLng(Solo_Numeros(strUso(intCount)))
            oDeco.Carga_valor DECODIFICADORA.REX_HOMOGENEIDAD, VALOR
            texto(4).Text = texto(4).Text & oDeco.getDESCRIPCION & SEPARADOR
          End If
    Next intCount
    'ESTABILIDAD
    strUso = Split(oMR.getESTABILIDAD, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then
            VALOR = CLng(Solo_Numeros(strUso(intCount)))
            oDeco.Carga_valor DECODIFICADORA.REX_ESTABILIDAD, VALOR
            texto(5).Text = texto(5).Text & oDeco.getDESCRIPCION & SEPARADOR
        End If
    Next intCount
    'PRODUCCIÓN
    strUso = Split(oMR.getPROCEDIMIENTO, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then
            VALOR = CLng(Solo_Numeros(strUso(intCount)))
            'M1332
            oDeco.Carga_valor DECODIFICADORA.REX_SIST_PRODUCCION, VALOR
            texto(6).Text = texto(6).Text & oDeco.getDESCRIPCION & SEPARADOR
        End If
    Next intCount
    'VALOR E INCERTIDUMBRE
    '------------------------------------------------
    texto(7).Text = oMR.getVALOR_PROPIEDAD & " " & oMR.getUNIDADES
    'M1332-I
    'texto(8).Text = oMR.getINCERTIDUMBRE & " " & oMR.getUNIDADES
    If Trim(oMR.getUNIDADES_INCERTIDUMBRE) = "" Then
        texto(8).Text = oMR.getINCERTIDUMBRE & " " & oMR.getUNIDADES
    Else
        texto(8).Text = oMR.getINCERTIDUMBRE & " " & oMR.getUNIDADES_INCERTIDUMBRE
    End If
    'M1332-F
    texto(9).Text = oMR.getPROC_ASIGNACION
    texto(10).Text = oMR.getVALIDEZ

    Set oDeco = Nothing
End Sub
Private Sub valoresComunes()
    'SEGUNDO BLOQUE DE VALORES: características adicionales (oBoteMR)
    '-----------------------------------------------------------------------
    Dim strUso() As String
    Dim intCount As Integer
    
    Dim oBote As New clsBotes_ex
    Dim oBoteMR As New clsTipos_bote_ex_mr       'Características adicionales para BOTES M.R.
    Dim oDeco As New clsDecodificadora           'Descodificadora
    Dim VALOR As Long
    
    oBote.CARGAR BOTE_EX_ID
    oBoteMR.CargaTipo CLng(oBote.getTIPO_BOTE_EX_ID) 'Carga por tipo (relación 1:1)
    'USO PREVISTO
    strUso = Split(oBoteMR.getUSO_PREVISTO, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then
            VALOR = CLng(Solo_Numeros(strUso(intCount)))
            oDeco.Carga_valor DECODIFICADORA.REX_USO_PREVISTO_TIPO, VALOR
             texto(11).Text = texto(11).Text & oDeco.getDESCRIPCION & SEPARADOR
        End If
    Next intCount
    strUso = Split(oBoteMR.getPERIODICIDAD, ";")
    For intCount = LBound(strUso) To UBound(strUso)
        If strUso(intCount) <> "" Then
            VALOR = CLng(Solo_Numeros(strUso(intCount)))
            oDeco.Carga_valor DECODIFICADORA.REX_USO_PREVISTO_PERIODICIDAD, VALOR
             texto(14).Text = texto(14).Text & oDeco.getDESCRIPCION & SEPARADOR
        End If
    Next intCount
    Set oBote = Nothing
    Set oBoteMR = Nothing
    Set oDeco = Nothing
End Sub
Private Sub cargarUsuario()
    Dim oBote As New clsBotes_ex
    Dim oUsuario As New clsUsuarios
    oBote.CARGAR BOTE_EX_ID
    If oBote.getCERTIFICADO = 1 Then
        texto(13) = oBote.getFECHA_CERTIFICACION
        fechaCertificado = oBote.getFECHA_CERTIFICACION
    Else
        Label2(1).Caption = "Usuario responsable de certf.:"
        texto(13) = ""
    End If
    
'    If oUsuario.CARGAR(CLng(oBote.getUSUARIO_CERTIFICADOR)) Then
'        texto(12) = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
'    End If
    If USUARIO_CERTIFICADOR <> 0 Then
        oUsuario.CARGAR USUARIO_CERTIFICADOR
        texto(12) = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
    End If
    'FIN-PROVISIONAL
    Set oUsuario = Nothing
    Set oBote = Nothing
End Sub

Private Sub cmdParametros_Click()
'Ventana para la inserción y modificación de los requisitos por defecto (TIPO) en la certificación
    Dim oBote As New clsBotes_ex
    oBote.CARGAR BOTE_EX_ID
    frmREX_Bote_Parametros.PK = oBote.getTIPO_BOTE_EX_ID
    frmREX_Bote_Parametros.Show 1
End Sub

Private Sub cmdcancel_Click()
    If CERTIFICADO = 0 Then
        If MsgBox("¿Esta seguro de salir sin informar la evaluación del material de referencia?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCertificadoExterno_Click()
   certificadoPDF
End Sub

Private Sub cmdok_Click()
On Error GoTo cmdok_Click_Error
'  DISCRIMINACIÓN POR TIPO
'  GUARDADO Y CIERRE DEL FORMULARIO
'  CARGA DE VALORES EN BBDD
'   VALORES:
'   2 --> M.R.
'   3 --> M.R.C
   Select Case tipo
   Case 2 To 3
      PNTA002 (True)
   Case 4 To 8
      PC (True)
   End Select
   Exit Sub
   
cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Evaluacion_Parametros"
End Sub
Private Sub PNTA002(mensaje As Boolean)
On Error GoTo PNTA002_Error
    Dim i As Integer
    Dim oEvaluacion As New clsRex_botes_certificados
    With oEvaluacion
        .setBOTE_EX_ID = BOTE_EX_ID
        .setC06_NOMBRE_MATERIAL = texto(0)
        .setC10_UTILIZACION_PREVISTA = texto(11)
        .setC11_INSTRUCCIONES_USO = texto(9)
        .setC13_NIVEL_HOMOGENEIDAD = texto(4)
        .setC14_CONCENTRACION = texto(10)
        .setC15_MATRIZ = texto(1)
        .setC17_CANTIDAD = texto(3)
        .setC18_ESTABILIDAD = texto(5)
        
        .setC23_TECNICO_RESPONSABLE = USUARIO_CERTIFICADOR
        .setC24_FECHA_EVALUACION = Format(Date, "dd-mm-yyyy")
        .setC08_FECHA_CERTIFICACION = Format(Date, "dd-mm-yyyy")
        '--------------------------------------------------------------------------------
        If op(11).Value = Checked Then
            .setC20_RESULTADO = 1
            chkConforme.Value = Checked
        Else
            .setC20_RESULTADO = 0
            chkConforme.Value = Unchecked
        End If
        .setC21_USO_PREVISTO = texto(11)
        If op(12).Value = Checked Then
            .setC22_CONFORME_PROPIEDAD = 1
        Else
            .setC22_CONFORME_PROPIEDAD = 0
        End If
        '--------------------------------------------------------------------------------
        If .Insertar = 0 Then
            .Modificar BOTE_EX_ID
            cargaParametros
            If mensaje = True Then
                MsgBox "Los datos de la certificación se han modificado correctamente.", vbInformation, App.Title
            End If
        Else
            cargaParametros
            If mensaje = True Then
                MsgBox "Los datos de la certificación se han insertado correctamente.", vbInformation, App.Title
            End If
        End If
'       Unload Me
    End With
   On Error GoTo 0
   Exit Sub

PNTA002_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Evaluacion"
End Sub
Private Sub cargaParametros()
    Dim oConforme As New clsRex_botes_certificados_conf
    Dim i As Integer
    With oConforme
        .Eliminar BOTE_EX_ID
        .setBOTE_EX_ID = BOTE_EX_ID
        For i = 0 To 10
            .setORDEN = i
            .setCONFORME = CStr(op(i).Value)
            .setVALOR = lblValor(i)
            .setTOLERANCIA = texto(i)
            .setUNIDADES = " "
            .Insertar
        Next i
    End With
End Sub
Private Sub certifica()
   ' If MsgBox("Asegúrese de aceptar primero los cambios efectuados en el formulario. ¿Desea certificar el reactivo?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oBote As New clsBotes_ex
        oBote.Certificar BOTE_EX_ID, chkConforme.Value, IIf(fechaCertificado.visible = True, fechaCertificado, Date)
        Set oBote = Nothing
        MsgBox "El reactivo ha sido certificado correctamente", vbInformation, App.Title
        Unload Me
    'End If
End Sub
Private Sub PC(mensaje As Boolean)
On Error GoTo PC_Error
    Dim i As Integer
    Dim oEvaluacion As New clsRex_botes_certificados
    With oEvaluacion
        .setBOTE_EX_ID = BOTE_EX_ID
        .setC23_TECNICO_RESPONSABLE = USUARIO.getID_EMPLEADO
        .setC24_FECHA_EVALUACION = Format(Date, "dd-mm-yyyy")
        .setC08_FECHA_CERTIFICACION = Format(Date, "dd-mm-yyyy")
        '--------------------------------------------------------------------------------
        .setC20_RESULTADO = 1
        chkConforme.Value = Checked
        For i = 0 To ELEMENTOS_LISTA - 1
            If x.Value(i, COLS.CONFORME) = "" Or x.Value(i, COLS.CONFORME) = "No" Then
                .setC20_RESULTADO = 0
                chkConforme.Value = Unchecked
            End If
        Next i
        .setC22_CONFORME_PROPIEDAD = 1
        '--------------------------------------------------------------------------------
        If .Insertar = 0 Then
            If .Modificar(BOTE_EX_ID) = True Then
                grid.Refresh
                Dim oConforme As New clsRex_botes_certificados_conf
                With oConforme
                    .Eliminar BOTE_EX_ID
                    .setBOTE_EX_ID = BOTE_EX_ID
                    For i = 0 To ELEMENTOS_LISTA - 1
                        .setORDEN = i
                        If x.Value(i, COLS.CONFORME) = "" Or x.Value(i, COLS.CONFORME) = "No" Then
                           .setCONFORME = "No"
                        Else
                           .setCONFORME = "Sí"
                        End If
                        .setVALOR = x(i, COLS.PARAMETRO)
                        .setTOLERANCIA = x(i, COLS.TOLERANCIA)
                        .setUNIDADES = x(i, COLS.unidades)
                        .Insertar
                    Next i
                End With
                If mensaje = True Then
                    MsgBox "Los datos de la certificación se han modificado correctamente.", vbInformation, App.Title
                End If
'                Unload Me
            End If
        Else
            If mensaje = True Then
                MsgBox "Los datos de la certificación se han insertado correctamente.", vbInformation, App.Title
            End If
'            Unload Me
        End If
    End With
   On Error GoTo 0
   Exit Sub
PC_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Evaluacion_Parametros"
End Sub

Private Sub Command1_Click()
   Dim consulta As String
   On Error GoTo Command1_Click_Error
        With frmReport
            Dim destino As String
            .iniciar
            .criterio = "{REX_BOTES_CERTIFICADOS.BOTE_EX_ID} = " & BOTE_EX_ID
            .informe = "\REX\rptcertificado"
            .consulta = consulta
             destino = App.Path & "\certificado.pdf"
            .pdf = destino
            .imprimir = False
            .generar
            .visible = False
            If Dir(destino) <> "" Then
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
            End If
        End With
   On Error GoTo 0
   Exit Sub
Command1_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PNTA002 of Formulario frmREX_Evaluacion_Parametros"
End Sub
Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim i As Integer
    Dim noconforme As Boolean
    noconforme = False
    With lblresultado(2)
        .Caption = "CONFORME"
        .BackColor = vbGreen
        For i = 0 To ELEMENTOS_LISTA - 1
            If x.Value(i, COLS.CONFORME) = "" Or x.Value(i, COLS.CONFORME) = "No" Then
                noconforme = True
            End If
        Next i
  
        If noconforme = True Then
           .Caption = "NO CONFORME"
           .BackColor = &H8080FF
        End If
    End With
    grid.Refresh
End Sub

Private Sub op_Click(Index As Integer)
On Error Resume Next
    If Index <> 11 And Index <> 12 Then
        texto(Index).Locked = True
    Else
        With lblresultado(2)
            If op(11).Value = Checked And op(12).Value = Checked Then
                .Caption = "CONFORME"
                .BackColor = vbGreen
            Else
                .Caption = "NO CONFORME"
                .BackColor = &H8080FF
            End If
        End With
    End If
End Sub

Public Sub certificadoPDF()
On Error GoTo fallo
    pdf1.LoadFile vbNullString
    Dim fichero As String
    Dim oAdjunto As New clsAdjuntos
    fichero = oAdjunto.CargarDocumentoUltimo(TOBJETO.TOBJETO_REX_CERTIFICADOS, CLng(BOTE_EX_ID), 0, False, ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_CERTIFICADO)
    mostrar_pdf fichero
    Set oAdjunto = Nothing

    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub mostrar_pdf(DOC As String)
    If DOC <> "" Then
        If UCase(Right(DOC, 3)) = "PDF" Then
            If Dir(DOC) <> "" Then
                With pdf1
                    .visible = True
                    .LoadFile DOC
                    .setShowScrollbars True
                    .setShowToolbar True
                    .setViewScroll Fit, Offset
                End With
            End If
        Else
            pdf1.visible = False
        End If
    End If
End Sub
Private Sub txtAsignacion_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtAsignacion.ToolTipText = txtAsignacion.Text
End Sub

Private Sub texto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    texto(Index).ToolTipText = texto(Index).Text
End Sub

Private Sub cmdCertificar_Click()
   Select Case tipo
   Case 2 To 3
      PNTA002 (False)
   Case 4 To 8
      PC (False)
   End Select
   certifica
End Sub
Private Sub cmdCertificar2_Click()
   Select Case tipo
   Case 2 To 3
      PNTA002 (False)
   Case 4 To 8
      PC (False)
   End Select
   certifica
End Sub

